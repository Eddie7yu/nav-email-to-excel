from __future__ import annotations

import datetime as dt
import gc
import hashlib
import importlib.util
import importlib.metadata
import platform
import sys
from pathlib import Path
from typing import Any

from nav_commit import ensure_process_exit, spreadsheet_app
from nav_config import ROOT, normalize_code, write_json_atomic
from nav_mail import decoded, fetch_authorized_messages
from nav_parse import NavRow, ParseError, choose_route_rows, rows_from_message
from nav_workbook import build_preview, validate_history
from runtime_secret import read_password


def _sender_id(sender: str) -> str:
    return hashlib.sha256(sender.lower().encode("utf-8")).hexdigest()[:12]


def collect_route_rows(
    config: dict[str, Any],
) -> tuple[dict[str, list[NavRow]], dict[str, Any]]:
    messages = fetch_authorized_messages(config)
    sender_routes: dict[str, list[dict[str, Any]]] = {}
    for route in config.get("routes") or []:
        sender_routes.setdefault(str(route["sender"]).lower(), []).append(route)
    output: dict[str, list[NavRow]] = {
        str(route["sheet"]): [] for route in config.get("routes") or []
    }
    route_reports: list[dict[str, Any]] = []
    errors: list[str] = []
    for sender, routes in sender_routes.items():
        route_message_counts = {str(route["sheet"]): 0 for route in routes}
        route_filtered_counts = {str(route["sheet"]): 0 for route in routes}
        for message in messages.get(sender, []):
            applicable = []
            for route in routes:
                subject_filter = (
                    str(route.get("subject_contains") or "").strip().casefold()
                )
                if (
                    subject_filter
                    and subject_filter not in decoded(message.get("Subject")).casefold()
                ):
                    route_filtered_counts[str(route["sheet"])] += 1
                else:
                    applicable.append(route)
            if not applicable:
                continue
            try:
                message_rows = rows_from_message(message)
            except ParseError:
                message_rows = []
            if not message_rows:
                errors.append("An in-scope authorized message could not be parsed")
                continue
            used_sheets: set[str] = set()
            for row in message_rows:
                matches: list[dict[str, Any]] = []
                for route in applicable:
                    expected = normalize_code(route.get("code"))
                    if expected and row.code == expected:
                        matches.append(route)
                    elif (
                        row.code is None
                        and len(applicable) == 1
                        and route.get("allow_sender_only")
                    ):
                        matches.append(route)
                    elif expected is None and len(applicable) == 1:
                        matches.append(route)
                if len(matches) != 1:
                    errors.append(
                        "An in-scope NAV row was not routed to exactly one authorized sheet"
                    )
                    continue
                route = matches[0]
                sheet = str(route["sheet"])
                output[sheet].append(row)
                used_sheets.add(sheet)
            if not used_sheets:
                errors.append(
                    "An in-scope authorized message produced no routed NAV rows"
                )
            for sheet in used_sheets:
                route_message_counts[sheet] += 1
        for route in routes:
            sheet = str(route["sheet"])
            try:
                selected = choose_route_rows(output[sheet], route, len(routes) == 1)
            except ParseError as exc:
                errors.append(f"{route['sheet']}: {exc}")
                selected = []
            output[sheet] = selected
            if selected:
                max_age = int(route.get("max_staleness_days", 14))
                if max(row.date for row in selected) < dt.date.today() - dt.timedelta(
                    days=max_age
                ):
                    errors.append(
                        f"{sheet}: latest routed NAV exceeds max_staleness_days"
                    )
            route_reports.append(
                {
                    "sheet": route["sheet"],
                    "sender_id": _sender_id(sender),
                    "messages_scanned": route_message_counts[sheet],
                    "messages_filtered_by_subject": route_filtered_counts[sheet],
                    "candidate_dates": [row.date.isoformat() for row in selected],
                }
            )
    report = {"passed": not errors, "routes": route_reports, "errors": errors}
    write_json_atomic(ROOT / "route-report.json", report)
    return output, report


def doctor(config: dict[str, Any]) -> dict[str, Any]:
    checks: list[dict[str, Any]] = []

    def add(name: str, passed: bool, detail: str) -> None:
        checks.append({"name": name, "passed": passed, "detail": detail})

    supported_python = (3, 11) <= sys.version_info[:2] <= (3, 14)
    add("python", supported_python, platform.python_version())
    workbook = Path(config["workbook_path"])
    add("workbook", workbook.is_file(), "present" if workbook.is_file() else "missing")
    add(
        "routes",
        bool(config.get("routes")),
        f"{len(config.get('routes') or [])} configured",
    )
    expected_packages = {
        "openpyxl": "3.1.5",
        "pypdf": "6.13.2",
        "xlrd": "2.0.2",
        "et_xmlfile": "2.0.0",
    }
    if sys.platform == "win32":
        expected_packages["pywin32"] = "312"
    for package, expected in expected_packages.items():
        module = "win32com" if package == "pywin32" else package
        available = importlib.util.find_spec(module) is not None
        try:
            installed = importlib.metadata.version(package) if available else "missing"
        except importlib.metadata.PackageNotFoundError:
            installed = "missing"
        add(
            f"package:{package}",
            available and installed == expected,
            f"{installed}; expected {expected}",
        )
    secret = bool(read_password(str(config["runtime_id"])))
    add("imap-secret", secret, "available" if secret else "not configured")
    if sys.platform == "win32":
        try:
            app, progid, process_id = spreadsheet_app()
            app.Quit()
            app = None
            gc.collect()
            ensure_process_exit(process_id)
            add("spreadsheet-com", True, progid)
        except Exception:
            add("spreadsheet-com", False, "Excel/WPS COM unavailable; preview only")
    else:
        add("spreadsheet-com", False, "formal commit is Windows-only")
    path_text = str(ROOT.resolve())
    local_path = not path_text.startswith("\\\\")
    add(
        "schedule-path",
        local_path,
        "local" if local_path else "UNC paths cannot be scheduled",
    )
    by_name = {item["name"]: item["passed"] for item in checks}
    base_names = {
        "python",
        "workbook",
        *{f"package:{name}" for name in expected_packages},
    }
    bootstrap_ready = all(by_name.get(name, False) for name in base_names)
    preview_ready = bootstrap_ready and by_name["routes"] and by_name["imap-secret"]
    commit_ready = preview_ready and by_name["spreadsheet-com"]
    schedule_ready = commit_ready and by_name["schedule-path"]
    blockers = [item["name"] for item in checks if not item["passed"]]
    return {
        "passed": preview_ready,
        "bootstrap_ready": bootstrap_ready,
        "preview_ready": preview_ready,
        "commit_ready": commit_ready,
        "schedule_ready": schedule_ready,
        "blockers": blockers,
        "checks": checks,
    }


def discover(config: dict[str, Any]) -> tuple[dict[str, list[NavRow]], dict[str, Any]]:
    return collect_route_rows(config)


def validate(
    config: dict[str, Any], rows: dict[str, list[NavRow]] | None = None
) -> dict[str, Any]:
    if rows is None:
        rows, discovery = collect_route_rows(config)
        if not discovery["passed"]:
            return {"passed": False, "routes": [], "errors": discovery["errors"]}
    return validate_history(config, rows)


def preview(
    config: dict[str, Any], rows: dict[str, list[NavRow]] | None = None
) -> dict[str, Any]:
    # A failed preview must never leave an older commit plan looking current.
    (ROOT / "plan.json").unlink(missing_ok=True)
    if rows is None:
        rows, discovery = collect_route_rows(config)
        if not discovery["passed"]:
            raise RuntimeError("Discovery failed; inspect route-report.json")
    validation = validate_history(config, rows)
    if not validation["passed"]:
        raise RuntimeError(
            "Historical validation failed; inspect validation-report.json"
        )
    return build_preview(config, rows)
