from __future__ import annotations

import datetime as dt
import gc
import hashlib
import importlib.util
import importlib.metadata
import platform
import sys
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Any

from nav_commit import (
    detected_spreadsheet_apps,
    ensure_process_exit,
    spreadsheet_app,
)
from nav_config import ROOT, active_routes, normalize_code, write_json_atomic
from nav_mail import (
    decoded,
    fetch_authorized_messages,
    fetch_candidate_messages,
    single_from_address,
)
from nav_parse import (
    NavRow,
    ParseError,
    choose_route_rows,
    deduplicate,
    rows_from_message,
)
from nav_workbook import build_preview, validate_history
from runtime_secret import read_password


def _sender_id(sender: str) -> str:
    return hashlib.sha256(sender.lower().encode("utf-8")).hexdigest()[:12]


def propose_routes(config: dict[str, Any]) -> dict[str, Any]:
    messages, scan = fetch_candidate_messages(config)
    groups: dict[str, dict[str, Any]] = {}
    ignored = 0
    for message in messages:
        sender = single_from_address(message)
        if not sender:
            ignored += 1
            continue
        try:
            rows = rows_from_message(message, "auto")
        except ParseError:
            rows = []
        if not rows:
            ignored += 1
            continue
        group = groups.setdefault(
            sender,
            {
                "subjects": [],
                "message_dates": [],
                "messages": 0,
                "observations": {},
            },
        )
        group["messages"] += 1
        subject = decoded(message.get("Subject")).strip()
        if subject and subject not in group["subjects"]:
            group["subjects"].append(subject)
        try:
            message_date = parsedate_to_datetime(str(message.get("Date"))).isoformat()
        except (TypeError, ValueError, OverflowError):
            message_date = None
        if message_date and message_date not in group["message_dates"]:
            group["message_dates"].append(message_date)
        for row in rows:
            key = (
                row.date.isoformat(),
                row.code or "",
                round(row.unit, 10),
                None if row.cumulative is None else round(row.cumulative, 10),
            )
            group["observations"][key] = {
                "date": row.date.isoformat(),
                "code": row.code,
                "unit": row.unit,
                "cumulative": row.cumulative,
                "source": row.source,
            }
    candidates: list[dict[str, Any]] = []
    for sender, group in sorted(groups.items()):
        observations = sorted(
            group["observations"].values(),
            key=lambda item: (item["date"], item["code"] or "", item["unit"]),
        )
        dates = sorted({item["date"] for item in observations})
        candidates.append(
            {
                "sender": sender,
                "message_count": group["messages"],
                "recent_message_times": group["message_dates"][:20],
                "subject_examples": group["subjects"][:10],
                "detected_codes": sorted(
                    {item["code"] for item in observations if item["code"]}
                ),
                "first_date": dates[0],
                "latest_date": dates[-1],
                "observations": observations[-100:],
            }
        )
    warnings: list[str] = []
    if scan["truncated"]:
        warnings.append(
            "邮箱候选扫描受到数量或总大小上限限制；AI 应在需要时缩小时间范围后重试"
        )
    if scan["skipped_oversize"]:
        warnings.append(f"已跳过 {scan['skipped_oversize']} 封超过单封大小上限的邮件")
    report = {
        "passed": bool(candidates),
        "scan": {**scan, "messages_ignored_as_non_nav": ignored},
        "candidates": candidates,
        "warnings": warnings,
        "errors": [] if candidates else ["没有发现可自动解析的净值邮件候选"],
    }
    write_json_atomic(ROOT / "route-proposals.json", report)
    return report


def collect_route_rows(
    config: dict[str, Any],
) -> tuple[dict[str, list[NavRow]], dict[str, Any]]:
    routes_config = active_routes(config)
    paused_routes = [
        route for route in config.get("routes") or [] if route.get("paused", False)
    ]
    warnings = [
        f"{route['sheet']}: paused ({str(route['pause_reason']).strip()})"
        for route in paused_routes
    ]
    if not routes_config:
        report = {
            "passed": False,
            "routes": [],
            "warnings": warnings,
            "errors": ["No active routes are configured"],
        }
        write_json_atomic(ROOT / "route-report.json", report)
        return {}, report
    messages = fetch_authorized_messages(config)
    sender_routes: dict[str, list[dict[str, Any]]] = {}
    for route in routes_config:
        sender_routes.setdefault(str(route["sender"]).lower(), []).append(route)
    output: dict[str, list[NavRow]] = {
        str(route["sheet"]): [] for route in routes_config
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
            message_rows: list[NavRow] = []
            parse_failed = False
            for parser_name in sorted(
                {str(route.get("parser", "auto")) for route in applicable}
            ):
                try:
                    parsed = rows_from_message(message, parser_name)
                except ParseError:
                    parsed = []
                if not parsed:
                    parse_failed = True
                    break
                message_rows.extend(parsed)
            if parse_failed or not message_rows:
                errors.append("An in-scope authorized message could not be parsed")
                continue
            message_rows = deduplicate(message_rows)
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
            if not selected:
                errors.append(
                    f"{sheet}: no routed NAV rows were found in the lookback window"
                )
            else:
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
    report = {
        "passed": not errors,
        "routes": route_reports,
        "warnings": warnings,
        "errors": errors,
    }
    write_json_atomic(ROOT / "route-report.json", report)
    return output, report


def doctor(config: dict[str, Any]) -> dict[str, Any]:
    checks: list[dict[str, Any]] = []
    detected_spreadsheets = detected_spreadsheet_apps()

    def add(name: str, passed: bool, detail: str) -> None:
        checks.append({"name": name, "passed": passed, "detail": detail})

    supported_python = (3, 11) <= sys.version_info[:2] <= (3, 14)
    add("python", supported_python, platform.python_version())
    workbook = Path(config["workbook_path"])
    add("workbook", workbook.is_file(), "present" if workbook.is_file() else "missing")
    configured_routes = config.get("routes") or []
    active = active_routes(config)
    add(
        "routes",
        bool(active),
        f"{len(active)} active; {len(configured_routes) - len(active)} paused",
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
        except Exception as exc:
            add("spreadsheet-com", False, str(exc))
    else:
        add(
            "spreadsheet-com",
            False,
            f"当前 Python 运行在 {sys.platform}，无法访问 Windows 上的 Excel/WPS COM；"
            "请从 Windows PowerShell 使用原生 Windows Python 部署",
        )
    path_text = str(ROOT.resolve())
    local_path = not path_text.startswith("\\\\")
    add(
        "schedule-path",
        local_path,
        "local" if local_path else "UNC paths cannot be scheduled",
    )
    schedules = config.get("schedule") or []
    add(
        "schedule-config",
        bool(schedules),
        (
            f"{len(schedules)} configured time(s)"
            if schedules
            else "not configured; ask the user for update frequency and local time"
        ),
    )
    by_name = {item["name"]: item["passed"] for item in checks}
    base_names = {
        "python",
        "workbook",
        *{f"package:{name}" for name in expected_packages},
    }
    bootstrap_ready = all(by_name.get(name, False) for name in base_names)
    mail_discovery_ready = bootstrap_ready and by_name["imap-secret"]
    preview_ready = bootstrap_ready and by_name["routes"] and by_name["imap-secret"]
    commit_ready = preview_ready and by_name["spreadsheet-com"]
    schedule_ready = (
        commit_ready and by_name["schedule-path"] and by_name["schedule-config"]
    )
    blockers = [item["name"] for item in checks if not item["passed"]]
    return {
        "passed": preview_ready,
        "bootstrap_ready": bootstrap_ready,
        "mail_discovery_ready": mail_discovery_ready,
        "preview_ready": preview_ready,
        "commit_ready": commit_ready,
        "schedule_ready": schedule_ready,
        "runtime_platform": sys.platform,
        "spreadsheet_apps_detected": detected_spreadsheets,
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
            return {
                "passed": False,
                "routes": [],
                "warnings": discovery.get("warnings", []),
                "errors": discovery["errors"],
            }
    return validate_history(config, rows)


def preview(
    config: dict[str, Any], rows: dict[str, list[NavRow]] | None = None
) -> dict[str, Any]:
    # A failed preview must never leave an older commit plan looking current.
    (ROOT / "plan.json").unlink(missing_ok=True)
    warnings: list[str] = []
    if rows is None:
        rows, discovery = collect_route_rows(config)
        if not discovery["passed"]:
            raise RuntimeError("Discovery failed; inspect route-report.json")
        warnings.extend(discovery.get("warnings") or [])
    validation = validate_history(config, rows)
    if not validation["passed"]:
        raise RuntimeError(
            "Historical validation failed; inspect validation-report.json"
        )
    warnings.extend(validation.get("warnings") or [])
    return build_preview(config, rows, warnings)
