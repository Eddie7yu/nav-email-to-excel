#!/usr/bin/env python3
from __future__ import annotations

import argparse
import datetime as dt
import gc
import json
import shutil
import subprocess
import sys
from email.message import EmailMessage
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Font


def check(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def create_book(path: Path) -> None:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Demo Fund"
    sheet.append(
        [
            "NAV Date",
            "Fund Return",
            "Product Name",
            "Unit NAV",
            "Product Code",
            "Cumulative NAV",
            "Audit Formula",
            "Benchmark Return",
            "Excess",
            "Benchmark Level",
        ]
    )
    sheet.append(
        [
            dt.date(2026, 1, 2),
            None,
            "Example Fund",
            1.0,
            "DEMO01",
            1.0,
            "=D2*2",
            None,
            None,
            100.0,
        ]
    )
    sheet.append(
        [
            dt.date(2026, 1, 9),
            "=F3/F2-1",
            "Example Fund",
            1.01,
            "DEMO01",
            1.01,
            "=D3*2",
            0.005,
            "=B3-H3",
            101.0,
        ]
    )
    sheet.append(
        [
            "TOTAL",
            "=F3/F2-1",
            None,
            None,
            None,
            None,
            None,
            "=PRODUCT(1+H2:H3)-1",
            "=B4-H4",
            None,
        ]
    )
    for row in (2, 3, 4):
        for column in (2, 8, 9):
            sheet.cell(row, column).font = Font(
                name="Arial", size=10, bold=True, color="00C00000"
            )
            sheet.cell(row, column).number_format = "0.00%"
    benchmark = workbook.create_sheet("Demo Benchmark")
    benchmark.append(["Date", "Benchmark Return"])
    for date, value in (
        (dt.date(2026, 1, 2), 0.0),
        (dt.date(2026, 1, 9), 0.005),
        (dt.date(2026, 1, 16), -0.002),
        (dt.date(2026, 1, 23), 0.003),
    ):
        benchmark.append([date, value])
    workbook.save(path)


def config_for(runtime: Path, book: Path) -> dict:
    return {
        "schema_version": 1,
        "runtime_id": "00000000-0000-4000-8000-000000000001",
        "workbook_path": str(book.resolve()),
        "imap": {
            "host": "imap.example.invalid",
            "port": 993,
            "user": "user@example.invalid",
            "mailbox": "INBOX",
            "lookback_days": 180,
        },
        "routes": [
            {
                "sender": "sender@example.invalid",
                "sheet": "Demo Fund",
                "code": "DEMO01",
                "parser": "auto",
                "allow_sender_only": False,
                "cumulative_policy": "require",
                "cumulative_offset": None,
                "return_basis": "cumulative",
                "return_frequency": "weekly",
                "series_start": "2026-01-02",
                "benchmark": {
                    "source_sheet": "Demo Benchmark",
                    "source_type": "aligned_return",
                    "source_date": "A",
                    "source_value": "B",
                },
            }
        ],
        "column_overrides": {},
        "style": {"mode": "infer", "zero_threshold": 0.00005},
        "schedule": [],
        "validation": {"minimum_history_dates": 2, "tolerance": 0.000001},
    }


def parser_tests(runtime: Path) -> None:
    sys.path.insert(0, str(runtime))
    import nav_mail
    from nav_parse import (
        ParseError,
        choose_route_rows,
        parse_number,
        rows_from_message,
        rows_from_text,
    )
    from nav_mail import (
        MailError,
        exact_from_matches,
        fetch_candidate_messages,
        imap_date,
        needs_imap_id,
        single_from_address,
    )

    check(
        imap_date(dt.date(2026, 7, 9)) == "09-Jul-2026",
        "IMAP dates must use fixed English month names",
    )
    check(
        needs_imap_id("IMAP.163.COM.")
        and needs_imap_id("imap.126.com")
        and needs_imap_id("imap.yeah.net")
        and not needs_imap_id("imap.qq.com"),
        "NetEase IMAP host detection is incorrect",
    )

    original_ssl = nav_mail.imaplib.IMAP4_SSL
    original_read_password = nav_mail.read_password
    events: list[tuple[Any, ...]] = []

    class FakeIMAP:
        id_status = "OK"

        def __init__(self, host, port, **_kwargs):
            events.append(("connect", host, port))

        def login(self, user, _password):
            events.append(("login", user))
            return "OK", [b""]

        def xatom(self, name, payload):
            events.append((name, payload))
            return self.id_status, [b""]

        def select(self, mailbox, readonly=False):
            events.append(("select", mailbox, readonly))
            return "OK", [b""]

        def logout(self):
            events.append(("logout",))
            return "BYE", [b""]

    mail_config = {
        "runtime_id": "00000000-0000-4000-8000-000000000001",
        "imap": {
            "host": "imap.163.com",
            "port": 993,
            "user": "user@example.invalid",
            "mailbox": "INBOX",
        },
    }
    try:
        nav_mail.imaplib.IMAP4_SSL = FakeIMAP
        nav_mail.read_password = lambda _runtime_id: "fixture-secret"
        nav_mail.connect(mail_config)
        check(
            [event[0] for event in events] == ["connect", "login", "ID", "select"],
            "NetEase IMAP ID must run after login and before mailbox selection",
        )
        check(
            "user@example.invalid" not in str(events[2])
            and "fixture-secret" not in str(events[2]),
            "IMAP ID leaked account information",
        )

        events.clear()
        mail_config["imap"]["host"] = "imap.qq.com"
        nav_mail.connect(mail_config)
        check(
            [event[0] for event in events] == ["connect", "login", "select"],
            "Non-NetEase IMAP received an unnecessary ID command",
        )

        events.clear()
        mail_config["imap"]["host"] = "imap.126.com"
        FakeIMAP.id_status = "BAD"
        try:
            nav_mail.connect(mail_config)
        except MailError as exc:
            check(
                "ID handshake" in str(exc),
                "Rejected NetEase IMAP ID did not return a specific error",
            )
        else:
            raise AssertionError("Rejected NetEase IMAP ID was accepted")
        check(events[-1][0] == "logout", "Failed IMAP connection was not closed")
    finally:
        FakeIMAP.id_status = "OK"
        nav_mail.imaplib.IMAP4_SSL = original_ssl
        nav_mail.read_password = original_read_password

    sender_message = EmailMessage()
    sender_message["From"] = "Example Sender <sender@example.invalid>"
    check(
        exact_from_matches(sender_message, "sender@example.invalid"),
        "exact From matching rejected a display name",
    )
    spoofed_message = EmailMessage()
    spoofed_message["From"] = "other@example.invalid"
    check(
        not exact_from_matches(spoofed_message, "sender@example.invalid"),
        "substring From matching was accepted",
    )
    candidate_message = EmailMessage()
    candidate_message["From"] = "NAV Desk <sender@example.invalid>"
    candidate_message["Subject"] = "NAV update"
    candidate_message.set_content(
        "Product Code | NAV Date | Unit NAV | Cumulative NAV\n"
        "DEMO01 | 2026-01-09 | 1.01 | 1.01"
    )
    candidate_payload = candidate_message.as_bytes()

    class CandidateIMAP:
        def uid(self, command, *_args):
            if command == "search":
                return "OK", [b"1"]
            query = str(_args[-1])
            if "RFC822.SIZE" in query:
                return "OK", [f"1 (RFC822.SIZE {len(candidate_payload)})".encode()]
            return "OK", [(b"1 (BODY[])", candidate_payload), b")"]

        def close(self):
            return "OK", [b""]

        def logout(self):
            return "BYE", [b""]

    original_connect = nav_mail.connect
    nav_mail.connect = lambda _config: CandidateIMAP()
    try:
        candidates, scan = fetch_candidate_messages(mail_config)
    finally:
        nav_mail.connect = original_connect
    check(
        len(candidates) == 1
        and single_from_address(candidates[0]) == "sender@example.invalid"
        and scan["messages_fetched"] == 1,
        "mailbox-wide sender discovery did not return the local NAV candidate",
    )
    check(
        parse_number("1.02%") is None, "percentage text must not be accepted as a NAV"
    )

    text = "Product Code | NAV Date | Unit NAV | Cumulative NAV\nDEMO01 | 2026-01-09 | 1.01 | 1.01"
    rows = rows_from_text(text, "body")
    check(
        len(rows) == 1 and rows[0].code == "DEMO01" and rows[0].unit == 1.01,
        "labelled table parsing failed",
    )

    attachment = openpyxl.Workbook()
    cover = attachment.active
    cover.title = "Cover"
    cover.append(["Cover page"])
    data = attachment.create_sheet("Data")
    for _ in range(10):
        data.append(["Note"])
    data.append(["Product Code", "NAV Date", "Cumulative NAV", "Unit NAV"])
    data.append(["DEMO01", "2026-01-16", 1.02, 1.02])
    import io

    buffer = io.BytesIO()
    attachment.save(buffer)
    message = EmailMessage()
    message["Subject"] = "NAV notice"
    message.set_content("See attachment")
    message.add_attachment(
        buffer.getvalue(),
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="report.xlsx",
    )
    parsed = rows_from_message(message)
    check(
        any(row.date == dt.date(2026, 1, 16) for row in parsed),
        "multi-sheet attachment parsing failed",
    )

    parser_dir = runtime / "parsers"
    parser_dir.mkdir(exist_ok=True)
    (parser_dir / "fixture.py").write_text(
        "from datetime import date\n"
        "from nav_parse import NavRow\n\n"
        "def parse_message(message):\n"
        "    return [NavRow(date(2026, 1, 30), 1.03, 1.03, 'DEMO01', 'local')]\n",
        encoding="utf-8",
    )
    local_rows = rows_from_message(EmailMessage(), "local:fixture")
    check(
        len(local_rows) == 1 and local_rows[0].date == dt.date(2026, 1, 30),
        "trusted local parser extension failed",
    )

    route = {"code": "DEMO01", "allow_sender_only": False, "sheet": "Demo Fund"}
    conflict = rows_from_text(
        "Product Code | NAV Date | Unit NAV\nDEMO01 | 2026-01-09 | 1.01\nDEMO01 | 2026-01-09 | 1.02",
        "body",
    )
    try:
        choose_route_rows(conflict, route, True)
    except ParseError:
        pass
    else:
        raise AssertionError("same-day value conflicts must fail closed")
    try:
        rows_from_text(
            "NAV Date: 2026-01-09\nNAV Date: 2026-01-16\nUnit NAV: 1.01", "body"
        )
    except ParseError:
        pass
    else:
        raise AssertionError("repeated labelled fields must fail closed")


def route_state_tests(runtime: Path) -> None:
    sys.path.insert(0, str(runtime))
    import nav_service

    book = runtime / "route-state-placeholder.xlsx"
    create_book(book)
    config = config_for(runtime, book)
    empty = json.loads(json.dumps(config))
    empty["routes"] = []
    proposal_message = EmailMessage()
    proposal_message["From"] = "NAV Desk <sender@example.invalid>"
    proposal_message["Subject"] = "Weekly NAV"
    proposal_message.set_content(
        "Product Code | NAV Date | Unit NAV | Cumulative NAV\n"
        "DEMO01 | 2026-01-09 | 1.01 | 1.01"
    )
    original_candidates = nav_service.fetch_candidate_messages
    nav_service.fetch_candidate_messages = lambda _config: (
        [proposal_message],
        {
            "messages_found": 1,
            "messages_fetched": 1,
            "bytes_fetched": 100,
            "skipped_oversize": 0,
            "truncated": False,
        },
    )
    try:
        proposal = nav_service.propose_routes(empty)
    finally:
        nav_service.fetch_candidate_messages = original_candidates
    check(
        proposal["passed"]
        and proposal["candidates"][0]["sender"] == "sender@example.invalid"
        and proposal["candidates"][0]["detected_codes"] == ["DEMO01"],
        "AI route proposal did not discover the sender and product code",
    )
    _, report = nav_service.collect_route_rows(empty)
    check(
        not report["passed"] and "No active routes" in report["errors"][0],
        "empty route configuration reported discovery success",
    )
    paused = json.loads(json.dumps(config))
    paused["routes"][0]["paused"] = True
    paused["routes"][0]["pause_reason"] = "fixture pause"
    _, report = nav_service.collect_route_rows(paused)
    check(
        not report["passed"] and report["warnings"],
        "paused route was not excluded with a visible warning",
    )
    original_fetch = nav_service.fetch_authorized_messages
    nav_service.fetch_authorized_messages = lambda _config: {
        "sender@example.invalid": []
    }
    try:
        _, report = nav_service.collect_route_rows(config)
    finally:
        nav_service.fetch_authorized_messages = original_fetch
    check(
        not report["passed"]
        and any("no routed NAV rows" in item for item in report["errors"]),
        "active route with no messages reported discovery success",
    )
    mixed = json.loads(json.dumps(config))
    mixed["routes"][0].update({"parser": "local:fixture", "max_staleness_days": 366})
    second_route = json.loads(json.dumps(mixed["routes"][0]))
    second_route.update({"sheet": "Second Fund", "code": "DEMO02", "parser": "auto"})
    mixed["routes"].append(second_route)
    mixed_message = EmailMessage()
    mixed_message.set_content(
        "Product Code | NAV Date | Unit NAV | Cumulative NAV\n"
        "DEMO02 | 2026-02-06 | 1.04 | 1.04"
    )
    nav_service.fetch_authorized_messages = lambda _config: {
        "sender@example.invalid": [mixed_message]
    }
    try:
        mixed_rows, report = nav_service.collect_route_rows(mixed)
    finally:
        nav_service.fetch_authorized_messages = original_fetch
    check(
        report["passed"]
        and len(mixed_rows["Demo Fund"]) == 1
        and len(mixed_rows["Second Fund"]) == 1,
        "multiple trusted parsers for one sender were not merged and routed safely",
    )


def lock_tests(runtime: Path) -> None:
    sys.path.insert(0, str(runtime))
    from nav_schedule import record_scheduled_run, status

    holder_code = (
        "import time\n"
        "from navctl import run_lock\n"
        "with run_lock():\n"
        "    print('ready', flush=True)\n"
        "    time.sleep(60)\n"
    )
    contender_code = (
        "import sys\n"
        "from navctl import run_lock\n"
        "try:\n"
        "    with run_lock():\n"
        "        print('acquired')\n"
        "except RuntimeError:\n"
        "    print('blocked')\n"
        "    raise SystemExit(3)\n"
    )
    holder = subprocess.Popen(
        [sys.executable, "-X", "utf8", "-c", holder_code],
        cwd=runtime,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
    )
    try:
        ready = holder.stdout.readline().strip() if holder.stdout else ""
        check(ready == "ready", "runtime lock holder did not start")
        blocked = subprocess.run(
            [sys.executable, "-X", "utf8", "-c", contender_code],
            cwd=runtime,
            capture_output=True,
            text=True,
            encoding="utf-8",
        )
        check(
            blocked.returncode == 3 and "blocked" in blocked.stdout,
            "runtime lock allowed concurrent access",
        )
        check(holder.poll() is None, "runtime lock probe interrupted the lock holder")
    finally:
        holder.terminate()
        holder.wait(timeout=10)
    recovered = subprocess.run(
        [sys.executable, "-X", "utf8", "-c", contender_code],
        cwd=runtime,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    check(
        recovered.returncode == 0 and "acquired" in recovered.stdout,
        "runtime lock did not recover after a crashed holder",
    )
    state = json.loads((runtime / "run.lock").read_text(encoding="utf-8"))
    check(state.get("status") == "idle", "runtime lock did not record idle state")
    scheduled = {
        "started": "2026-01-01T09:30:00",
        "finished": "2026-01-01T09:30:05",
        "passed": False,
        "exit_code": 2,
        "error": "fixture failure",
    }
    record_scheduled_run(scheduled)
    check(
        status().get("last_run") == scheduled,
        "schedule status did not expose the latest run result",
    )
    check(
        "scheduled-update" in (runtime / "run-update.cmd").read_text(encoding="utf-8"),
        "scheduled wrapper does not run automatic updates",
    )


def workbook_tests(runtime: Path, use_com: bool) -> str | None:
    sys.path.insert(0, str(runtime))
    from nav_automation import (
        AutomationError,
        approve,
        automatic_update,
        revoke,
        status as automation_status,
    )
    from nav_commit import CommitError, _process_id, commit, ensure_process_exit
    from nav_config import ConfigError, load_config, validate_config
    from nav_parse import NavRow
    from nav_service import preview as service_preview
    from nav_workbook import WorkbookError, build_preview, validate_history

    book = runtime / "脱敏 示例.xlsx"
    create_book(book)
    config = config_for(runtime, book)
    revoke()
    check(
        not automation_status(config)["approved"],
        "automatic updates started approved before the first reviewed commit",
    )
    approve(config)
    check(
        automation_status(config)["approved"],
        "reviewed configuration could not approve future automatic updates",
    )
    changed_config = json.loads(json.dumps(config))
    changed_config["routes"][0]["return_frequency"] = "daily"
    check(
        not automation_status(changed_config)["approved"],
        "automatic approval survived a write-rule configuration change",
    )
    revoke()
    invalid = json.loads(json.dumps(config))
    invalid["routes"][0]["series_start"] = "2026-99-99"
    invalid["routes"][0]["parser"] = "local:../unsafe"
    invalid["unexpected"] = True
    try:
        validate_config(invalid)
    except ConfigError:
        pass
    else:
        raise AssertionError("invalid dates and unknown config fields must be rejected")
    local_parser_config = json.loads(json.dumps(config))
    local_parser_config["routes"][0]["parser"] = "local:fixture"
    validate_config(local_parser_config)
    mixed_parser_config = json.loads(json.dumps(config))
    second_route = json.loads(json.dumps(mixed_parser_config["routes"][0]))
    second_route.update(
        {"sheet": "Second Fund", "code": "DEMO02", "parser": "local:fixture"}
    )
    mixed_parser_config["routes"].append(second_route)
    validate_config(mixed_parser_config)
    paused_config = json.loads(json.dumps(config))
    paused_config["routes"][0]["paused"] = True
    paused_config["routes"][0]["pause_reason"] = "fixture pause"
    validate_config(paused_config)
    invalid_append = json.loads(json.dumps(config))
    invalid_append["routes"][0].update(
        {"sheet_mode": "append", "code": None, "product_name": None}
    )
    try:
        validate_config(invalid_append)
    except ConfigError:
        pass
    else:
        raise AssertionError("append mode accepted a route without product identity")
    invalid_formula_name = json.loads(json.dumps(config))
    invalid_formula_name["routes"][0]["product_name"] = (
        '=HYPERLINK("https://example.invalid")'
    )
    try:
        validate_config(invalid_formula_name)
    except ConfigError:
        pass
    else:
        raise AssertionError("route accepted a formula-like product name")
    (runtime / "config.json").write_text(
        json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    config = load_config(runtime / "config.json")
    rows = {
        "Demo Fund": [
            NavRow(dt.date(2026, 1, 2), 1.0, 1.0, "DEMO01", "fixture"),
            NavRow(dt.date(2026, 1, 9), 1.01, 1.01, "DEMO01", "fixture"),
            NavRow(dt.date(2026, 1, 16), 1.02, 1.02, "DEMO01", "fixture"),
            NavRow(dt.date(2026, 1, 23), 1.015, 1.015, "DEMO01", "fixture"),
        ]
    }
    unsafe_book = runtime / "unsafe-summary.xlsx"
    create_book(unsafe_book)
    unsafe_workbook = openpyxl.load_workbook(unsafe_book)
    try:
        unsafe_workbook["Demo Fund"]["G4"] = "=SUM(D2:D3)"
        unsafe_workbook.save(unsafe_book)
    finally:
        unsafe_workbook.close()
    unsafe_config = config_for(runtime, unsafe_book)
    try:
        build_preview(unsafe_config, rows)
    except WorkbookError as exc:
        check(
            "summary formula" in str(exc),
            "unsafe summary formula failed for the wrong reason",
        )
    else:
        raise AssertionError("unmanaged summary formula must fail closed")

    validation = validate_history(config, rows)
    check(validation["passed"], f"historical validation failed: {validation['errors']}")
    plan = build_preview(config, rows)
    check(
        len(plan["sheets"]) == 1
        and plan["sheets"][0]["new_dates"] == ["2026-01-16", "2026-01-23"],
        "catch-up plan is incomplete",
    )
    preview = openpyxl.load_workbook(plan["preview_path"], data_only=False)
    try:
        sheet = preview["Demo Fund"]
        check(
            sheet["A4"].value == dt.datetime(2026, 1, 16, 0, 0)
            or sheet["A4"].value == dt.date(2026, 1, 16),
            "first catch-up date is wrong",
        )
        check(
            sheet["A5"].value == dt.datetime(2026, 1, 23, 0, 0)
            or sheet["A5"].value == dt.date(2026, 1, 23),
            "second catch-up date is wrong",
        )
        check(
            isinstance(sheet["B4"].value, str) and sheet["B4"].value.startswith("="),
            "weekly return formula is missing",
        )
        check(
            sheet["H4"].value == "='Demo Benchmark'!B4",
            "benchmark date mapping is wrong",
        )
        check(sheet["I4"].value == "=B4-H4", "excess formula is wrong")
        check(sheet["G4"].value == "=D4*2", "unmanaged formula translation failed")
        check(
            sheet["H2"].value is None and sheet["I2"].value is None,
            "existing historical benchmark cells must remain authoritative",
        )
        check(
            sheet["B6"].value == "=F5/F2-1",
            "weekly product summary does not use completed-period anchors",
        )
        check(
            sheet["B3"].value == "=F3/F2-1" and sheet["H3"].value == 0.005,
            "historical formulas were rewritten",
        )
    finally:
        preview.close()

    master_before = book.read_bytes()
    tampered = openpyxl.load_workbook(plan["preview_path"])
    tampered["Demo Fund"]["D4"] = 9.99
    tampered.save(plan["preview_path"])
    tampered.close()
    try:
        commit(config)
    except CommitError:
        pass
    else:
        raise AssertionError("commit accepted a preview changed after review")
    check(
        book.read_bytes() == master_before,
        "preview tampering changed the master workbook",
    )

    gap_rows = {
        "Demo Fund": rows["Demo Fund"]
        + [NavRow(dt.date(2026, 1, 5), 1.005, 1.005, "DEMO01", "fixture")]
    }
    try:
        build_preview(config, gap_rows)
    except WorkbookError:
        check(
            not (runtime / "plan.json").exists(),
            "failed preview left a committable plan",
        )
    else:
        raise AssertionError("internal historical gaps must require supervised repair")

    plan = build_preview(config, rows)

    application = None
    if use_com:
        from win32com.client import DispatchEx

        master_before_locked_commit = book.read_bytes()
        backups_before = set((runtime / "backups").glob("*"))
        blocking_app = DispatchEx("Excel.Application")
        blocking_app.Visible = False
        blocking_app.DisplayAlerts = False
        blocking_pid = _process_id(blocking_app)
        blocking_book = blocking_app.Workbooks.Open(str(book.resolve()))
        try:
            locked = subprocess.run(
                [
                    sys.executable,
                    "-X",
                    "utf8",
                    "navctl.py",
                    "commit",
                    "--yes-reviewed-preview",
                ],
                cwd=runtime,
                capture_output=True,
                text=True,
                encoding="utf-8",
            )
            locked_report = json.loads(locked.stdout)
            check(
                locked.returncode == 2
                and not locked_report["passed"]
                and "关闭" in locked_report["error"]
                and "Traceback" not in locked.stderr,
                "open-workbook conflict did not return a readable JSON error",
            )
            check(
                book.read_bytes() == master_before_locked_commit,
                "open-workbook conflict changed the master",
            )
            check(
                set((runtime / "backups").glob("*")) == backups_before,
                "failed open-workbook commit retained a backup",
            )
        finally:
            blocking_book.Close(False)
            blocking_app.Quit()
            blocking_book = None
            blocking_app = None
            gc.collect()
            ensure_process_exit(blocking_pid)
        result = commit(config)
        application = str(result["application"])
        check(
            result["changed"] and result["rows"] == 2,
            "COM commit did not apply both rows",
        )
        calculated = openpyxl.load_workbook(book, data_only=True)
        try:
            check(
                abs(calculated["Demo Fund"]["B4"].value - (1.02 / 1.01 - 1)) < 1e-10,
                "COM product return value is wrong",
            )
            check(
                abs(calculated["Demo Fund"]["H4"].value - (-0.002)) < 1e-12,
                "COM benchmark value is wrong",
            )
            check(
                abs(calculated["Demo Fund"]["I4"].value - ((1.02 / 1.01 - 1) + 0.002))
                < 1e-10,
                "COM excess value is wrong",
            )
        finally:
            calculated.close()
    else:
        shutil.copy2(plan["preview_path"], book)

    second_validation = validate_history(config, rows)
    check(second_validation["passed"], "post-commit history validation failed")
    second = build_preview(config, rows)
    check(
        not second["sheets"] and second["preview_path"] is None,
        "second run must be idempotent and leave no workbook copy",
    )
    check(
        not (runtime / "plan.json").exists(),
        "a no-op preview must not leave a committable plan",
    )

    level_book = runtime / "level-benchmark.xlsx"
    create_book(level_book)
    level_workbook = openpyxl.load_workbook(level_book)
    try:
        for row, value in enumerate((100.0, 101.0, 102.0, 103.0), 2):
            level_workbook["Demo Benchmark"].cell(row, 2).value = value
        level_workbook.save(level_book)
    finally:
        level_workbook.close()
    level_config = config_for(runtime, level_book)
    level_config["routes"][0]["benchmark"]["source_type"] = "level"
    level_plan = build_preview(level_config, rows)
    level_preview = openpyxl.load_workbook(level_plan["preview_path"], data_only=False)
    try:
        level_sheet = level_preview["Demo Fund"]
        check(
            level_sheet["J4"].value == "='Demo Benchmark'!B4",
            "benchmark level mapping is wrong",
        )
        check(
            level_sheet["H4"].value == "=J4/J3-1",
            "benchmark level anchor return is wrong",
        )
        check(
            level_sheet["I4"].value == "=B4-H4", "level-based excess formula is wrong"
        )
        check(
            level_sheet["H6"].value == "='Demo Benchmark'!B5/'Demo Benchmark'!B2-1",
            "benchmark summary anchors are wrong",
        )
    finally:
        level_preview.close()

    plain_book = runtime / "no-return-column.xlsx"
    plain_workbook = openpyxl.Workbook()
    plain_sheet = plain_workbook.active
    plain_sheet.title = "Plain Fund"
    plain_sheet.append(["NAV Date", "Unit NAV", "Cumulative NAV"])
    plain_sheet.append([dt.date(2026, 1, 2), 1.0, 1.0])
    plain_sheet.append([dt.date(2026, 1, 9), 1.01, 1.01])
    plain_sheet.append(["TOTAL", None, None])
    plain_workbook.save(plain_book)
    plain_workbook.close()
    plain_config = config_for(runtime, plain_book)
    plain_config["routes"][0].update(
        {"sheet": "Plain Fund", "code": None, "benchmark": None}
    )
    plain_rows = {
        "Plain Fund": [
            NavRow(dt.date(2026, 1, 2), 1.0, 1.0, None, "fixture"),
            NavRow(dt.date(2026, 1, 9), 1.01, 1.01, None, "fixture"),
            NavRow(dt.date(2026, 1, 16), 1.02, 1.02, None, "fixture"),
        ]
    }
    plain_plan = build_preview(plain_config, plain_rows)
    check(
        plain_plan["sheets"][0]["new_dates"] == ["2026-01-16"],
        "workbook without a return column could not produce a preview",
    )

    header_book = runtime / "append-header-only.xlsx"
    header_workbook = openpyxl.Workbook()
    header_sheet = header_workbook.active
    header_sheet.title = "Header Fund"
    header_sheet.append(["净值日期", "产品名称", "单位净值"])
    header_workbook.save(header_book)
    header_workbook.close()
    header_config = config_for(runtime, header_book)
    header_config["routes"][0].update(
        {
            "sheet": "Header Fund",
            "sheet_mode": "append",
            "code": None,
            "product_name": "Example Name Only Fund",
            "benchmark": None,
            "return_basis": "unit",
        }
    )
    header_config["column_overrides"] = {}
    header_rows = {
        "Header Fund": [NavRow(dt.date(2026, 1, 2), 1.0, None, None, "fixture")]
    }
    header_validation = validate_history(header_config, header_rows)
    check(
        header_validation["passed"] and header_validation["warnings"],
        "name-only header sheet could not enter append cold start",
    )
    header_plan = build_preview(
        header_config, header_rows, header_validation["warnings"]
    )
    header_preview = openpyxl.load_workbook(
        header_plan["preview_path"], data_only=False
    )
    try:
        check(
            header_preview["Header Fund"]["B2"].value == "Example Name Only Fund"
            and header_preview["Header Fund"]["C2"].value == 1.0,
            "name-only header sheet did not receive the minimum record",
        )
    finally:
        header_preview.close()

    append_book = runtime / "append-cold-start.xlsx"
    append_workbook = openpyxl.Workbook()
    append_sheet = append_workbook.active
    append_sheet.title = "New Fund"
    analysis_sheet = append_workbook.create_sheet("Analysis")
    analysis_sheet["A1"] = "Example metric"
    analysis_sheet["B1"] = "=MAX('New Fund'!D:D)"
    append_workbook.save(append_book)
    append_workbook.close()
    append_config = config_for(runtime, append_book)
    append_config["routes"][0].update(
        {
            "sheet": "New Fund",
            "sheet_mode": "append",
            "code": "NEW01",
            "product_name": "Example New Fund",
            "benchmark": None,
            "cumulative_policy": "unit",
            "return_basis": "unit",
        }
    )
    append_config["column_overrides"] = {}
    validate_config(append_config)
    append_rows = {
        "New Fund": [
            NavRow(dt.date(2026, 1, 2), 1.0, None, "NEW01", "fixture"),
            NavRow(dt.date(2026, 1, 9), 1.01, None, "NEW01", "fixture"),
        ]
    }
    append_validation = validate_history(append_config, append_rows)
    check(
        append_validation["passed"]
        and append_validation["warnings"]
        and append_validation["routes"][0]["cold_start"],
        "append cold start did not downgrade missing history to a visible warning",
    )
    append_plan = service_preview(append_config, append_rows)
    check(
        append_plan["sheets"][0]["sheet_mode"] == "append"
        and append_plan["sheets"][0]["new_dates"] == ["2026-01-02", "2026-01-09"],
        "append cold start plan is incomplete",
    )
    check(
        append_plan["warnings"],
        "append cold start warning was not preserved in the preview plan",
    )
    append_preview = openpyxl.load_workbook(
        append_plan["preview_path"], data_only=False
    )
    try:
        new_sheet = append_preview["New Fund"]
        check(
            [new_sheet.cell(1, column).value for column in range(1, 6)]
            == ["净值日期", "产品代码", "产品名称", "单位净值", "累计单位净值"],
            "blank append sheet did not receive canonical headers",
        )
        check(
            new_sheet["B2"].value == "NEW01"
            and new_sheet["C2"].value == "Example New Fund"
            and new_sheet["D3"].value == 1.01
            and new_sheet["E3"].value == 1.01,
            "append preview did not write product identity and NAV values",
        )
        check(
            new_sheet["A4"].value is None
            and append_preview["Analysis"]["B1"].value == "=MAX('New Fund'!D:D)",
            "append preview added a summary row or changed an analysis sheet",
        )
    finally:
        append_preview.close()
    if use_com:
        append_result = commit(append_config)
        application = str(append_result["application"])
        check(
            append_result["changed"] and append_result["rows"] == 2,
            "COM append commit did not apply both cold-start rows",
        )
    else:
        shutil.copy2(append_plan["preview_path"], append_book)
    approve(append_config)
    check(
        automation_status(append_config)["approved"],
        "first reviewed append commit did not enable automatic updates",
    )
    committed_append = openpyxl.load_workbook(append_book, data_only=False)
    try:
        check(
            committed_append["New Fund"]["A4"].value is None
            and committed_append["Analysis"]["B1"].value == "=MAX('New Fund'!D:D)",
            "append commit added a summary row or changed the analysis sheet",
        )
    finally:
        committed_append.close()
    append_second_validation = validate_history(append_config, append_rows)
    check(
        append_second_validation["passed"] and not append_second_validation["warnings"],
        "append mode did not leave a verifiable history after first commit",
    )
    append_second = build_preview(append_config, append_rows)
    check(
        not append_second["sheets"] and append_second["preview_path"] is None,
        "append mode is not idempotent after first commit",
    )
    append_rows["New Fund"].append(
        NavRow(dt.date(2026, 1, 16), 1.02, None, "NEW01", "fixture")
    )
    append_next_validation = validate_history(append_config, append_rows)
    check(
        append_next_validation["passed"] and not append_next_validation["warnings"],
        "existing append table did not validate before a later update",
    )
    if use_com:
        altered = json.loads(json.dumps(append_config))
        altered["routes"][0]["product_name"] = "Changed Name"
        try:
            automatic_update(altered, append_rows)
        except AutomationError:
            pass
        else:
            raise AssertionError(
                "automatic update accepted a configuration changed after approval"
            )
        append_next_result = automatic_update(append_config, append_rows)
        application = str(append_next_result["application"])
        check(
            append_next_result["changed"]
            and append_next_result["rows"] == 1
            and not (runtime / "plan.json").exists(),
            "automatic update did not write one row and remove its staging plan",
        )
    else:
        append_next = build_preview(append_config, append_rows)
        check(
            append_next["sheets"][0]["new_dates"] == ["2026-01-16"]
            and append_next["sheets"][0]["copy_template_rows"],
            "existing append table did not plan exactly one later date",
        )
        shutil.copy2(append_next["preview_path"], append_book)
    append_updated = openpyxl.load_workbook(append_book, data_only=False)
    try:
        check(
            append_updated["New Fund"]["D4"].value == 1.02
            and append_updated["New Fund"]["A5"].value is None
            and append_updated["Analysis"]["B1"].value == "=MAX('New Fund'!D:D)",
            "later append update added a footer or changed the analysis sheet",
        )
    finally:
        append_updated.close()
    return application


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--runtime", required=True)
    parser.add_argument("--com", action="store_true")
    args = parser.parse_args()
    runtime = Path(args.runtime).resolve()
    print("[1/4] 检查自动解析、本地解析器、精确发件人和冲突拦截")
    parser_tests(runtime)
    print("      PASS")
    print("[2/4] 检查空路由、暂停路由和无邮件失败关闭")
    route_state_tests(runtime)
    print("      PASS")
    print("[3/4] 检查崩溃恢复和并发运行锁")
    lock_tests(runtime)
    print("      PASS")
    print("[4/4] 检查严格表、空表冷启动、分析页保留、补录、基准和幂等性")
    application = workbook_tests(runtime, args.com)
    print("      PASS")
    if args.com:
        print(f"      Excel/WPS COM、文件占用提示及缓存数值：PASS（{application}）")
    else:
        print("      未启用 COM；如需验证正式写入，请添加 --com")
    print("selftest_driver: PASS（全程仅使用虚构数据）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
