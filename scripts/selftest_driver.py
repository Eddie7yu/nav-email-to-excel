#!/usr/bin/env python3
from __future__ import annotations

import argparse
import datetime as dt
import gc
import io
import json
import shutil
import subprocess
import sys
from copy import deepcopy
from email.message import EmailMessage
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Border, Font, PatternFill, Side


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
                "data_frequency": "auto",
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
    import navctl
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
    from runtime_secret import (
        MASKED_INPUT_PROMPT,
        SecretInputCancelled,
        _read_interactive_secret,
        _read_masked,
    )

    try:
        _read_interactive_secret(
            lambda: (_ for _ in ()).throw(
                AssertionError("non-terminal input reader was called")
            ),
            io.StringIO(),
            io.StringIO(),
        )
    except RuntimeError as exc:
        check(
            str(exc) == "此命令需要用户在真实终端中运行",
            "non-terminal secret input returned the wrong error",
        )
    else:
        raise AssertionError("non-terminal secret input was allowed to wait for input")

    class FakeTerminal:
        def isatty(self):
            return True

    cancelled_keys = iter(["a", "\x03"])
    cancelled_mask = io.StringIO()
    try:
        _read_interactive_secret(
            lambda: next(cancelled_keys), FakeTerminal(), cancelled_mask
        )
    except SecretInputCancelled:
        pass
    else:
        raise AssertionError("Ctrl+C did not cancel interactive secret input")
    check(
        cancelled_mask.getvalue().endswith("*\n"),
        "Ctrl+C did not finish the masked input line cleanly",
    )

    keys = iter(["a", "b", "\b", "C", "\x00", "K", "\r"])
    masked_output = io.StringIO()
    masked_value = _read_masked(lambda: next(keys), masked_output)
    check(
        masked_value == "aC"
        and masked_output.getvalue().startswith(f"{MASKED_INPUT_PROMPT}\n")
        and masked_output.getvalue().endswith("**\b \b*\n"),
        "Windows masked secret input did not echo stars or handle backspace",
    )
    original_set_password = navctl.set_password
    navctl.set_password = lambda _runtime_id: runtime / "secret.json"
    secret_stdout = io.StringIO()
    secret_stderr = io.StringIO()
    original_stdout, original_stderr = sys.stdout, sys.stderr
    try:
        sys.stdout, sys.stderr = secret_stdout, secret_stderr
        result = navctl.command_secret(
            {"runtime_id": "00000000-0000-4000-8000-000000000001"},
            argparse.Namespace(secret_action="set"),
        )
    finally:
        sys.stdout, sys.stderr = original_stdout, original_stderr
        navctl.set_password = original_set_password
    check(
        result == 0
        and secret_stderr.getvalue() == "已加密保存\n"
        and json.loads(secret_stdout.getvalue())["stored"],
        "secret set did not print the encrypted-save confirmation",
    )
    original_launch_secret_prompt = navctl.launch_secret_prompt
    navctl.launch_secret_prompt = lambda: 4321
    launch_stdout = io.StringIO()
    try:
        sys.stdout = launch_stdout
        launched = navctl.command_secret(
            {"runtime_id": "00000000-0000-4000-8000-000000000001"},
            argparse.Namespace(secret_action="launch"),
        )
    finally:
        sys.stdout = original_stdout
        navctl.launch_secret_prompt = original_launch_secret_prompt
    launch_report = json.loads(launch_stdout.getvalue())
    check(
        launched == 0
        and launch_report["launched"]
        and launch_report["process_id"] == 4321
        and "secret status" in launch_report["next"],
        "secret launch did not open a dedicated prompt and require verification",
    )
    navctl.set_password = lambda _runtime_id: (_ for _ in ()).throw(
        SecretInputCancelled
    )
    cancelled_stdout = io.StringIO()
    cancelled_stderr = io.StringIO()
    original_stdout, original_stderr = sys.stdout, sys.stderr
    try:
        sys.stdout, sys.stderr = cancelled_stdout, cancelled_stderr
        cancelled = navctl.command_secret(
            {"runtime_id": "00000000-0000-4000-8000-000000000001"},
            argparse.Namespace(secret_action="set"),
        )
    finally:
        sys.stdout, sys.stderr = original_stdout, original_stderr
        navctl.set_password = original_set_password
    check(
        cancelled == 130
        and cancelled_stderr.getvalue() == ""
        and json.loads(cancelled_stdout.getvalue())
        == {"passed": False, "error": "已取消"},
        "Ctrl+C during secret input did not exit cleanly",
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
    candidate_calls: list[tuple[str, tuple[Any, ...]]] = []

    class CandidateIMAP:
        def uid(self, command, *_args):
            candidate_calls.append((command, _args))
            if command == "search":
                return "OK", [b"1 2"]
            query = str(_args[-1])
            if "RFC822.SIZE" in query:
                return "OK", [
                    f"1 (UID 1 RFC822.SIZE {len(candidate_payload)})".encode(),
                    f"2 (UID 2 RFC822.SIZE {len(candidate_payload)})".encode(),
                ]
            uid = _args[0]
            return "OK", [(uid + b" (BODY[])", candidate_payload), b")"]

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
        len(candidates) == 2
        and single_from_address(candidates[0]) == "sender@example.invalid"
        and scan["messages_fetched"] == 2
        and len(
            [
                call
                for call in candidate_calls
                if call[0] == "fetch" and "RFC822.SIZE" in str(call[1][-1])
            ]
        )
        == 1,
        "mailbox-wide sender discovery did not return the local NAV candidate",
    )

    class DisconnectingIMAP(CandidateIMAP):
        def uid(self, command, *_args):
            if command == "fetch" and "BODY.PEEK" in str(_args[-1]):
                raise nav_mail.imaplib.IMAP4.abort("fixture disconnect")
            return super().uid(command, *_args)

    nav_mail.connect = lambda _config: DisconnectingIMAP()
    try:
        fetch_candidate_messages(mail_config)
    except MailError as exc:
        check(
            "意外断开" in str(exc) and "fixture disconnect" not in str(exc),
            "IMAP disconnect did not return a readable, sanitized stage error",
        )
    else:
        raise AssertionError("IMAP disconnect was not converted to MailError")
    finally:
        nav_mail.connect = original_connect
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
    invalid["routes"][0]["data_frequency"] = "hourly"
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

    dense_rows = {
        "Demo Fund": [
            NavRow(date, value, value, "DEMO01", "fixture")
            for date, value in (
                (dt.date(2026, 1, 2), 1.0),
                (dt.date(2026, 1, 5), 1.002),
                (dt.date(2026, 1, 6), 1.004),
                (dt.date(2026, 1, 7), 1.006),
                (dt.date(2026, 1, 8), 1.008),
                (dt.date(2026, 1, 9), 1.01),
                (dt.date(2026, 1, 12), 1.012),
                (dt.date(2026, 1, 13), 1.014),
                (dt.date(2026, 1, 14), 1.016),
                (dt.date(2026, 1, 15), 1.018),
                (dt.date(2026, 1, 16), 1.02),
                (dt.date(2026, 1, 19), 1.019),
                (dt.date(2026, 1, 20), 1.018),
                (dt.date(2026, 1, 21), 1.017),
                (dt.date(2026, 1, 22), 1.016),
                (dt.date(2026, 1, 23), 1.015),
            )
        ]
    }
    dense_validation = validate_history(config, dense_rows)
    check(
        dense_validation["passed"]
        and dense_validation["routes"][0]["data_frequency"] == "weekly"
        and dense_validation["routes"][0]["data_frequency_source"]
        == "workbook-history",
        "weekly template frequency was not inferred from existing history",
    )
    dense_plan = build_preview(config, dense_rows)
    check(
        dense_plan["sheets"][0]["new_dates"] == ["2026-01-16", "2026-01-23"]
        and dense_plan["sheets"][0]["data_frequency"] == "weekly",
        "daily email history was not reduced to the weekly template dates",
    )
    dense_preview = openpyxl.load_workbook(dense_plan["preview_path"], data_only=False)
    try:
        dense_sheet = dense_preview["Demo Fund"]
        check(
            dense_sheet.max_row == 6
            and dense_sheet["A6"].value == "TOTAL"
            and dense_sheet["B4"].value == "=F4/F3-1",
            "weekly preview did not preserve the summary row and formulas",
        )
    finally:
        dense_preview.close()

    forced_daily = json.loads(json.dumps(config))
    forced_daily["routes"][0]["data_frequency"] = "daily"
    try:
        build_preview(forced_daily, dense_rows)
    except WorkbookError as exc:
        check("conflicts" in str(exc), "frequency mismatch failed unclearly")
    else:
        raise AssertionError("an explicit daily override replaced a weekly template")

    gap_book = runtime / "daily-gap.xlsx"
    gap_workbook = openpyxl.Workbook()
    gap_sheet = gap_workbook.active
    gap_sheet.title = "Daily Fund"
    gap_sheet.append(["NAV Date", "Product Code", "Unit NAV", "Cumulative NAV"])
    gap_sheet.append([dt.date(2026, 1, 2), "DEMO01", 1.0, 1.0])
    gap_sheet.append([dt.date(2026, 1, 4), "DEMO01", 1.02, 1.02])
    gap_sheet.append(["TOTAL", None, None, None])
    gap_workbook.save(gap_book)
    gap_workbook.close()
    gap_config = config_for(runtime, gap_book)
    gap_config["routes"][0].update(
        {
            "sheet": "Daily Fund",
            "product_name": None,
            "benchmark": None,
            "return_frequency": "daily",
        }
    )
    gap_rows = {
        "Daily Fund": [
            NavRow(dt.date(2026, 1, 2), 1.0, 1.0, "DEMO01", "fixture"),
            NavRow(dt.date(2026, 1, 3), 1.01, 1.01, "DEMO01", "fixture"),
            NavRow(dt.date(2026, 1, 4), 1.02, 1.02, "DEMO01", "fixture"),
        ]
    }
    try:
        build_preview(gap_config, gap_rows)
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
    review_path = Path(second["review_path"])
    check(
        not second["sheets"]
        and second["preview_path"] is None
        and second["approval_kind"] == "validated-no-change"
        and review_path.is_file()
        and "待写入日期数：0" in review_path.read_text(encoding="utf-8"),
        "second run must produce a reviewable no-change baseline without a workbook copy",
    )
    check(
        (runtime / "plan.json").is_file(),
        "a validated no-change preview must leave an approvable plan",
    )
    no_change_master = book.read_bytes()
    no_change_backups = set((runtime / "backups").glob("*"))
    review_path.write_text(
        review_path.read_text(encoding="utf-8") + "tampered\n", encoding="utf-8"
    )
    try:
        commit(config)
    except CommitError:
        pass
    else:
        raise AssertionError("commit accepted a changed no-change baseline report")
    check(
        book.read_bytes() == no_change_master,
        "tampered no-change approval changed the master workbook",
    )
    second = build_preview(config, rows)
    reviewed = subprocess.run(
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
    reviewed_report = json.loads(reviewed.stdout)
    check(
        reviewed.returncode == 0
        and not reviewed_report["changed"]
        and reviewed_report["approved_baseline"]
        and reviewed_report["automatic_updates"]["approved"],
        "reviewed no-change baseline did not approve automatic updates",
    )
    check(
        book.read_bytes() == no_change_master
        and set((runtime / "backups").glob("*")) == no_change_backups,
        "no-change baseline approval wrote the workbook or created a backup",
    )
    automatic_noop = automatic_update(config, rows)
    check(
        not automatic_noop["changed"]
        and automatic_noop["rows"] == 0
        and not (runtime / "plan.json").exists(),
        "approved scheduled no-op did not finish successfully and clean its staging plan",
    )
    revoke()

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

    reserved_book = runtime / "summary-reserved-cold-start.xlsx"
    reserved_workbook = openpyxl.Workbook()
    reserved_sheet = reserved_workbook.active
    reserved_sheet.title = "Reserved Fund"
    reserved_sheet.append(
        ["Product Code", "NAV Date", "Unit NAV", "Cumulative NAV", "Weekly Return"]
    )
    reserved_sheet.append(["DEMO43", dt.date(2026, 1, 1), None, None, None])
    reserved_sheet.append(["累计", None, None, None, None])
    reserved_workbook.save(reserved_book)
    reserved_workbook.close()
    reserved_config = config_for(runtime, reserved_book)
    reserved_config["routes"][0].update(
        {
            "sheet": "Reserved Fund",
            "code": "DEMO43",
            "benchmark": None,
            "return_basis": "cumulative",
        }
    )
    reserved_rows = {
        "Reserved Fund": [
            NavRow(date, value, value, "DEMO43", "fixture")
            for date, value in (
                (dt.date(2026, 1, 2), 1.0),
                (dt.date(2026, 1, 5), 1.002),
                (dt.date(2026, 1, 6), 1.004),
                (dt.date(2026, 1, 7), 1.006),
                (dt.date(2026, 1, 8), 1.008),
                (dt.date(2026, 1, 9), 1.01),
                (dt.date(2026, 1, 12), 1.012),
                (dt.date(2026, 1, 13), 1.014),
                (dt.date(2026, 1, 14), 1.016),
                (dt.date(2026, 1, 15), 1.018),
                (dt.date(2026, 1, 16), 1.02),
            )
        ]
    }
    reserved_validation = validate_history(reserved_config, reserved_rows)
    check(
        reserved_validation["passed"]
        and reserved_validation["warnings"]
        and reserved_validation["routes"][0]["cold_start_kind"]
        == "summary-reserved-row"
        and reserved_validation["routes"][0]["data_frequency"] == "weekly",
        "reserved summary row was not recognized as a safe cold start",
    )
    reserved_plan = build_preview(
        reserved_config, reserved_rows, reserved_validation["warnings"]
    )
    reserved_sheet_plan = reserved_plan["sheets"][0]
    check(
        reserved_sheet_plan["insert_count"] == 2
        and reserved_sheet_plan["populated_count"] == 3
        and reserved_sheet_plan["filled_existing_rows"] == [2]
        and reserved_sheet_plan["new_dates"]
        == ["2026-01-02", "2026-01-09", "2026-01-16"],
        "reserved-row plan did not distinguish the filled row from inserted rows",
    )
    reserved_preview = openpyxl.load_workbook(
        reserved_plan["preview_path"], data_only=False
    )
    try:
        preview_sheet = reserved_preview["Reserved Fund"]
        check(
            (
                preview_sheet["B2"].value == dt.datetime(2026, 1, 2)
                or preview_sheet["B2"].value == dt.date(2026, 1, 2)
            )
            and preview_sheet["C2"].value == 1.0
            and preview_sheet["D2"].value == 1.0
            and preview_sheet["C4"].value == 1.02,
            "reserved first row or later email history was not populated",
        )
        check(
            preview_sheet["A5"].value == "累计"
            and preview_sheet["E4"].value == "=D4/D3-1"
            and preview_sheet["E5"].value == "=D4/D2-1",
            "summary row or managed formulas were not preserved after cold start",
        )
    finally:
        reserved_preview.close()
    if use_com:
        reserved_result = commit(reserved_config)
        application = str(reserved_result["application"])
        check(
            reserved_result["changed"] and reserved_result["rows"] == 3,
            "COM reserved-row cold start did not report all populated dates",
        )
    else:
        shutil.copy2(reserved_plan["preview_path"], reserved_book)
    reserved_second_validation = validate_history(reserved_config, reserved_rows)
    check(
        reserved_second_validation["passed"]
        and not reserved_second_validation["warnings"],
        "reserved-row cold start did not become ordinary verified history",
    )
    reserved_second = build_preview(reserved_config, reserved_rows)
    check(
        not reserved_second["sheets"] and reserved_second["preview_path"] is None,
        "reserved-row cold start was not idempotent after the first write",
    )

    unsafe_reserved_book = runtime / "unsafe-reserved-row.xlsx"
    unsafe_reserved_workbook = openpyxl.Workbook()
    unsafe_reserved_sheet = unsafe_reserved_workbook.active
    unsafe_reserved_sheet.title = "Reserved Fund"
    unsafe_reserved_sheet.append(
        ["Product Code", "NAV Date", "Unit NAV", "Cumulative NAV", "Note"]
    )
    unsafe_reserved_sheet.append(
        ["DEMO43", dt.date(2026, 1, 1), None, None, "keep this content"]
    )
    unsafe_reserved_sheet.append(["累计", None, None, None, None])
    unsafe_reserved_workbook.save(unsafe_reserved_book)
    unsafe_reserved_workbook.close()
    unsafe_reserved_config = config_for(runtime, unsafe_reserved_book)
    unsafe_reserved_config["routes"][0].update(
        {"sheet": "Reserved Fund", "code": "DEMO43", "benchmark": None}
    )
    unsafe_reserved_validation = validate_history(unsafe_reserved_config, reserved_rows)
    check(
        not unsafe_reserved_validation["passed"]
        and not unsafe_reserved_validation["routes"][0]["cold_start"],
        "a row containing extra business content was treated as a placeholder",
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


def template_tests(runtime: Path, use_com: bool) -> str | None:
    from nav_automation import approve, automatic_update
    from nav_commit import commit, ensure_process_exit, spreadsheet_app
    from nav_config import validate_config
    from nav_parse import NavRow
    from nav_template import TemplateError, init_template
    from nav_workbook import build_preview, validate_history

    target = runtime / "全新脱敏模板.xlsx"
    route_specs = [
        ("周度无指数", "WEEK00", "weekly", None),
        ("周度有指数", "WEEK01", "weekly", "示例指数"),
        ("日度无指数", "DAY000", "daily", None),
        ("日度有指数", "DAY001", "daily", "示例指数"),
    ]
    routes = []
    for index, (sheet, code, frequency, benchmark_name) in enumerate(route_specs, 1):
        benchmark = None
        if benchmark_name:
            benchmark = {
                "source_sheet": "指数数据",
                "source_type": "level",
                "source_date": "A",
                "source_value": "B",
                "display_name": benchmark_name,
            }
        routes.append(
            {
                "sender": f"sender{index}@example.invalid",
                "subject_contains": code,
                "sheet": sheet,
                "sheet_mode": "template",
                "code": code,
                "product_name": f"示例产品{index}",
                "parser": "auto",
                "allow_sender_only": False,
                "cumulative_policy": "require",
                "return_basis": "cumulative",
                "return_frequency": frequency,
                "data_frequency": frequency,
                "series_start": "2026-01-02",
                "benchmark": benchmark,
            }
        )
    config = {
        "schema_version": 1,
        "runtime_id": "00000000-0000-4000-8000-000000000170",
        "workbook_path": str(target.resolve()),
        "workbook_mode": "bundled-template",
        "imap": {
            "host": "imap.example.invalid",
            "port": 993,
            "user": "user@example.invalid",
            "mailbox": "INBOX",
            "lookback_days": 180,
        },
        "routes": routes,
        "column_overrides": {},
        "style": {"mode": "cn-red-up-green-down", "zero_threshold": 0.00005},
        "schedule": [{"days": ["MON", "WED", "FRI"], "time": "09:00"}],
        "validation": {"minimum_history_dates": 2, "tolerance": 0.000001},
    }
    validate_config(config)
    initialized = init_template(config)
    check(
        initialized["product_sheets"]
        == ["周度无指数", "周度有指数", "日度无指数", "日度有指数"]
        and initialized["benchmark_source_sheets"] == ["指数数据"]
        and target.is_file(),
        "template initializer did not create four product sheets and one shared source sheet",
    )
    try:
        init_template(config)
    except TemplateError as exc:
        check("refusing to overwrite" in str(exc), "overwrite refusal is unclear")
    else:
        raise AssertionError("template initializer overwrote an existing workbook")

    workbook = openpyxl.load_workbook(target, data_only=False)
    dates_by_frequency = {
        "weekly": [
            dt.date(2026, 1, 2),
            dt.date(2026, 1, 9),
            dt.date(2026, 1, 16),
        ],
        "daily": [
            dt.date(2026, 1, 2),
            dt.date(2026, 1, 5),
            dt.date(2026, 1, 6),
        ],
    }
    try:
        check(
            workbook.sheetnames
            == ["周度无指数", "周度有指数", "日度无指数", "日度有指数", "指数数据"],
            "generated workbook sheet order is wrong",
        )
        expected_headers = {
            "周度无指数": [
                "产品代码",
                "产品名称",
                "单位净值",
                "累计单位净值",
                "净值日期",
                "周收益",
            ],
            "周度有指数": [
                "产品代码",
                "产品名称",
                "单位净值",
                "累计单位净值",
                "净值日期",
                "周收益",
                "示例指数",
                "指数收益(周度)",
                "超额(周度)",
            ],
            "日度无指数": [
                "产品代码",
                "产品名称",
                "单位净值",
                "累计单位净值",
                "净值日期",
                "日收益",
                "周收益",
            ],
            "日度有指数": [
                "产品代码",
                "产品名称",
                "单位净值",
                "累计单位净值",
                "净值日期",
                "日收益",
                "周收益",
                "示例指数收益(日度)",
                "超额(日度)",
            ],
        }
        for name, headers in expected_headers.items():
            sheet = workbook[name]
            check(
                [sheet.cell(2, column).value for column in range(1, len(headers) + 1)]
                == headers,
                f"{name}: generated headers are wrong",
            )
            check(
                sheet["A1"].value is None
                and list(sheet.merged_cells.ranges)
                and sheet["B2"].font.name == "等线"
                and sheet["B2"].font.sz == 10
                and sheet["F3"].fill.fgColor.rgb.endswith("FFF2CC")
                and sum(len(item.rules) for item in sheet.conditional_formatting) >= 2,
                f"{name}: template style or dynamic return coloring is incomplete",
            )
        check(
            workbook["日度有指数"]["I3"].fill.fgColor.rgb.endswith("FFFF00"),
            "daily excess column is not bright yellow",
        )
        source = workbook["指数数据"]
        check(
            [source.cell(1, column).value for column in range(1, 4)]
            == ["日期", "示例指数点位", "来源"],
            "shared index source headers are wrong",
        )
        source_points = {
            dt.date(2026, 1, 2): 100.0,
            dt.date(2026, 1, 5): 100.2,
            dt.date(2026, 1, 6): 100.3,
            dt.date(2026, 1, 9): 101.0,
            dt.date(2026, 1, 16): 100.5,
        }
        for row, (date, level) in enumerate(sorted(source_points.items()), 2):
            source.cell(row, 1).value = date
            source.cell(row, 2).value = level
            source.cell(row, 3).value = "https://example.invalid/index"
        workbook.save(target)
    finally:
        workbook.close()

    values = {
        "周度无指数": [1.0, 1.02, 1.03],
        "周度有指数": [1.0, 1.01, 1.02],
        "日度无指数": [1.0, 0.99, 1.01],
        "日度有指数": [1.0, 0.98, 0.97],
    }

    def rows(count: int) -> dict[str, list[NavRow]]:
        return {
            sheet: [
                NavRow(date, value, value, code, "fixture")
                for date, value in zip(
                    dates_by_frequency[frequency][:count], values[sheet][:count]
                )
            ]
            for sheet, code, frequency, _benchmark in route_specs
        }

    first_rows = rows(1)
    first_validation = validate_history(config, first_rows)
    check(
        first_validation["passed"]
        and len(first_validation["warnings"]) == 4
        and all(
            item["cold_start_kind"] == "bundled-template"
            for item in first_validation["routes"]
        ),
        "one-date bundled template cold start was not accepted with visible warnings",
    )
    first_plan = build_preview(config, first_rows, first_validation["warnings"])
    check(
        len(first_plan["sheets"]) == 4
        and all(item["new_dates"] == ["2026-01-02"] for item in first_plan["sheets"]),
        "one-date template preview is incomplete",
    )
    application = None
    if use_com:
        application = str(commit(config)["application"])
    else:
        shutil.copy2(first_plan["preview_path"], target)
    one_date_validation = validate_history(config, first_rows)
    check(
        one_date_validation["passed"] and one_date_validation["warnings"],
        "template did not keep the cold-start warning until two dates were verified",
    )

    second_rows = rows(2)
    second_validation = validate_history(config, second_rows)
    check(
        second_validation["passed"] and second_validation["warnings"],
        "second date was not allowed to enter the reviewed template preview",
    )
    second_plan = build_preview(config, second_rows, second_validation["warnings"])
    preview = openpyxl.load_workbook(second_plan["preview_path"], data_only=False)
    try:
        weekly = preview["周度有指数"]
        daily = preview["日度有指数"]
        check(
            weekly["F4"].value == "=D4/D3-1"
            and weekly["G3"].value == "='指数数据'!B2"
            and weekly["G4"].value == "='指数数据'!B5"
            and weekly["H4"].value == "=G4/G3-1"
            and weekly["I4"].value == "=F4-H4",
            "weekly benchmark return or excess formulas are wrong: "
            f"{[weekly[cell].value for cell in ('F4', 'G3', 'G4', 'H4', 'I4')]}",
        )
        check(
            daily["F4"].value == "=D4/D3-1"
            and daily["G4"].value == "=D4/D3-1"
            and daily["H4"].value == "='指数数据'!B3/'指数数据'!B2-1"
            and daily["I4"].value == "=F4-H4"
            and daily["F5"].value == "=D4/D3-1"
            and daily["G5"].value is None,
            "daily/weekly product returns or daily excess formulas are wrong",
        )
        check(
            set(second_plan["sheets"][3]["return_columns"]) == {6, 7, 8, 9},
            "daily benchmark style plan does not include all return columns",
        )
    finally:
        preview.close()
    if use_com:
        application = str(commit(config)["application"])
        app, _progid, process_id = spreadsheet_app()
        book = None
        try:
            book = app.Workbooks.Open(
                str(target.resolve()), ReadOnly=True, UpdateLinks=0
            )
            app.CalculateFull()
            red = int(book.Worksheets("周度无指数").Range("F4").Font.Color)
            green = int(book.Worksheets("日度有指数").Range("F4").Font.Color)
            check(
                red == 255 and green == 176 * 256 + 80 * 65536,
                "Excel/WPS did not display red-up/green-down fixed colors",
            )
        finally:
            if book is not None:
                book.Close(SaveChanges=False)
            app.Quit()
            app = None
            gc.collect()
            ensure_process_exit(process_id)
    else:
        shutil.copy2(second_plan["preview_path"], target)
    strict = validate_history(config, second_rows)
    check(
        strict["passed"] and not strict["warnings"],
        "two verified template dates did not restore strict validation",
    )
    no_op = build_preview(config, second_rows)
    check(
        not no_op["sheets"] and no_op["preview_path"] is None,
        "template update is not idempotent after two verified dates",
    )

    third_rows = rows(3)
    third_validation = validate_history(config, third_rows)
    check(
        third_validation["passed"] and not third_validation["warnings"],
        "later template increment did not pass strict history validation",
    )
    if use_com:
        approve(config)
        result = automatic_update(config, third_rows)
        application = str(result["application"])
        check(
            result["changed"] and result["rows"] == 4,
            "scheduled automatic update did not write one date to every product sheet",
        )
    else:
        third_plan = build_preview(config, third_rows)
        expected_increment = {
            "周度无指数": ["2026-01-16"],
            "周度有指数": ["2026-01-16"],
            "日度无指数": ["2026-01-06"],
            "日度有指数": ["2026-01-06"],
        }
        check(
            all(
                item["new_dates"] == expected_increment[item["sheet"]]
                for item in third_plan["sheets"]
            ),
            "later template preview did not contain exactly one incremental date",
        )
        shutil.copy2(third_plan["preview_path"], target)
    final = validate_history(config, third_rows)
    check(
        final["passed"]
        and not final["warnings"]
        and not build_preview(config, third_rows)["sheets"],
        "template workflow is not strict and idempotent after the later increment",
    )
    return application


def product_lifecycle_tests(runtime: Path, use_com: bool) -> str | None:
    from nav_automation import approve, status as approval_status
    from nav_config import load_config, validate_config
    from nav_parse import NavRow
    from nav_product_workbook import prepare_clone_spec
    from nav_products import add, adopt, clone, pause, resume, status, sync
    from nav_template import init_template
    from nav_workbook import validate_history

    target = runtime / "产品生命周期模板.xlsx"
    config_path = runtime / "product-lifecycle-config.json"
    initial_route = {
        "sender": "initial@example.invalid",
        "subject_contains": "BASE001",
        "sheet": "初始产品",
        "sheet_mode": "template",
        "code": "BASE001",
        "product_name": "初始示例产品",
        "parser": "auto",
        "paused": False,
        "allow_sender_only": False,
        "cumulative_policy": "require",
        "return_basis": "cumulative",
        "return_frequency": "weekly",
        "data_frequency": "weekly",
        "max_staleness_days": 14,
        "benchmark": {
            "source_sheet": "指数数据",
            "source_type": "level",
            "source_date": "A",
            "source_value": "B",
            "display_name": "示例指数",
        },
    }
    config = {
        "schema_version": 1,
        "runtime_id": "00000000-0000-4000-8000-000000000171",
        "workbook_path": str(target.resolve()),
        "workbook_mode": "bundled-template",
        "imap": {
            "host": "imap.example.invalid",
            "port": 993,
            "user": "user@example.invalid",
            "mailbox": "INBOX",
            "lookback_days": 180,
        },
        "routes": [initial_route],
        "column_overrides": {},
        "style": {"mode": "cn-red-up-green-down", "zero_threshold": 0.00005},
        "schedule": [{"days": ["MON", "WED", "FRI"], "time": "09:00"}],
        "validation": {"minimum_history_dates": 2, "tolerance": 0.000001},
    }
    validate_config(config)
    init_template(config)
    config_path.write_text(
        json.dumps(config, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    proposal = {
        "passed": True,
        "scan": {},
        "candidates": [
            {
                "sender": "new@example.invalid",
                "message_count": 3,
                "recent_message_times": [],
                "subject_examples": ["NEW002 周净值"],
                "detected_codes": ["NEW002"],
                "first_date": "2026-01-02",
                "latest_date": "2026-01-16",
                "observations": [
                    {
                        "date": "2026-01-02",
                        "code": "NEW002",
                        "unit": 1.01,
                        "cumulative": 1.01,
                        "source": "body",
                    },
                    {
                        "date": "2026-01-09",
                        "code": "NEW002",
                        "unit": 1.02,
                        "cumulative": 1.02,
                        "source": "body",
                    },
                ],
            }
        ],
        "warnings": [],
        "errors": [],
    }
    (runtime / "route-proposals.json").write_text(
        json.dumps(proposal, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    before = sync(config, refresh=False)
    check(
        before["passed"]
        and len(before["new_candidates"]) == 1
        and before["new_candidates"][0]["code"] == "NEW002",
        "products sync did not identify the new candidate",
    )
    result = add(
        config,
        config_path,
        proposal_index=1,
        sheet="新增产品",
        frequency="weekly",
        product_name="新增示例产品",
        benchmark_source_sheet="新增指数数据",
        benchmark_display_name="新增示例指数",
    )
    check(
        result["changed"]
        and result["requires_preview_approval"]
        and result["template"]["product_sheet"] == "新增产品"
        and Path(result["template"]["backup"]).is_file(),
        "products add did not create a backed-up template page",
    )
    updated = load_config(config_path)
    workbook = openpyxl.load_workbook(target, data_only=False)
    try:
        check(
            workbook.sheetnames == ["初始产品", "新增产品", "指数数据", "新增指数数据"]
            and workbook["新增产品"]["F3"].fill.fgColor.rgb.endswith("FFF2CC")
            and workbook["新增产品"]["G2"].value == "新增示例指数"
            and workbook["新增指数数据"]["B1"].value == "新增示例指数点位"
            and sum(
                len(item.rules) for item in workbook["新增产品"].conditional_formatting
            )
            >= 2,
            "new template product page has wrong order or style",
        )
    finally:
        workbook.close()
    check(
        len(updated["routes"]) == 2
        and not approval_status(updated)["approved"]
        and not sync(updated, refresh=False)["new_candidates"],
        "new product route or approval state is wrong",
    )

    approve(updated)
    paused_result = pause(
        updated, config_path, sheet="新增产品", reason="示例产品暂时停更"
    )
    paused = load_config(config_path)
    check(
        paused_result["approval_preserved"]
        and approval_status(paused)["approved"]
        and paused["routes"][1]["paused"]
        and target.is_file(),
        "pausing one product did not preserve history and narrowed approval",
    )
    resumed_result = resume(paused, config_path, sheet="新增产品")
    resumed = load_config(config_path)
    check(
        resumed_result["requires_preview_approval"]
        and not approval_status(resumed)["approved"]
        and not resumed["routes"][1]["paused"],
        "resuming a product did not require renewed preview approval",
    )
    report = status(resumed)
    check(
        report["active"] == 2
        and report["paused"] == 0
        and not report["workbook_missing_sheets"],
        "products status returned the wrong lifecycle state",
    )
    cli_status = subprocess.run(
        [
            sys.executable,
            "-X",
            "utf8",
            "navctl.py",
            "--config",
            str(config_path),
            "products",
            "status",
        ],
        cwd=runtime,
        capture_output=True,
        text=True,
        encoding="utf-8",
    )
    check(
        cli_status.returncode == 0 and json.loads(cli_status.stdout)["active"] == 2,
        "navctl products status CLI is not wired correctly",
    )

    workbook_before = target.read_bytes()
    config_before = config_path.read_bytes()
    try:
        add(
            resumed,
            config_path,
            proposal_index=1,
            sheet="新增产品",
            frequency="weekly",
        )
    except (RuntimeError, ValueError):
        pass
    else:
        raise AssertionError("products add accepted a duplicate managed worksheet")
    check(
        target.read_bytes() == workbook_before
        and config_path.read_bytes() == config_before,
        "duplicate product rejection changed the workbook or configuration",
    )

    existing_target = runtime / "产品生命周期已有表.xlsx"
    existing_book = openpyxl.Workbook()
    source_sheet = existing_book.active
    source_sheet.title = "参考产品"
    source_sheet.append(["参考产品费率说明", None, None, None, None, None])
    source_sheet.append(
        ["产品代码", "产品名称", "单位净值", "累计单位净值", "净值日期", "收益（周度）"]
    )
    source_sheet.append(
        ["BASE001", "参考示例产品", 1.0, 1.0, dt.date(2025, 12, 26), "/"]
    )
    source_sheet.append(
        ["BASE001", "参考示例产品", 1.01, 1.01, dt.date(2026, 1, 2), "=D4/D3-1"]
    )
    source_sheet.append(["累计", None, None, None, None, "=D4/D3-1"])
    thin = Side(style="thin", color="000000")
    for row in range(1, 6):
        for column in range(1, 7):
            cell = source_sheet.cell(row, column)
            cell.font = Font(name="等线", size=10, bold=row == 2)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            if row == 2:
                cell.fill = PatternFill("solid", fgColor="9BC2E6")
            if column == 6 and row >= 3:
                cell.fill = PatternFill("solid", fgColor="FFF2CC")
                cell.number_format = "0.00%"
    source_sheet.column_dimensions["B"].width = 28
    prepared = existing_book.create_sheet("用户新建页")
    prepared.append(["用户自行填写的产品说明", None, None, None, None, None])
    prepared.append(
        ["产品代码", "产品名称", "单位净值", "累计单位净值", "净值日期", "收益（周度）"]
    )
    prepared.append(
        ["NEW002", "新增示例产品", 1.01, None, dt.date(2026, 1, 2), None]
    )
    prepared.append(["累计", None, None, None, None, None])
    existing_book.create_sheet("分析页")["A1"] = "用户分析"
    existing_book.save(existing_target)
    existing_book.close()
    existing_config_path = runtime / "existing-product-config.json"
    existing_config = deepcopy(config)
    existing_config["runtime_id"] = "00000000-0000-4000-8000-000000000172"
    existing_config["workbook_path"] = str(existing_target.resolve())
    existing_config["workbook_mode"] = "existing"
    existing_config["style"] = {"mode": "infer", "zero_threshold": 0.00005}
    existing_config["routes"] = [
        {
            **initial_route,
            "sheet": "参考产品",
            "sheet_mode": "summary",
            "data_frequency": "auto",
            "benchmark": None,
        }
    ]
    validate_config(existing_config)
    existing_config_path.write_text(
        json.dumps(existing_config, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    adopt_proposal = deepcopy(proposal)
    adopt_proposal["candidates"][0]["detected_codes"] = ["NEW002", "OTHER999"]
    adopt_proposal["candidates"][0]["observations"].append(
        {
            "date": "2026-01-02",
            "code": "OTHER999",
            "unit": 2.01,
            "cumulative": 2.01,
            "source": "body",
        }
    )
    (runtime / "route-proposals.json").write_text(
        json.dumps(adopt_proposal, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    existing_before = existing_target.read_bytes()
    existing_result = adopt(
        existing_config,
        existing_config_path,
        proposal_index=1,
        sheet="用户新建页",
    )
    existing_updated = load_config(existing_config_path)
    existing_status = status(existing_updated)
    check(
        existing_result["changed"]
        and existing_result["action"] == "adopted-existing-sheet",
        "products adopt did not report a successful adoption",
    )
    check(
        existing_result["inference"]["sheet_mode"] == "summary"
        and existing_result["inference"]["frequency"] == "weekly",
        f"products adopt inferred the wrong layout: {existing_result['inference']}",
    )
    check(
        existing_result["inference"]["code"] == "NEW002",
        "products adopt did not use the worksheet to resolve a multi-code sender",
    )
    check(
        existing_target.read_bytes() == existing_before,
        "products adopt changed the user workbook",
    )
    check(
        len(existing_updated["routes"]) == 2
        and existing_status["unmanaged_workbook_sheets"] == ["分析页"],
        f"products adopt saved the wrong route status: {existing_status}",
    )
    adopted_validation = validate_history(
        existing_updated,
        {
            "参考产品": [
                NavRow(
                    dt.date(2025, 12, 26), 1.0, 1.0, "BASE001", "fixture"
                ),
                NavRow(dt.date(2026, 1, 2), 1.01, 1.01, "BASE001", "fixture"),
            ],
            "用户新建页": [
                NavRow(dt.date(2026, 1, 2), 1.01, 1.01, "NEW002", "fixture"),
                NavRow(dt.date(2026, 1, 9), 1.02, 1.02, "NEW002", "fixture"),
            ],
        },
    )
    adopted_report = next(
        item
        for item in adopted_validation["routes"]
        if item["sheet"] == "用户新建页"
    )
    check(
        adopted_validation["passed"]
        and adopted_report["cold_start_kind"] == "summary-reserved-row",
        f"partial user-prepared onboarding row was not accepted safely: {adopted_validation}",
    )
    spec = prepare_clone_spec(
        existing_updated, existing_updated["routes"][0], "照参考新增"
    )
    check(
        spec.header_row == 2
        and spec.data_row == 3
        and spec.target_summary_row == 4,
        "clone specification did not preserve the reference sheet structure",
    )

    application = None
    if use_com:
        clone_proposal = deepcopy(proposal)
        clone_proposal["candidates"][0] = {
            "sender": "clone@example.invalid",
            "message_count": 2,
            "recent_message_times": [],
            "subject_examples": ["CLONE003 周净值"],
            "detected_codes": ["CLONE003"],
            "first_date": "2026-02-06",
            "latest_date": "2026-02-13",
            "observations": [
                {
                    "date": "2026-02-06",
                    "code": "CLONE003",
                    "unit": 1.1,
                    "cumulative": 1.1,
                    "source": "body",
                },
                {
                    "date": "2026-02-13",
                    "code": "CLONE003",
                    "unit": 1.2,
                    "cumulative": 1.2,
                    "source": "body",
                },
            ],
        }
        (runtime / "route-proposals.json").write_text(
            json.dumps(clone_proposal, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        clone_result = clone(
            existing_updated,
            existing_config_path,
            proposal_index=1,
            sheet="照参考新增",
            copy_from="参考产品",
            product_name="复制新增示例产品",
        )
        application = clone_result["workbook"]["application"]
        cloned = openpyxl.load_workbook(existing_target, data_only=False)
        try:
            check(
                cloned.sheetnames
                == ["参考产品", "照参考新增", "用户新建页", "分析页"]
                and cloned["照参考新增"]["A1"].value is None
                and cloned["照参考新增"]["A2"].value == "产品代码"
                and cloned["照参考新增"]["A3"].value == "CLONE003"
                and cloned["照参考新增"]["B3"].value == "复制新增示例产品"
                and cloned["照参考新增"]["C3"].value is None
                and cloned["照参考新增"]["D3"].value is None
                and cloned["照参考新增"]["E3"].value.date()
                == dt.date(2026, 2, 6)
                and cloned["照参考新增"]["F3"].value is None
                and cloned["照参考新增"]["A4"].value == "累计"
                and cloned["照参考新增"]["F4"].value is None
                and cloned["参考产品"]["A1"].value == "参考产品费率说明",
                "products clone did not preserve the format or clear source business data",
            )
        finally:
            cloned.close()
        cloned_config = load_config(existing_config_path)
        check(
            len(cloned_config["routes"]) == 3
            and cloned_config["routes"][-1]["sheet"] == "照参考新增"
            and Path(clone_result["workbook"]["backup"]).is_file(),
            "products clone did not save its route or backup",
        )
    return application


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--runtime", required=True)
    parser.add_argument("--com", action="store_true")
    args = parser.parse_args()
    runtime = Path(args.runtime).resolve()
    print("[1/6] 检查自动解析、本地解析器、精确发件人和冲突拦截")
    parser_tests(runtime)
    print("      PASS")
    print("[2/6] 检查空路由、暂停路由和无邮件失败关闭")
    route_state_tests(runtime)
    print("      PASS")
    print("[3/6] 检查崩溃恢复和并发运行锁")
    lock_tests(runtime)
    print("      PASS")
    print("[4/6] 检查严格表、空表冷启动、分析页保留、补录、基准和幂等性")
    application = workbook_tests(runtime, args.com)
    print("      PASS")
    print("[5/6] 检查四类脱敏模板、共享指数、冷启动、增量和拒绝覆盖")
    template_application = template_tests(runtime, args.com)
    application = template_application or application
    print("      PASS")
    print("[6/6] 检查产品发现、接管已建 Sheet、照参考新建、暂停、恢复和备份")
    product_application = product_lifecycle_tests(runtime, args.com)
    application = product_application or application
    print("      PASS")
    if args.com:
        print(f"      Excel/WPS COM、文件占用提示及缓存数值：PASS（{application}）")
    else:
        print("      未启用 COM；如需验证正式写入，请添加 --com")
    print("selftest_driver: PASS（全程仅使用虚构数据）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
