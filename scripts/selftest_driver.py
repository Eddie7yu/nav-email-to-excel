#!/usr/bin/env python3
from __future__ import annotations

import argparse
import datetime as dt
import json
import shutil
import sys
from email.message import EmailMessage
from pathlib import Path

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
    from nav_parse import (
        ParseError,
        choose_route_rows,
        parse_number,
        rows_from_message,
        rows_from_text,
    )
    from nav_mail import exact_from_matches, imap_date

    check(
        imap_date(dt.date(2026, 7, 9)) == "09-Jul-2026",
        "IMAP dates must use fixed English month names",
    )
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


def workbook_tests(runtime: Path, use_com: bool) -> None:
    sys.path.insert(0, str(runtime))
    from nav_commit import CommitError, commit
    from nav_config import ConfigError, load_config, validate_config
    from nav_parse import NavRow
    from nav_workbook import WorkbookError, build_preview, validate_history

    book = runtime / "脱敏 示例.xlsx"
    create_book(book)
    config = config_for(runtime, book)
    invalid = json.loads(json.dumps(config))
    invalid["routes"][0]["series_start"] = "2026-99-99"
    invalid["unexpected"] = True
    try:
        validate_config(invalid)
    except ConfigError:
        pass
    else:
        raise AssertionError("invalid dates and unknown config fields must be rejected")
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

    if use_com:
        result = commit(config)
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


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--runtime", required=True)
    parser.add_argument("--com", action="store_true")
    args = parser.parse_args()
    runtime = Path(args.runtime).resolve()
    parser_tests(runtime)
    workbook_tests(runtime, args.com)
    print("selftest_driver: PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
