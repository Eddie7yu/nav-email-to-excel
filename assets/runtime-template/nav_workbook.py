from __future__ import annotations

import datetime as dt
import json
import re
import shutil
import statistics
import uuid
from collections import Counter
from copy import copy
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string, get_column_letter, quote_sheetname

from nav_config import (
    ROOT,
    STATE_ROOT,
    active_routes,
    normalize_code,
    write_json_atomic,
)
from nav_parse import NavRow, parse_date, parse_number


HEADER_WORDS = {
    "date": ("净值日期", "估值日期", "业务日期", "nav date", "date", "日期"),
    "code": ("产品代码", "基金代码", "product code", "fund code", "代码"),
    "name": ("产品名称", "基金名称", "product name", "fund name", "名称"),
    "unit": ("单位净值", "份额净值", "unit nav"),
    "cumulative": ("累计单位净值", "累计净值", "cumulative nav"),
    "return": ("产品收益", "基金收益", "周收益", "日收益", "return"),
    "benchmark_level": ("指数点位", "基准点位", "benchmark level", "index level"),
    "benchmark_return": ("指数收益", "基准收益", "benchmark return", "index return"),
    "excess": ("超额", "excess", "alpha"),
}
APPEND_HEADER_LABELS = {field: words[0] for field, words in HEADER_WORDS.items()}
TOTAL_WORDS = {"累计", "合计", "total", "cumulative"}


class WorkbookError(RuntimeError):
    pass


@dataclass
class Layout:
    sheet: str
    header_row: int
    data_start: int
    summary_row: int
    last_data_row: int
    columns: dict[str, int]
    mode: str = "summary"
    headers_to_write: dict[int, str] | None = None


def _norm(value: Any) -> str:
    return re.sub(r"[\s_:/：()（）\[\]-]", "", str(value or "")).lower()


def _field(value: Any) -> str | None:
    text = _norm(value)
    if not text:
        return None
    # Match date before unit NAV: a header such as "NAV Date" contains both
    # concepts, and treating it as NAV would make otherwise valid layouts fail.
    order = (
        "cumulative",
        "date",
        "code",
        "name",
        "benchmark_return",
        "benchmark_level",
        "excess",
        "return",
        "unit",
    )
    for field in order:
        if any(_norm(word) in text for word in HEADER_WORDS[field]):
            return field
    return None


def _column(value: Any) -> int:
    if isinstance(value, int) and value > 0:
        return value
    text = str(value or "").strip().upper()
    if text.isdigit() and int(text) > 0:
        return int(text)
    if re.fullmatch(r"[A-Z]{1,3}", text):
        return column_index_from_string(text)
    raise WorkbookError(f"Invalid column override: {value!r}")


def _has_values(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    start_row: int = 1,
    end_row: int | None = None,
) -> bool:
    end = sheet.max_row if end_row is None else end_row
    for row in range(start_row, end + 1):
        for column in range(1, sheet.max_column + 1):
            value = sheet.cell(row, column).value
            if value is not None and value != "":
                return True
    return False


def _append_columns(route: dict[str, Any], override: dict[str, Any]) -> dict[str, int]:
    explicit = {
        field: _column(override[field]) for field in HEADER_WORDS if field in override
    }
    if explicit:
        return explicit
    fields = ["date"]
    if normalize_code(route.get("code")):
        fields.append("code")
    if str(route.get("product_name") or "").strip():
        fields.append("name")
    fields.extend(("unit", "cumulative"))
    if route.get("benchmark"):
        fields.append("return")
        if str(route["benchmark"].get("source_type", "level")) == "level":
            fields.append("benchmark_level")
        fields.extend(("benchmark_return", "excess"))
    return {field: index for index, field in enumerate(fields, 1)}


def _require_append_identity(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    columns: dict[str, int],
    route: dict[str, Any],
) -> None:
    code_identity = bool(columns.get("code") and normalize_code(route.get("code")))
    name_identity = bool(
        columns.get("name") and str(route.get("product_name") or "").strip()
    )
    if not code_identity and not name_identity:
        raise WorkbookError(
            f"{sheet.title}: append mode needs a writable product code or product name column"
        )


def discover_layout(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    override: dict[str, Any] | None = None,
    route: dict[str, Any] | None = None,
) -> Layout:
    override = override or {}
    route = route or {}
    mode = str(route.get("sheet_mode", "summary"))
    if mode not in {"summary", "append"}:
        raise WorkbookError(f"{sheet.title}: unsupported sheet mode {mode}")
    header_row = int(override.get("header_row") or 0)
    columns: dict[str, int] = {}
    sheet_is_blank = not _has_values(sheet)
    if mode == "append" and sheet_is_blank:
        header_row = header_row or 1
        columns = _append_columns(route, override)
        if not {"date", "unit"} <= columns.keys():
            raise WorkbookError(
                f"{sheet.title}: append mode requires explicit date and unit NAV columns"
            )
        _require_append_identity(sheet, columns, route)
        headers = {
            column: APPEND_HEADER_LABELS[field] for field, column in columns.items()
        }
        return Layout(
            sheet.title,
            header_row,
            header_row + 1,
            header_row + 1,
            header_row,
            columns,
            mode,
            headers,
        )
    if header_row:
        for field in HEADER_WORDS:
            if field in override:
                columns[field] = _column(override[field])
    else:
        best_score = -1
        for row in range(1, min(sheet.max_row, 30) + 1):
            candidate: dict[str, int] = {}
            for column in range(1, min(sheet.max_column, 80) + 1):
                field = _field(sheet.cell(row, column).value)
                if field and field not in candidate:
                    candidate[field] = column
            score = len(candidate) + (3 if {"date", "unit"} <= candidate.keys() else 0)
            if score > best_score:
                best_score = score
                header_row = row
                columns = candidate
        for field in HEADER_WORDS:
            if field in override:
                columns[field] = _column(override[field])

    if not {"date", "unit"} <= columns.keys():
        raise WorkbookError(
            f"{sheet.title}: date and unit NAV columns are not unambiguous"
        )
    data_start = header_row + 1
    headers_to_write: dict[int, str] = {}
    if mode == "append":
        _require_append_identity(sheet, columns, route)
        headers_to_write = {
            column: APPEND_HEADER_LABELS[field]
            for field, column in columns.items()
            if sheet.cell(header_row, column).value in {None, ""}
        }
    dated_rows: list[int] = []
    for row in range(data_start, sheet.max_row + 1):
        if parse_date(sheet.cell(row, columns["date"]).value):
            dated_rows.append(row)
    if not dated_rows:
        if mode == "summary":
            raise WorkbookError(f"{sheet.title}: no dated NAV rows found")
        if _has_values(sheet, data_start):
            raise WorkbookError(
                f"{sheet.title}: append mode found content below the header but no dated NAV rows"
            )
        return Layout(
            sheet.title,
            header_row,
            data_start,
            data_start,
            header_row,
            columns,
            mode,
            headers_to_write,
        )
    last_data_row = max(dated_rows)
    if mode == "append":
        if _has_values(sheet, last_data_row + 1):
            raise WorkbookError(
                f"{sheet.title}: append mode requires no footer content below the last dated row"
            )
        return Layout(
            sheet.title,
            header_row,
            data_start,
            last_data_row + 1,
            last_data_row,
            columns,
            mode,
            headers_to_write,
        )
    summary_row = 0
    for row in range(last_data_row + 1, sheet.max_row + 2):
        values = {
            _norm(sheet.cell(row, column).value)
            for column in range(1, min(sheet.max_column, 6) + 1)
        }
        if values & TOTAL_WORDS:
            summary_row = row
            break
    if not summary_row:
        raise WorkbookError(f"{sheet.title}: summary row was not found after the data")
    if summary_row != last_data_row + 1:
        raise WorkbookError(
            f"{sheet.title}: summary row must immediately follow the last dated row"
        )
    return Layout(
        sheet.title,
        header_row,
        data_start,
        summary_row,
        last_data_row,
        columns,
        mode,
        headers_to_write,
    )


def effective_cumulative(row: NavRow, route: dict[str, Any]) -> float:
    policy = str(route.get("cumulative_policy", "require"))
    if row.cumulative is not None:
        return row.cumulative
    if policy == "unit":
        return row.unit
    if policy == "offset":
        return row.unit + float(route["cumulative_offset"])
    raise WorkbookError(
        f"{route['sheet']}: cumulative NAV is required for {row.date.isoformat()}"
    )


def existing_rows(
    sheet: openpyxl.worksheet.worksheet.Worksheet, layout: Layout
) -> dict[dt.date, dict[str, Any]]:
    output: dict[dt.date, dict[str, Any]] = {}
    ordered_dates: list[dt.date] = []
    for row in range(layout.data_start, layout.summary_row):
        date = parse_date(sheet.cell(row, layout.columns["date"]).value)
        if not date:
            continue
        if date in output:
            raise WorkbookError(f"{sheet.title}: duplicate date {date.isoformat()}")
        ordered_dates.append(date)
        output[date] = {
            "row": row,
            "unit": parse_number(sheet.cell(row, layout.columns["unit"]).value),
            "cumulative": parse_number(
                sheet.cell(row, layout.columns.get("cumulative", 0)).value
            )
            if layout.columns.get("cumulative")
            else None,
            "code": normalize_code(sheet.cell(row, layout.columns.get("code", 0)).value)
            if layout.columns.get("code")
            else None,
        }
    if ordered_dates != sorted(ordered_dates):
        raise WorkbookError(f"{sheet.title}: dated NAV rows are not in ascending order")
    return output


def _summary_reserved_row(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    layout: Layout,
    route: dict[str, Any],
    existing: dict[dt.date, dict[str, Any]],
) -> tuple[dt.date, int] | None:
    """Identify a strictly bounded summary-mode cold-start placeholder.

    A real data row is never treated as a placeholder.  The only accepted shape is
    one dated row immediately above the summary row, with a matching configured
    identity and no content other than that identity and the date.
    """

    if layout.mode != "summary" or len(existing) != 1:
        return None
    date, values = next(iter(existing.items()))
    row = int(values["row"])
    if row != layout.data_start or row != layout.summary_row - 1:
        return None

    expected_code = normalize_code(route.get("code"))
    expected_name = str(route.get("product_name") or "").strip()
    identity_columns: set[int] = set()
    if expected_code and layout.columns.get("code"):
        code_column = layout.columns["code"]
        code_value = normalize_code(sheet.cell(row, code_column).value)
        if code_value:
            if code_value != expected_code:
                return None
            identity_columns.add(code_column)
    if expected_name and layout.columns.get("name"):
        name_column = layout.columns["name"]
        name_value = str(sheet.cell(row, name_column).value or "").strip()
        if name_value:
            if name_value != expected_name:
                return None
            identity_columns.add(name_column)
    if not identity_columns:
        return None

    allowed_columns = {layout.columns["date"], *identity_columns}
    for column in range(1, sheet.max_column + 1):
        value = sheet.cell(row, column).value
        if column not in allowed_columns and value is not None and value != "":
            return None
    return date, row


def _header_data_frequency(
    sheet: openpyxl.worksheet.worksheet.Worksheet, layout: Layout
) -> str | None:
    return_column = layout.columns.get("return")
    if not return_column:
        return None
    header = _norm(sheet.cell(layout.header_row, return_column).value)
    if "周" in header or "week" in header:
        return "weekly"
    if "日收益" in header or "日度" in header or "daily" in header:
        return "daily"
    return None


def _weekly_like_gaps(gaps: list[int]) -> bool:
    if not gaps:
        return False
    weekly = sum(1 for gap in gaps if gap >= 5 and min(gap % 7, (7 - gap % 7) % 7) <= 1)
    return weekly / len(gaps) >= 0.7


def _data_frequency(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    layout: Layout,
    route: dict[str, Any],
    existing: dict[dt.date, dict[str, Any]],
    reserved: tuple[dt.date, int] | None,
) -> tuple[str, str, int | None]:
    configured = str(route.get("data_frequency", "auto"))
    history_dates = sorted(
        date for date in existing if reserved is None or date != reserved[0]
    )
    inferred: str | None = None
    source = "email-source"
    preferred_weekday: int | None = None
    if len(history_dates) >= 2:
        gaps = [
            (current - previous).days
            for previous, current in zip(history_dates, history_dates[1:])
        ]
        median_gap = float(statistics.median(gaps))
        if median_gap <= 4:
            inferred = "daily"
        elif median_gap <= 10 or _weekly_like_gaps(gaps):
            inferred = "weekly"
        else:
            raise WorkbookError(
                f"{sheet.title}: existing dates do not prove a daily or weekly data frequency"
            )
        source = "workbook-history"
        if inferred == "weekly":
            counts = Counter(date.weekday() for date in history_dates)
            last_seen = {
                weekday: max(
                    index
                    for index, date in enumerate(history_dates)
                    if date.weekday() == weekday
                )
                for weekday in counts
            }
            preferred_weekday = max(
                counts, key=lambda weekday: (counts[weekday], last_seen[weekday])
            )
    else:
        inferred = _header_data_frequency(sheet, layout)
        if inferred:
            source = "workbook-header"

    if configured != "auto":
        if inferred and configured != inferred:
            raise WorkbookError(
                f"{sheet.title}: configured data_frequency {configured} conflicts with "
                f"the {inferred} workbook template"
            )
        return configured, "route-config", preferred_weekday
    return inferred or "daily", source, preferred_weekday


def _select_data_rows(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    layout: Layout,
    route: dict[str, Any],
    existing: dict[dt.date, dict[str, Any]],
    reserved: tuple[dt.date, int] | None,
    candidates: list[NavRow],
) -> tuple[list[NavRow], str, str]:
    frequency, source, preferred_weekday = _data_frequency(
        sheet, layout, route, existing, reserved
    )
    if frequency == "daily":
        return candidates, frequency, source

    groups: dict[tuple[int, int], list[NavRow]] = {}
    for candidate in candidates:
        iso = candidate.date.isocalendar()
        groups.setdefault((iso.year, iso.week), []).append(candidate)
    current_iso = dt.date.today().isocalendar()
    current_week = (current_iso.year, current_iso.week)
    existing_dates = set(existing)
    selected: list[NavRow] = []
    for week in sorted(groups):
        rows = sorted(groups[week], key=lambda item: item.date)
        exact = [row for row in rows if row.date in existing_dates]
        if exact:
            selected.extend(exact)
            continue
        if week >= current_week:
            continue
        if preferred_weekday is None:
            selected.append(rows[-1])
            continue
        on_or_before = [row for row in rows if row.date.weekday() <= preferred_weekday]
        selected.append(on_or_before[-1] if on_or_before else rows[0])
    return selected, frequency, source


def validate_history(
    config: dict[str, Any], route_rows: dict[str, list[NavRow]]
) -> dict[str, Any]:
    path = Path(config["workbook_path"])
    workbook = openpyxl.load_workbook(
        path, data_only=True, read_only=False, keep_vba=path.suffix.lower() == ".xlsm"
    )
    minimum = int((config.get("validation") or {}).get("minimum_history_dates", 2))
    tolerance = float((config.get("validation") or {}).get("tolerance", 1e-6))
    max_future_days = int((config.get("validation") or {}).get("max_future_days", 0))
    max_period_change = float(
        (config.get("validation") or {}).get("max_period_change", 0.5)
    )
    reports: list[dict[str, Any]] = []
    errors: list[str] = []
    warnings: list[str] = []
    try:
        for route in active_routes(config):
            sheet_name = str(route["sheet"])
            if sheet_name not in workbook.sheetnames:
                errors.append(f"Missing managed sheet: {sheet_name}")
                continue
            layout = discover_layout(
                workbook[sheet_name],
                (config.get("column_overrides") or {}).get(sheet_name),
                route,
            )
            existing = existing_rows(workbook[sheet_name], layout)
            matches = 0
            conflicts = 0
            start = parse_date(route.get("series_start")) or dt.date.min
            candidates = sorted(
                (
                    candidate
                    for candidate in route_rows.get(sheet_name, [])
                    if candidate.date >= start
                ),
                key=lambda item: item.date,
            )
            reserved = _summary_reserved_row(
                workbook[sheet_name], layout, route, existing
            )
            selected_candidates, data_frequency, frequency_source = _select_data_rows(
                workbook[sheet_name],
                layout,
                route,
                existing,
                reserved,
                candidates,
            )
            selected_dates = {candidate.date for candidate in selected_candidates}
            reserved_date = reserved[0] if reserved else None
            candidate_dates = selected_dates
            reserved_source_date = (
                min(candidate_dates) if reserved and candidate_dates else None
            )
            known_units = {
                date: values["unit"]
                for date, values in existing.items()
                if values["unit"] is not None
            }
            for candidate in candidates:
                expected_code = normalize_code(route.get("code"))
                if expected_code and candidate.code != expected_code:
                    errors.append(
                        f"{sheet_name}: candidate code mismatch on {candidate.date.isoformat()}"
                    )
                if candidate.date > dt.date.today() + dt.timedelta(
                    days=max_future_days
                ):
                    errors.append(
                        f"{sheet_name}: NAV date {candidate.date.isoformat()} exceeds max_future_days"
                    )
                if (
                    reserved_source_date is not None
                    and candidate.date == reserved_source_date
                ):
                    if layout.columns.get("cumulative"):
                        effective_cumulative(candidate, route)
                    known_units[candidate.date] = candidate.unit
                    continue
                if candidate.date not in existing:
                    prior = [
                        (date, unit)
                        for date, unit in known_units.items()
                        if date < candidate.date
                    ]
                    if prior:
                        previous_unit = max(prior)[1]
                        if (
                            abs(candidate.unit / float(previous_unit) - 1)
                            > max_period_change
                        ):
                            errors.append(
                                f"{sheet_name}: NAV change exceeds max_period_change on {candidate.date.isoformat()}"
                            )
                    known_units[candidate.date] = candidate.unit
                if (
                    candidate.date < start
                    or candidate.date not in existing
                    or candidate.date not in selected_dates
                ):
                    continue
                observed = existing[candidate.date]
                unit_ok = (
                    observed["unit"] is not None
                    and abs(float(observed["unit"]) - candidate.unit) <= tolerance
                )
                cumulative_ok = True
                if layout.columns.get("cumulative"):
                    expected_cumulative = effective_cumulative(candidate, route)
                    cumulative_ok = (
                        observed["cumulative"] is not None
                        and abs(float(observed["cumulative"]) - expected_cumulative)
                        <= tolerance
                    )
                code_ok = (
                    not route.get("code")
                    or not layout.columns.get("code")
                    or observed["code"] == normalize_code(route.get("code"))
                )
                if unit_ok and cumulative_ok and code_ok:
                    matches += 1
                else:
                    conflicts += 1
                    errors.append(
                        f"{sheet_name}: historical value conflict on {candidate.date.isoformat()}"
                    )
            summary_cold_start = reserved_date is not None
            if summary_cold_start:
                if len(candidate_dates) < minimum:
                    errors.append(
                        f"{sheet_name}: summary-mode cold start found only {len(candidate_dates)} email dates; {minimum} required"
                    )
                if candidate_dates and len(candidate_dates) >= minimum:
                    warnings.append(
                        f"{sheet_name}: detected a reserved summary row cold start; replace the empty placeholder with the earliest email NAV, keep the summary row, and review the first preview"
                    )
            elif matches < minimum:
                message = f"{sheet_name}: only {matches} verified historical dates; {minimum} required"
                if layout.mode == "append":
                    warnings.append(
                        f"{message}; append-mode cold start requires review"
                    )
                else:
                    errors.append(message)
            reports.append(
                {
                    "sheet": sheet_name,
                    "sheet_mode": layout.mode,
                    "cold_start": summary_cold_start
                    or (layout.mode == "append" and matches < minimum),
                    "cold_start_kind": (
                        "summary-reserved-row"
                        if summary_cold_start
                        else "append"
                        if layout.mode == "append" and matches < minimum
                        else None
                    ),
                    "matched_history_dates": matches,
                    "conflicts": conflicts,
                    "data_frequency": data_frequency,
                    "data_frequency_source": frequency_source,
                }
            )
    finally:
        workbook.close()
    report = {
        "passed": not errors,
        "routes": reports,
        "warnings": warnings,
        "errors": errors,
    }
    write_json_atomic(STATE_ROOT / "validation-report.json", report)
    return report


def _copy_row(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    source: int,
    target: int,
    max_column: int,
) -> None:
    sheet.row_dimensions[target].height = sheet.row_dimensions[source].height
    for column in range(1, max_column + 1):
        original = sheet.cell(source, column)
        cell = sheet.cell(target, column)
        if original.has_style:
            cell._style = copy(original._style)
        cell.number_format = original.number_format
        cell.alignment = copy(original.alignment)
        cell.protection = copy(original.protection)
        if isinstance(original.value, str) and original.value.startswith("="):
            try:
                cell.value = Translator(
                    original.value, origin=original.coordinate
                ).translate_formula(cell.coordinate)
            except Exception as exc:
                raise WorkbookError(
                    f"{sheet.title}: could not translate formula at {original.coordinate}"
                ) from exc
        else:
            cell.value = original.value


def _date_rows(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    layout: Layout,
    summary_row: int,
    start: dt.date,
) -> list[tuple[dt.date, int]]:
    rows: list[tuple[dt.date, int]] = []
    for row in range(layout.data_start, summary_row):
        date = parse_date(sheet.cell(row, layout.columns["date"]).value)
        if date and date >= start:
            rows.append((date, row))
    return sorted(rows)


def _set_return_formulas(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    layout: Layout,
    route: dict[str, Any],
    summary_row: int,
    new_rows: set[int],
    changed: set[tuple[int, int]],
    write_summary: bool = True,
) -> tuple[list[int], set[int]]:
    return_column = layout.columns.get("return")
    if not return_column:
        return [], set()
    basis_name = (
        "cumulative"
        if route.get("return_basis", "cumulative") == "cumulative"
        else "unit"
    )
    basis_column = layout.columns.get(basis_name)
    if not basis_column:
        raise WorkbookError(
            f"{sheet.title}: return basis column {basis_name} is missing"
        )
    start = parse_date(route.get("series_start")) or dt.date.min
    rows = _date_rows(sheet, layout, summary_row, start)
    if not rows:
        return [], set()
    letter = get_column_letter(basis_column)
    frequency = str(route.get("return_frequency", "weekly"))
    period_rows: list[int] = []
    affected_rows: set[int] = set()
    if frequency == "daily":
        for index, (_, row) in enumerate(rows):
            period_rows.append(row)
            if row in new_rows:
                value = (
                    None
                    if index == 0
                    else f"={letter}{row}/{letter}{rows[index - 1][1]}-1"
                )
                sheet.cell(row, return_column).value = value
                changed.add((row, return_column))
                affected_rows.add(row)
    else:
        groups: dict[tuple[int, int], list[tuple[dt.date, int]]] = {}
        for date, row in rows:
            iso = date.isocalendar()
            groups.setdefault((iso.year, iso.week), []).append((date, row))
        current_iso = dt.date.today().isocalendar()
        completed = [
            key for key in sorted(groups) if key < (current_iso.year, current_iso.week)
        ]
        for key in completed:
            period_rows.append(max(groups[key])[1])
        new_week_keys = {
            (date.isocalendar().year, date.isocalendar().week)
            for date, row in rows
            if row in new_rows
        }
        for key in new_week_keys:
            week_rows = [row for _, row in groups[key]]
            affected_rows.update(week_rows)
            target = max(groups[key])[1] if key in completed else None
            previous = (
                period_rows[period_rows.index(target) - 1]
                if target in period_rows and period_rows.index(target) > 0
                else None
            )
            for row in week_rows:
                value = (
                    f"={letter}{row}/{letter}{previous}-1"
                    if row == target and previous is not None
                    else None
                )
                if sheet.cell(row, return_column).value != value:
                    sheet.cell(row, return_column).value = value
                    changed.add((row, return_column))
    if write_summary:
        summary_value = None
        if frequency == "daily" and period_rows:
            summary_value = f"={letter}{period_rows[-1]}/{letter}{period_rows[0]}-1"
        elif frequency == "weekly" and period_rows:
            summary_value = f"={letter}{period_rows[-1]}/{letter}{period_rows[0]}-1"
        sheet.cell(summary_row, return_column).value = summary_value
        changed.add((summary_row, return_column))
    return period_rows, affected_rows


def _benchmark_source(
    workbook: openpyxl.Workbook, benchmark: dict[str, Any]
) -> tuple[dict[dt.date, tuple[int, Any]], int]:
    source_name = str(benchmark["source_sheet"])
    if source_name not in workbook.sheetnames:
        raise WorkbookError(f"Missing benchmark source sheet: {source_name}")
    sheet = workbook[source_name]
    date_column = _column(benchmark["source_date"])
    value_column = _column(benchmark["source_value"])
    values: dict[dt.date, tuple[int, Any]] = {}
    for row in range(1, sheet.max_row + 1):
        date = parse_date(sheet.cell(row, date_column).value)
        value = sheet.cell(row, value_column).value
        if date and value is not None:
            if date in values:
                raise WorkbookError(
                    f"{source_name}: duplicate benchmark date {date.isoformat()}"
                )
            values[date] = (row, value)
    return values, value_column


def _product_formula(column: int, rows: list[int]) -> str:
    expressions = [f"1+{get_column_letter(column)}{row}" for row in rows]
    while len(expressions) > 200:
        expressions = [
            f"PRODUCT({','.join(expressions[index : index + 200])})"
            for index in range(0, len(expressions), 200)
        ]
    return f"=PRODUCT({','.join(expressions)})-1"


def _set_benchmark_formulas(
    workbook: openpyxl.Workbook,
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    layout: Layout,
    route: dict[str, Any],
    summary_row: int,
    period_rows: list[int],
    affected_rows: set[int],
    changed: set[tuple[int, int]],
    write_summary: bool = True,
) -> None:
    benchmark = route.get("benchmark")
    if not benchmark:
        return
    source_values, source_column = _benchmark_source(workbook, benchmark)
    source_sheet = quote_sheetname(str(benchmark["source_sheet"]))
    source_letter = get_column_letter(source_column)
    target_return = layout.columns.get("benchmark_return")
    excess = layout.columns.get("excess")
    product_return = layout.columns.get("return")
    if not target_return or not excess or not product_return:
        raise WorkbookError(
            f"{sheet.title}: benchmark return, product return, and excess columns are required"
        )
    start = parse_date(route.get("series_start")) or dt.date.min
    rows = _date_rows(sheet, layout, summary_row, start)
    by_row = {row: date for date, row in rows}
    target_rows = [
        row for row in period_rows if sheet.cell(row, product_return).value is not None
    ]
    affected_targets = [row for row in target_rows if row in affected_rows]
    source_type = str(benchmark.get("source_type", "level"))
    if source_type not in {"aligned_return", "level"}:
        raise WorkbookError("benchmark.source_type must be aligned_return or level")
    level_column = layout.columns.get("benchmark_level")
    for row in affected_rows:
        if row in affected_targets:
            continue
        for column in (
            target_return,
            excess,
            level_column if source_type == "level" else None,
        ):
            if column and sheet.cell(row, column).value is not None:
                sheet.cell(row, column).value = None
                changed.add((row, column))
    if source_type == "aligned_return":
        for row in affected_targets:
            date = by_row[row]
            source = source_values.get(date)
            if not source:
                raise WorkbookError(
                    f"{sheet.title}: benchmark is missing for {date.isoformat()}"
                )
            source_row = source[0]
            sheet.cell(
                row, target_return
            ).value = f"={source_sheet}!{source_letter}{source_row}"
            changed.add((row, target_return))
            sheet.cell(
                row, excess
            ).value = f"={get_column_letter(product_return)}{row}-{get_column_letter(target_return)}{row}"
            changed.add((row, excess))
    else:
        if not level_column:
            raise WorkbookError(f"{sheet.title}: benchmark level column is required")
        for row in affected_targets:
            index = period_rows.index(row)
            if index == 0:
                raise WorkbookError(
                    f"{sheet.title}: benchmark level has no prior period anchor"
                )
            previous_level_row = period_rows[index - 1]
            for anchor_row in (previous_level_row, row):
                date = by_row[anchor_row]
                source = source_values.get(date)
                if not source:
                    raise WorkbookError(
                        f"{sheet.title}: benchmark is missing for {date.isoformat()}"
                    )
                if (
                    anchor_row == row
                    or sheet.cell(anchor_row, level_column).value is None
                ):
                    sheet.cell(
                        anchor_row, level_column
                    ).value = f"={source_sheet}!{source_letter}{source[0]}"
                    changed.add((anchor_row, level_column))
            level_letter = get_column_letter(level_column)
            sheet.cell(
                row, target_return
            ).value = f"={level_letter}{row}/{level_letter}{previous_level_row}-1"
            changed.add((row, target_return))
            sheet.cell(
                row, excess
            ).value = f"={get_column_letter(product_return)}{row}-{get_column_letter(target_return)}{row}"
            changed.add((row, excess))

    if not write_summary:
        return
    if source_type == "level" and period_rows:
        first_date, last_date = by_row[period_rows[0]], by_row[period_rows[-1]]
        first_source, last_source = (
            source_values.get(first_date),
            source_values.get(last_date),
        )
        if not first_source or not last_source:
            raise WorkbookError(f"{sheet.title}: benchmark summary anchors are missing")
        summary_value = f"={source_sheet}!{source_letter}{last_source[0]}/{source_sheet}!{source_letter}{first_source[0]}-1"
    elif period_rows:
        summary_value = _product_formula(target_return, period_rows)
    else:
        summary_value = None
    sheet.cell(summary_row, target_return).value = summary_value
    sheet.cell(summary_row, excess).value = (
        f"={get_column_letter(product_return)}{summary_row}-{get_column_letter(target_return)}{summary_row}"
        if summary_value is not None
        else None
    )
    changed.update({(summary_row, target_return), (summary_row, excess)})


def _ensure_summary_formula_safety(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    layout: Layout,
    route: dict[str, Any],
) -> None:
    managed = {column for column in (layout.columns.get("return"),) if column}
    if route.get("benchmark"):
        managed.update(
            column
            for column in (
                layout.columns.get("benchmark_return"),
                layout.columns.get("excess"),
            )
            if column
        )
    for column in range(1, sheet.max_column + 1):
        value = sheet.cell(layout.summary_row, column).value
        if isinstance(value, str) and value.startswith("=") and column not in managed:
            coordinate = sheet.cell(layout.summary_row, column).coordinate
            raise WorkbookError(
                f"{sheet.title}: summary formula at {coordinate} is not managed; "
                "automatic insertion cannot prove that its range will expand safely"
            )


def _write_nav_values(
    sheet: openpyxl.worksheet.worksheet.Worksheet,
    row: int,
    layout: Layout,
    route: dict[str, Any],
    nav: NavRow,
    changed: set[tuple[int, int]] | None = None,
) -> None:
    values: list[tuple[int, Any]] = [
        (layout.columns["date"], nav.date),
        (layout.columns["unit"], nav.unit),
    ]
    if layout.columns.get("cumulative"):
        values.append((layout.columns["cumulative"], effective_cumulative(nav, route)))
    if layout.columns.get("code") and route.get("code"):
        values.append((layout.columns["code"], normalize_code(route["code"])))
    if layout.columns.get("name") and route.get("product_name"):
        values.append((layout.columns["name"], str(route["product_name"]).strip()))
    for column, value in values:
        cell = sheet.cell(row, column)
        same = (
            parse_date(cell.value) == value
            if column == layout.columns["date"]
            else cell.value == value
        )
        if not same:
            cell.value = value
            if changed is not None:
                changed.add((row, column))


def build_preview(
    config: dict[str, Any],
    route_rows: dict[str, list[NavRow]],
    warnings: list[str] | None = None,
) -> dict[str, Any]:
    (STATE_ROOT / "plan.json").unlink(missing_ok=True)
    master = Path(config["workbook_path"])
    preview_dir = ROOT / "previews"
    preview_dir.mkdir(exist_ok=True)
    keep = int((config.get("retention") or {}).get("preview_count", 10))
    old_previews = sorted(
        (
            path
            for path in preview_dir.glob("preview-*")
            if path.suffix.lower() in {".xlsx", ".xlsm"}
        ),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    for old in old_previews[max(keep - 1, 0) :]:
        old.unlink()
    preview = (
        preview_dir
        / f"preview-{dt.datetime.now():%Y%m%d-%H%M%S}-{uuid.uuid4().hex[:8]}{master.suffix.lower()}"
    )
    shutil.copy2(master, preview)
    keep_vba = master.suffix.lower() == ".xlsm"
    workbook = openpyxl.load_workbook(preview, data_only=False, keep_vba=keep_vba)
    plan_sheets: list[dict[str, Any]] = []
    try:
        for route in active_routes(config):
            sheet_name = str(route["sheet"])
            sheet = workbook[sheet_name]
            layout = discover_layout(
                sheet, (config.get("column_overrides") or {}).get(sheet_name), route
            )
            current = existing_rows(sheet, layout)
            start = parse_date(route.get("series_start")) or dt.date.min
            candidates = sorted(
                (row for row in route_rows.get(sheet_name, []) if row.date >= start),
                key=lambda row: row.date,
            )
            reserved = _summary_reserved_row(sheet, layout, route, current)
            candidates, data_frequency, frequency_source = _select_data_rows(
                sheet, layout, route, current, reserved, candidates
            )
            reserved_date = reserved[0] if reserved else None
            reserved_row = reserved[1] if reserved else None
            reserved_nav: NavRow | None = None
            if reserved_date is not None:
                distinct_dates = {row.date for row in candidates}
                minimum = int(
                    (config.get("validation") or {}).get("minimum_history_dates", 2)
                )
                if len(distinct_dates) < minimum:
                    raise WorkbookError(
                        f"{sheet_name}: summary-mode cold start needs at least {minimum} email NAV dates"
                    )
                reserved_nav = candidates[0]
            additions = (
                candidates[1:]
                if reserved_nav is not None
                else [row for row in candidates if row.date not in current]
            )
            if not additions and reserved_nav is None:
                continue
            if layout.mode == "summary":
                _ensure_summary_formula_safety(sheet, layout, route)
            latest_existing = max(current) if current else None
            gaps = (
                [row.date for row in additions if row.date <= latest_existing]
                if latest_existing
                else []
            )
            if gaps:
                dates = ", ".join(date.isoformat() for date in gaps)
                raise WorkbookError(
                    f"{sheet_name}: internal historical gaps require supervised repair: {dates}"
                )
            old_summary = layout.summary_row
            count = len(additions)
            max_column = max([sheet.max_column, *layout.columns.values()])
            changed: set[tuple[int, int]] = set()
            filled_existing_rows: list[int] = []
            for column, label in (layout.headers_to_write or {}).items():
                cell = sheet.cell(layout.header_row, column)
                cell.value = label
                font = copy(cell.font)
                font.bold = True
                cell.font = font
                changed.add((layout.header_row, column))
            if reserved_nav is not None and reserved_row is not None:
                _write_nav_values(
                    sheet,
                    reserved_row,
                    layout,
                    route,
                    reserved_nav,
                    changed,
                )
                filled_existing_rows.append(reserved_row)
            sheet.insert_rows(old_summary, count)
            for offset, nav in enumerate(additions):
                target = old_summary + offset
                if current or layout.mode == "summary":
                    template = old_summary - 1 if offset == 0 else target - 1
                    _copy_row(sheet, template, target, max_column)
                _write_nav_values(sheet, target, layout, route, nav)
                if layout.mode == "append" and not current:
                    sheet.cell(
                        target, layout.columns["date"]
                    ).number_format = "yyyy-mm-dd"
                    sheet.cell(
                        target, layout.columns["unit"]
                    ).number_format = "0.000000"
                if layout.columns.get("cumulative"):
                    if layout.mode == "append" and not current:
                        sheet.cell(
                            target, layout.columns["cumulative"]
                        ).number_format = "0.000000"
            new_summary = old_summary + count
            new_rows = set(range(old_summary, new_summary))
            managed_rows = new_rows | set(filled_existing_rows)
            period_rows, affected_rows = _set_return_formulas(
                sheet,
                layout,
                route,
                new_summary,
                managed_rows,
                changed,
                write_summary=layout.mode == "summary",
            )
            _set_benchmark_formulas(
                workbook,
                sheet,
                layout,
                route,
                new_summary,
                period_rows,
                affected_rows,
                changed,
                write_summary=layout.mode == "summary",
            )
            plan_sheets.append(
                {
                    "sheet": sheet_name,
                    "sheet_mode": layout.mode,
                    "data_frequency": data_frequency,
                    "data_frequency_source": frequency_source,
                    "header_row": layout.header_row,
                    "insert_before": old_summary,
                    "insert_count": count,
                    "populated_count": count + len(filled_existing_rows),
                    "new_rows": list(range(old_summary, new_summary)),
                    "filled_existing_rows": filled_existing_rows,
                    "summary_row": new_summary,
                    "copy_template_rows": bool(current) or layout.mode == "summary",
                    "format_rows": sorted(
                        {
                            *([layout.header_row] if layout.headers_to_write else []),
                            *(
                                range(old_summary, new_summary)
                                if layout.mode == "append" and not current
                                else []
                            ),
                        }
                    ),
                    "changed_cells": sorted(
                        [{"row": row, "column": column} for row, column in changed],
                        key=lambda item: (item["row"], item["column"]),
                    ),
                    "new_dates": [
                        row.date.isoformat()
                        for row in (
                            ([reserved_nav] if reserved_nav is not None else [])
                            + additions
                        )
                    ],
                    "return_columns": [
                        column
                        for column in (
                            layout.columns.get("return"),
                            layout.columns.get("benchmark_return"),
                            layout.columns.get("excess"),
                        )
                        if column
                    ],
                }
            )
        workbook.save(preview)
    except Exception:
        workbook.close()
        preview.unlink(missing_ok=True)
        (STATE_ROOT / "plan.json").unlink(missing_ok=True)
        raise
    else:
        workbook.close()
    plan = {
        "schema_version": 1,
        "plan_id": str(uuid.uuid4()),
        "created": dt.datetime.now().isoformat(timespec="seconds"),
        "config_sha256": payload_sha256(config),
        "master_path": str(master.resolve()),
        "master_sha256": file_sha256(master),
        "preview_path": str(preview.resolve()) if plan_sheets else None,
        "preview_sha256": file_sha256(preview) if plan_sheets else None,
        "warnings": list(warnings or []),
        "sheets": plan_sheets,
    }
    if not plan_sheets:
        preview.unlink(missing_ok=True)
        (STATE_ROOT / "plan.json").unlink(missing_ok=True)
        return plan
    try:
        validate_preview(config, plan)
    except Exception:
        preview.unlink(missing_ok=True)
        (STATE_ROOT / "plan.json").unlink(missing_ok=True)
        raise
    write_json_atomic(STATE_ROOT / "plan.json", plan)
    return plan


def validate_preview(config: dict[str, Any], plan: dict[str, Any]) -> None:
    master = Path(plan["master_path"])
    preview = Path(plan["preview_path"])
    keep_vba = master.suffix.lower() == ".xlsm"
    original = openpyxl.load_workbook(
        master, data_only=False, read_only=False, keep_vba=keep_vba
    )
    candidate = openpyxl.load_workbook(
        preview, data_only=False, read_only=False, keep_vba=keep_vba
    )
    try:
        if original.sheetnames != candidate.sheetnames:
            raise WorkbookError("Preview changed the workbook sheet topology")
        plans = {item["sheet"]: item for item in plan["sheets"]}
        routes = {str(route["sheet"]): route for route in active_routes(config)}
        for name in original.sheetnames:
            left, right = original[name], candidate[name]
            sheet_plan = plans.get(name)
            if sheet_plan is None:
                if left.max_row != right.max_row or left.max_column != right.max_column:
                    raise WorkbookError(f"{name}: preview changed an unmanaged sheet")
                for row in range(1, left.max_row + 1):
                    for column in range(1, left.max_column + 1):
                        if (
                            left.cell(row, column).value
                            != right.cell(row, column).value
                        ):
                            raise WorkbookError(
                                f"{name}: preview changed an unmanaged cell"
                            )
                continue
            insert_before = int(sheet_plan["insert_before"])
            new_summary = int(sheet_plan["summary_row"])
            allowed = {
                (int(item["row"]), int(item["column"]))
                for item in sheet_plan.get("changed_cells") or []
            }
            max_column = max(left.max_column, right.max_column)
            for row in range(1, insert_before):
                for column in range(1, max_column + 1):
                    if (row, column) not in allowed and left.cell(
                        row, column
                    ).value != right.cell(row, column).value:
                        raise WorkbookError(
                            f"{name}: preview changed an unapproved historical cell"
                        )
            for column in range(1, max_column + 1):
                if (new_summary, column) not in allowed and left.cell(
                    insert_before, column
                ).value != right.cell(new_summary, column).value:
                    raise WorkbookError(
                        f"{name}: preview changed an unapproved summary cell"
                    )
        for sheet_plan in plan["sheets"]:
            sheet = candidate[sheet_plan["sheet"]]
            layout = discover_layout(
                sheet,
                (config.get("column_overrides") or {}).get(sheet.title),
                routes[sheet.title],
            )
            dates = []
            for row in range(layout.data_start, layout.summary_row):
                date = parse_date(sheet.cell(row, layout.columns["date"]).value)
                if date:
                    dates.append(date)
            if len(dates) != len(set(dates)):
                raise WorkbookError(f"{sheet.title}: preview contains duplicate dates")
            if not set(sheet_plan["new_dates"]) <= {date.isoformat() for date in dates}:
                raise WorkbookError(f"{sheet.title}: preview is missing proposed dates")
            for row in range(1, sheet.max_row + 1):
                for cell in sheet[row]:
                    if (
                        isinstance(cell.value, str)
                        and cell.value.startswith("=")
                        and "#REF!" in cell.value.upper()
                    ):
                        raise WorkbookError(
                            f"{sheet.title}: preview contains a broken formula at {cell.coordinate}"
                        )
    finally:
        original.close()
        candidate.close()


def file_sha256(path: Path) -> str:
    import hashlib

    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def payload_sha256(payload: Any) -> str:
    import hashlib

    encoded = json.dumps(
        payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()
