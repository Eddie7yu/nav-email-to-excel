from __future__ import annotations

import datetime as dt
import json
import re
import uuid
from pathlib import Path
from typing import Any


STATE_ROOT = Path(__file__).resolve().parent
ROOT = STATE_ROOT.parent if STATE_ROOT.name == "app" else STATE_ROOT
CONFIG_PATH = STATE_ROOT / "config.json"
DAY_NAMES = {"MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"}
TOP_FIELDS = {
    "schema_version",
    "runtime_id",
    "workbook_path",
    "workbook_mode",
    "imap",
    "routes",
    "column_overrides",
    "style",
    "schedule",
    "validation",
    "retention",
}
IMAP_FIELDS = {
    "host",
    "port",
    "user",
    "mailbox",
    "lookback_days",
    "max_messages",
    "max_header_messages",
    "max_message_bytes",
    "max_total_bytes",
}
ROUTE_FIELDS = {
    "sender",
    "subject_contains",
    "sheet",
    "sheet_mode",
    "code",
    "product_name",
    "parser",
    "paused",
    "pause_reason",
    "allow_sender_only",
    "cumulative_policy",
    "cumulative_offset",
    "return_basis",
    "return_frequency",
    "data_frequency",
    "series_start",
    "max_staleness_days",
    "benchmark",
    "benchmark_review_only",
}
BENCHMARK_FIELDS = {
    "source_sheet",
    "source_type",
    "source_date",
    "source_value",
    "display_name",
}
COLUMN_FIELDS = {
    "header_row",
    "date",
    "code",
    "name",
    "unit",
    "cumulative",
    "return",
    "daily_return",
    "weekly_return",
    "benchmark_level",
    "benchmark_return",
    "excess",
}


class ConfigError(ValueError):
    pass


def load_config(path: Path = CONFIG_PATH) -> dict[str, Any]:
    try:
        config = json.loads(path.read_text(encoding="utf-8"))
    except FileNotFoundError as exc:
        raise ConfigError(f"Missing runtime configuration: {path}") from exc
    except json.JSONDecodeError as exc:
        raise ConfigError(f"Invalid JSON in {path.name}: {exc}") from exc
    validate_config(config)
    return config


def validate_config(config: dict[str, Any]) -> None:
    if not isinstance(config, dict):
        raise ConfigError("config root must be an object")
    errors: list[str] = []
    for field in sorted(set(config) - TOP_FIELDS):
        errors.append(f"unknown top-level field: {field}")
    if config.get("schema_version") != 1:
        errors.append("schema_version must be 1")
    try:
        uuid.UUID(str(config.get("runtime_id", "")))
    except (ValueError, TypeError, AttributeError):
        errors.append("runtime_id must be a UUID")

    workbook = Path(str(config.get("workbook_path", ""))).expanduser()
    if not workbook.is_absolute():
        errors.append("workbook_path must be absolute")
    elif workbook.suffix.lower() not in {".xlsx", ".xlsm"}:
        errors.append("workbook_path must end in .xlsx or .xlsm")
    workbook_mode = str(config.get("workbook_mode", "existing"))
    if workbook_mode not in {"existing", "bundled-template"}:
        errors.append("workbook_mode must be existing or bundled-template")
    if workbook_mode == "bundled-template" and workbook.suffix.lower() != ".xlsx":
        errors.append("bundled-template workbook_path must end in .xlsx")

    imap = config.get("imap") or {}
    if not isinstance(imap, dict):
        errors.append("imap must be an object")
        imap = {}
    for field in sorted(set(imap) - IMAP_FIELDS):
        errors.append(f"unknown imap field: {field}")
    for field in ("host", "user", "mailbox"):
        if not str(imap.get(field, "")).strip():
            errors.append(f"imap.{field} is required")
    if re.search(r"\s", str(imap.get("host", "")).strip()):
        errors.append("imap.host cannot contain whitespace")
    try:
        port = int(imap.get("port", 993))
        if not 1 <= port <= 65535:
            raise ValueError
    except (TypeError, ValueError):
        errors.append("imap.port must be between 1 and 65535")
    for field, default, minimum, maximum in (
        ("lookback_days", 180, 1, 3650),
        ("max_messages", 2000, 1, 10000),
        ("max_header_messages", 20000, 1, 100000),
        ("max_message_bytes", 25 * 1024 * 1024, 1024, 100 * 1024 * 1024),
        ("max_total_bytes", 100 * 1024 * 1024, 1024, 1024 * 1024 * 1024),
    ):
        try:
            value = int(imap.get(field, default))
            if not minimum <= value <= maximum:
                raise ValueError
        except (TypeError, ValueError):
            errors.append(f"imap.{field} must be between {minimum} and {maximum}")
    try:
        if int(imap.get("max_header_messages", 20000)) < int(
            imap.get("max_messages", 2000)
        ):
            errors.append("imap.max_header_messages must be at least imap.max_messages")
    except (TypeError, ValueError):
        pass

    routes = config.get("routes")
    if not isinstance(routes, list):
        errors.append("routes must be a list")
        routes = []
    route_keys: set[tuple[str, str, str]] = set()
    senders: dict[str, list[dict[str, Any]]] = {}
    managed_sheets: set[str] = set()
    for index, route in enumerate(routes, 1):
        prefix = f"route #{index}"
        if not isinstance(route, dict):
            errors.append(f"{prefix} must be an object")
            continue
        for field in sorted(set(route) - ROUTE_FIELDS):
            errors.append(f"{prefix} has unknown field: {field}")
        sender = str(route.get("sender", "")).strip().lower()
        sheet = str(route.get("sheet", "")).strip()
        code = normalize_code(route.get("code"))
        if not re.fullmatch(
            r"[A-Za-z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Za-z0-9](?:[A-Za-z0-9.-]*[A-Za-z0-9])?",
            sender,
        ):
            errors.append(f"{prefix}.sender must be an email address")
        if route.get("code") is not None and not isinstance(route.get("code"), str):
            errors.append(
                f"{prefix}.code must be a quoted string so leading zeros are preserved"
            )
        if not sheet:
            errors.append(f"{prefix}.sheet is required")
        elif sheet in managed_sheets:
            errors.append(f"{prefix}.sheet is already managed by another route")
        if sheet:
            managed_sheets.add(sheet)
        sheet_mode = str(route.get("sheet_mode", "summary"))
        if sheet_mode not in {"summary", "append", "template"}:
            errors.append(f"{prefix}.sheet_mode must be summary, append, or template")
        if workbook_mode == "bundled-template" and sheet_mode != "template":
            errors.append(
                f"{prefix}.sheet_mode must be template for a bundled-template workbook"
            )
        if workbook_mode == "existing" and sheet_mode == "template":
            errors.append(
                f"{prefix}.sheet_mode template is reserved for bundled workbooks"
            )
        product_name = route.get("product_name")
        if product_name is not None and (
            not isinstance(product_name, str)
            or not product_name.strip()
            or len(product_name.strip()) > 200
            or "\n" in product_name
            or "\r" in product_name
            or product_name.lstrip().startswith("=")
        ):
            errors.append(
                f"{prefix}.product_name must be a non-empty, non-formula single-line string of at most 200 characters"
            )
        if (
            sheet_mode in {"append", "template"}
            and not code
            and not (isinstance(product_name, str) and product_name.strip())
        ):
            errors.append(
                f"{prefix} in {sheet_mode} mode requires code or product_name"
            )
        subject_contains = route.get("subject_contains")
        if subject_contains is not None and (
            not isinstance(subject_contains, str)
            or not subject_contains.strip()
            or len(subject_contains) > 200
        ):
            errors.append(
                f"{prefix}.subject_contains must be a non-empty string of at most 200 characters"
            )
        parser_name = str(route.get("parser", "auto"))
        if parser_name != "auto" and not re.fullmatch(
            r"local:[a-z][a-z0-9_-]{0,63}", parser_name
        ):
            errors.append(
                f"{prefix}.parser must be auto or local:<trusted-parser-name>"
            )
        paused = route.get("paused", False)
        if not isinstance(paused, bool):
            errors.append(f"{prefix}.paused must be true or false")
        pause_reason = route.get("pause_reason")
        if paused and (
            not isinstance(pause_reason, str)
            or not pause_reason.strip()
            or len(pause_reason.strip()) > 200
        ):
            errors.append(
                f"{prefix}.pause_reason is required when paused and must be at most 200 characters"
            )
        if not paused and pause_reason is not None:
            errors.append(f"{prefix}.pause_reason requires paused: true")
        if not isinstance(route.get("allow_sender_only", False), bool):
            errors.append(f"{prefix}.allow_sender_only must be true or false")
        benchmark_review_only = route.get("benchmark_review_only", False)
        if not isinstance(benchmark_review_only, bool):
            errors.append(f"{prefix}.benchmark_review_only must be true or false")
        if benchmark_review_only and (
            workbook_mode != "existing" or sheet_mode != "summary"
        ):
            errors.append(
                f"{prefix}.benchmark_review_only is only valid for an existing summary sheet"
            )
        policy = str(route.get("cumulative_policy", "require"))
        if policy not in {"require", "unit", "offset"}:
            errors.append(f"{prefix}.cumulative_policy is invalid")
        if policy == "offset" and (
            not isinstance(route.get("cumulative_offset"), (int, float))
            or isinstance(route.get("cumulative_offset"), bool)
        ):
            errors.append(f"{prefix}.cumulative_offset is required for offset policy")
        if str(route.get("return_basis", "cumulative")) not in {"unit", "cumulative"}:
            errors.append(f"{prefix}.return_basis is invalid")
        if str(route.get("return_frequency", "weekly")) not in {"daily", "weekly"}:
            errors.append(f"{prefix}.return_frequency is invalid")
        if str(route.get("data_frequency", "auto")) not in {
            "auto",
            "daily",
            "weekly",
        }:
            errors.append(f"{prefix}.data_frequency is invalid")
        if (
            sheet_mode == "template"
            and str(route.get("data_frequency", "auto")) == "auto"
        ):
            errors.append(
                f"{prefix}.data_frequency must be daily or weekly for template initialization"
            )
        if sheet_mode == "template" and str(
            route.get("return_frequency", "weekly")
        ) != str(route.get("data_frequency", "auto")):
            errors.append(
                f"{prefix}.return_frequency must match data_frequency in template mode"
            )
        if route.get("series_start"):
            try:
                dt.date.fromisoformat(str(route["series_start"]))
            except ValueError:
                errors.append(
                    f"{prefix}.series_start must be a real date in YYYY-MM-DD form"
                )
        try:
            max_staleness = int(route.get("max_staleness_days", 14))
            if (
                isinstance(route.get("max_staleness_days", 14), bool)
                or not 1 <= max_staleness <= 366
            ):
                raise ValueError
        except (TypeError, ValueError):
            errors.append(f"{prefix}.max_staleness_days must be between 1 and 366")
        benchmark = route.get("benchmark")
        if benchmark_review_only and benchmark is not None:
            errors.append(
                f"{prefix}.benchmark_review_only cannot be combined with benchmark"
            )
        if benchmark is not None:
            if not isinstance(benchmark, dict):
                errors.append(f"{prefix}.benchmark must be null or an object")
            else:
                for field in sorted(set(benchmark) - BENCHMARK_FIELDS):
                    errors.append(f"{prefix}.benchmark has unknown field: {field}")
                for field in ("source_sheet", "source_date", "source_value"):
                    if not str(benchmark.get(field, "")).strip():
                        errors.append(f"{prefix}.benchmark.{field} is required")
                if benchmark.get("source_type", "level") not in {
                    "level",
                    "aligned_return",
                }:
                    errors.append(
                        f"{prefix}.benchmark.source_type must be level or aligned_return"
                    )
                display_name = benchmark.get("display_name")
                if display_name is not None and (
                    not isinstance(display_name, str)
                    or not display_name.strip()
                    or len(display_name.strip()) > 100
                    or "\n" in display_name
                    or "\r" in display_name
                    or display_name.lstrip().startswith("=")
                ):
                    errors.append(
                        f"{prefix}.benchmark.display_name must be a non-empty, non-formula single-line string of at most 100 characters"
                    )
        key = (sender, sheet, code or "")
        if key in route_keys:
            errors.append(f"{prefix} duplicates an earlier route")
        route_keys.add(key)
        senders.setdefault(sender, []).append(route)

    for sender, sender_routes in senders.items():
        if len(sender_routes) > 1:
            for route in sender_routes:
                if not normalize_code(route.get("code")):
                    errors.append(
                        f"sender {sender} has multiple routes; every route needs an exact code"
                    )
                if route.get("allow_sender_only"):
                    errors.append(
                        f"sender {sender} has multiple routes; allow_sender_only is unsafe"
                    )
        codes = [normalize_code(route.get("code")) for route in sender_routes]
        if len([code for code in codes if code]) != len(
            set(code for code in codes if code)
        ):
            errors.append(f"sender {sender} has duplicate route codes")

    for index, item in enumerate(config.get("schedule") or [], 1):
        if not isinstance(item, dict):
            errors.append(f"schedule #{index} must be an object")
            continue
        for field in sorted(set(item) - {"days", "time"}):
            errors.append(f"schedule #{index} has unknown field: {field}")
        days = {str(day).upper() for day in item.get("days") or []}
        time = str(item.get("time", ""))
        if not days or not days <= DAY_NAMES:
            errors.append(f"schedule #{index}.days is invalid")
        if not re.fullmatch(r"(?:[01]\d|2[0-3]):[0-5]\d", time):
            errors.append(f"schedule #{index}.time is invalid")

    validation = config.get("validation") or {}
    if not isinstance(validation, dict):
        errors.append("validation must be an object")
        validation = {}
    for field in sorted(
        set(validation)
        - {"minimum_history_dates", "tolerance", "max_future_days", "max_period_change"}
    ):
        errors.append(f"unknown validation field: {field}")
    try:
        minimum = int(validation.get("minimum_history_dates", 2))
        if (
            isinstance(validation.get("minimum_history_dates", 2), bool)
            or not 2 <= minimum <= 100
        ):
            raise ValueError
    except (TypeError, ValueError):
        errors.append("validation.minimum_history_dates must be between 2 and 100")
    try:
        tolerance = float(validation.get("tolerance", 1e-6))
        if (
            isinstance(validation.get("tolerance", 1e-6), bool)
            or not 0 < tolerance <= 0.01
        ):
            raise ValueError
    except (TypeError, ValueError):
        errors.append("validation.tolerance must be greater than 0 and at most 0.01")
    try:
        max_future = int(validation.get("max_future_days", 0))
        if (
            isinstance(validation.get("max_future_days", 0), bool)
            or not 0 <= max_future <= 7
        ):
            raise ValueError
    except (TypeError, ValueError):
        errors.append("validation.max_future_days must be between 0 and 7")
    try:
        max_change = float(validation.get("max_period_change", 0.5))
        if (
            isinstance(validation.get("max_period_change", 0.5), bool)
            or not 0.001 <= max_change <= 10
        ):
            raise ValueError
    except (TypeError, ValueError):
        errors.append("validation.max_period_change must be between 0.001 and 10")

    style = config.get("style") or {}
    if not isinstance(style, dict):
        errors.append("style must be an object")
        style = {}
    for field in sorted(set(style) - {"mode", "zero_threshold"}):
        errors.append(f"unknown style field: {field}")
    if style.get("mode", "infer") not in {"infer", "cn-red-up-green-down"}:
        errors.append("style.mode must be infer or cn-red-up-green-down")
    try:
        zero_threshold = float(style.get("zero_threshold", 0.00005))
        if (
            isinstance(style.get("zero_threshold", 0.00005), bool)
            or not 0 <= zero_threshold <= 0.1
        ):
            raise ValueError
    except (TypeError, ValueError):
        errors.append("style.zero_threshold must be between 0 and 0.1")

    overrides = config.get("column_overrides") or {}
    if not isinstance(overrides, dict):
        errors.append("column_overrides must be an object")
    else:
        for sheet, fields in overrides.items():
            if not isinstance(fields, dict):
                errors.append(f"column_overrides.{sheet} must be an object")
                continue
            for field in sorted(set(fields) - COLUMN_FIELDS):
                errors.append(f"column_overrides.{sheet} has unknown field: {field}")

    retention = config.get("retention") or {}
    if not isinstance(retention, dict):
        errors.append("retention must be an object")
        retention = {}
    for field in sorted(set(retention) - {"backup_count", "preview_count", "log_days"}):
        errors.append(f"unknown retention field: {field}")
    for field, default, maximum in (
        ("backup_count", 10, 100),
        ("preview_count", 10, 100),
        ("log_days", 30, 365),
    ):
        try:
            value = int(retention.get(field, default))
            if (
                isinstance(retention.get(field, default), bool)
                or not 1 <= value <= maximum
            ):
                raise ValueError
        except (TypeError, ValueError):
            errors.append(f"retention.{field} must be between 1 and {maximum}")

    if errors:
        raise ConfigError("; ".join(errors))


def normalize_code(value: Any) -> str | None:
    text = re.sub(r"[^A-Za-z0-9]", "", str(value or "")).upper()
    return text or None


def active_routes(config: dict[str, Any]) -> list[dict[str, Any]]:
    return [
        route
        for route in config.get("routes") or []
        if isinstance(route, dict) and not route.get("paused", False)
    ]


def write_json_atomic(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    temporary.replace(path)
