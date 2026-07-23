from __future__ import annotations

import datetime as dt
import gc
import hashlib
import importlib.util
import importlib.metadata
import json
import platform
import re
import sys
import time
from copy import deepcopy
from email.utils import parsedate_to_datetime
from itertools import combinations
from pathlib import Path
from typing import Any, Callable

from nav_commit import (
    detected_spreadsheet_apps,
    ensure_process_exit,
    spreadsheet_app,
)
from nav_config import (
    ROOT,
    STATE_ROOT,
    active_routes,
    benchmark_requires_review,
    benchmark_review_issue,
    normalize_code,
    write_json_atomic,
)
from nav_mail import (
    decoded,
    fetch_authorized_messages,
    fetch_candidate_headers,
    fetch_candidate_messages,
    non_nav_subject_category,
    single_from_address,
)
from nav_parse import (
    NavRow,
    ParseError,
    choose_route_rows,
    consume_attachment_parse_diagnostics,
    consume_parse_library_warnings,
    deduplicate,
    rows_from_message,
)
from nav_workbook import build_preview, validate_history
from runtime_secret import read_password


def _sender_id(sender: str) -> str:
    return hashlib.sha256(sender.lower().encode("utf-8")).hexdigest()[:12]


def _short_id(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8", errors="replace")).hexdigest()[:12]


DISCOVERY_SNAPSHOT = STATE_ROOT / "discovery-snapshot.json"
DISCOVERY_ROUTE_FIELDS = (
    "sender",
    "subject_contains",
    "subject_excludes",
    "subject_product_code",
    "sheet",
    "code",
    "code_aliases",
    "parser",
    "paused",
    "pause_reason",
    "allow_sender_only",
    "series_start",
    "baseline_overlap",
    "max_staleness_days",
)


def _source_hash(path: Path) -> str:
    if not path.is_file():
        return "missing"
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _library_version(name: str) -> str:
    try:
        return importlib.metadata.version(name)
    except importlib.metadata.PackageNotFoundError:
        return "missing"


def _discovery_config_fingerprint(config: dict[str, Any]) -> str:
    imap = config.get("imap") or {}
    routes = [
        {
            field: route.get(field)
            for field in DISCOVERY_ROUTE_FIELDS
            if field in route
        }
        for route in config.get("routes") or []
        if isinstance(route, dict)
    ]
    routes.sort(key=lambda route: str(route.get("sheet") or ""))
    source_hashes = {
        name: _source_hash(STATE_ROOT / name)
        for name in ("nav_service.py", "nav_mail.py", "nav_parse.py")
    }
    local_parsers = sorted(
        {
            str(route.get("parser")).split(":", 1)[1]
            for route in active_routes(config)
            if str(route.get("parser") or "").startswith("local:")
        }
    )
    source_hashes.update(
        {
            f"parsers/{name}.py": _source_hash(
                STATE_ROOT / "parsers" / f"{name}.py"
            )
            for name in local_parsers
        }
    )
    payload = {
        "imap": {
            field: imap.get(field)
            for field in (
                "host",
                "port",
                "user",
                "mailbox",
                "lookback_days",
                "max_messages",
                "max_header_messages",
                "max_message_bytes",
                "max_total_bytes",
            )
        },
        "routes": routes,
        "sources": source_hashes,
        "libraries": {
            name: _library_version(name)
            for name in ("openpyxl", "pypdf", "xlrd")
        },
    }
    encoded = json.dumps(
        payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    )
    return hashlib.sha256(encoded.encode("utf-8")).hexdigest()


def _snapshot_rows(rows: dict[str, list[NavRow]]) -> dict[str, list[dict[str, Any]]]:
    return {
        sheet: [
            {
                "date": row.date.isoformat(),
                "unit": row.unit,
                "cumulative": row.cumulative,
                "code": row.code,
            }
            for row in values
        ]
        for sheet, values in rows.items()
    }


def _restore_snapshot_rows(
    payload: dict[str, Any],
) -> dict[str, list[NavRow]]:
    restored: dict[str, list[NavRow]] = {}
    for sheet, values in payload.items():
        if not isinstance(sheet, str) or not isinstance(values, list):
            raise RuntimeError("发现快照结构无效；请重新运行完整 preview")
        restored[sheet] = [
            NavRow(
                dt.date.fromisoformat(str(item["date"])),
                float(item["unit"]),
                (
                    None
                    if item.get("cumulative") is None
                    else float(item["cumulative"])
                ),
                normalize_code(item.get("code")),
                "verified-discovery-snapshot",
            )
            for item in values
            if isinstance(item, dict)
        ]
        if len(restored[sheet]) != len(values):
            raise RuntimeError("发现快照记录无效；请重新运行完整 preview")
    return restored


def _save_discovery_snapshot(
    config: dict[str, Any],
    rows: dict[str, list[NavRow]],
    report: dict[str, Any],
    mail_scope_fingerprint: str | None,
) -> None:
    if not report.get("passed") or not mail_scope_fingerprint:
        DISCOVERY_SNAPSHOT.unlink(missing_ok=True)
        return
    write_json_atomic(
        DISCOVERY_SNAPSHOT,
        {
            "schema_version": 1,
            "created_at": dt.datetime.now(dt.timezone.utc).isoformat(),
            "config_fingerprint": _discovery_config_fingerprint(config),
            "mail_scope_fingerprint": mail_scope_fingerprint,
            "rows": _snapshot_rows(rows),
            "report": report,
        },
    )


def reuse_discovery_snapshot(
    config: dict[str, Any],
) -> tuple[dict[str, list[NavRow]], dict[str, Any]]:
    try:
        snapshot = json.loads(DISCOVERY_SNAPSHOT.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError) as exc:
        raise RuntimeError(
            "没有可复用的完整发现快照；请先运行一次普通 preview"
        ) from exc
    if (
        snapshot.get("schema_version") != 1
        or snapshot.get("config_fingerprint")
        != _discovery_config_fingerprint(config)
    ):
        raise RuntimeError(
            "路由、邮箱范围或解析器已变化，发现快照不能复用；"
            "请运行普通 preview"
        )
    current_scope = fetch_authorized_messages(config, load_bodies=False)
    if (
        not current_scope.scope_fingerprint
        or current_scope.scope_fingerprint
        != snapshot.get("mail_scope_fingerprint")
    ):
        raise RuntimeError(
            "邮箱范围自上次完整解析后已变化，发现快照不能复用；"
            "请运行普通 preview"
        )
    rows = _restore_snapshot_rows(snapshot.get("rows") or {})
    report = deepcopy(snapshot.get("report") or {})
    if not report.get("passed"):
        raise RuntimeError("发现快照不是成功结果；请运行普通 preview")
    report["discovery_reused"] = True
    report["mail_scope_reverified"] = True
    report.setdefault("warnings", []).append(
        "已用轻量邮件头、UID 和大小指纹确认邮箱范围未变化，"
        "本次复用上次完整解析结果"
    )
    write_json_atomic(STATE_ROOT / "route-report.json", report)
    return rows, report


def _series_start(route: dict[str, Any]) -> dt.date:
    value = str(route.get("series_start") or "").strip()
    return dt.date.fromisoformat(value) if value else dt.date.min


def _accepted_route_codes(route: dict[str, Any]) -> set[str]:
    expected = normalize_code(route.get("code"))
    accepted = {expected} if expected else set()
    accepted.update(
        normalized
        for alias in route.get("code_aliases") or []
        if (normalized := normalize_code(alias))
    )
    return accepted


def _matching_routes_for_row(
    row: NavRow, routes: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    matches: list[dict[str, Any]] = []
    for route in routes:
        expected = normalize_code(route.get("code"))
        if expected and row.code in _accepted_route_codes(route):
            matches.append(route)
        elif (
            row.code is None
            and len(routes) == 1
            and route.get("allow_sender_only")
        ):
            matches.append(route)
        elif expected is None and len(routes) == 1:
            matches.append(route)
    return matches


def _route_ref(route: dict[str, Any], sender: str) -> dict[str, str]:
    sheet = str(route["sheet"])
    identity = "\x1f".join(
        (
            sender.lower(),
            sheet,
            str(route.get("code") or ""),
            str(route.get("parser", "auto")),
        )
    )
    return {"route_id": _short_id(identity), "sheet": sheet}


def _route_diagnostic_ref(route: dict[str, Any], sender: str) -> dict[str, Any]:
    result: dict[str, Any] = {
        **_route_ref(route, sender),
        "parser": str(route.get("parser", "auto")),
        "subject_filter_missing": not bool(
            str(route.get("subject_contains") or "").strip()
        ),
    }
    subject_filter = str(route.get("subject_contains") or "").strip().casefold()
    if subject_filter:
        result["subject_filter_id"] = _short_id(subject_filter)
    subject_excludes = [
        str(item).strip().casefold()
        for item in route.get("subject_excludes") or []
        if str(item).strip()
    ]
    if subject_excludes:
        result["subject_exclude_ids"] = [
            _short_id(item) for item in subject_excludes
        ]
    code = normalize_code(route.get("code"))
    if code:
        result["code_id"] = _short_id(code)
    return result


def _message_ids(message: Any, sender: str) -> tuple[str, str]:
    subject = decoded(message.get("Subject")).strip()
    identity = "\x1f".join(
        (
            sender.lower(),
            str(message.get("Message-ID") or ""),
            str(message.get("Date") or ""),
            subject,
        )
    )
    return _short_id(identity), _short_id(subject.casefold())


def _record_route_overlap(
    overlaps: dict[tuple[str, tuple[str, ...]], dict[str, Any]],
    *,
    reason: str,
    routes: list[dict[str, Any]],
    sender: str,
    message: Any,
) -> None:
    refs = sorted(
        (_route_diagnostic_ref(route, sender) for route in routes),
        key=lambda item: item["route_id"],
    )
    route_ids = tuple(item["route_id"] for item in refs)
    key = (reason, route_ids)
    message_id, subject_id = _message_ids(message, sender)
    item = overlaps.setdefault(
        key,
        {
            "reason": reason,
            "routes": refs,
            "message_count": 0,
            "message_ids": [],
            "subject_ids": [],
            "suggested_filter": {
                "action": "make-route-filters-mutually-exclusive",
                "fields": [
                    "subject_contains",
                    "subject_excludes",
                    "parser",
                    "code",
                    "code_aliases",
                ],
                "rerun_sheets": [ref["sheet"] for ref in refs],
            },
        },
    )
    item["message_count"] += 1
    if message_id not in item["message_ids"] and len(item["message_ids"]) < 5:
        item["message_ids"].append(message_id)
    if subject_id not in item["subject_ids"] and len(item["subject_ids"]) < 5:
        item["subject_ids"].append(subject_id)


def _safe_source_type(source: str) -> str:
    lowered = source.casefold()
    for token in ("xlsx", "xls", "csv", "pdf", "html", "body", "text"):
        if token in lowered:
            return token
    if "subject-product-code" in lowered:
        return "subject-product-code"
    return "parser-output"


def _row_signature(row: NavRow) -> tuple[float, float | None]:
    return (
        round(row.unit, 10),
        None if row.cumulative is None else round(row.cumulative, 10),
    )


def _correction_subject_marker(subject: str) -> str | None:
    text = re.sub(r"\s+", "", subject).casefold()
    if any(
        marker in text
        for marker in (
            "无需更正",
            "不作更正",
            "不用更正",
            "非更正",
            "未更正",
            "更正前",
            "修正前",
        )
    ):
        return None
    if any(
        marker in text
        for marker in (
            "以此为准",
            "以本邮件为准",
            "以本通知为准",
            "更正通知",
            "更正版",
            "净值更正",
            "数据更正",
            "修正通知",
            "修正版",
        )
    ):
        return "explicit-authoritative-correction"
    if re.search(r"(?<!非)(?:更正|修正)", text):
        return "explicit-correction"
    if re.search(r"\b(?:corrected|correction|revised|revision)\b", subject.casefold()):
        return "explicit-correction"
    return None


def _message_order(message: Any) -> tuple[str, int] | None:
    uid = getattr(message, "_nav_source_uid", None)
    if isinstance(uid, int) and uid > 0:
        return "mailbox-uid", uid
    try:
        sent = parsedate_to_datetime(str(message.get("Date") or ""))
    except (TypeError, ValueError, OverflowError):
        return None
    if sent is None or sent.tzinfo is None:
        return None
    return "message-date", int(sent.timestamp() * 1_000_000)


def _message_evidence(message: Any, sender: str) -> dict[str, Any]:
    message_id, subject_id = _message_ids(message, sender)
    return {
        "message_id": message_id,
        "subject_id": subject_id,
        "order": _message_order(message),
        "correction_marker": _correction_subject_marker(
            decoded(message.get("Subject")).strip()
        ),
    }


def _candidate_fingerprint(
    row: NavRow, evidence: dict[str, Any]
) -> dict[str, str]:
    value_id = _short_id(
        "\x1f".join(
            (
                row.date.isoformat(),
                f"{row.unit:.10f}",
                "" if row.cumulative is None else f"{row.cumulative:.10f}",
            )
        )
    )
    return {
        "message_id": str(evidence["message_id"]),
        "subject_id": str(evidence["subject_id"]),
        "value_id": value_id,
        "source_id": _short_id(str(row.source)),
        "source_type": _safe_source_type(str(row.source)),
    }


def _apply_explicit_corrections(
    candidates: list[tuple[NavRow, dict[str, Any]]],
    route: dict[str, Any],
    sender: str,
) -> tuple[list[NavRow], list[dict[str, Any]]]:
    by_date: dict[dt.date, list[tuple[NavRow, dict[str, Any]]]] = {}
    for row, evidence in candidates:
        by_date.setdefault(row.date, []).append((row, evidence))
    selected: list[NavRow] = []
    applied: list[dict[str, Any]] = []
    for date, date_candidates in sorted(by_date.items()):
        signatures = {_row_signature(row) for row, _ in date_candidates}
        if len({signature[0] for signature in signatures}) == 1 and len(
            {signature[1] for signature in signatures if signature[1] is not None}
        ) <= 1:
            selected.extend(row for row, _ in date_candidates)
            continue
        by_message: dict[
            str, list[tuple[NavRow, dict[str, Any]]]
        ] = {}
        for row, evidence in date_candidates:
            by_message.setdefault(str(evidence["message_id"]), []).append(
                (row, evidence)
            )
        correction_messages: list[
            tuple[tuple[str, int], list[tuple[NavRow, dict[str, Any]]]]
        ] = []
        correction_is_ambiguous = False
        for group in by_message.values():
            evidence = group[0][1]
            if not evidence.get("correction_marker"):
                continue
            if len({_row_signature(row) for row, _ in group}) != 1:
                correction_is_ambiguous = True
                break
            order = evidence.get("order")
            if order is None:
                correction_is_ambiguous = True
                break
            correction_messages.append((order, group))
        order_kinds = {
            str(evidence["order"][0])
            for _, evidence in date_candidates
            if evidence.get("order") is not None
        }
        all_ordered = all(
            evidence.get("order") is not None for _, evidence in date_candidates
        )
        if (
            correction_is_ambiguous
            or not correction_messages
            or not all_ordered
            or len(order_kinds) != 1
        ):
            selected.extend(row for row, _ in date_candidates)
            continue
        latest_order, replacement_group = max(
            correction_messages, key=lambda item: item[0][1]
        )
        other_candidates = [
            (row, evidence)
            for row, evidence in date_candidates
            if str(evidence["message_id"])
            != str(replacement_group[0][1]["message_id"])
        ]
        if not other_candidates or any(
            evidence["order"][1] >= latest_order[1]
            for _, evidence in other_candidates
        ):
            selected.extend(row for row, _ in date_candidates)
            continue
        replacement_row, replacement_evidence = sorted(
            replacement_group,
            key=lambda item: item[0].cumulative is not None,
            reverse=True,
        )[0]
        selected.append(replacement_row)
        applied.append(
            {
                "route": _route_diagnostic_ref(route, sender),
                "date": date.isoformat(),
                "correction_marker": replacement_evidence["correction_marker"],
                "order_basis": latest_order[0],
                "replacement": _candidate_fingerprint(
                    replacement_row, replacement_evidence
                ),
                "replaced": [
                    _candidate_fingerprint(row, evidence)
                    for row, evidence in other_candidates
                ],
            }
        )
    return selected, applied


def _date_conflict_diagnostics(
    rows: list[NavRow], route: dict[str, Any], sender: str
) -> list[dict[str, Any]]:
    by_date: dict[dt.date, list[NavRow]] = {}
    for row in rows:
        by_date.setdefault(row.date, []).append(row)
    route_ref = _route_diagnostic_ref(route, sender)
    conflicts: list[dict[str, Any]] = []
    for date, candidates in sorted(by_date.items()):
        signatures = {_row_signature(item) for item in candidates}
        compatible_units = len({value[0] for value in signatures}) == 1
        cumulative_values = {value[1] for value in signatures if value[1] is not None}
        if compatible_units and len(cumulative_values) <= 1:
            continue
        fingerprints = []
        seen: set[tuple[str, str]] = set()
        for row in candidates:
            value_id = _short_id(
                "\x1f".join(
                    (
                        row.date.isoformat(),
                        f"{row.unit:.10f}",
                        "" if row.cumulative is None else f"{row.cumulative:.10f}",
                    )
                )
            )
            source_id = _short_id(str(row.source))
            key = (value_id, source_id)
            if key in seen:
                continue
            seen.add(key)
            fingerprints.append(
                {
                    "value_id": value_id,
                    "source_id": source_id,
                    "source_type": _safe_source_type(str(row.source)),
                }
            )
        conflicts.append(
            {
                "route": route_ref,
                "date": date.isoformat(),
                "candidate_count": len(candidates),
                "candidate_sources": fingerprints,
                "suggested_filter": {
                    "action": "narrow-source-before-accepting-one-value",
                    "fields": ["subject_contains", "parser"],
                    "rerun_sheet": str(route["sheet"]),
                },
            }
        )
    return conflicts


def _attachment_metadata(message: Any) -> list[dict[str, Any]]:
    grouped: dict[tuple[str, str], dict[str, Any]] = {}
    for part in message.walk():
        raw_name = part.get_filename()
        if not raw_name:
            continue
        name = decoded(raw_name).strip()
        suffix = Path(name).suffix.lower()
        if not re.fullmatch(r"\.[a-z0-9]{1,10}", suffix):
            suffix = "(none)"
        content_type = str(part.get_content_type() or "application/octet-stream")
        if not re.fullmatch(r"[a-z0-9.+-]+/[a-z0-9.+-]+", content_type.lower()):
            content_type = "application/octet-stream"
        key = (suffix, content_type.lower())
        item = grouped.setdefault(
            key,
            {
                "suffix": suffix,
                "content_type": content_type.lower(),
                "count": 0,
                "name_ids": [],
            },
        )
        item["count"] += 1
        name_id = _short_id(name.casefold())
        if name_id not in item["name_ids"]:
            item["name_ids"].append(name_id)
    return list(grouped.values())


def _parse_diagnostic(
    message: Any,
    sender: str,
    routes: list[dict[str, Any]],
    parser_name: str,
    exc: BaseException | None,
    attachment_diagnostics: list[dict[str, str]] | None = None,
    *,
    message_parsed: bool = False,
) -> dict[str, Any]:
    subject = decoded(message.get("Subject")).strip()
    attachments = _attachment_metadata(message)
    message_identity = "\x1f".join(
        (
            sender.lower(),
            str(message.get("Message-ID") or ""),
            str(message.get("Date") or ""),
            subject,
        )
    )
    attachment_diagnostics = list(attachment_diagnostics or [])
    nav_subject_signal = bool(
        re.search(r"(?:\bnav\b|净值|估值)", subject.casefold())
    )
    supported_attachment = any(
        item["suffix"] in {".xlsx", ".xlsm", ".xls", ".csv", ".txt", ".pdf"}
        for item in attachments
    )
    built_in_non_nav_category = non_nav_subject_category(subject)
    likely_non_nav = not message_parsed and (
        built_in_non_nav_category is not None
        or (not nav_subject_signal and not supported_attachment)
    )
    if message_parsed:
        error_type = "AttachmentCoverageGap"
        root_error_type = "ControlledAttachmentDiagnostic"
        safe_message = "The message produced NAV rows, but one or more attachments produced no usable NAV records"
    elif exc is None:
        error_type = "NoRowsParsed"
        root_error_type = "NoRowsParsed"
        safe_message = "Parser returned no NAV rows"
    else:
        error_type = (
            "LocalParserError"
            if parser_name.startswith("local:")
            else "AutomaticParserError"
        )
        root_error_type = "ControlledParseFailure"
        safe_message = "Parser failed; use the controlled error types and local message identifiers"
    return {
        "stage": "message_parse",
        "message_id": _short_id(message_identity),
        "subject_id": _short_id(subject.casefold()),
        "routes": [_route_ref(route, sender) for route in routes],
        "parser": parser_name,
        "attachment_count": sum(item["count"] for item in attachments),
        "attachment_types": attachments,
        "attachment_diagnostics": attachment_diagnostics,
        "notice_classification": (
            built_in_non_nav_category
            or ("likely-non-nav-notice" if likely_non_nav else "nav-parser-gap")
        ),
        "suggested_filter": (
            {
                "action": "confirm-non-nav-then-add-subject-excludes",
                "field": "subject_excludes",
            }
            if likely_non_nav
            else {
                "action": "inspect-attachment-or-message-parser",
                "field": "parser",
            }
        ),
        "error_type": error_type,
        "root_error_type": root_error_type,
        "safe_message": safe_message,
    }


def _warning_summary(items: list[dict[str, str]]) -> list[dict[str, Any]]:
    counts: dict[tuple[str, str, str, str], int] = {}
    for item in items:
        key = (
            str(item.get("code") or "parser-library-warning"),
            str(item.get("library") or "unknown"),
            str(item.get("source_type") or "unknown"),
            str(item.get("category") or "Warning"),
        )
        counts[key] = counts.get(key, 0) + 1
    return [
        {
            "code": key[0],
            "library": key[1],
            "source_type": key[2],
            "category": key[3],
            "count": count,
        }
        for key, count in sorted(counts.items())
    ]


PROPOSAL_PARTIAL = STATE_ROOT / "route-proposals.partial.json"
PROPOSAL_PROGRESS = STATE_ROOT / "route-proposal-progress.json"


def _proposal_progress(
    scope_id: str,
    status: str,
    phase: str,
    completed: int = 0,
    total: int = 0,
    *,
    resume_before_uid: int | None = None,
    attachments_parsed: int = 0,
    sink: Callable[[dict[str, Any]], None] | None = None,
) -> None:
    payload = {
        "status": status,
        "scope_id": scope_id,
        "phase": phase,
        "completed": completed,
        "total": total,
        "attachments_parsed": attachments_parsed,
        "resume_available": resume_before_uid is not None,
        "resume_before_uid": resume_before_uid,
        "updated_at": dt.datetime.now().isoformat(timespec="seconds"),
    }
    write_json_atomic(PROPOSAL_PROGRESS, payload)
    if sink:
        sink(payload)


def _merge_proposal_candidates(
    previous: list[dict[str, Any]], current: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    merged: dict[str, dict[str, Any]] = {}
    for candidate in [*previous, *current]:
        sender = str(candidate.get("sender") or "").strip().casefold()
        if not sender:
            continue
        target = merged.setdefault(
            sender,
            {
                "sender": sender,
                "message_count": 0,
                "recent_message_times": [],
                "subject_examples": [],
                "observations": {},
                "selection": candidate.get("selection"),
            },
        )
        target["message_count"] += int(candidate.get("message_count") or 0)
        for field in ("recent_message_times", "subject_examples"):
            for value in candidate.get(field) or []:
                if value not in target[field]:
                    target[field].append(value)
        for item in candidate.get("observations") or []:
            key = (
                item.get("date"),
                item.get("code"),
                item.get("unit"),
                item.get("cumulative"),
            )
            target["observations"][key] = item
    result: list[dict[str, Any]] = []
    for sender, target in sorted(merged.items()):
        observations = sorted(
            target.pop("observations").values(),
            key=lambda item: (item["date"], item.get("code") or "", item["unit"]),
        )
        dates = sorted({item["date"] for item in observations})
        target["detected_codes"] = sorted(
            {item["code"] for item in observations if item.get("code")}
        )
        target["first_date"] = dates[0]
        target["latest_date"] = dates[-1]
        target["observations"] = observations[-100:]
        target["recent_message_times"] = target["recent_message_times"][:20]
        target["subject_examples"] = target["subject_examples"][:10]
        result.append(target)
    return result


def _merge_warning_summaries(
    previous: list[dict[str, Any]], current: list[dict[str, Any]]
) -> list[dict[str, Any]]:
    counts: dict[tuple[str, str, str], int] = {}
    for item in [*previous, *current]:
        key = (
            str(item.get("stage") or ""),
            str(item.get("source_type") or ""),
            str(item.get("category") or ""),
        )
        counts[key] = counts.get(key, 0) + int(item.get("count") or 0)
    return [
        {
            "stage": key[0],
            "source_type": key[1],
            "category": key[2],
            "count": count,
        }
        for key, count in sorted(counts.items())
    ]


def _bind_subject_product_code(
    rows: list[NavRow], code: str
) -> tuple[list[NavRow], bool]:
    normalized = normalize_code(code)
    if not normalized:
        raise ParseError("Subject product code binding is empty")
    present = {row.code for row in rows if row.code}
    if present and (present != {normalized} or any(row.code is None for row in rows)):
        raise ParseError(
            "Subject product code binding conflicts with mixed or different row codes"
        )
    if present:
        return rows, False
    by_date: dict[dt.date, set[tuple[float, float | None]]] = {}
    for row in rows:
        by_date.setdefault(row.date, set()).add(
            (
                round(row.unit, 10),
                None
                if row.cumulative is None
                else round(row.cumulative, 10),
            )
        )
    if any(len(values) != 1 for values in by_date.values()):
        raise ParseError(
            "Subject product code binding found conflicting values for one date"
        )
    return (
        [
            NavRow(
                date=row.date,
                unit=row.unit,
                cumulative=row.cumulative,
                code=normalized,
                source=f"{row.source}:subject-product-code",
            )
            for row in rows
        ],
        True,
    )


def propose_headers(config: dict[str, Any], limit: int = 25) -> dict[str, Any]:
    headers, scan = fetch_candidate_headers(config, limit)
    groups: dict[str, dict[str, Any]] = {}
    ignored = 0
    for header in headers:
        sender = single_from_address(header)
        if not sender:
            ignored += 1
            continue
        group = groups.setdefault(sender, {"subjects": [], "messages": 0})
        group["messages"] += 1
        subject = decoded(header.get("Subject")).strip()
        if subject and subject not in group["subjects"]:
            group["subjects"].append(subject)
    candidates = [
        {
            "sender": sender,
            "message_count": group["messages"],
            "subject_examples": group["subjects"][:20],
        }
        for sender, group in sorted(groups.items())
    ]
    report = {
        "passed": bool(candidates),
        "header_only": True,
        "scan": {**scan, "headers_ignored_without_single_sender": ignored},
        "candidates": candidates,
        "warnings": (
            ["候选邮件头报告受到数量上限限制；需要时扩大 header_limit 后重试"]
            if scan["truncated"]
            else []
        ),
        "errors": [] if candidates else ["候选邮件头中没有可唯一识别的发件人"],
    }
    write_json_atomic(STATE_ROOT / "route-proposal-headers.json", report)
    return report


def propose_routes(
    config: dict[str, Any],
    *,
    sender: str | None = None,
    subject_contains: str | None = None,
    subject_product_code: str | None = None,
    lookback_days: int | None = None,
    batch_messages: int | None = None,
    time_budget_seconds: int | None = None,
    resume: bool = False,
    progress_sink: Callable[[dict[str, Any]], None] | None = None,
) -> dict[str, Any]:
    previous: dict[str, Any] | None = None
    before_uid = None
    scoped_config = deepcopy(config)
    if resume:
        try:
            previous = json.loads(PROPOSAL_PARTIAL.read_text(encoding="utf-8"))
        except (FileNotFoundError, json.JSONDecodeError) as exc:
            raise RuntimeError("没有可恢复的候选扫描；请重新开始选择式 propose") from exc
        selection = previous.get("selection") or {}
        sender = str(selection.get("sender") or "").strip() or None
        subject_contains = (
            str(selection.get("subject_contains") or "").strip() or None
        )
        subject_product_code = (
            str(selection.get("subject_product_code") or "").strip() or None
        )
        previous_scan = previous.get("scan") or {}
        lookback_days = int(previous_scan.get("effective_lookback_days") or 30)
        before_uid = previous_scan.get("resume_before_uid")
        if before_uid is None:
            raise RuntimeError("候选扫描检查点没有可恢复游标；请重新开始")
        before_uid = int(before_uid)
    if lookback_days is not None:
        if isinstance(lookback_days, bool) or not 1 <= int(lookback_days) <= 3650:
            raise ValueError("本次候选回看天数必须在 1 到 3650 之间")
        scoped_config["imap"]["lookback_days"] = int(lookback_days)
    effective_lookback = int(scoped_config["imap"].get("lookback_days", 180))
    binding_code = normalize_code(subject_product_code)
    if binding_code:
        if not sender or not subject_contains:
            raise ValueError(
                "--subject-product-code 必须同时限定 --sender 和 --subject-contains"
            )
        if normalize_code(subject_contains) != binding_code:
            raise ValueError(
                "--subject-product-code 必须与本次唯一主题代码完全一致"
            )
    if time_budget_seconds is not None and (
        isinstance(time_budget_seconds, bool)
        or not 30 <= int(time_budget_seconds) <= 1800
    ):
        raise ValueError("单范围时间预算必须在 30 到 1800 秒之间")
    deadline = (
        time.monotonic() + int(time_budget_seconds)
        if time_budget_seconds is not None
        else None
    )
    scope_id = _short_id(
        f"{str(sender or '').casefold()}|{str(subject_contains or '').casefold()}|"
        f"{effective_lookback}|{binding_code or ''}"
    )
    _proposal_progress(scope_id, "running", "search", sink=progress_sink)

    def progress(phase: str, completed: int, total: int) -> None:
        _proposal_progress(
            scope_id, "running", phase, completed, total, sink=progress_sink
        )

    fetch_options: dict[str, Any] = {}
    if sender or subject_contains:
        fetch_options.update(sender=sender, subject_contains=subject_contains)
    if batch_messages is not None:
        fetch_options["batch_messages"] = batch_messages
    if before_uid is not None:
        fetch_options["before_uid"] = before_uid
    if deadline is not None:
        fetch_options["deadline"] = deadline
    if any(
        value is not None
        for value in (batch_messages, time_budget_seconds, before_uid)
    ):
        fetch_options["progress"] = progress
    messages, scan = fetch_candidate_messages(scoped_config, **fetch_options)
    selection = scan.get("selection") or {
        "mode": "all",
        "sender": None,
        "subject_contains": None,
    }
    selection["subject_product_code"] = binding_code
    groups: dict[str, dict[str, Any]] = {}
    ignored = 0
    library_warnings: list[dict[str, str]] = []
    parsed_messages = 0
    parsed_attachments = 0
    parse_timed_out = False
    last_parsed_uid = None
    subject_binding_messages = 0
    parsed_records = 0
    messages_without_records = 0
    plain_body_messages = 0
    html_body_messages = 0
    attachment_types: dict[str, int] = {}
    first_failure_reason = None
    parse_deadline = (
        time.monotonic() + int(time_budget_seconds)
        if time_budget_seconds is not None
        else None
    )
    for message in messages:
        if parse_deadline is not None and time.monotonic() >= parse_deadline:
            parse_timed_out = True
            break
        parsed_attachments += sum(
            1
            for part in message.walk()
            if part.get_filename() or part.get_content_disposition() == "attachment"
        )
        if any(
            part.get_content_type() == "text/plain"
            for part in message.walk()
        ):
            plain_body_messages += 1
        if any(
            part.get_content_type() == "text/html"
            for part in message.walk()
        ):
            html_body_messages += 1
        for part in message.walk():
            if part.get_filename() or part.get_content_disposition() == "attachment":
                content_type = str(part.get_content_type() or "application/octet-stream")
                attachment_types[content_type] = attachment_types.get(content_type, 0) + 1
        sender = single_from_address(message)
        if not sender:
            ignored += 1
            continue
        try:
            rows = rows_from_message(message, "auto")
        except ParseError:
            rows = []
            first_failure_reason = first_failure_reason or "parse-error"
        finally:
            library_warnings.extend(consume_parse_library_warnings())
        if not rows:
            ignored += 1
            messages_without_records += 1
            first_failure_reason = first_failure_reason or "no-nav-records"
            parsed_messages += 1
            last_parsed_uid = getattr(message, "_nav_source_uid", last_parsed_uid)
            _proposal_progress(
                scope_id,
                "running",
                "parse",
                parsed_messages,
                len(messages),
                attachments_parsed=parsed_attachments,
                sink=progress_sink,
            )
            continue
        if binding_code:
            try:
                rows, binding_used = _bind_subject_product_code(rows, binding_code)
            except ParseError as exc:
                raise RuntimeError(
                    "主题产品代码绑定发现混合代码或同日冲突；拒绝生成候选"
                ) from exc
            if binding_used:
                subject_binding_messages += 1
        parsed_records += len(rows)
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
        parsed_messages += 1
        last_parsed_uid = getattr(message, "_nav_source_uid", last_parsed_uid)
        if parse_deadline is not None:
            _proposal_progress(
                scope_id,
                "running",
                "parse",
                parsed_messages,
                len(messages),
                attachments_parsed=parsed_attachments,
                sink=progress_sink,
            )
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
                "selection": selection,
            }
        )
    warnings: list[str] = []
    if scan["truncated"]:
        warnings.append(
            "邮箱候选扫描受到数量或总大小上限限制；AI 应在需要时缩小时间范围后重试"
        )
    if scan["skipped_oversize"]:
        warnings.append(f"已跳过 {scan['skipped_oversize']} 封超过单封大小上限的邮件")
    warning_summary = _warning_summary(library_warnings)
    if previous:
        warning_summary = _merge_warning_summaries(
            previous.get("parser_library_warnings") or [], warning_summary
        )
    if warning_summary:
        warnings.append(
            f"解析库产生 {sum(item['count'] for item in warning_summary)} 条结构化警告；"
            "请在 parser_library_warnings 中检查类别后再确认解析结果"
        )
    incremental = any(
        value is not None for value in (batch_messages, time_budget_seconds, before_uid)
    )
    if parse_timed_out:
        scan["timed_out"] = True
        scan["range_complete"] = False
        if last_parsed_uid is not None:
            scan["resume_before_uid"] = int(last_parsed_uid)
        else:
            source_uids = [
                int(uid)
                for message in messages
                if (uid := getattr(message, "_nav_source_uid", None)) is not None
            ]
            if source_uids:
                scan["resume_before_uid"] = max(source_uids) + 1
            elif before_uid is not None:
                scan["resume_before_uid"] = before_uid
    previous_scan = (previous or {}).get("scan") or {}
    scan["effective_lookback_days"] = effective_lookback
    scan["messages_fetched_total"] = int(
        previous_scan.get("messages_fetched_total") or 0
    ) + int(scan.get("messages_fetched") or 0)
    scan["messages_parsed_total"] = int(
        previous_scan.get("messages_parsed_total") or 0
    ) + parsed_messages
    scan["subject_code_binding_messages"] = int(
        previous_scan.get("subject_code_binding_messages") or 0
    ) + subject_binding_messages
    scan["parsed_records"] = int(previous_scan.get("parsed_records") or 0) + parsed_records
    scan["attachments_parsed"] = int(
        previous_scan.get("attachments_parsed") or 0
    ) + parsed_attachments
    scan["chunks_completed"] = int(
        previous_scan.get("chunks_completed") or 0
    ) + 1
    if previous:
        candidates = _merge_proposal_candidates(
            previous.get("candidates") or [], candidates
        )
    previous_gap = (previous or {}).get("parse_gap_summary") or {}
    combined_attachment_types = dict(previous_gap.get("attachment_types") or {})
    for content_type, count in attachment_types.items():
        combined_attachment_types[content_type] = int(
            combined_attachment_types.get(content_type) or 0
        ) + count
    headers_matched = int(
        previous_gap.get("headers_matched")
        or scan.get("matching_messages_in_range")
        or scan.get("messages_selected")
        or 0
    )
    parse_gap_summary = {
        "headers_matched": headers_matched,
        "messages_without_records": int(
            previous_gap.get("messages_without_records") or 0
        )
        + messages_without_records,
        "parsed_records": scan["parsed_records"],
        "body_types": {
            "plain_messages": int(
                (previous_gap.get("body_types") or {}).get("plain_messages") or 0
            )
            + plain_body_messages,
            "html_messages": int(
                (previous_gap.get("body_types") or {}).get("html_messages") or 0
            )
            + html_body_messages,
        },
        "attachment_types": dict(sorted(combined_attachment_types.items())),
        "first_failure_reason": previous_gap.get("first_failure_reason")
        or first_failure_reason,
        "local_parser_recommended": False,
    }
    incomplete = incremental and not bool(scan.get("range_complete", True))
    parse_gap_summary["local_parser_recommended"] = bool(
        not incomplete
        and parse_gap_summary["headers_matched"] > 0
        and parse_gap_summary["parsed_records"] == 0
    )
    if parse_gap_summary["local_parser_recommended"]:
        warnings.append(
            "范围内邮件头已匹配但通用解析器未得到净值记录；请检查 parse_gap_summary 并评估本地 parser"
        )
    report = {
        "passed": bool(candidates) and not incomplete,
        "partial": incomplete,
        "resume_available": incomplete and scan.get("resume_before_uid") is not None,
        "scan": {
            **scan,
            "messages_ignored_as_non_nav": int(
                previous_scan.get("messages_ignored_as_non_nav") or 0
            )
            + ignored,
        },
        "selection": selection,
        "candidates": candidates,
        "warnings": warnings,
        "parser_library_warnings": warning_summary,
        "subject_code_binding": (
            {
                "enabled": True,
                "code": binding_code,
                "messages": scan["subject_code_binding_messages"],
                "evidence": "exact-sender-and-subject-code",
            }
            if binding_code
            else {"enabled": False, "messages": 0}
        ),
        "parse_gap_summary": parse_gap_summary,
        "errors": (
            ["候选范围尚未完成；使用 propose --resume 从本地游标继续"]
            if incomplete
            else []
            if candidates
            else ["没有发现可自动解析的净值邮件候选"]
        ),
    }
    if incomplete:
        write_json_atomic(PROPOSAL_PARTIAL, report)
        _proposal_progress(
            scope_id,
            "paused",
            "time-budget" if scan.get("timed_out") else "batch-complete",
            parsed_messages,
            len(messages),
            resume_before_uid=scan.get("resume_before_uid"),
            attachments_parsed=scan["attachments_parsed"],
            sink=progress_sink,
        )
    else:
        write_json_atomic(STATE_ROOT / "route-proposals.json", report)
        PROPOSAL_PARTIAL.unlink(missing_ok=True)
        _proposal_progress(
            scope_id,
            "complete",
            "complete",
            parsed_messages,
            len(messages),
            attachments_parsed=scan["attachments_parsed"],
            sink=progress_sink,
        )
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
            "diagnostics": [],
            "route_overlaps": [],
            "date_conflicts": [],
            "corrections_applied": [],
            "parser_library_warnings": [],
            "ignored_unconfigured_rows": 0,
            "messages_with_ignored_unconfigured_rows": 0,
            "pre_managed_rows_ignored": 0,
            "excluded_non_nav_messages": 0,
            "excluded_non_nav_reasons": {},
            "imap_reconnects": 0,
        }
        write_json_atomic(STATE_ROOT / "route-report.json", report)
        DISCOVERY_SNAPSHOT.unlink(missing_ok=True)
        return {}, report
    messages = fetch_authorized_messages(config)
    mail_scope_fingerprint = getattr(messages, "scope_fingerprint", None)
    excluded_non_nav_messages = int(
        getattr(messages, "excluded_non_nav_messages", 0)
    )
    excluded_non_nav_reasons = dict(
        getattr(messages, "excluded_non_nav_reasons", {})
    )
    imap_reconnects = int(getattr(messages, "reconnect_count", 0))
    sender_routes: dict[str, list[dict[str, Any]]] = {}
    for route in routes_config:
        sender_routes.setdefault(str(route["sender"]).lower(), []).append(route)
    output: dict[str, list[NavRow]] = {
        str(route["sheet"]): [] for route in routes_config
    }
    routed_candidates: dict[str, list[tuple[NavRow, dict[str, Any]]]] = {
        str(route["sheet"]): [] for route in routes_config
    }
    pre_managed_candidates: dict[str, list[tuple[NavRow, dict[str, Any]]]] = {
        str(route["sheet"]): [] for route in routes_config
    }
    route_reports: list[dict[str, Any]] = []
    errors: list[str] = []
    diagnostics: list[dict[str, Any]] = []
    route_overlaps: dict[tuple[str, tuple[str, ...]], dict[str, Any]] = {}
    date_conflicts: list[dict[str, Any]] = []
    corrections_applied: list[dict[str, Any]] = []
    library_warnings: list[dict[str, str]] = []
    ignored_unconfigured_rows = 0
    messages_with_ignored_unconfigured_rows = 0
    pre_managed_rows_ignored = 0
    for sender, routes in sender_routes.items():
        route_message_counts = {str(route["sheet"]): 0 for route in routes}
        route_filtered_counts = {str(route["sheet"]): 0 for route in routes}
        route_binding_counts = {str(route["sheet"]): 0 for route in routes}
        for message in messages.get(sender, []):
            message_evidence = _message_evidence(message, sender)
            applicable = []
            message_subject = decoded(message.get("Subject")).casefold()
            for route in routes:
                subject_filter = (
                    str(route.get("subject_contains") or "").strip().casefold()
                )
                subject_excludes = [
                    str(item).strip().casefold()
                    for item in route.get("subject_excludes") or []
                ]
                if (
                    subject_filter
                    and subject_filter not in message_subject
                ) or any(
                    excluded in message_subject for excluded in subject_excludes
                ):
                    route_filtered_counts[str(route["sheet"])] += 1
                else:
                    applicable.append(route)
            if not applicable:
                continue
            if len(applicable) > 1:
                for left, right in combinations(applicable, 2):
                    _record_route_overlap(
                        route_overlaps,
                        reason="overlapping-message-scope",
                        routes=[left, right],
                        sender=sender,
                        message=message,
                    )
            message_rows: list[NavRow] = []
            parse_failed = False
            for parser_name in sorted(
                {str(route.get("parser", "auto")) for route in applicable}
            ):
                parse_error: ParseError | None = None
                attachment_diagnostics: list[dict[str, str]] = []
                try:
                    parsed = rows_from_message(message, parser_name)
                except ParseError as exc:
                    parsed = []
                    parse_error = exc
                finally:
                    library_warnings.extend(consume_parse_library_warnings())
                    attachment_diagnostics = consume_attachment_parse_diagnostics()
                if not parsed:
                    diagnostics.append(
                        _parse_diagnostic(
                            message,
                            sender,
                            [
                                route
                                for route in applicable
                                if str(route.get("parser", "auto")) == parser_name
                            ],
                            parser_name,
                            parse_error,
                            attachment_diagnostics,
                        )
                    )
                    parse_failed = True
                    break
                if attachment_diagnostics:
                    diagnostics.append(
                        _parse_diagnostic(
                            message,
                            sender,
                            [
                                route
                                for route in applicable
                                if str(route.get("parser", "auto")) == parser_name
                            ],
                            parser_name,
                            None,
                            attachment_diagnostics,
                            message_parsed=True,
                        )
                    )
                message_rows.extend(parsed)
            if parse_failed or not message_rows:
                errors.append("An in-scope authorized message could not be parsed")
                continue
            message_rows = deduplicate(message_rows)
            if any(row.code is None for row in message_rows):
                binding_routes = [
                    route for route in applicable if route.get("subject_product_code")
                ]
                managed_unbound_rows = [
                    row
                    for row in message_rows
                    if row.code is None
                    and any(
                        row.date >= _series_start(route) for route in applicable
                    )
                ]
                eligible_binding_routes = [
                    route
                    for route in binding_routes
                    if managed_unbound_rows
                    and all(
                        [
                            candidate
                            for candidate in applicable
                            if row.date >= _series_start(candidate)
                        ]
                        == [route]
                        for row in managed_unbound_rows
                    )
                ]
                if managed_unbound_rows and binding_routes:
                    if len(eligible_binding_routes) != 1:
                        errors.append(
                            "Subject product code binding did not resolve to exactly one route"
                        )
                        continue
                    try:
                        message_rows, binding_used = _bind_subject_product_code(
                            message_rows,
                            str(
                                eligible_binding_routes[0][
                                    "subject_product_code"
                                ]
                            ),
                        )
                    except ParseError:
                        errors.append(
                            "Subject product code binding conflicted with parsed rows"
                        )
                        continue
                    if binding_used:
                        route_binding_counts[
                            str(eligible_binding_routes[0]["sheet"])
                        ] += 1
            used_sheets: set[str] = set()
            message_ignored_unconfigured_rows = 0
            message_routing_errors = 0
            message_pre_managed_rows = 0
            for row in message_rows:
                date_applicable = [
                    route for route in applicable if row.date >= _series_start(route)
                ]
                pre_managed_applicable = [
                    route for route in applicable if row.date < _series_start(route)
                ]
                pre_managed_matches = _matching_routes_for_row(
                    row, pre_managed_applicable
                )
                if not date_applicable:
                    if len(pre_managed_matches) == 1:
                        pre_managed_candidates[
                            str(pre_managed_matches[0]["sheet"])
                        ].append((row, message_evidence))
                    pre_managed_rows_ignored += 1
                    message_pre_managed_rows += 1
                    continue
                matches = _matching_routes_for_row(row, date_applicable)
                if not matches and row.code is not None:
                    if len(pre_managed_matches) == 1:
                        pre_managed_candidates[
                            str(pre_managed_matches[0]["sheet"])
                        ].append((row, message_evidence))
                        pre_managed_rows_ignored += 1
                        message_pre_managed_rows += 1
                        continue
                    ignored_unconfigured_rows += 1
                    message_ignored_unconfigured_rows += 1
                    continue
                if len(matches) != 1:
                    if len(matches) > 1:
                        for left, right in combinations(matches, 2):
                            _record_route_overlap(
                                route_overlaps,
                                reason="multiple-row-route-matches",
                                routes=[left, right],
                                sender=sender,
                                message=message,
                            )
                    errors.append(
                        "An in-scope NAV row was not routed to exactly one authorized sheet"
                    )
                    message_routing_errors += 1
                    continue
                route = matches[0]
                sheet = str(route["sheet"])
                output[sheet].append(row)
                routed_candidates[sheet].append((row, message_evidence))
                used_sheets.add(sheet)
            if message_ignored_unconfigured_rows:
                messages_with_ignored_unconfigured_rows += 1
            if (
                not used_sheets
                and not message_ignored_unconfigured_rows
                and not message_routing_errors
                and not message_pre_managed_rows
            ):
                errors.append(
                    "An in-scope authorized message produced no routed NAV rows"
                )
            for sheet in used_sheets:
                route_message_counts[sheet] += 1
        for route in routes:
            sheet = str(route["sheet"])
            managed_rows, applied = _apply_explicit_corrections(
                routed_candidates[sheet], route, sender
            )
            corrections_applied.extend(applied)
            date_conflicts.extend(
                _date_conflict_diagnostics(managed_rows, route, sender)
            )
            try:
                selected = choose_route_rows(managed_rows, route, len(routes) == 1)
            except ParseError as exc:
                errors.append(f"{route['sheet']}: {exc}")
                selected = []
            managed_candidate_dates = [row.date.isoformat() for row in selected]
            verification_anchor: NavRow | None = None
            start = _series_start(route)
            if (
                route.get("baseline_overlap") == "last_existing_point"
                and start > dt.date.min
            ):
                anchor_date = start - dt.timedelta(days=1)
                anchor_rows = [
                    row
                    for row, _evidence in pre_managed_candidates[sheet]
                    if row.date == anchor_date
                ]
                if anchor_rows:
                    try:
                        resolved_anchor = choose_route_rows(
                            anchor_rows, route, len(routes) == 1
                        )
                    except ParseError:
                        resolved_anchor = []
                    if len(resolved_anchor) == 1:
                        verification_anchor = resolved_anchor[0]
                        selected = [verification_anchor, *selected]
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
                    "subject_code_binding_messages": route_binding_counts[sheet],
                    "subject_code_binding_evidence": (
                        "exact-sender-and-subject-code"
                        if route_binding_counts[sheet]
                        else None
                    ),
                    "candidate_dates": managed_candidate_dates,
                    "verification_anchor_date": (
                        verification_anchor.date.isoformat()
                        if verification_anchor
                        else None
                    ),
                    "pre_managed_rows_seen": len(pre_managed_candidates[sheet]),
                }
            )
    binding_messages = sum(
        int(item.get("subject_code_binding_messages") or 0)
        for item in route_reports
    )
    if binding_messages:
        warnings.append(
            f"已按精确发件人和主题产品代码证据绑定无代码净值行：{binding_messages} 封邮件"
        )
    if ignored_unconfigured_rows:
        warnings.append(
            "已忽略带有明确代码、但不属于任何活动路由的净值行："
            f"{ignored_unconfigured_rows} 行，涉及 "
            f"{messages_with_ignored_unconfigured_rows} 封邮件"
        )
    if pre_managed_rows_ignored:
        warnings.append(
            "已保留 series_start 之前的邮件记录且未纳入受管冲突或写入："
            f"{pre_managed_rows_ignored} 行"
        )
    if excluded_non_nav_messages:
        warnings.append(
            "已在下载正文前排除主题明确属于非正式净值类别的邮件："
            f"{excluded_non_nav_messages} 封；请检查 excluded_non_nav_reasons"
        )
    if corrections_applied:
        warnings.append(
            "已依据后发且主题明确标注的更正通知替代原始同日值："
            f"{len(corrections_applied)} 个净值日；请检查 corrections_applied"
        )
    warning_summary = _warning_summary(library_warnings)
    if warning_summary:
        warnings.append(
            f"解析库产生 {sum(item['count'] for item in warning_summary)} 条结构化警告；"
            "正式批准前应检查 parser_library_warnings"
        )
    report = {
        "passed": not errors,
        "routes": route_reports,
        "warnings": warnings,
        "errors": errors,
        "diagnostics": diagnostics,
        "route_overlaps": list(route_overlaps.values()),
        "date_conflicts": date_conflicts,
        "corrections_applied": corrections_applied,
        "parser_library_warnings": warning_summary,
        "ignored_unconfigured_rows": ignored_unconfigured_rows,
        "messages_with_ignored_unconfigured_rows": (
            messages_with_ignored_unconfigured_rows
        ),
        "pre_managed_rows_ignored": pre_managed_rows_ignored,
        "excluded_non_nav_messages": excluded_non_nav_messages,
        "excluded_non_nav_reasons": excluded_non_nav_reasons,
        "imap_reconnects": imap_reconnects,
    }
    write_json_atomic(STATE_ROOT / "route-report.json", report)
    _save_discovery_snapshot(
        config, output, report, mail_scope_fingerprint
    )
    return output, report


def doctor(config: dict[str, Any]) -> dict[str, Any]:
    checks: list[dict[str, Any]] = []
    detected_spreadsheets = detected_spreadsheet_apps()

    def add(name: str, passed: bool, detail: str) -> None:
        checks.append({"name": name, "passed": passed, "detail": detail})

    supported_python = (3, 11) <= sys.version_info[:2] <= (3, 14)
    add("python", supported_python, platform.python_version())
    workbook = Path(config["workbook_path"])
    workbook_exists = workbook.is_file()
    workbook_mode = str(config.get("workbook_mode", "existing"))
    add("workbook", workbook_exists, "present" if workbook_exists else "missing")
    template_target_ready = (
        workbook_mode == "bundled-template"
        and workbook.suffix.lower() == ".xlsx"
        and workbook.parent.is_dir()
    )
    if workbook_mode == "bundled-template":
        add(
            "workbook-template-target",
            template_target_ready,
            "initialized" if workbook_exists else "ready for workbook init-template",
        )
    configured_routes = config.get("routes") or []
    active = active_routes(config)
    add(
        "routes",
        bool(active),
        f"{len(active)} active; {len(configured_routes) - len(active)} paused",
    )
    unresolved_reviews = [
        {
            "sheet": str(route["sheet"]),
            "issue": str(benchmark_review_issue(route)),
        }
        for route in active
        if benchmark_requires_review(route)
    ]
    add(
        "write-rules-resolved",
        not unresolved_reviews,
        (
            "resolved"
            if not unresolved_reviews
            else (
                f"{len(unresolved_reviews)} sheet(s) require benchmark review: "
                + ", ".join(
                    f"{item['sheet']}={item['issue']}" for item in unresolved_reviews
                )
            )
        ),
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
    base_names = {"python", *{f"package:{name}" for name in expected_packages}}
    base_names.add(
        "workbook-template-target"
        if workbook_mode == "bundled-template"
        else "workbook"
    )
    bootstrap_ready = all(by_name.get(name, False) for name in base_names)
    mail_discovery_ready = bootstrap_ready and by_name["imap-secret"]
    preview_ready = (
        bootstrap_ready
        and workbook_exists
        and by_name["routes"]
        and by_name["imap-secret"]
    )
    commit_ready = (
        preview_ready
        and by_name["spreadsheet-com"]
        and by_name["write-rules-resolved"]
    )
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


def _config_for_sheets(
    config: dict[str, Any], sheets: list[str] | None, command: str
) -> tuple[dict[str, Any], list[str]]:
    if not sheets:
        return config, []
    requested = list(dict.fromkeys(str(sheet).strip() for sheet in sheets))
    if any(not sheet for sheet in requested):
        raise ValueError(f"{command} --sheet 不能为空")
    by_sheet = {str(route["sheet"]): route for route in active_routes(config)}
    unknown = [sheet for sheet in requested if sheet not in by_sheet]
    if unknown:
        raise ValueError(
            f"{command} --sheet 只能选择活动路由；未找到：" + "、".join(unknown)
        )
    scoped_config = deepcopy(config)
    scoped_config["routes"] = [deepcopy(by_sheet[sheet]) for sheet in requested]
    return scoped_config, requested


def discover(
    config: dict[str, Any], sheets: list[str] | None = None
) -> tuple[dict[str, list[NavRow]], dict[str, Any]]:
    scoped_config, scope_sheets = _config_for_sheets(config, sheets, "discover")
    rows, report = collect_route_rows(scoped_config)
    report["scoped"] = bool(scope_sheets)
    report["scope_sheets"] = scope_sheets
    report["scope_route_count"] = len(active_routes(scoped_config))
    report["final_full_preview_required"] = bool(scope_sheets)
    write_json_atomic(STATE_ROOT / "route-report.json", report)
    return rows, report


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
    config: dict[str, Any],
    rows: dict[str, list[NavRow]] | None = None,
    *,
    mail_config: dict[str, Any] | None = None,
    sheets: list[str] | None = None,
    reuse_discovery: bool = False,
) -> dict[str, Any]:
    # A failed preview must never leave an older commit plan looking current.
    (STATE_ROOT / "plan.json").unlink(missing_ok=True)
    (STATE_ROOT / "concurrency-report.json").unlink(missing_ok=True)
    preview_config, scope_sheets = _config_for_sheets(config, sheets, "preview")
    effective_mail_config = mail_config or config
    if scope_sheets:
        effective_mail_config, _ = _config_for_sheets(
            effective_mail_config, scope_sheets, "preview"
        )
    if reuse_discovery and scope_sheets:
        raise ValueError("局部 preview 不能复用完整发现快照")
    warnings: list[str] = []
    discovery_reused = False
    if rows is None:
        if reuse_discovery:
            rows, discovery = reuse_discovery_snapshot(
                effective_mail_config
            )
            discovery_reused = True
        else:
            rows, discovery = collect_route_rows(effective_mail_config)
        if not discovery["passed"]:
            raise RuntimeError("Discovery failed; inspect route-report.json")
        warnings.extend(discovery.get("warnings") or [])
    validation = validate_history(preview_config, rows)
    if not validation["passed"]:
        raise RuntimeError(
            "Historical validation failed; inspect validation-report.json"
        )
    warnings.extend(validation.get("warnings") or [])
    plan = build_preview(
        preview_config,
        rows,
        warnings,
        diagnostic_only=bool(scope_sheets),
    )
    plan["scoped"] = bool(scope_sheets)
    plan["scope_sheets"] = scope_sheets
    plan["final_full_preview_required"] = bool(scope_sheets)
    plan["discovery_reused"] = discovery_reused
    if scope_sheets:
        write_json_atomic(STATE_ROOT / "plan.json", plan)
    return plan
