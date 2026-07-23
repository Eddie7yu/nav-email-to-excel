from __future__ import annotations

import datetime as dt
import hashlib
import json
from pathlib import Path
from typing import Any

from nav_config import ROOT, STATE_ROOT, write_json_atomic
from nav_workbook import make_file_writable


APPROVAL = STATE_ROOT / "automation-approval.json"
APPROVED_CONFIG_FIELDS = (
    "schema_version",
    "runtime_id",
    "workbook_path",
    "imap",
    "routes",
    "column_overrides",
    "style",
    "validation",
)


class AutomationError(RuntimeError):
    pass


def approval_fingerprint(config: dict[str, Any]) -> str:
    payload = {field: config.get(field) for field in APPROVED_CONFIG_FIELDS}
    encoded = json.dumps(
        payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def _read() -> dict[str, Any] | None:
    try:
        value = json.loads(APPROVAL.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return None
    return value if isinstance(value, dict) else None


def status(config: dict[str, Any]) -> dict[str, Any]:
    state = _read()
    approved = bool(
        state
        and state.get("schema_version") == 1
        and state.get("runtime_id") == config.get("runtime_id")
        and state.get("config_fingerprint") == approval_fingerprint(config)
    )
    return {
        "approved": approved,
        "approved_at": state.get("approved_at") if approved and state else None,
        "reason": (
            None if approved else "首次预览尚未批准，或写表配置在批准后发生了变化"
        ),
    }


def approve(config: dict[str, Any]) -> dict[str, Any]:
    state = {
        "schema_version": 1,
        "runtime_id": str(config["runtime_id"]),
        "config_fingerprint": approval_fingerprint(config),
        "approved_at": dt.datetime.now().isoformat(timespec="seconds"),
    }
    write_json_atomic(APPROVAL, state)
    return status(config)


def require_approved(config: dict[str, Any]) -> None:
    current = status(config)
    if not current["approved"]:
        raise AutomationError(
            "自动更新尚未启用：请先检查第一次预览或零新增基线报告并完成一次人工批准；"
            "如果配置刚被修改，请重新预览和批准一次"
        )


def revoke() -> dict[str, Any]:
    existed = APPROVAL.exists()
    APPROVAL.unlink(missing_ok=True)
    return {"revoked": existed}


def discard_staging(plan: dict[str, Any] | None) -> None:
    if plan:
        for field in ("preview_path", "review_path"):
            value = plan.get(field)
            if not value:
                continue
            preview = Path(str(value)).resolve()
            if preview.parent == (ROOT / "previews").resolve():
                make_file_writable(preview)
                preview.unlink(missing_ok=True)
    (STATE_ROOT / "plan.json").unlink(missing_ok=True)


def automatic_update(
    config: dict[str, Any], rows: dict[str, list[Any]] | None = None
) -> dict[str, Any]:
    require_approved(config)
    plan: dict[str, Any] | None = None
    try:
        from nav_commit import commit
        from nav_service import preview

        plan = preview(config, rows)
        if not plan["sheets"]:
            return {
                "changed": False,
                "rows": 0,
                "sheets": 0,
                "warnings": plan.get("warnings") or [],
            }
        result = commit(config)
        result["warnings"] = plan.get("warnings") or []
        return result
    finally:
        discard_staging(plan)
