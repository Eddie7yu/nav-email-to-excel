from __future__ import annotations

import datetime as dt
import json
import os
import re
import shutil
import stat
import subprocess
import sys
import uuid
from pathlib import Path
from typing import Any

from nav_config import STATE_ROOT

DEMO_ROOT = STATE_ROOT / "demo-runs"
RUN_ID_PATTERN = re.compile(r"demo-\d{8}-\d{6}-[0-9a-f]{8}")
FIXTURE_VALUE = "-".join(("fixture", "only"))
DEMO_FILES = (
    "nav_commit.py",
    "nav_config.py",
    "nav_demo_worker.py",
    "nav_mail.py",
    "nav_parse.py",
    "nav_service.py",
    "nav_workbook.py",
    "runtime_secret.py",
)


def _navctl_command(arguments: str) -> str:
    python = r".\.venv\Scripts\python.exe" if os.name == "nt" else ".venv/bin/python"
    return f"{python} navctl.py {arguments}"


def _run_path(run_id: str) -> Path:
    if not RUN_ID_PATTERN.fullmatch(run_id):
        raise RuntimeError("演练 ID 格式无效")
    path = (DEMO_ROOT / run_id).resolve()
    if path.parent != DEMO_ROOT.resolve():
        raise RuntimeError("演练目录不在允许范围内")
    return path


def _remove_tree(path: Path, *, ignore_errors: bool = False) -> None:
    if path.is_dir():
        for item in path.rglob("*"):
            if item.is_file():
                try:
                    item.chmod(item.stat().st_mode | stat.S_IWRITE)
                except OSError:
                    if not ignore_errors:
                        raise
    shutil.rmtree(path, ignore_errors=ignore_errors)


def _run_worker(path: Path, action: str) -> dict[str, Any]:
    environment = dict(
        os.environ,
        PYTHONUTF8="1",
        NAV_EMAIL_PASSWORD=FIXTURE_VALUE,
    )
    result = subprocess.run(
        [sys.executable, "-X", "utf8", "nav_demo_worker.py", action],
        cwd=path,
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=environment,
    )
    try:
        payload = json.loads(result.stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError("离线演练没有返回有效结果") from exc
    if result.returncode or not payload.get("passed"):
        detail = str(payload.get("error") or result.stderr.strip() or "未知错误")
        raise RuntimeError(f"离线演练失败：{detail}")
    return payload


def prepare() -> dict[str, Any]:
    DEMO_ROOT.mkdir(exist_ok=True)
    run_id = f"demo-{dt.datetime.now():%Y%m%d-%H%M%S}-{uuid.uuid4().hex[:8]}"
    path = _run_path(run_id)
    path.mkdir()
    try:
        for name in DEMO_FILES:
            source = STATE_ROOT / name
            if not source.is_file():
                raise RuntimeError(f"演练所需文件缺失：{name}")
            shutil.copy2(source, path / name)
        payload = _run_worker(path, "prepare")
    except Exception:
        _remove_tree(path, ignore_errors=True)
        raise
    payload.update(
        {
            "run_id": run_id,
            "demo_directory": str(path),
            "next_command": _navctl_command(
                f"demo commit --run-id {run_id} --yes-reviewed-preview"
            ),
        }
    )
    return payload


def commit(run_id: str) -> dict[str, Any]:
    path = _run_path(run_id)
    if not path.is_dir() or not (path / "demo-state.json").is_file():
        raise RuntimeError("找不到该演练，或演练尚未完成预览")
    payload = _run_worker(path, "commit")
    payload.update(
        {
            "run_id": run_id,
            "demo_directory": str(path),
            "next_command": _navctl_command(f"demo remove --run-id {run_id}"),
        }
    )
    return payload


def remove(run_id: str) -> dict[str, Any]:
    path = _run_path(run_id)
    existed = path.is_dir()
    if existed:
        _remove_tree(path)
    try:
        DEMO_ROOT.rmdir()
    except OSError:
        pass
    return {
        "passed": True,
        "demo": True,
        "run_id": run_id,
        "removed": existed,
        "real_data_used": False,
    }


def list_runs() -> dict[str, Any]:
    runs: list[dict[str, Any]] = []
    if DEMO_ROOT.is_dir():
        for path in sorted(DEMO_ROOT.iterdir()):
            if not path.is_dir() or not RUN_ID_PATTERN.fullmatch(path.name):
                continue
            state_path = path / "demo-state.json"
            state: dict[str, Any] = {}
            try:
                state = json.loads(state_path.read_text(encoding="utf-8"))
            except (FileNotFoundError, json.JSONDecodeError):
                pass
            runs.append(
                {
                    "run_id": path.name,
                    "prepared": bool(state),
                    "committed": bool(state.get("committed")),
                }
            )
    return {
        "passed": True,
        "demo": True,
        "runs": runs,
        "real_data_used": False,
    }
