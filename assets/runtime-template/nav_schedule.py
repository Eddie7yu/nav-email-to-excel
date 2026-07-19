from __future__ import annotations

import json
import os
import subprocess
import datetime as dt
from typing import Any

from nav_config import ROOT, write_json_atomic


STATE = ROOT / "scheduled_tasks.json"


class ScheduleError(RuntimeError):
    pass


def _state() -> dict[str, Any]:
    try:
        return json.loads(STATE.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return {"tasks": []}


def _delete(name: str) -> None:
    result = subprocess.run(
        ["schtasks", "/Delete", "/TN", name, "/F"], capture_output=True, text=True
    )
    if result.returncode:
        query = subprocess.run(
            ["schtasks", "/Query", "/TN", name], capture_output=True, text=True
        )
        if query.returncode == 0:
            raise ScheduleError(f"Could not remove recorded task: {name}")


def remove() -> dict[str, Any]:
    if os.name != "nt":
        raise ScheduleError("Windows Task Scheduler is required")
    names = [str(name) for name in _state().get("tasks") or []]
    remaining: list[str] = []
    for name in names:
        try:
            _delete(name)
        except ScheduleError:
            remaining.append(name)
    if remaining:
        write_json_atomic(STATE, {"tasks": remaining})
        raise ScheduleError(
            "Some recorded tasks could not be removed; state was preserved"
        )
    STATE.unlink(missing_ok=True)
    return {"removed": names}


def install(config: dict[str, Any]) -> dict[str, Any]:
    if os.name != "nt":
        raise ScheduleError("Windows Task Scheduler is required")
    if str(ROOT.resolve()).startswith("\\\\"):
        raise ScheduleError("Scheduled runtimes must use a local path, not UNC")
    schedules = config.get("schedule") or []
    if not schedules:
        raise ScheduleError("config.schedule is empty")
    previous = [str(name) for name in _state().get("tasks") or []]
    wrapper = ROOT / "run-preview.cmd"
    if not wrapper.is_file():
        raise ScheduleError("run-preview.cmd is missing")
    generation = dt.datetime.now().strftime("%Y%m%d%H%M%S")
    prefix = f"NAV-{str(config['runtime_id'])[:8]}-{generation}"
    created: list[str] = []
    try:
        for index, item in enumerate(schedules, 1):
            name = f"{prefix}-{index:02d}"
            days = ",".join(str(day).upper() for day in item["days"])
            command = [
                "schtasks",
                "/Create",
                "/TN",
                name,
                "/TR",
                f'"{wrapper.resolve()}"',
                "/SC",
                "WEEKLY",
                "/D",
                days,
                "/ST",
                str(item["time"]),
                "/RL",
                "LIMITED",
                "/IT",
                "/F",
            ]
            result = subprocess.run(command, capture_output=True, text=True)
            if result.returncode:
                raise ScheduleError(f"Failed to create schedule item #{index}")
            verify = subprocess.run(
                ["schtasks", "/Query", "/TN", name], capture_output=True, text=True
            )
            if verify.returncode:
                raise ScheduleError(f"Could not verify schedule item #{index}")
            created.append(name)
    except Exception:
        for name in created:
            _delete(name)
        raise
    try:
        for name in previous:
            _delete(name)
    except ScheduleError:
        write_json_atomic(STATE, {"tasks": list(dict.fromkeys(previous + created))})
        raise ScheduleError(
            "New preview tasks were created, but old tasks could not all be removed; combined state was preserved"
        )
    write_json_atomic(STATE, {"tasks": created})
    return {
        "created": created,
        "mode": "preview-only",
        "requires_logged_in_session": True,
    }


def status() -> dict[str, Any]:
    names = [str(name) for name in _state().get("tasks") or []]
    if os.name != "nt":
        return {"tasks": names, "available": False}
    results = []
    for name in names:
        result = subprocess.run(
            ["schtasks", "/Query", "/TN", name, "/FO", "LIST"],
            capture_output=True,
            text=True,
        )
        results.append({"name": name, "exists": result.returncode == 0})
    return {"tasks": results, "available": True}
