from __future__ import annotations

import datetime as dt
import json
import os
import subprocess
from typing import Any

from nav_config import ROOT, write_json_atomic


STATE = ROOT / "scheduled_tasks.json"
LAST_RUN = ROOT / "last-scheduled-run.json"


class ScheduleError(RuntimeError):
    pass


def _state() -> dict[str, Any]:
    try:
        return json.loads(STATE.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return {"tasks": []}


def record_scheduled_run(payload: dict[str, Any]) -> None:
    write_json_atomic(LAST_RUN, payload)


def _last_run() -> dict[str, Any] | None:
    try:
        value = json.loads(LAST_RUN.read_text(encoding="utf-8"))
    except (FileNotFoundError, json.JSONDecodeError):
        return None
    return value if isinstance(value, dict) else None


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
        return {"tasks": names, "available": False, "last_run": _last_run()}
    service = None
    folder = None
    try:
        from win32com.client import Dispatch

        service = Dispatch("Schedule.Service")
        service.Connect()
        folder = service.GetFolder("\\")
    except Exception:
        service = None
        folder = None
    state_names = {
        0: "unknown",
        1: "disabled",
        2: "queued",
        3: "ready",
        4: "running",
    }
    results: list[dict[str, Any]] = []
    for name in names:
        if folder is not None:
            try:
                task = folder.GetTask(name)
                state = int(task.State)
                results.append(
                    {
                        "name": name,
                        "exists": True,
                        "state": state_names.get(state, "unknown"),
                        "last_run_time": _task_time(task.LastRunTime),
                        "next_run_time": _task_time(task.NextRunTime),
                        "last_result": int(task.LastTaskResult),
                    }
                )
                continue
            except Exception:
                pass
        result = subprocess.run(
            ["schtasks", "/Query", "/TN", name],
            capture_output=True,
            text=True,
        )
        results.append({"name": name, "exists": result.returncode == 0})
    return {"tasks": results, "available": True, "last_run": _last_run()}


def _task_time(value: Any) -> str | None:
    try:
        if int(value.year) <= 1900:
            return None
        try:
            return value.isoformat(timespec="seconds")
        except TypeError:
            return value.isoformat()
    except (AttributeError, TypeError, ValueError):
        return None
