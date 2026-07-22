from __future__ import annotations

import datetime as dt
import json
import os
import subprocess
from typing import Any

from nav_automation import require_approved
from nav_config import ROOT, STATE_ROOT, write_json_atomic


STATE = STATE_ROOT / "scheduled_tasks.json"
LAST_RUN = STATE_ROOT / "last-scheduled-run.json"


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
    require_approved(config)
    schedules = config.get("schedule") or []
    if not schedules:
        raise ScheduleError("config.schedule is empty")
    previous = [str(name) for name in _state().get("tasks") or []]
    wrapper = STATE_ROOT / "run-update.cmd"
    if not wrapper.is_file():
        raise ScheduleError("run-update.cmd is missing")
    generation = dt.datetime.now().strftime("%Y%m%d%H%M%S")
    prefix = f"NAV-AUTO-{str(config['runtime_id'])[:8]}-{generation}"
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
            "New automatic-update tasks were created, but old tasks could not all be removed; combined state was preserved"
        )
    write_json_atomic(STATE, {"tasks": created, "mode": "automatic-update"})
    return {
        "created": created,
        "mode": "automatic-update-with-backup",
        "requires_logged_in_session": True,
    }


def status() -> dict[str, Any]:
    names = [str(name) for name in _state().get("tasks") or []]
    timezone = _local_timezone()
    if os.name != "nt":
        return {
            "tasks": names,
            "available": False,
            "last_run": _last_run(),
            "mode": _state().get("mode"),
            "time_basis": "host-local wall-clock time",
            "local_timezone": timezone,
        }
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
                        "last_run_time_local": _task_time_detail(
                            task.LastRunTime, timezone
                        ),
                        "next_run_time_local": _task_time_detail(
                            task.NextRunTime, timezone
                        ),
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
    return {
        "tasks": results,
        "available": True,
        "last_run": _last_run(),
        "mode": _state().get("mode"),
        "time_basis": "Windows Task Scheduler local wall-clock time",
        "local_timezone": timezone,
    }


def _task_time(value: Any) -> str | None:
    try:
        parts = _task_time_parts(value)
        if parts is None:
            return None
        return dt.datetime(*parts).isoformat(timespec="seconds")
    except (AttributeError, TypeError, ValueError):
        return None


def _task_time_parts(value: Any) -> tuple[int, int, int, int, int, int] | None:
    try:
        parts = (
            int(value.year),
            int(value.month),
            int(value.day),
            int(value.hour),
            int(value.minute),
            int(value.second),
        )
    except (AttributeError, TypeError, ValueError):
        return None
    return None if parts[0] <= 1900 else parts


def _utc_offset(value: dt.timedelta | None) -> str | None:
    if value is None:
        return None
    total_minutes = int(value.total_seconds() // 60)
    sign = "+" if total_minutes >= 0 else "-"
    total_minutes = abs(total_minutes)
    hours, minutes = divmod(total_minutes, 60)
    return f"{sign}{hours:02d}:{minutes:02d}"


def _windows_timezone_id() -> str | None:
    if os.name != "nt":
        return None
    try:
        import winreg

        path = r"SYSTEM\CurrentControlSet\Control\TimeZoneInformation"
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
            for value_name in ("TimeZoneKeyName", "StandardName"):
                try:
                    value, _ = winreg.QueryValueEx(key, value_name)
                except OSError:
                    continue
                identifier = str(value).rstrip("\x00").strip()
                if identifier:
                    return identifier
    except (ImportError, OSError):
        pass
    return None


def _local_timezone() -> dict[str, str | None]:
    now = dt.datetime.now().astimezone()
    windows_id = _windows_timezone_id()
    identifier = windows_id or str(now.tzinfo or "local")
    offset = _utc_offset(now.utcoffset())
    display = identifier if offset is None else f"{identifier} (UTC{offset})"
    return {
        "id": identifier,
        "windows_id": windows_id,
        "current_utc_offset": offset,
        "display": display,
    }


def _task_time_detail(
    value: Any, timezone: dict[str, str | None]
) -> dict[str, str | None] | None:
    parts = _task_time_parts(value)
    if parts is None:
        return None
    local_value = dt.datetime(*parts)
    offset = _utc_offset(local_value.astimezone().utcoffset())
    timezone_id = timezone.get("id")
    display_zone = timezone_id if offset is None else f"{timezone_id} (UTC{offset})"
    return {
        "datetime": local_value.isoformat(timespec="seconds"),
        "timezone_id": timezone_id,
        "utc_offset": offset,
        "display": f"{local_value:%Y-%m-%d %H:%M:%S} {display_zone}",
        "source": "Windows Task Scheduler local wall-clock time",
    }
