from __future__ import annotations

import argparse
import datetime as dt
import json
import os
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterator

from nav_commit import commit
from nav_config import CONFIG_PATH, ConfigError, ROOT, load_config
from nav_demo import commit as commit_demo
from nav_demo import list_runs as list_demo_runs
from nav_demo import prepare as prepare_demo
from nav_demo import remove as remove_demo
from nav_schedule import install as install_schedule
from nav_schedule import remove as remove_schedule
from nav_schedule import status as schedule_status
from nav_service import discover, doctor, preview, validate
from runtime_secret import read_password, remove_password, set_password


def emit(payload: Any) -> None:
    print(json.dumps(payload, ensure_ascii=False, indent=2))


def prune_logs(config: dict[str, Any]) -> None:
    directory = ROOT / "logs"
    if not directory.is_dir():
        return
    cutoff = (
        dt.datetime.now().timestamp()
        - int((config.get("retention") or {}).get("log_days", 30)) * 86400
    )
    for path in directory.glob("*.log"):
        if path.is_file() and path.stat().st_mtime < cutoff:
            path.unlink()


@contextmanager
def run_lock() -> Iterator[None]:
    path = ROOT / "run.lock"
    descriptor = None
    if path.exists():
        try:
            state = json.loads(path.read_text(encoding="utf-8"))
            pid = int(state["pid"])
            created = dt.datetime.fromisoformat(state["created"])
        except (FileNotFoundError, KeyError, ValueError, json.JSONDecodeError):
            age = (
                dt.datetime.now() - created
                if "created" in locals()
                else dt.timedelta(days=2)
            )
            if age > dt.timedelta(hours=24):
                path.unlink(missing_ok=True)
        else:
            try:
                os.kill(pid, 0)
            except OSError:
                path.unlink(missing_ok=True)
    try:
        descriptor = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        payload = json.dumps(
            {
                "pid": os.getpid(),
                "created": dt.datetime.now().isoformat(timespec="seconds"),
            }
        )
        os.write(descriptor, payload.encode("utf-8"))
        os.close(descriptor)
        descriptor = None
    except FileExistsError as exc:
        raise RuntimeError(
            "Another runtime process is active; refusing a concurrent run"
        ) from exc
    try:
        yield
    finally:
        if descriptor is not None:
            os.close(descriptor)
        if path.exists():
            path.unlink()


def command_doctor(config: dict[str, Any], _args: argparse.Namespace) -> int:
    report = doctor(config)
    emit(report)
    return 0 if report["passed"] else 2


def command_secret(config: dict[str, Any], args: argparse.Namespace) -> int:
    if args.secret_action == "set":
        path = set_password(str(config["runtime_id"]))
        emit({"stored": True, "location": str(path.parent), "value_echoed": False})
        return 0
    if args.secret_action == "remove":
        emit(
            {
                "removed": remove_password(str(config["runtime_id"])),
                "value_echoed": False,
            }
        )
        return 0
    available = bool(read_password(str(config["runtime_id"])))
    emit(
        {
            "available": available,
            "source": "environment-or-local-secure-store" if available else "none",
        }
    )
    return 0 if available else 2


def command_discover(config: dict[str, Any], _args: argparse.Namespace) -> int:
    with run_lock():
        _, report = discover(config)
    emit(report)
    return 0 if report["passed"] else 2


def command_validate(config: dict[str, Any], _args: argparse.Namespace) -> int:
    with run_lock():
        report = validate(config)
    emit(report)
    return 0 if report["passed"] else 2


def command_preview(config: dict[str, Any], _args: argparse.Namespace) -> int:
    with run_lock():
        plan = preview(config)
    emit(
        {
            "preview_path": plan["preview_path"],
            "sheets": [
                {"sheet": item["sheet"], "new_dates": item["new_dates"]}
                for item in plan["sheets"]
            ],
            "master_unchanged": True,
        }
    )
    return 0


def command_commit(config: dict[str, Any], args: argparse.Namespace) -> int:
    if not args.yes_reviewed_preview:
        raise RuntimeError("Refusing commit without --yes-reviewed-preview")
    with run_lock():
        emit(commit(config))
    return 0


def command_schedule(config: dict[str, Any], args: argparse.Namespace) -> int:
    if args.schedule_action == "install":
        emit(install_schedule(config))
    elif args.schedule_action == "remove":
        emit(remove_schedule())
    else:
        emit(schedule_status())
    return 0


def command_demo(args: argparse.Namespace) -> int:
    if args.demo_action == "prepare":
        emit(prepare_demo())
    elif args.demo_action == "commit":
        if not args.yes_reviewed_preview:
            raise RuntimeError(
                "拒绝演练写入；检查虚构预览后添加 --yes-reviewed-preview"
            )
        emit(commit_demo(args.run_id))
    elif args.demo_action == "remove":
        emit(remove_demo(args.run_id))
    else:
        emit(list_demo_runs())
    return 0


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(
        description="Safe local IMAP-to-Excel NAV automation"
    )
    result.add_argument(
        "--config", default=str(CONFIG_PATH), help="Runtime config.json path"
    )
    commands = result.add_subparsers(dest="command", required=True)
    commands.add_parser("doctor")
    secret = commands.add_parser("secret")
    secret.add_argument("secret_action", choices=("set", "status", "remove"))
    commands.add_parser("discover")
    commands.add_parser("validate")
    commands.add_parser("preview")
    commit_parser = commands.add_parser("commit")
    commit_parser.add_argument("--yes-reviewed-preview", action="store_true")
    schedule = commands.add_parser("schedule")
    schedule.add_argument("schedule_action", choices=("install", "remove", "status"))
    demo = commands.add_parser(
        "demo", help="使用虚构邮箱和工作簿进行完全离线的安全演练"
    )
    demo_actions = demo.add_subparsers(dest="demo_action", required=True)
    demo_actions.add_parser("prepare", help="生成虚构预览并停下来等待检查")
    demo_actions.add_parser("list", help="列出本运行目录中的离线演练")
    demo_commit = demo_actions.add_parser(
        "commit", help="对已检查的虚构预览执行 Excel/WPS COM 演练"
    )
    demo_commit.add_argument("--run-id", required=True)
    demo_commit.add_argument("--yes-reviewed-preview", action="store_true")
    demo_remove = demo_actions.add_parser("remove", help="删除指定的虚构演练")
    demo_remove.add_argument("--run-id", required=True)
    return result


def main() -> int:
    args = parser().parse_args()
    try:
        if args.command == "demo":
            return command_demo(args)
        config = load_config(Path(args.config).resolve())
        prune_logs(config)
        commands = {
            "doctor": command_doctor,
            "secret": command_secret,
            "discover": command_discover,
            "validate": command_validate,
            "preview": command_preview,
            "commit": command_commit,
            "schedule": command_schedule,
        }
        return commands[args.command](config, args)
    except (ConfigError, RuntimeError, ValueError) as exc:
        emit({"passed": False, "error": str(exc)})
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
