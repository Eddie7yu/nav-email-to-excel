from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import sys
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterator

from nav_automation import (
    approve as approve_automation,
    automatic_update,
    revoke as revoke_automation,
    status as automation_status,
)
from nav_commit import commit
from nav_config import CONFIG_PATH, STATE_ROOT, ConfigError, ROOT, load_config
from nav_demo import commit as commit_demo
from nav_demo import list_runs as list_demo_runs
from nav_demo import prepare as prepare_demo
from nav_demo import remove as remove_demo
from nav_products import add as add_product
from nav_products import adopt as adopt_product
from nav_products import clone as clone_product
from nav_products import pause as pause_product
from nav_products import resume as resume_product
from nav_products import status as product_status
from nav_products import sync as sync_products
from nav_schedule import install as install_schedule
from nav_schedule import record_scheduled_run
from nav_schedule import remove as remove_schedule
from nav_schedule import status as schedule_status
from nav_service import discover, doctor, preview, propose_routes, validate
from nav_template import init_template
from runtime_secret import (
    SecretInputCancelled,
    launch_secret_prompt,
    read_password,
    remove_password,
    set_password,
)


def emit(payload: Any) -> None:
    print(json.dumps(payload, ensure_ascii=False, indent=2))


def configure_output() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure:
            reconfigure(errors="backslashreplace")


def record_scheduled_run_safely(payload: dict[str, Any]) -> None:
    try:
        record_scheduled_run(payload)
    except OSError:
        pass


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
    path = STATE_ROOT / "run.lock"
    descriptor = _acquire_runtime_lock(path)
    try:
        _write_lock_state(
            descriptor,
            {
                "status": "active",
                "pid": os.getpid(),
                "created": dt.datetime.now().isoformat(timespec="seconds"),
            },
        )
        yield
    finally:
        try:
            _write_lock_state(
                descriptor,
                {
                    "status": "idle",
                    "pid": os.getpid(),
                    "released": dt.datetime.now().isoformat(timespec="seconds"),
                },
            )
        except OSError:
            pass
        os.close(descriptor)


def _write_lock_state(descriptor: int, state: dict[str, Any]) -> None:
    payload = (json.dumps(state, ensure_ascii=False) + "\n").encode("utf-8")
    os.lseek(descriptor, 0, os.SEEK_SET)
    os.ftruncate(descriptor, 0)
    os.write(descriptor, payload)
    os.fsync(descriptor)


def _acquire_runtime_lock(path: Path) -> int:
    if os.name == "nt":
        import ctypes
        import msvcrt
        from ctypes import wintypes

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.CreateFileW.argtypes = (
            wintypes.LPCWSTR,
            wintypes.DWORD,
            wintypes.DWORD,
            wintypes.LPVOID,
            wintypes.DWORD,
            wintypes.DWORD,
            wintypes.HANDLE,
        )
        kernel32.CreateFileW.restype = wintypes.HANDLE
        kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
        kernel32.CloseHandle.restype = wintypes.BOOL
        handle = kernel32.CreateFileW(
            str(path),
            0x80000000 | 0x40000000,
            0,
            None,
            4,
            0x80,
            None,
        )
        invalid = ctypes.c_void_p(-1).value
        if handle == invalid:
            error = ctypes.get_last_error()
            if error in {32, 33}:
                raise RuntimeError(
                    "Another runtime process is active; refusing a concurrent run"
                )
            raise ctypes.WinError(error)
        try:
            return msvcrt.open_osfhandle(int(handle), os.O_RDWR)
        except Exception:
            kernel32.CloseHandle(handle)
            raise

    import fcntl

    descriptor = os.open(path, os.O_CREAT | os.O_RDWR, 0o600)
    try:
        fcntl.flock(descriptor, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except BlockingIOError as exc:
        os.close(descriptor)
        raise RuntimeError(
            "Another runtime process is active; refusing a concurrent run"
        ) from exc
    return descriptor


def command_doctor(config: dict[str, Any], _args: argparse.Namespace) -> int:
    report = doctor(config)
    emit(report)
    return 0 if report["passed"] else 2


def command_secret(config: dict[str, Any], args: argparse.Namespace) -> int:
    if args.secret_action == "launch":
        process_id = launch_secret_prompt()
        emit(
            {
                "launched": True,
                "process_id": process_id,
                "next": "请在新窗口粘贴授权码并回车；随后必须运行 secret status 确认保存成功",
            }
        )
        return 0
    if args.secret_action == "set":
        try:
            path = set_password(str(config["runtime_id"]))
        except SecretInputCancelled:
            emit({"passed": False, "error": "已取消"})
            return 130
        print("已加密保存", file=sys.stderr, flush=True)
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


def command_propose(config: dict[str, Any], _args: argparse.Namespace) -> int:
    with run_lock():
        report = propose_routes(config)
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
            "review_path": plan.get("review_path"),
            "approval_kind": plan.get("approval_kind"),
            "warnings": plan.get("warnings") or [],
            "sheets": [
                {"sheet": item["sheet"], "new_dates": item["new_dates"]}
                for item in plan["sheets"]
            ],
            "master_unchanged": True,
        }
    )
    return 0


def command_scheduled_update(config: dict[str, Any], args: argparse.Namespace) -> int:
    with run_lock():
        result = automatic_update(config)
    payload = {
        "started": args.scheduled_started,
        "finished": dt.datetime.now().isoformat(timespec="seconds"),
        "passed": True,
        "exit_code": 0,
        "changed": bool(result.get("changed")),
        "sheets": int(result.get("sheets", 0)),
        "new_rows": int(result.get("rows", 0)),
        "backup": result.get("backup"),
        "warnings": result.get("warnings") or [],
    }
    record_scheduled_run_safely(payload)
    emit(payload)
    return 0


def command_commit(config: dict[str, Any], args: argparse.Namespace) -> int:
    if not args.yes_reviewed_preview:
        raise RuntimeError("Refusing commit without --yes-reviewed-preview")
    with run_lock():
        result = commit(config)
        result["automatic_updates"] = approve_automation(config)
        emit(result)
    return 0


def command_automation(config: dict[str, Any], args: argparse.Namespace) -> int:
    if args.automation_action == "revoke":
        emit(revoke_automation())
    else:
        emit(automation_status(config))
    return 0


def command_schedule(config: dict[str, Any], args: argparse.Namespace) -> int:
    if args.schedule_action == "install":
        emit(install_schedule(config))
    elif args.schedule_action == "remove":
        emit(remove_schedule())
    else:
        report = schedule_status()
        report["automatic_updates"] = automation_status(config)
        emit(report)
    return 0


def command_workbook(config: dict[str, Any], args: argparse.Namespace) -> int:
    if args.workbook_action == "init-template":
        emit(init_template(config))
        return 0
    raise RuntimeError(f"Unsupported workbook action: {args.workbook_action}")


def command_products(config: dict[str, Any], args: argparse.Namespace) -> int:
    config_path = Path(args.config).expanduser().resolve()
    if args.products_action == "status":
        emit(product_status(config))
        return 0
    with run_lock():
        if args.products_action == "sync":
            result = sync_products(config, refresh=not args.use_existing_proposals)
        elif args.products_action == "add":
            result = add_product(
                config,
                config_path,
                proposal_index=args.proposal_index,
                sheet=args.sheet,
                frequency=args.frequency,
                code=args.code,
                product_name=args.product_name,
                subject_contains=args.subject_contains,
                sheet_mode=args.sheet_mode,
                parser=args.parser,
                cumulative_policy=args.cumulative_policy,
                cumulative_offset=args.cumulative_offset,
                return_basis=args.return_basis,
                series_start=args.series_start,
                max_staleness_days=args.max_staleness_days,
                benchmark_source_sheet=args.benchmark_source_sheet,
                benchmark_source_type=args.benchmark_source_type,
                benchmark_source_date=args.benchmark_source_date,
                benchmark_source_value=args.benchmark_source_value,
                benchmark_display_name=args.benchmark_display_name,
            )
        elif args.products_action == "adopt":
            result = adopt_product(
                config,
                config_path,
                proposal_index=args.proposal_index,
                sheet=args.sheet,
                code=args.code,
                product_name=args.product_name,
                subject_contains=args.subject_contains,
            )
        elif args.products_action == "clone":
            result = clone_product(
                config,
                config_path,
                proposal_index=args.proposal_index,
                sheet=args.sheet,
                copy_from=args.copy_from,
                code=args.code,
                product_name=args.product_name,
                subject_contains=args.subject_contains,
                inherit_benchmark=args.inherit_benchmark,
            )
        elif args.products_action == "pause":
            result = pause_product(
                config, config_path, sheet=args.sheet, reason=args.reason
            )
        elif args.products_action == "resume":
            result = resume_product(config, config_path, sheet=args.sheet)
        else:
            raise RuntimeError(f"Unsupported products action: {args.products_action}")
    emit(result)
    return 0 if result.get("passed", True) else 2


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
    result = argparse.ArgumentParser(description="Local IMAP-to-Excel NAV automation")
    result.add_argument(
        "--config", default=str(CONFIG_PATH), help="Runtime config.json path"
    )
    commands = result.add_subparsers(dest="command", required=True)
    commands.add_parser("doctor")
    secret = commands.add_parser("secret")
    secret.add_argument("secret_action", choices=("launch", "set", "status", "remove"))
    commands.add_parser("propose")
    commands.add_parser("discover")
    commands.add_parser("validate")
    commands.add_parser("preview")
    commands.add_parser("scheduled-update", help=argparse.SUPPRESS)
    commit_parser = commands.add_parser("commit")
    commit_parser.add_argument("--yes-reviewed-preview", action="store_true")
    schedule = commands.add_parser("schedule")
    schedule.add_argument("schedule_action", choices=("install", "remove", "status"))
    automation = commands.add_parser("automation")
    automation.add_argument("automation_action", choices=("status", "revoke"))
    workbook = commands.add_parser("workbook")
    workbook.add_argument("workbook_action", choices=("init-template",))
    products = commands.add_parser("products")
    product_actions = products.add_subparsers(dest="products_action", required=True)
    product_actions.add_parser("status", help="查看当前活动、暂停及工作表状态")
    product_sync = product_actions.add_parser(
        "sync", help="扫描邮箱并比较新增候选与当前产品"
    )
    product_sync.add_argument("--use-existing-proposals", action="store_true")
    product_add = product_actions.add_parser(
        "add", help="从 products sync 的候选新增受管产品"
    )
    product_add.add_argument("--proposal-index", type=int, required=True)
    product_add.add_argument("--sheet", required=True)
    product_add.add_argument("--frequency", choices=("daily", "weekly"), required=True)
    product_add.add_argument("--code")
    product_add.add_argument("--product-name")
    product_add.add_argument("--subject-contains")
    product_add.add_argument("--sheet-mode", choices=("summary", "append"))
    product_add.add_argument("--parser", default="auto")
    product_add.add_argument(
        "--cumulative-policy", choices=("require", "unit", "offset"), default="require"
    )
    product_add.add_argument("--cumulative-offset", type=float)
    product_add.add_argument(
        "--return-basis", choices=("unit", "cumulative"), default="cumulative"
    )
    product_add.add_argument("--series-start")
    product_add.add_argument("--max-staleness-days", type=int, default=14)
    product_add.add_argument("--benchmark-source-sheet")
    product_add.add_argument(
        "--benchmark-source-type",
        choices=("level", "aligned_return"),
        default="level",
    )
    product_add.add_argument("--benchmark-source-date", default="A")
    product_add.add_argument("--benchmark-source-value", default="B")
    product_add.add_argument("--benchmark-display-name")
    product_adopt = product_actions.add_parser(
        "adopt", help="接管用户已经建好的产品 Sheet"
    )
    product_adopt.add_argument("--proposal-index", type=int, required=True)
    product_adopt.add_argument("--sheet", required=True)
    product_adopt.add_argument("--code")
    product_adopt.add_argument("--product-name")
    product_adopt.add_argument("--subject-contains")
    product_clone = product_actions.add_parser(
        "clone", help="照现有受管 Sheet 的格式新建并接管产品"
    )
    product_clone.add_argument("--proposal-index", type=int, required=True)
    product_clone.add_argument("--sheet", required=True)
    product_clone.add_argument("--copy-from", required=True)
    product_clone.add_argument("--code")
    product_clone.add_argument("--product-name")
    product_clone.add_argument("--subject-contains")
    product_clone.add_argument("--inherit-benchmark", action="store_true")
    product_pause = product_actions.add_parser(
        "pause", help="暂停产品但保留工作表和历史"
    )
    product_pause.add_argument("--sheet", required=True)
    product_pause.add_argument("--reason", required=True)
    product_resume = product_actions.add_parser("resume", help="恢复已暂停产品")
    product_resume.add_argument("--sheet", required=True)
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
    configure_output()
    args = parser().parse_args()
    args.scheduled_started = dt.datetime.now().isoformat(timespec="seconds")
    try:
        if args.command == "demo":
            return command_demo(args)
        config = load_config(Path(args.config).resolve())
        prune_logs(config)
        commands = {
            "doctor": command_doctor,
            "secret": command_secret,
            "propose": command_propose,
            "discover": command_discover,
            "validate": command_validate,
            "preview": command_preview,
            "scheduled-update": command_scheduled_update,
            "commit": command_commit,
            "schedule": command_schedule,
            "automation": command_automation,
            "workbook": command_workbook,
            "products": command_products,
        }
        return commands[args.command](config, args)
    except (ConfigError, RuntimeError, ValueError, OSError) as exc:
        error = (
            str(exc)
            if not isinstance(exc, OSError)
            else f"系统操作失败：{exc.strerror or type(exc).__name__}"
        )
        if args.command == "scheduled-update":
            record_scheduled_run_safely(
                {
                    "started": args.scheduled_started,
                    "finished": dt.datetime.now().isoformat(timespec="seconds"),
                    "passed": False,
                    "exit_code": 2,
                    "error": error,
                }
            )
        emit({"passed": False, "error": error})
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
