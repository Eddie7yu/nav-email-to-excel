from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
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
from nav_config import (
    CONFIG_PATH,
    STATE_ROOT,
    ConfigError,
    ROOT,
    active_routes,
    benchmark_review_issue,
    load_config,
)
from nav_demo import commit as commit_demo
from nav_demo import list_runs as list_demo_runs
from nav_demo import prepare as prepare_demo
from nav_demo import remove as remove_demo
from nav_products import add as add_product
from nav_products import add_code_alias
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
from nav_service import (
    discover,
    doctor,
    preview,
    propose_headers,
    propose_routes,
    validate,
)
from nav_template import init_template
from runtime_secret import (
    SecretInputCancelled,
    launch_secret_prompt,
    read_password,
    remove_password,
    set_password,
)

MAX_UTF8_ARGFILE_BYTES = 1024 * 1024


def emit(payload: Any) -> None:
    print(json.dumps(payload, ensure_ascii=False, indent=2))


def emit_proposal_progress(payload: dict[str, Any]) -> None:
    print(
        json.dumps({"proposal_progress": payload}, ensure_ascii=False),
        file=sys.stderr,
        flush=True,
    )


def configure_output() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure:
            reconfigure(errors="backslashreplace")


def expand_utf8_argfiles(argv: list[str]) -> tuple[list[str], bool]:
    expanded: list[str] = []
    used = False
    for value in argv:
        if not value.startswith("@") or len(value) == 1:
            expanded.append(value)
            continue
        path = Path(value[1:]).expanduser().resolve()
        if not path.is_file():
            raise ValueError(f"UTF-8 参数文件不存在：{path}")
        if path.stat().st_size > MAX_UTF8_ARGFILE_BYTES:
            raise ValueError("UTF-8 参数文件超过 1 MiB 上限")
        try:
            text = path.read_text(encoding="utf-8-sig")
        except UnicodeDecodeError as exc:
            raise ValueError("参数文件必须使用 UTF-8 编码") from exc
        for line in text.splitlines():
            argument = line.strip()
            if not argument or argument.startswith("#"):
                continue
            if argument.startswith("@"):
                raise ValueError("UTF-8 参数文件不支持嵌套 @ 文件")
            if "\x00" in argument:
                raise ValueError("UTF-8 参数文件包含非法 NUL 字符")
            expanded.append(argument)
        used = True
    return expanded, used


def _recovered_utf8_text(value: str) -> str | None:
    for encoding in ("gbk", "cp1252", "latin1"):
        try:
            recovered = value.encode(encoding).decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            continue
        if recovered != value and re.search(r"[\u3400-\u9fff]", recovered):
            return recovered
    return None


def validate_filter_encoding(value: str | None, label: str) -> None:
    text = str(value or "")
    if not text:
        return
    if "\ufffd" in text or any(0xDC80 <= ord(character) <= 0xDCFF for character in text):
        raise ValueError(
            f"{label} 含编码替换字符；请改用 UTF-8 参数文件 @<路径>"
        )
    if _recovered_utf8_text(text) is not None:
        raise ValueError(
            f"{label} 疑似被终端错误解码；请改用 UTF-8 参数文件 @<路径>"
        )


def record_scheduled_run_safely(payload: dict[str, Any]) -> None:
    try:
        record_scheduled_run(payload)
    except OSError:
        pass


def deployment_status(
    config: dict[str, Any],
    schedule_report: dict[str, Any],
    *,
    authorization_available: bool | None = None,
) -> dict[str, Any]:
    routes = [
        route for route in config.get("routes") or [] if isinstance(route, dict)
    ]
    active = active_routes(config)
    review_issues = [
        issue
        for route in active
        if (issue := benchmark_review_issue(route)) is not None
    ]
    approval = automation_status(config)
    last_run = schedule_report.get("last_run")
    if not isinstance(last_run, dict):
        last_run = {}
    if authorization_available is None:
        authorization_available = bool(read_password(str(config["runtime_id"])))
    return {
        "authorization": {"available": bool(authorization_available)},
        "routes": {
            "configured": len(routes),
            "active": len(active),
            "paused": len(routes) - len(active),
        },
        "benchmark_reviews": {
            "blocking": len(review_issues),
            "source_unresolved": review_issues.count(
                "benchmark-source-unresolved"
            ),
            "license_unresolved": review_issues.count(
                "benchmark-license-unresolved"
            ),
        },
        "automatic_updates": {
            "approved": bool(approval.get("approved")),
            "approved_at": approval.get("approved_at"),
        },
        "scheduled_tasks": len(schedule_report.get("tasks") or []),
        "last_update": {
            "available": bool(last_run),
            "finished_at": last_run.get("finished"),
            "passed": last_run.get("passed"),
            "changed": last_run.get("changed"),
            "new_rows": int(last_run.get("new_rows") or 0),
        },
    }


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


def command_discover(config: dict[str, Any], args: argparse.Namespace) -> int:
    with run_lock():
        _, report = discover(config, args.sheets)
    emit(report)
    return 0 if report["passed"] else 2


def command_propose(config: dict[str, Any], args: argparse.Namespace) -> int:
    validate_filter_encoding(args.subject_contains, "--subject-contains")
    with run_lock():
        if args.headers_only:
            if (
                args.sender
                or args.subject_contains
                or args.subject_product_code
                or args.allow_unscoped_full_scan
                or args.resume
                or args.lookback_days is not None
            ):
                raise ValueError(
                    "--headers-only 不能与完整解析的范围参数同时使用"
                )
            report = propose_headers(config, args.header_limit)
        else:
            if args.resume and (
                args.sender
                or args.subject_contains
                or args.subject_product_code
                or args.allow_unscoped_full_scan
                or args.lookback_days is not None
            ):
                raise ValueError("--resume 使用本地检查点，不能同时改写扫描范围")
            if (
                not args.resume
                and not args.sender
                and not args.allow_unscoped_full_scan
            ):
                raise ValueError(
                    "完整候选解析必须至少指定 --sender；先读取头部报告，"
                    "主题不确定时只限定发件人即可。仅在有明确诊断理由并核实资源边界后，"
                    "才能显式添加 --allow-unscoped-full-scan"
                )
            configured_lookback = int(config["imap"].get("lookback_days", 180))
            effective_lookback = args.lookback_days
            if effective_lookback is None and not args.resume:
                effective_lookback = min(configured_lookback, 30)
            report = propose_routes(
                config,
                sender=args.sender,
                subject_contains=args.subject_contains,
                subject_product_code=args.subject_product_code,
                lookback_days=effective_lookback,
                batch_messages=args.batch_messages,
                time_budget_seconds=args.time_budget_seconds,
                resume=args.resume,
                progress_sink=emit_proposal_progress,
            )
    if args.subject_contains and isinstance(report.get("scan"), dict):
        report["scan"]["filter_input_encoding"] = {
            "source": (
                "utf8-argfile"
                if getattr(args, "utf8_argfile_used", False)
                else "native-unicode-argv"
            ),
            "encoding_check": "passed",
        }
    emit(report)
    return 0 if report["passed"] else 2


def command_validate(config: dict[str, Any], _args: argparse.Namespace) -> int:
    with run_lock():
        report = validate(config)
    emit(report)
    return 0 if report["passed"] else 2


def command_preview(config: dict[str, Any], args: argparse.Namespace) -> int:
    if args.reuse_discovery and (args.lookback_days is not None or args.sheets):
        raise ValueError(
            "--reuse-discovery 不能与 --lookback-days 或 --sheet 同时使用"
        )
    mail_config = None
    if args.lookback_days is not None:
        if not 1 <= args.lookback_days <= 3650:
            raise ValueError("本次预览回看天数必须在 1 到 3650 之间")
        mail_config = json.loads(json.dumps(config))
        mail_config["imap"]["lookback_days"] = args.lookback_days
    with run_lock():
        plan = preview(
            config,
            mail_config=mail_config,
            sheets=args.sheets,
            reuse_discovery=args.reuse_discovery,
        )
    emit(
        {
            "preview_path": plan["preview_path"],
            "review_path": plan.get("review_path"),
            "preview_display_name": plan.get("preview_display_name"),
            "preview_read_only": bool(plan.get("preview_read_only")),
            "approval_kind": plan.get("approval_kind"),
            "warnings": plan.get("warnings") or [],
            "sheets": [
                {"sheet": item["sheet"], "new_dates": item["new_dates"]}
                for item in plan["sheets"]
            ],
            "master_unchanged": True,
            "scoped": bool(plan.get("scoped")),
            "scope_sheets": plan.get("scope_sheets") or [],
            "committable": plan.get("committable", True),
            "final_full_preview_required": bool(
                plan.get("final_full_preview_required")
            ),
            "discovery_reused": bool(plan.get("discovery_reused")),
            "effective_lookback_days": (
                args.lookback_days
                if args.lookback_days is not None
                else int(config["imap"].get("lookback_days", 60))
            ),
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
        report["deployment"] = deployment_status(config, report)
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
            validate_filter_encoding(
                args.subject_contains, "--subject-contains"
            )
            if args.use_existing_proposals:
                if (
                    args.sender
                    or args.subject_contains
                    or args.allow_unscoped_full_scan
                ):
                    raise ValueError(
                        "--use-existing-proposals 不能与邮箱扫描范围参数同时使用"
                    )
                result = sync_products(config, refresh=False)
            else:
                if not args.sender and not args.allow_unscoped_full_scan:
                    raise ValueError(
                        "products sync 刷新邮箱时必须至少指定 --sender；"
                        "已有完整候选报告时使用 --use-existing-proposals。"
                        "只有受控诊断才可显式添加 --allow-unscoped-full-scan"
                    )
                result = sync_products(
                    config,
                    refresh=True,
                    sender=args.sender,
                    subject_contains=args.subject_contains,
                )
            if args.subject_contains:
                result["filter_input_encoding"] = {
                    "source": (
                        "utf8-argfile"
                        if getattr(args, "utf8_argfile_used", False)
                        else "native-unicode-argv"
                    ),
                    "encoding_check": "passed",
                }
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
                history_scope=args.history_scope,
                inspect_only=args.inspect_only,
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
        elif args.products_action == "alias":
            result = add_code_alias(
                config, config_path, sheet=args.sheet, code=args.code
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
    propose = commands.add_parser("propose")
    propose.add_argument(
        "--headers-only",
        action="store_true",
        help="只读取有界的 From/Subject，不下载正文或附件",
    )
    propose.add_argument(
        "--header-limit",
        type=int,
        default=25,
        help="头部预筛选读取的最近邮件数量（默认 25）",
    )
    propose.add_argument("--sender", help="完整解析时限定精确发件人")
    propose.add_argument("--subject-contains", help="完整解析时限定稳定主题片段")
    propose.add_argument(
        "--subject-product-code",
        help="显式把唯一主题代码绑定到附件中的无代码净值行",
    )
    propose.add_argument(
        "--allow-unscoped-full-scan",
        action="store_true",
        help="受控诊断时显式允许无发件人范围的完整扫描",
    )
    propose.add_argument(
        "--lookback-days",
        type=int,
        help="本次完整候选解析的回看天数；默认不超过 30 天",
    )
    propose.add_argument(
        "--batch-messages",
        type=int,
        default=50,
        help="单次最多解析的完整邮件数（默认 50）",
    )
    propose.add_argument(
        "--time-budget-seconds",
        type=int,
        default=120,
        help="下载与解析阶段各自的软时间预算（默认各 120 秒）",
    )
    propose.add_argument(
        "--resume",
        action="store_true",
        help="从本地部分报告与 UID 游标继续同一范围",
    )
    discover_parser = commands.add_parser("discover")
    discover_parser.add_argument(
        "--sheet",
        action="append",
        dest="sheets",
        help="只读复跑指定活动 Sheet；可重复使用，最终仍需完整 preview",
    )
    commands.add_parser("validate")
    preview_parser = commands.add_parser("preview")
    preview_parser.add_argument(
        "--lookback-days",
        type=int,
        help="本次历史补录预览临时使用的回看天数，不改变日常配置",
    )
    preview_parser.add_argument(
        "--sheet",
        action="append",
        dest="sheets",
        help="生成指定活动 Sheet 的只读诊断预览；可重复，不能提交",
    )
    preview_parser.add_argument(
        "--reuse-discovery",
        action="store_true",
        help=(
            "轻量复核邮箱范围未变化后复用上次完整解析；"
            "路由、解析器或邮箱变化时自动拒绝"
        ),
    )
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
    product_sync.add_argument("--sender", help="刷新候选时限定精确发件人")
    product_sync.add_argument(
        "--subject-contains", help="刷新候选时限定稳定主题片段"
    )
    product_sync.add_argument(
        "--allow-unscoped-full-scan",
        action="store_true",
        help="受控诊断时显式允许无发件人范围的完整扫描",
    )
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
    product_adopt.add_argument(
        "--history-scope",
        choices=("tail", "mail-history"),
        default="tail",
        help=(
            "内部接管范围：tail 默认只续接表尾；mail-history "
            "允许把可证明的邮箱历史放入首次受保护预览"
        ),
    )
    product_adopt.add_argument(
        "--inspect-only",
        action="store_true",
        help="只读判断接管方式和基准审查状态，不修改配置",
    )
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
    product_alias = product_actions.add_parser(
        "alias", help="确认并保存邮件产品代码的精确别名"
    )
    product_alias.add_argument("--sheet", required=True)
    product_alias.add_argument("--code", required=True)
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
    try:
        expanded_argv, utf8_argfile_used = expand_utf8_argfiles(sys.argv[1:])
    except (OSError, ValueError) as exc:
        emit({"passed": False, "error": str(exc)})
        return 2
    args = parser().parse_args(expanded_argv)
    args.utf8_argfile_used = utf8_argfile_used
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
        failure = {"passed": False, "error": error}
        if args.command == "scheduled-update":
            failure["exit_code"] = 2
            record_scheduled_run_safely(
                {
                    "started": args.scheduled_started,
                    "finished": dt.datetime.now().isoformat(timespec="seconds"),
                    **failure,
                }
            )
        emit(failure)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
