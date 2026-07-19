#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import uuid
import venv
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "assets" / "runtime-template"
RUNTIME_FILES = (
    ".gitignore",
    "navctl.py",
    "nav_commit.py",
    "nav_config.py",
    "nav_mail.py",
    "nav_parse.py",
    "nav_schedule.py",
    "nav_service.py",
    "nav_workbook.py",
    "requirements.lock",
    "run-preview.cmd",
    "runtime_secret.py",
)


def runtime_python(root: Path) -> Path:
    return root / ".venv" / ("Scripts/python.exe" if os.name == "nt" else "bin/python")


def validate_inputs(args: argparse.Namespace) -> tuple[Path, Path]:
    if not (3, 11) <= sys.version_info[:2] <= (3, 14):
        raise RuntimeError("Python 3.11 through 3.14 is required")
    destination = Path(args.destination).expanduser().resolve()
    workbook = Path(args.workbook).expanduser().resolve()
    if destination.exists():
        raise RuntimeError(
            f"Destination already exists; choose a new directory: {destination}"
        )
    if not destination.parent.is_dir():
        raise RuntimeError(f"Destination parent does not exist: {destination.parent}")
    if not os.access(destination.parent, os.W_OK):
        raise RuntimeError(f"Destination parent is not writable: {destination.parent}")
    if not workbook.is_file() or workbook.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise RuntimeError("Workbook must be an existing .xlsx or .xlsm file")
    if not TEMPLATE.is_dir():
        raise RuntimeError("Bundled runtime template is missing")
    args.email = args.email.strip()
    args.imap_host = args.imap_host.strip()
    args.mailbox = args.mailbox.strip()
    if not re.fullmatch(r"[^\s@]+@[^\s@]+", args.email):
        raise RuntimeError("Email account is invalid")
    if not args.imap_host or re.search(r"\s", args.imap_host):
        raise RuntimeError("IMAP host is required and cannot contain whitespace")
    if not args.mailbox:
        raise RuntimeError("IMAP mailbox cannot be empty")
    if not 1 <= args.imap_port <= 65535:
        raise RuntimeError("IMAP port must be between 1 and 65535")
    if not 1 <= args.lookback_days <= 3650:
        raise RuntimeError("Lookback days must be between 1 and 3650")
    return destination, workbook


def config_payload(args: argparse.Namespace, workbook: Path) -> dict:
    return {
        "schema_version": 1,
        "runtime_id": str(uuid.uuid4()),
        "workbook_path": str(workbook),
        "imap": {
            "host": args.imap_host,
            "port": args.imap_port,
            "user": args.email,
            "mailbox": args.mailbox,
            "lookback_days": args.lookback_days,
            "max_messages": 2000,
            "max_message_bytes": 25 * 1024 * 1024,
            "max_total_bytes": 100 * 1024 * 1024,
        },
        "routes": [],
        "column_overrides": {},
        "style": {"mode": "infer", "zero_threshold": 0.00005},
        "schedule": [],
        "validation": {
            "minimum_history_dates": 2,
            "tolerance": 0.000001,
            "max_future_days": 0,
            "max_period_change": 0.5,
        },
        "retention": {"backup_count": 10, "preview_count": 10, "log_days": 30},
    }


def create_runtime(args: argparse.Namespace, destination: Path, workbook: Path) -> None:
    staging = (
        destination.parent / f".{destination.name}.bootstrap-{uuid.uuid4().hex[:8]}"
    )
    try:
        staging.mkdir()
        for name in RUNTIME_FILES:
            source = TEMPLATE / name
            if not source.is_file():
                raise RuntimeError(f"Bundled runtime file is missing: {name}")
            shutil.copy2(source, staging / name)
        builder = venv.EnvBuilder(
            with_pip=True, clear=False, system_site_packages=args.skip_deps
        )
        builder.create(staging / ".venv")
        python = runtime_python(staging)
        if not args.skip_deps:
            result = subprocess.run(
                [
                    str(python),
                    "-m",
                    "pip",
                    "install",
                    "--disable-pip-version-check",
                    "-r",
                    str(staging / "requirements.lock"),
                ],
                cwd=staging,
            )
            if result.returncode:
                raise RuntimeError(
                    "Dependency installation failed; no runtime was installed"
                )
        config = config_payload(args, workbook)
        (staging / "config.json").write_text(
            json.dumps(config, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
        )
        for command in staging.glob("*.cmd"):
            text = (
                command.read_text(encoding="utf-8")
                .replace("\r\n", "\n")
                .replace("\n", "\r\n")
            )
            command.write_bytes(text.encode("utf-8"))
        staging.rename(destination)
    except Exception:
        if staging.exists():
            shutil.rmtree(staging)
        raise


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(
        description="Create an isolated local NAV email runtime"
    )
    result.add_argument("--destination", required=True)
    result.add_argument("--workbook", required=True)
    result.add_argument("--email", required=True)
    result.add_argument("--imap-host", required=True)
    result.add_argument("--imap-port", type=int, default=993)
    result.add_argument("--mailbox", default="INBOX")
    result.add_argument("--lookback-days", type=int, default=180)
    result.add_argument("--skip-deps", action="store_true", help=argparse.SUPPRESS)
    result.add_argument(
        "--validate-only", action="store_true", help="Validate without creating files"
    )
    return result


def main() -> int:
    args = parser().parse_args()
    try:
        destination, workbook = validate_inputs(args)
        if args.validate_only:
            print("Bootstrap validation passed; no files were created.")
            return 0
        create_runtime(args, destination, workbook)
        print(f"Runtime created: {destination}")
        print("Next: configure authorized routes, then run navctl.py doctor.")
        return 0
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
