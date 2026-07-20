#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "assets" / "runtime-template"
DRIVER = ROOT / "scripts" / "selftest_driver.py"
BOOTSTRAP = ROOT / "scripts" / "bootstrap.py"


def bootstrap_test(temporary: Path, use_com: bool) -> None:
    import openpyxl

    workbook_path = temporary / "输入 工作簿.xlsx"
    workbook = openpyxl.Workbook()
    workbook.active.append(["NAV Date", "Unit NAV"])
    workbook.save(workbook_path)
    destination = temporary / "部署 运行时"
    command = [
        sys.executable,
        "-X",
        "utf8",
        str(BOOTSTRAP),
        "--destination",
        str(destination),
        "--workbook",
        str(workbook_path),
        "--email",
        "user@example.invalid",
        "--imap-host",
        "imap.example.invalid",
        "--skip-deps",
    ]
    environment = dict(os.environ, PYTHONUTF8="1")
    validation = subprocess.run(
        command + ["--validate-only"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=environment,
    )
    if validation.returncode or destination.exists():
        raise AssertionError(
            f"bootstrap validation-only mode failed: {validation.stderr}"
        )
    invalid_host = command.copy()
    invalid_host[invalid_host.index("imap.example.invalid")] = ""
    rejected = subprocess.run(
        invalid_host + ["--validate-only"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=environment,
    )
    if rejected.returncode == 0 or destination.exists():
        raise AssertionError("bootstrap accepted an empty IMAP host")
    rejected = subprocess.run(
        command + ["--lookback-days", "-1", "--validate-only"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=environment,
    )
    if rejected.returncode == 0 or destination.exists():
        raise AssertionError("bootstrap accepted a negative lookback window")
    if os.name == "nt":
        long_parent = temporary / "long-path-check"
        long_parent.mkdir()
        minimum_length = 121
        leaf_length = max(10, minimum_length - len(str(long_parent.resolve())) - 1)
        long_destination = long_parent / ("x" * leaf_length)
        long_command = command.copy()
        long_command[long_command.index(str(destination))] = str(long_destination)
        rejected = subprocess.run(
            long_command + ["--validate-only"],
            capture_output=True,
            text=True,
            encoding="utf-8",
            env=environment,
        )
        if (
            rejected.returncode != 2
            or long_destination.exists()
            or "安装目录过长" not in rejected.stderr
        ):
            raise AssertionError(
                "bootstrap did not reject a Windows destination beyond its supported path budget"
            )
    installation = subprocess.run(
        command, capture_output=True, text=True, encoding="utf-8", env=environment
    )
    if installation.returncode:
        raise AssertionError(f"bootstrap installation failed: {installation.stderr}")
    unwanted = [
        path
        for path in destination.rglob("*")
        if ".venv" not in path.relative_to(destination).parts
        and (path.name == "__pycache__" or path.suffix.lower() in {".pyc", ".pyo"})
    ]
    if unwanted:
        raise AssertionError("bootstrap copied build caches into the deployed runtime")
    if not (destination / "parsers").is_dir():
        raise AssertionError(
            "bootstrap did not create the trusted local parser directory"
        )
    config_path = destination / "config.json"
    original_config = config_path.read_bytes()
    config = json.loads(config_path.read_text(encoding="utf-8"))
    if config["workbook_path"] != str(workbook_path.resolve()) or config["routes"]:
        raise AssertionError("bootstrap generated an unsafe initial configuration")
    runtime_python = (
        destination
        / ".venv"
        / ("Scripts/python.exe" if os.name == "nt" else "bin/python")
    )
    if not runtime_python.is_file():
        raise AssertionError("bootstrap did not create the isolated Python runtime")
    for wrapper in destination.glob("*.cmd"):
        payload = wrapper.read_bytes()
        if b"\n" in payload.replace(b"\r\n", b""):
            raise AssertionError(
                f"{wrapper.name} does not use Windows CRLF line endings"
            )
    doctor = subprocess.run(
        [str(runtime_python), "-X", "utf8", "navctl.py", "doctor"],
        cwd=destination,
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=environment,
    )
    report = json.loads(doctor.stdout)
    if (
        doctor.returncode != 2
        or not report["bootstrap_ready"]
        or report["mail_discovery_ready"]
        or report["preview_ready"]
    ):
        raise AssertionError(
            f"bootstrapped runtime readiness report is wrong: {doctor.stdout}"
        )
    scheduled_update = subprocess.run(
        [str(runtime_python), "-X", "utf8", "navctl.py", "scheduled-update"],
        cwd=destination,
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=environment,
    )
    if scheduled_update.returncode != 2:
        raise AssertionError("unapproved scheduled update did not fail closed")
    schedule_status = subprocess.run(
        [str(runtime_python), "-X", "utf8", "navctl.py", "schedule", "status"],
        cwd=destination,
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=environment,
    )
    status_report = json.loads(schedule_status.stdout)
    if (
        schedule_status.returncode
        or status_report.get("last_run", {}).get("passed") is not False
        or status_report.get("last_run", {}).get("exit_code") != 2
    ):
        raise AssertionError(
            "schedule status did not report the latest automatic-update failure"
        )
    demo = subprocess.run(
        [str(runtime_python), "-X", "utf8", "navctl.py", "demo", "prepare"],
        cwd=destination,
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=environment,
    )
    try:
        demo_report = json.loads(demo.stdout)
    except json.JSONDecodeError as exc:
        raise AssertionError(
            f"offline demo returned invalid JSON: {demo.stdout}"
        ) from exc
    if (
        demo.returncode
        or not demo_report["passed"]
        or demo_report["real_data_used"]
        or not Path(demo_report["preview_path"]).is_file()
    ):
        raise AssertionError(f"offline demo preparation failed: {demo.stdout}")
    run_id = str(demo_report["run_id"])
    try:
        refused = subprocess.run(
            [
                str(runtime_python),
                "-X",
                "utf8",
                "navctl.py",
                "demo",
                "commit",
                "--run-id",
                run_id,
            ],
            cwd=destination,
            capture_output=True,
            text=True,
            encoding="utf-8",
            env=environment,
        )
        if refused.returncode != 2:
            raise AssertionError("offline demo commit did not require preview approval")
        if use_com:
            committed = subprocess.run(
                [
                    str(runtime_python),
                    "-X",
                    "utf8",
                    "navctl.py",
                    "demo",
                    "commit",
                    "--run-id",
                    run_id,
                    "--yes-reviewed-preview",
                ],
                cwd=destination,
                capture_output=True,
                text=True,
                encoding="utf-8",
                env=environment,
            )
            commit_report = json.loads(committed.stdout)
            if (
                committed.returncode
                or not commit_report["passed"]
                or commit_report["rows_written"] != 1
            ):
                raise AssertionError(
                    f"offline demo COM commit failed: {committed.stdout}"
                )
    finally:
        removed = subprocess.run(
            [
                str(runtime_python),
                "-X",
                "utf8",
                "navctl.py",
                "demo",
                "remove",
                "--run-id",
                run_id,
            ],
            cwd=destination,
            capture_output=True,
            text=True,
            encoding="utf-8",
            env=environment,
        )
        if removed.returncode:
            raise AssertionError(f"offline demo cleanup failed: {removed.stdout}")
    if config_path.read_bytes() != original_config:
        raise AssertionError("offline demo changed the real runtime configuration")
    if os.name == "nt":
        unicode_destination = temporary / "unicode-😀-runtime"
        unicode_command = [
            sys.executable,
            str(BOOTSTRAP),
            "--destination",
            str(unicode_destination),
            "--workbook",
            str(workbook_path),
            "--email",
            "user@example.invalid",
            "--imap-host",
            "imap.example.invalid",
            "--skip-deps",
        ]
        unicode_environment = dict(os.environ, PYTHONUTF8="0")
        unicode_result = subprocess.run(
            unicode_command,
            capture_output=True,
            env=unicode_environment,
        )
        if unicode_result.returncode or not unicode_destination.is_dir():
            raise AssertionError(
                "bootstrap failed after creating a valid non-console-codepage path"
            )


def main() -> int:
    parser = argparse.ArgumentParser(description="完全使用虚构数据的离线回归测试")
    parser.add_argument(
        "--com",
        action="store_true",
        help="同时使用临时工作簿验证 Excel/WPS COM 正式写入",
    )
    args = parser.parse_args()
    with tempfile.TemporaryDirectory(prefix="nav-skill-") as temporary:
        runtime = Path(temporary) / "中文 路径" / "runtime"
        runtime.parent.mkdir()
        shutil.copytree(TEMPLATE, runtime)
        command = [sys.executable, "-X", "utf8", str(DRIVER), "--runtime", str(runtime)]
        if args.com:
            command.append("--com")
        print("开始核心离线回归；不会读取真实邮箱、密钥或工作簿。", flush=True)
        result = subprocess.run(
            command, cwd=runtime, env=dict(os.environ, PYTHONUTF8="1")
        )
        if result.returncode:
            return result.returncode
        print("检查全新部署、就绪状态和 navctl 离线演练命令。", flush=True)
        bootstrap_test(Path(temporary), args.com)
    print("selftest: PASS（临时文件已清理，未使用真实资料）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
