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

from bootstrap import (
    DEFAULT_RUNTIME_DIRECTORY,
    MAX_WINDOWS_DESTINATION_CHARS,
    pip_install_command,
)


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "assets" / "runtime-template"
WORKBOOK_TEMPLATE = ROOT / "assets" / "workbook-templates" / "nav-standard-cn.xlsx"
DRIVER = ROOT / "scripts" / "selftest_driver.py"
BOOTSTRAP = ROOT / "scripts" / "bootstrap.py"


def bootstrap_test(temporary: Path, use_com: bool) -> None:
    import openpyxl

    workbook_parent = temporary / "部署 空格 & 括号 (测试)"
    workbook_parent.mkdir()
    workbook_path = workbook_parent / "输入 工作簿.xlsx"
    workbook = openpyxl.Workbook()
    workbook.active.append(["NAV Date", "Unit NAV"])
    workbook.save(workbook_path)
    destination = workbook_parent / DEFAULT_RUNTIME_DIRECTORY
    command = [
        sys.executable,
        "-X",
        "utf8",
        str(BOOTSTRAP),
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
    explicit_destination = temporary / "显式部署目录"
    explicit_validation = subprocess.run(
        command + ["--destination", str(explicit_destination), "--validate-only"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=environment,
    )
    if explicit_validation.returncode or explicit_destination.exists():
        raise AssertionError("bootstrap rejected an explicit destination override")
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
    mirror_url = "https://pypi.tuna.tsinghua.edu.cn/simple"
    accepted = subprocess.run(
        command + ["--index-url", mirror_url, "--validate-only"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=environment,
    )
    if accepted.returncode or destination.exists():
        raise AssertionError("bootstrap rejected a valid one-time HTTPS package index")
    rejected = subprocess.run(
        command
        + [
            "--index-url",
            "http://pypi.tuna.tsinghua.edu.cn/simple",
            "--validate-only",
        ],
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=environment,
    )
    if rejected.returncode != 2 or destination.exists():
        raise AssertionError("bootstrap accepted an insecure package index")
    install_command = pip_install_command(
        Path("python"), Path("requirements.lock"), mirror_url
    )
    if install_command[-4:] != ["--index-url", mirror_url, "-r", "requirements.lock"]:
        raise AssertionError(
            "bootstrap did not scope the selected index to pip install"
        )
    if os.name == "nt":
        long_parent = temporary / "long-path-check"
        long_parent.mkdir()
        minimum_length = MAX_WINDOWS_DESTINATION_CHARS + 1
        leaf_length = max(10, minimum_length - len(str(long_parent.resolve())) - 1)
        long_destination = long_parent / ("x" * leaf_length)
        long_command = command + ["--destination", str(long_destination)]
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
    if not destination.is_dir():
        raise AssertionError(
            "bootstrap did not derive the workbook-adjacent runtime directory"
        )
    collision = subprocess.run(
        command + ["--validate-only"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=environment,
    )
    if collision.returncode != 2 or "already exists" not in collision.stderr:
        raise AssertionError(
            "bootstrap did not fail closed when the derived runtime already existed"
        )
    unwanted = [
        path
        for path in destination.rglob("*")
        if ".venv" not in path.relative_to(destination).parts
        and (path.name == "__pycache__" or path.suffix.lower() in {".pyc", ".pyo"})
    ]
    if unwanted:
        raise AssertionError("bootstrap copied build caches into the deployed runtime")
    app = destination / "app"
    if not (app / "parsers").is_dir():
        raise AssertionError(
            "bootstrap did not create the trusted local parser directory"
        )
    expected_user_entries = {
        ".gitignore",
        "app",
        "backups",
        "logs",
        "previews",
        "使用说明.txt",
        "手动更新.bat",
        "查看状态.bat",
        "首次授权.bat",
    }
    if {path.name for path in destination.iterdir()} != expected_user_entries:
        raise AssertionError("bootstrap did not keep the user-facing root concise")
    if list(destination.glob("*.py")) or list(destination.glob("*.json")):
        raise AssertionError(
            "bootstrap exposed internal code or state in the user root"
        )
    for launcher_name in ("首次授权.bat", "查看状态.bat", "手动更新.bat"):
        launcher = destination / launcher_name
        launcher_text = launcher.read_text(encoding="utf-8")
        if (
            not launcher.is_file()
            or not launcher.read_bytes().isascii()
            or "app\\.venv\\Scripts\\python.exe" not in launcher_text
            or "app\\navctl.py" not in launcher_text
        ):
            raise AssertionError(
                f"bootstrap generated an invalid user launcher: {launcher_name}"
            )
    secret_launcher = destination / "首次授权.bat"
    if (
        not secret_launcher.is_file()
        or "secret set" not in secret_launcher.read_text(encoding="utf-8")
        or not secret_launcher.read_bytes().isascii()
    ):
        raise AssertionError(
            "bootstrap did not create a cmd-compatible visible secret prompt"
        )
    config_path = app / "config.json"
    original_config = config_path.read_bytes()
    config = json.loads(config_path.read_text(encoding="utf-8"))
    if (
        config["workbook_path"] != str(workbook_path.resolve())
        or config["workbook_mode"] != "existing"
        or config["style"]["mode"] != "infer"
        or config["routes"]
    ):
        raise AssertionError("bootstrap generated an unsafe initial configuration")
    if not (app / "assets" / "nav-standard-cn.xlsx").is_file():
        raise AssertionError(
            "bootstrap did not install the sanitized workbook template"
        )
    runtime_python = (
        app / ".venv" / ("Scripts/python.exe" if os.name == "nt" else "bin/python")
    )
    if not runtime_python.is_file():
        raise AssertionError("bootstrap did not create the isolated Python runtime")
    for wrapper in (
        path
        for path in destination.rglob("*")
        if path.is_file() and path.suffix.lower() in {".bat", ".cmd"}
    ):
        payload = wrapper.read_bytes()
        if b"\n" in payload.replace(b"\r\n", b""):
            raise AssertionError(
                f"{wrapper.name} does not use Windows CRLF line endings"
            )
    if os.name == "nt":
        command_processor = Path(
            os.environ.get("COMSPEC")
            or Path(os.environ.get("SystemRoot", r"C:\Windows"))
            / "System32"
            / "cmd.exe"
        )
        if not command_processor.is_file():
            raise AssertionError("Windows command processor was not found")
        expected_codes = {
            "查看状态.bat": 0,
            "手动更新.bat": 2,
            "首次授权.bat": 2,
        }
        for launcher_name, expected_code in expected_codes.items():
            launched = subprocess.run(
                [
                    str(command_processor),
                    "/d",
                    "/s",
                    "/c",
                    launcher_name,
                ],
                cwd=destination,
                input="\n\n",
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=environment,
                timeout=20,
            )
            output = f"{launched.stdout}\n{launched.stderr}".lower()
            if (
                launched.returncode != expected_code
                or "not recognized as an internal or external command" in output
                or "不是内部或外部命令" in output
            ):
                raise AssertionError(
                    f"Windows launcher failed from a special-character path: "
                    f"{launcher_name}: {output}"
                )
    doctor = subprocess.run(
        [str(runtime_python), "-X", "utf8", "navctl.py", "doctor"],
        cwd=app,
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=environment,
    )
    report = json.loads(doctor.stdout)
    checks = {item["name"]: item["passed"] for item in report["checks"]}
    if (
        doctor.returncode != 2
        or not report["bootstrap_ready"]
        or report["mail_discovery_ready"]
        or report["preview_ready"]
        or report["schedule_ready"]
        or checks.get("schedule-config") is not False
        or report.get("runtime_platform") != sys.platform
        or not isinstance(report.get("spreadsheet_apps_detected"), list)
    ):
        raise AssertionError(
            f"bootstrapped runtime readiness report is wrong: {doctor.stdout}"
        )
    if os.name == "nt":
        piped_secret = subprocess.run(
            [str(runtime_python), "-X", "utf8", "navctl.py", "secret", "set"],
            cwd=app,
            input="",
            capture_output=True,
            text=True,
            encoding="utf-8",
            env=environment,
            timeout=10,
        )
        piped_report = json.loads(piped_secret.stdout)
        if (
            piped_secret.returncode != 2
            or piped_report.get("error") != "此命令需要用户在真实终端中运行"
            or piped_secret.stderr
        ):
            raise AssertionError(
                "piped secret input did not fail immediately and cleanly"
            )
    scheduled_update = subprocess.run(
        [str(runtime_python), "-X", "utf8", "navctl.py", "scheduled-update"],
        cwd=app,
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=environment,
    )
    if scheduled_update.returncode != 2:
        raise AssertionError("unapproved scheduled update did not fail closed")
    schedule_status = subprocess.run(
        [str(runtime_python), "-X", "utf8", "navctl.py", "schedule", "status"],
        cwd=app,
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
        cwd=app,
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
            cwd=app,
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
                cwd=app,
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
            cwd=app,
            capture_output=True,
            text=True,
            encoding="utf-8",
            env=environment,
        )
        if removed.returncode:
            raise AssertionError(f"offline demo cleanup failed: {removed.stdout}")
    if config_path.read_bytes() != original_config:
        raise AssertionError("offline demo changed the real runtime configuration")
    if list(destination.glob("*.py")) or list(destination.glob("*.json")):
        raise AssertionError(
            "runtime commands leaked internal files into the user root"
        )
    if (
        not (app / "last-scheduled-run.json").is_file()
        or not (app / "run.lock").is_file()
    ):
        raise AssertionError("runtime state was not kept inside app")
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

    new_workbook_parent = temporary / "新建工作簿场景"
    new_workbook_parent.mkdir()
    new_destination = new_workbook_parent / DEFAULT_RUNTIME_DIRECTORY
    new_workbook = new_workbook_parent / "尚未创建的净值表.xlsx"
    new_command = [
        sys.executable,
        "-X",
        "utf8",
        str(BOOTSTRAP),
        "--new-workbook",
        str(new_workbook),
        "--email",
        "user@example.invalid",
        "--imap-host",
        "imap.example.invalid",
        "--skip-deps",
    ]
    new_installation = subprocess.run(
        new_command,
        capture_output=True,
        text=True,
        encoding="utf-8",
        env=environment,
    )
    if new_installation.returncode or new_workbook.exists():
        raise AssertionError(
            f"new-workbook bootstrap failed or created the workbook too early: {new_installation.stderr}"
        )
    new_config = json.loads(
        (new_destination / "app" / "config.json").read_text(encoding="utf-8")
    )
    if (
        new_config["workbook_mode"] != "bundled-template"
        or new_config["workbook_path"] != str(new_workbook.resolve())
        or new_config["style"]["mode"] != "cn-red-up-green-down"
    ):
        raise AssertionError("new-workbook bootstrap wrote the wrong safety profile")


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
        (runtime / "assets").mkdir()
        shutil.copy2(WORKBOOK_TEMPLATE, runtime / "assets" / WORKBOOK_TEMPLATE.name)
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
