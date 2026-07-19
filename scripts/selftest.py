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


def bootstrap_test(temporary: Path) -> None:
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
    config = json.loads((destination / "config.json").read_text(encoding="utf-8"))
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
        or report["preview_ready"]
    ):
        raise AssertionError(
            f"bootstrapped runtime readiness report is wrong: {doctor.stdout}"
        )


def main() -> int:
    parser = argparse.ArgumentParser(description="Offline sanitized regression test")
    parser.add_argument(
        "--com",
        action="store_true",
        help="Also test a temporary Excel/WPS formal commit",
    )
    args = parser.parse_args()
    with tempfile.TemporaryDirectory(prefix="nav-skill-") as temporary:
        runtime = Path(temporary) / "中文 路径" / "runtime"
        runtime.parent.mkdir()
        shutil.copytree(TEMPLATE, runtime)
        command = [sys.executable, "-X", "utf8", str(DRIVER), "--runtime", str(runtime)]
        if args.com:
            command.append("--com")
        result = subprocess.run(
            command, cwd=runtime, env=dict(os.environ, PYTHONUTF8="1")
        )
        if result.returncode:
            return result.returncode
        bootstrap_test(Path(temporary))
    print("selftest: PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
