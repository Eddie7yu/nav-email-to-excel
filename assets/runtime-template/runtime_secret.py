from __future__ import annotations

import base64
import ctypes
import json
import os
import subprocess
import sys
from ctypes import wintypes
from pathlib import Path
from typing import Callable, TextIO

from nav_config import ROOT

MASKED_INPUT_PROMPT = "请右键粘贴授权码，屏幕只显示星号，粘贴后回车"


class SecretInputCancelled(Exception):
    pass


class DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ("cbData", wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_ubyte)),
    ]


def secret_path(runtime_id: str) -> Path:
    if os.name != "nt":
        raise RuntimeError(
            "Persistent secret storage is Windows-only; use NAV_EMAIL_PASSWORD"
        )
    root = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    safe_id = "".join(ch for ch in runtime_id if ch.isalnum() or ch in "-_")
    return root / "nav-email-to-excel" / safe_id / "secret.json"


def read_password(runtime_id: str) -> str:
    environment = os.environ.get("NAV_EMAIL_PASSWORD", "")
    if environment:
        return environment
    if os.name != "nt":
        return ""
    path = secret_path(runtime_id)
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
        blob = base64.b64decode(payload["blob"])
        return _unprotect(blob).decode("utf-8")
    except (FileNotFoundError, KeyError, ValueError, OSError, json.JSONDecodeError):
        return ""


def set_password(runtime_id: str, password: str | None = None) -> Path:
    if os.name != "nt":
        raise RuntimeError(
            "Set NAV_EMAIL_PASSWORD in the current shell on non-Windows systems"
        )
    value = password if password is not None else _read_windows_secret()
    if not value:
        raise ValueError("Authorization code cannot be empty")
    path = secret_path(runtime_id)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_suffix(".tmp")
    payload = {
        "protected": "windows-dpapi",
        "blob": base64.b64encode(_protect(value.encode("utf-8"))).decode("ascii"),
    }
    temporary.write_text(json.dumps(payload), encoding="utf-8")
    temporary.replace(path)
    return path


def launch_secret_prompt() -> int:
    if os.name != "nt":
        raise RuntimeError("可见授权码窗口只支持 Windows")
    launcher = ROOT / "首次授权.bat"
    if not launcher.is_file():
        raise RuntimeError("授权码窗口启动脚本缺失")
    system_root = Path(os.environ.get("SystemRoot", r"C:\Windows"))
    command_processor = Path(
        os.environ.get("COMSPEC") or system_root / "System32" / "cmd.exe"
    )
    if not command_processor.is_file():
        raise RuntimeError("找不到 Windows 命令处理程序，无法打开授权码窗口")
    process = subprocess.Popen(
        [str(command_processor), "/d", "/s", "/c", launcher.name],
        cwd=launcher.parent,
        creationflags=getattr(subprocess, "CREATE_NEW_CONSOLE", 0x00000010),
        close_fds=True,
    )
    return int(process.pid)


def _read_windows_secret() -> str:
    import msvcrt

    return _read_interactive_secret(msvcrt.getwch, sys.stdin, sys.stderr)


def _read_interactive_secret(
    get_character: Callable[[], str], input_stream: TextIO | None, output: TextIO
) -> str:
    _require_interactive_terminal(input_stream)
    try:
        return _read_masked(get_character, output)
    except KeyboardInterrupt:
        output.write("\n")
        output.flush()
        raise SecretInputCancelled from None


def _require_interactive_terminal(input_stream: TextIO | None) -> None:
    try:
        interactive = bool(input_stream and input_stream.isatty())
    except (AttributeError, OSError, ValueError):
        interactive = False
    if not interactive:
        raise RuntimeError("此命令需要用户在真实终端中运行")


def _read_masked(get_character: Callable[[], str], output: TextIO) -> str:
    print(MASKED_INPUT_PROMPT, file=output, flush=True)
    characters: list[str] = []
    while True:
        character = get_character()
        if character in {"\r", "\n"}:
            output.write("\n")
            output.flush()
            return "".join(characters)
        if character == "\x03":
            raise KeyboardInterrupt
        if character in {"\x00", "\xe0"}:
            get_character()
            continue
        if character == "\b":
            if characters:
                characters.pop()
                output.write("\b \b")
                output.flush()
            continue
        if len(character) != 1 or ord(character) < 32:
            continue
        characters.append(character)
        output.write("*")
        output.flush()


def remove_password(runtime_id: str) -> bool:
    if os.name != "nt":
        raise RuntimeError(
            "Unset NAV_EMAIL_PASSWORD in the current shell on non-Windows systems"
        )
    path = secret_path(runtime_id)
    existed = path.exists()
    path.unlink(missing_ok=True)
    try:
        path.parent.rmdir()
    except OSError:
        pass
    return existed


def _blob(data: bytes) -> tuple[DATA_BLOB, ctypes.Array[ctypes.c_char]]:
    buffer = ctypes.create_string_buffer(data)
    value = DATA_BLOB(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_ubyte)))
    return value, buffer


def _protect(data: bytes) -> bytes:
    source, keepalive = _blob(data)
    target = DATA_BLOB()
    ok = ctypes.windll.crypt32.CryptProtectData(
        ctypes.byref(source),
        "nav-email-to-excel",
        None,
        None,
        None,
        0x1,
        ctypes.byref(target),
    )
    if not ok:
        raise OSError(ctypes.get_last_error(), "Windows DPAPI encryption failed")
    try:
        return ctypes.string_at(target.pbData, target.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(target.pbData)


def _unprotect(data: bytes) -> bytes:
    source, keepalive = _blob(data)
    target = DATA_BLOB()
    ok = ctypes.windll.crypt32.CryptUnprotectData(
        ctypes.byref(source),
        None,
        None,
        None,
        None,
        0x1,
        ctypes.byref(target),
    )
    if not ok:
        raise OSError(ctypes.get_last_error(), "Windows DPAPI decryption failed")
    try:
        return ctypes.string_at(target.pbData, target.cbData)
    finally:
        ctypes.windll.kernel32.LocalFree(target.pbData)
