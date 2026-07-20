#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import subprocess
from dataclasses import dataclass
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
IGNORED_DIRS = {".git", ".pytest_cache", ".ruff_cache", ".venv", "__pycache__"}
TEXT_SUFFIXES = {
    ".py",
    ".md",
    ".txt",
    ".yaml",
    ".yml",
    ".cmd",
    ".ps1",
    ".lock",
    ".gitignore",
    "",
}
DENIED_SUFFIXES = {
    ".xlsx",
    ".xlsm",
    ".xls",
    ".csv",
    ".pdf",
    ".eml",
    ".zip",
    ".log",
    ".png",
    ".jpg",
    ".jpeg",
}
DENIED_NAMES = {
    "config.json",
    "route-report.json",
    "validation-report.json",
    "plan.json",
    "scheduled_tasks.json",
    "last-scheduled-run.json",
    "run.lock",
    "secret.json",
}
PATTERNS = {
    "absolute-user-path": re.compile(r"(?i)[A-Z]:[\\/]Users[\\/][^\\/\s]+"),
    "sync-or-desktop-path": re.compile(
        r"(?i)(OneDrive|Dropbox|Google Drive|Desktop)[\\/]"
    ),
    "non-example-email": re.compile(r"(?i)\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b"),
    "credential-literal": re.compile(
        r"(?i)(password|passwd|api[_-]?key|secret|token)\s*[:=]\s*['\"][^'\"]{4,}['\"]"
    ),
    "private-key": re.compile(r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----"),
}


@dataclass(frozen=True)
class Finding:
    category: str
    path: str
    line: int


def allowed_email(value: str) -> bool:
    domain = value.rsplit("@", 1)[-1].lower()
    return domain in {"example.com", "example.invalid", "users.noreply.github.com"}


def scan_text(path: str, text: str) -> list[Finding]:
    findings: list[Finding] = []
    for number, line in enumerate(text.splitlines(), 1):
        for category, pattern in PATTERNS.items():
            matches = list(pattern.finditer(line))
            if category == "non-example-email":
                matches = [
                    match for match in matches if not allowed_email(match.group(0))
                ]
            if matches:
                findings.append(Finding(category, path, number))
    return findings


def tree_files() -> list[Path]:
    return sorted(
        path
        for path in ROOT.rglob("*")
        if path.is_file() and not (set(path.relative_to(ROOT).parts) & IGNORED_DIRS)
    )


def scan_tree() -> list[Finding]:
    findings: list[Finding] = []
    for path in tree_files():
        relative = path.relative_to(ROOT).as_posix()
        findings.extend(scan_text(relative, relative))
        if path.name in DENIED_NAMES:
            findings.append(Finding("runtime-artifact", relative, 0))
            continue
        if path.suffix.lower() in DENIED_SUFFIXES:
            findings.append(Finding("binary-or-data-artifact", relative, 0))
            continue
        if path.suffix.lower() not in TEXT_SUFFIXES and path.name not in {
            "LICENSE",
            ".gitignore",
        }:
            findings.append(Finding("unexpected-file-type", relative, 0))
            continue
        try:
            text = path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            findings.append(Finding("non-utf8-text", relative, 0))
            continue
        findings.extend(scan_text(relative, text))
    return findings


def scan_history() -> list[Finding]:
    if not (ROOT / ".git").exists():
        return [Finding("git-history-unavailable", ".git", 0)]
    findings: list[Finding] = []
    commits = subprocess.run(
        ["git", "-C", str(ROOT), "rev-list", "--all"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        check=True,
    ).stdout.splitlines()
    for commit in commits:
        raw_paths = subprocess.run(
            ["git", "-C", str(ROOT), "ls-tree", "-rz", "--name-only", commit],
            capture_output=True,
            check=True,
        ).stdout
        paths = [item.decode("utf-8") for item in raw_paths.split(b"\0") if item]
        for relative in paths:
            path = Path(relative)
            findings.extend(scan_text(f"{commit[:8]}:{relative}", relative))
            if path.name in DENIED_NAMES or path.suffix.lower() in DENIED_SUFFIXES:
                findings.append(
                    Finding("historical-data-artifact", f"{commit[:8]}:{relative}", 0)
                )
                continue
            if path.suffix.lower() not in TEXT_SUFFIXES and path.name not in {
                "LICENSE",
                ".gitignore",
            }:
                findings.append(
                    Finding(
                        "historical-unexpected-file-type", f"{commit[:8]}:{relative}", 0
                    )
                )
                continue
            result = subprocess.run(
                ["git", "-C", str(ROOT), "show", f"{commit}:{relative}"],
                capture_output=True,
                check=False,
            )
            if result.returncode:
                continue
            try:
                text = result.stdout.decode("utf-8")
            except UnicodeDecodeError:
                findings.append(
                    Finding("historical-non-utf8-text", f"{commit[:8]}:{relative}", 0)
                )
                continue
            findings.extend(scan_text(f"{commit[:8]}:{relative}", text))
    metadata = subprocess.run(
        ["git", "-C", str(ROOT), "log", "--all", "--format=%H%x00%an%x00%ae%x00%s"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        check=True,
    ).stdout
    findings.extend(scan_text("git-commit-metadata", metadata))
    refs = subprocess.run(
        ["git", "-C", str(ROOT), "for-each-ref", "--format=%(refname)"],
        capture_output=True,
        text=True,
        encoding="utf-8",
        check=True,
    ).stdout
    findings.extend(scan_text("git-refs", refs))
    return findings


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Redacted privacy gate for the skill tree and Git history"
    )
    parser.add_argument(
        "--history", action="store_true", help="Scan every reachable Git commit"
    )
    args = parser.parse_args()
    findings = scan_tree()
    if args.history:
        findings.extend(scan_history())
    unique = sorted(
        set(findings), key=lambda item: (item.category, item.path, item.line)
    )
    if unique:
        for finding in unique:
            suffix = f":{finding.line}" if finding.line else ""
            print(f"ERROR {finding.category}: {finding.path}{suffix}")
        print(f"Privacy audit failed with {len(unique)} redacted finding(s).")
        return 2
    print(
        f"Privacy audit passed: {len(tree_files())} files scanned; no sensitive artifacts detected."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
