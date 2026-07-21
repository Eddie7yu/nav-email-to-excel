#!/usr/bin/env python3
from __future__ import annotations

import argparse
import io
import re
import subprocess
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parents[1]
IGNORED_DIRS = {
    ".git",
    ".pytest_cache",
    ".ruff_cache",
    ".venv",
    "__pycache__",
    "node_modules",
}
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
    "route-proposals.json",
    "automation-approval.json",
    "scheduled_tasks.json",
    "last-scheduled-run.json",
    "run.lock",
    "secret.json",
}
ALLOWED_WORKBOOK = "assets/workbook-templates/nav-standard-cn.xlsx"
EXPECTED_TEMPLATE_SHEETS = [
    "模板-周度-无指数",
    "模板-周度-有指数",
    "模板-日度-无指数",
    "模板-日度-有指数",
    "模板-指数数据",
]
ALLOWED_TEMPLATE_TEXT = {
    "产品代码",
    "产品名称",
    "单位净值",
    "累计单位净值",
    "净值日期",
    "周收益",
    "基准指数",
    "指数收益(周度)",
    "超额(周度)",
    "日收益",
    "指数收益(日度)",
    "超额(日度)",
    "累计",
    "日期",
    "指数点位/收益",
    "来源",
}
DENIED_XLSX_PARTS = {
    "connections.xml",
    "vbaproject.bin",
    "externallinks",
    "customxml",
    "comments",
    "threadedcomments",
    "persons",
    "embeddings",
    "media",
    "printersettings",
    "macrosheets",
    "ctrlprops",
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


def _xml_text(root: ET.Element) -> str:
    return "".join(root.itertext())


def scan_template_workbook(path: str, payload: bytes) -> list[Finding]:
    findings: list[Finding] = []
    try:
        archive = zipfile.ZipFile(io.BytesIO(payload))
    except (OSError, zipfile.BadZipFile):
        return [Finding("invalid-template-workbook", path, 0)]
    with archive:
        names = archive.namelist()
        lowered = [name.lower() for name in names]
        for name, lower in zip(names, lowered):
            if any(part in lower for part in DENIED_XLSX_PARTS):
                findings.append(Finding("forbidden-xlsx-part", f"{path}!{name}", 0))
            if lower.endswith((".bin", ".vml")):
                findings.append(Finding("forbidden-xlsx-binary", f"{path}!{name}", 0))
        required = {
            "xl/workbook.xml",
            "xl/sharedStrings.xml",
            *{f"xl/worksheets/sheet{index}.xml" for index in range(1, 6)},
        }
        for name in sorted(required - set(names)):
            findings.append(Finding("missing-template-part", f"{path}!{name}", 0))
        if findings:
            return findings
        try:
            workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
            shared_root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        except ET.ParseError:
            return [Finding("invalid-template-xml", path, 0)]
        namespace = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        sheets = workbook_root.findall(".//x:sheet", namespace)
        sheet_names = [str(sheet.attrib.get("name", "")) for sheet in sheets]
        if sheet_names != EXPECTED_TEMPLATE_SHEETS:
            findings.append(Finding("unexpected-template-sheets", path, 0))
        if any(sheet.attrib.get("state", "visible") != "visible" for sheet in sheets):
            findings.append(Finding("hidden-template-sheet", path, 0))
        shared = [_xml_text(item) for item in shared_root.findall("x:si", namespace)]
        if any(value not in ALLOWED_TEMPLATE_TEXT for value in shared):
            findings.append(Finding("unexpected-template-text", path, 0))
        for name in names:
            if not name.lower().endswith((".xml", ".rels")):
                continue
            raw = archive.read(name)
            try:
                text = raw.decode("utf-8")
                root = ET.fromstring(raw)
            except (UnicodeDecodeError, ET.ParseError):
                findings.append(Finding("invalid-template-xml", f"{path}!{name}", 0))
                continue
            findings.extend(scan_text(f"{path}!{name}", text))
            if 'TargetMode="External"' in text or "TargetMode='External'" in text:
                findings.append(
                    Finding("external-xlsx-relationship", f"{path}!{name}", 0)
                )
            if name.startswith("xl/worksheets/"):
                if root.findall(".//x:f", namespace):
                    findings.append(Finding("template-formula", f"{path}!{name}", 0))
                if root.findall(".//x:hyperlink", namespace):
                    findings.append(Finding("template-hyperlink", f"{path}!{name}", 0))
                hidden_rows = [
                    row
                    for row in root.findall(".//x:row", namespace)
                    if row.attrib.get("hidden") in {"1", "true"}
                ]
                hidden_columns = [
                    column
                    for column in root.findall(".//x:col", namespace)
                    if column.attrib.get("hidden") in {"1", "true"}
                ]
                if hidden_rows or hidden_columns:
                    findings.append(
                        Finding("hidden-template-data", f"{path}!{name}", 0)
                    )
                for cell in root.findall(".//x:c", namespace):
                    value = cell.find("x:v", namespace)
                    inline = cell.find("x:is", namespace)
                    if value is None and inline is None:
                        continue
                    cell_type = cell.attrib.get("t")
                    if cell_type == "s" and value is not None:
                        try:
                            cell_text = shared[int(value.text or "")]
                        except (ValueError, IndexError):
                            findings.append(
                                Finding("invalid-shared-string", f"{path}!{name}", 0)
                            )
                            continue
                    elif cell_type == "inlineStr" and inline is not None:
                        cell_text = _xml_text(inline)
                    else:
                        cell_text = str(value.text or "") if value is not None else ""
                    if cell_text and cell_text not in ALLOWED_TEMPLATE_TEXT:
                        findings.append(
                            Finding("template-business-data", f"{path}!{name}", 0)
                        )
        for core_name in ("docProps/core.xml", "docProps/custom.xml"):
            if (
                core_name in names
                and _xml_text(ET.fromstring(archive.read(core_name))).strip()
            ):
                findings.append(
                    Finding("template-author-metadata", f"{path}!{core_name}", 0)
                )
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
        if relative == ALLOWED_WORKBOOK:
            findings.extend(scan_template_workbook(relative, path.read_bytes()))
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
            if relative == ALLOWED_WORKBOOK:
                result = subprocess.run(
                    ["git", "-C", str(ROOT), "show", f"{commit}:{relative}"],
                    capture_output=True,
                    check=False,
                )
                if result.returncode:
                    findings.append(
                        Finding(
                            "historical-template-unreadable",
                            f"{commit[:8]}:{relative}",
                            0,
                        )
                    )
                else:
                    findings.extend(
                        scan_template_workbook(
                            f"{commit[:8]}:{relative}", result.stdout
                        )
                    )
                continue
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
