#!/usr/bin/env python3
from __future__ import annotations

import re
from pathlib import Path

import yaml


ROOT = Path(__file__).resolve().parents[1]
EXPECTED_NAME = "nav-email-to-excel"


def fail(message: str) -> None:
    raise SystemExit(f"ERROR: {message}")


def main() -> int:
    skill_path = ROOT / "SKILL.md"
    text = skill_path.read_text(encoding="utf-8")
    match = re.match(r"\A---\n(.*?)\n---\n", text, re.DOTALL)
    if not match:
        fail("SKILL.md must start with YAML frontmatter")
    metadata = yaml.safe_load(match.group(1))
    if not isinstance(metadata, dict) or set(metadata) != {"name", "description"}:
        fail("SKILL.md frontmatter must contain only name and description")
    if metadata["name"] != EXPECTED_NAME:
        fail("skill name does not match the package directory")
    description = str(metadata["description"])
    if (
        not description.strip()
        or len(description) > 1024
        or "<" in description
        or ">" in description
    ):
        fail("skill description is empty or invalid")
    for target in re.findall(r"\[[^]]+\]\(([^)]+)\)", text):
        if "://" not in target and not (ROOT / target).is_file():
            fail(f"SKILL.md references a missing local file: {target}")

    interface = yaml.safe_load(
        (ROOT / "agents" / "openai.yaml").read_text(encoding="utf-8")
    )
    if not isinstance(interface, dict) or set(interface) != {"interface"}:
        fail("agents/openai.yaml has unexpected top-level fields")
    fields = interface["interface"]
    required = {"display_name", "short_description", "default_prompt"}
    if not isinstance(fields, dict) or not required <= set(fields):
        fail("agents/openai.yaml is missing interface fields")
    if f"${EXPECTED_NAME}" not in str(fields["default_prompt"]):
        fail("default_prompt must explicitly invoke the skill")
    print("Package metadata check passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
