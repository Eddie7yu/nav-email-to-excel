# Portable AI handoff

Give the repository URL and the local workbook path to an AI that can read files and run Python. Never include the IMAP authorization code.

Use this request:

```text
Read SKILL.md and references/configuration.md completely. Use the bundled scripts to create a local NAV email automation for my existing workbook.

Start read-only. Do not write the master workbook, send messages, or install scheduled tasks without my explicit approval. Do not ask me to paste an authorization code into chat; tell me when to run the hidden local secret command myself.

Configure only the senders and sheets I authorize. Discover columns semantically, validate at least two historical dates per route, fail closed on ambiguity or value conflict, and produce a preview plus a sanitized acceptance report. Formal writing must use the guarded Excel/WPS COM commit path after I approve the preview.
```

Clone only the clean public repository. Do not copy an already configured runtime.

Codex on Windows:

```powershell
git clone <repo-url> "$env:USERPROFILE\.codex\skills\nav-email-to-excel"
codex
```

Then ask Codex: `Use $nav-email-to-excel to deploy a preview-first runtime for this workbook: <absolute-path>`.

Claude Code on Windows:

```powershell
git clone <repo-url> "$env:USERPROFILE\.claude\skills\nav-email-to-excel"
claude
```

Then ask Claude Code to use the `nav-email-to-excel` skill and provide the workbook's absolute local path.

Cursor:

```powershell
git clone <repo-url> "D:\tools\nav-email-to-excel"
cursor "D:\tools\nav-email-to-excel"
```

Tell Cursor Agent to read the absolute `SKILL.md` path and to operate only on the separately supplied absolute workbook path. Confirm that the workspace permission includes the workbook's directory before discovery; otherwise Cursor cannot inspect it.

Default installation locations:

- Codex personal skill: `%USERPROFILE%\.codex\skills\nav-email-to-excel`
- Claude Code personal skill: `%USERPROFILE%\.claude\skills\nav-email-to-excel`
- Cursor: any dedicated local tools directory; direct Agent to the cloned `SKILL.md`.

Update an existing clean installation with `git -C <skill-directory> pull --ff-only`. Keep runtimes elsewhere so an update cannot overwrite local configuration.

Do not copy a configured runtime into an AI skill directory. Share only the clean skill repository.
