---
name: nav-email-to-excel
description: Deploy, configure, validate, operate, or repair a local IMAP-to-Excel fund NAV workflow. Use when an AI must map recurring NAV emails or Excel/CSV/PDF attachments into an existing workbook, preserve formulas and formatting, catch up missed dates, validate historical values, generate previews, perform guarded Excel/WPS writes, or install isolated Windows scheduled tasks.
---

# Deploy NAV email automation

Treat the existing workbook as the authority for layout, formulas, and visual style. Keep all mailbox data, workbook copies, credentials, discovery reports, and runtime configuration on the user's machine.

## Enforce the safety boundary

- Start read-only. Do not write the master workbook, send messages, or install tasks without explicit user approval.
- Never request an IMAP authorization code in chat or a command-line argument. Ask the user to run the hidden local secret command themselves.
- Fail closed on ambiguous products, conflicting dates or values, missing cumulative NAV, unknown columns, incomplete historical validation, missing benchmarks, or workbook structure changes.
- Do not place real names, senders, product names, product codes, workbook names, email content, attachments, paths, logs, or credentials in this skill, tests, commits, or responses.
- Keep the runtime outside the skill directory so skill upgrades cannot overwrite local configuration or data.

## Bootstrap a local runtime

Collect only the destination directory, workbook path, IMAP account, host, and mailbox. Then run from the skill root:

```powershell
python scripts/bootstrap.py --destination "D:\nav-runtime" --workbook "D:\data\nav.xlsx" --email "user@example.com" --imap-host "imap.example.com"
```

Bootstrap creates an isolated virtual environment, installs locked versions, and writes local configuration.

Ask the user to set the mailbox secret locally:

```powershell
cd D:\nav-runtime
.\.venv\Scripts\python.exe navctl.py secret set
```

On macOS/Linux, keep the authorization code only in the current shell; the runtime does not persist a plaintext secret:

```bash
cd /opt/nav-runtime
read -rsp "IMAP authorization code: " NAV_EMAIL_PASSWORD && export NAV_EMAIL_PASSWORD && printf '\n'
```

This version supports IMAP over SSL with an app password/authorization code. It does not implement OAuth-only mailbox login. PDF parsing is text-only and does not perform OCR.

## Configure routes semantically

Read [references/configuration.md](references/configuration.md) before editing `config.json`. Add only authorized senders and managed sheets.

- Prefer `parser: auto` for common labelled bodies and Excel/CSV/PDF tables.
- Require an exact product code when one sender can carry more than one product.
- Use explicit column overrides only when semantic header discovery is insufficient.
- Set the cumulative NAV policy per route. Keep `require` unless historical evidence proves `unit` or a fixed `offset`.
- Set `series_start` when the investment or analysis basis changes. Never continue cumulative results across a new series.
- Map benchmarks to a verified workbook source sheet by exact date. Do not fetch or guess a public index implicitly.

Keep unique parser code in the local runtime. Add a sanitized regression fixture before promoting any generally useful parser back into the public skill.

## Validate before previewing

Run:

```powershell
.\.venv\Scripts\python.exe navctl.py doctor
.\.venv\Scripts\python.exe navctl.py discover
.\.venv\Scripts\python.exe navctl.py validate
.\.venv\Scripts\python.exe navctl.py preview
```

On macOS/Linux, replace `.\.venv\Scripts\python.exe` with `.venv/bin/python`. Those platforms support discovery, validation, and preview only; formal workbook commit remains Windows-only.

Accept configuration only when:

1. Every managed sheet has unambiguous date and unit NAV columns.
2. Each route matches at least two distinct historical dates by default.
3. Code, date, unit NAV, and email-provided cumulative NAV match the workbook within configured tolerance.
4. New tail dates are proposed once, in order, with no duplicate date or same-day value conflict. An internal historical gap stops for supervised repair.
5. Weekly returns appear only on the last available date of a completed week; daily returns use the prior valid date.
6. Benchmark and excess cells are both populated or both left blank.
7. The preview preserves workbook topology, contains every proposed date, and passes the bundled formula and idempotency regression tests.

`doctor` reports separate `bootstrap_ready`, `preview_ready`, `commit_ready`, and `schedule_ready` states. Do not treat dependency-only readiness as permission to read email or write the workbook.

The preview is a local copy. Review its sheet list, inserted rows, formulas, and formatting without modifying the master.

## Commit through Excel or WPS

Formal writes require Windows and a working Microsoft Excel or WPS Spreadsheet COM interface. After the user explicitly approves the reviewed preview, run:

```powershell
.\.venv\Scripts\python.exe navctl.py commit --yes-reviewed-preview
```

Commit must create a backup, apply the plan to a same-directory temporary copy through COM, validate the temporary result, close the spreadsheet process, and atomically replace the master. Any failure must leave the master hash unchanged.

Do not use an openpyxl-only fallback for the formal workbook. If COM is unavailable, stop at preview and report that formal writing is not verified.

## Enable logged-in preview scheduling

Install preview-only Windows tasks only after a successful manual preview and explicit user approval:

```powershell
.\.venv\Scripts\python.exe navctl.py schedule install
```

Tasks require the user to be logged in because they use Task Scheduler interactive mode. Sleep, shutdown, or logout can delay or skip a run. Scheduled tasks never commit the master workbook; a person must review the generated preview and run the guarded commit command manually.

Task names include the runtime ID. Re-running installation replaces only tasks recorded for this runtime. Remove them with:

```powershell
.\.venv\Scripts\python.exe navctl.py schedule remove
```

Do not schedule a runtime stored on a UNC/network path. Do not reuse Python paths or task definitions copied from another machine.

Before deleting a runtime, remove its tasks and DPAPI secret:

```powershell
.\.venv\Scripts\python.exe navctl.py schedule remove
.\.venv\Scripts\python.exe navctl.py secret remove
```

Preview copies, backups, and logs are sensitive local files. Retention is bounded by `config.json`; delete the runtime directory only after any required backup has been moved to an approved location.

## Hand off to other AI tools

The same `SKILL.md` follows the Agent Skills format used by Codex and Claude Code. For Cursor or another local agent, have it read this file and [references/portable-use.md](references/portable-use.md), then operate the bundled deterministic scripts rather than rewriting the workflow.

Before publishing any change, run:

```powershell
python -X utf8 scripts/privacy_audit.py
python -X utf8 scripts/selftest.py
python -X utf8 scripts/package_check.py
```

Also run the official skill validator supplied by the target agent environment when it is available.
