# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A de-identified reference implementation that extracts semi-structured NAV (net asset value) data from custody/brokerage **emails** (body tables, `.xls(x)` attachments, PDF weekly reports) and writes it into a **human-maintained Excel workbook** safely and reversibly. All real fund names, product codes, and email addresses are placeholders (`基金01`, `DEMO04`, `*.example.com`).

The methodology behind the code is the real point — read `METHODOLOGY.md` (中文) / `METHODOLOGY.en.md` (English) before making design decisions. The eight design principles there (preview-before-commit, staged pipeline, validation as a first-class step, never trust hand-maintained data, format follows the document not the code, idempotent + fault-tolerant, config-driven, one-time migrations kept separate from the cron job) are load-bearing and should not be eroded.

## Commands

All scripts live in `src/` and must be run from there (paths resolve relative to the script dir).

```bash
cd src
pip install -r requirements.txt

# First-time / new-machine setup (writes auth code to a non-synced location, builds registry + index)
python setup_machine.py --auth <IMAP授权码> [--master <workbook name|abs path>] [--user <email>]

# The pipeline, stage by stage (each runs standalone and is independently debuggable):
python build_index.py     # fetch email headers per sender -> index.json (no bodies fetched)
python fill_index.py      # fetch benchmark index closes -> index_cache.json (self-validating)
python validate.py        # regression-check: re-parse last N recorded rows, prove they reproduce
python propose.py         # compute rows that SHOULD be added (writes nothing) -> proposed_updates.csv
python write.py           # PREVIEW: write a *_自动更新预览.xlsx copy (master untouched)
python write.py --commit  # back up master (timestamped), then write master in place
python notify.py --dry    # compose weekly summary email, print only
python notify.py          # send via SMTP

# Orchestrated runs (used by the scheduler):
python run_weekly.py                 # preview only (master untouched, email printed not sent)
python run_weekly.py --commit        # write master (with backup) + send summary email
python run_weekly.py --commit --no-notify   # write silently, no email (mid-week runs)
```

There is **no test suite, linter, or build step**. The closest thing to a test is `validate.py` (regression against known-good recorded data) — run it before trusting new writes. To generate a de-identified sample workbook for experimenting: `python examples/make_sample_workbook.py`.

Windows scheduling: `setup_tasks.bat` / `delete_tasks.bat` (run as Administrator) create/remove five `schtasks` jobs that call `run_weekly.bat` (full run + email) and `run_quiet.bat` (write, no email, with self-diagnostics to `logs/bat_quiet.txt`).

### Testing knobs (env vars)

- `NAV_TEST_ALLWEEKS=1` — lift the "never write the unfinished current week" guard so current-week rows become writable (see `propose.py:CURRENT_WK`).
- `NAV_QQ_PW` — IMAP/SMTP authorization code, highest-priority source (see `navlib.get_password`).

## Pipeline data flow

```
build_index → fill_index → validate → propose → write → notify
  index.json   index_cache  validation   (in-mem)  master.xlsx  email
                .json        .json                  + changelog.csv
                                                    + last_run.json
```

`run_weekly.py` chains these. `build_index.py` and `write.py` are **fatal** if they fail; `fill_index.py`, `validate.py`, `notify.py` are **non-fatal** (logged, run continues). `validate.py` only runs on even ISO weeks. Every run logs to `logs/run_YYYYMMDD_HHMMSS.log`.

State files are runtime artifacts (gitignored): `index.json`, `index_cache.json`, `validation.json`, `last_run.json`, `proposed_updates.csv`, `changelog.csv`.

## Architecture

### Two ingestion phases, routed by sender format

`config.json` maps each sender email → a **format key**. Format keys split into two families:

- **Phase 1** (`navlib.py`): NAV arrives in the email **subject/body**. Parsers in `navlib.parse_subject` (fast, header-only, used by `build_index`) and `navlib.parse_body` (full, used by validate/propose). Format keys: `gtht`, `citics`, `htsc_glrfw`, `htsc_incos`, `xyzq`. Routed to `scope_sheets`.
- **Phase 2** (`phase2.py`): NAV arrives in an **`.xls/.xlsx` attachment** or a **PDF weekly report**. `phase2.parse_rows` handles tabular, labeled (key/value), and PDF layouts via header-synonym matching. Format keys in `phase2.FMTS` (`cms`, `dwzq`, `csc`, `htsc_leap`, `gtht_ta`, `yiyuan`). Routed to `scope_sheets_p2`.

Phase-2 routing (`phase2.route`) is **restricted to phase-2 sheets** on purpose: some custodians mail both phase-1 and phase-2 products, and a foreign product code in the subject must NOT fall back to manager-name routing (it would write the wrong product's NAV). This guardrail is deliberate — preserve it.

**Adding a new source = write a parser + register the sender in `config.json`.** Do not touch the orchestrator.

### Two config files (both gitignored; `.example.json` templates are committed)

- `config.json` — IMAP/SMTP creds, `master_path`, `senders` (email→format), `scope_sheets`, `scope_sheets_p2`, `notify`.
- `registry.json` — per-sheet structure: `header_row`, `data_start`, `last_data_row`, `max_col`, `codes`, `names`, `return_base` (whether the weekly-return formula divides C or D), last date. **Regenerate with `python build_registry.py`** whenever the workbook layout changes (product added/removed, rows shifted). `setup_machine.py` does this automatically.

`navlib._resolve_master()` locates the workbook portably (configured abs path → name under the parent dir → the single/"净值"-named `.xlsx`), so the tool moves between machines without path edits.

### Workbook columns (convention across sheets)

`A`=code, `B`=name, `C`=单位净值 (unit NAV), `D`=累计单位净值 (cumulative NAV), `E`=date as Excel serial (epoch 1899-12-30), `F`=weekly return formula, `G`=index level, `H`=index return, `I`=excess return. A `累计` summary row sits directly below the data; rows beyond it must be empty for `write.py` to consider the sheet safe to append to.

### Writer internals (`write.py`) — the trickiest code

This is where the "format follows the document" principle is enforced. Read it carefully before editing:

- New rows **copy cell styles from the row above** (`copy(...)._style`); fonts/colors/number formats are never hard-coded. The one exception is the 累计-row profit/loss recoloring (`_recolor_accum`: profit=red `FFFF0000`, loss=green `FF008000`), which re-derives sign by actually evaluating the formula (`_eval_accum`) rather than reading Excel's cached value.
- The `累计` summary row is captured as a template, then rewritten at its new position; `ArrayFormula` ranges (e.g. `=PRODUCT(1+H4:H39)-1`) are extended via `_bump_array` (bump the range *end* only, preserve per-column start rows, re-anchor `ref`).
- **Index/benchmark handling** has four modes per sheet: `None` (no index col), `fill` (validated closes available → write G/H/I), `blank` (index rejected → write NAV, leave G blank + flag), `wait` (cache missing → **defer** the row, retry next run). Deferral keeps rows contiguous and never writes bad data.
- "Type B" index columns (header like `500指数`) are cross-sheet references to a source 中证 sheet's H column, auto-filled by date in `fill_crossref_index`.
- Hand-maintained-data defenses: cells with only whitespace count as empty; values are `strip()`-ed on write to avoid leaking trailing `\n` (which renders as garbage boxes in Excel).

### Proposal safety guards (`propose.py`)

`compute_proposals()` decides what *should* be written and is reused by `write.py`. Guards, all of which should be preserved:

- Only **completed weeks** are written; the current (unfinished) week goes to `preview` (unless `NAV_TEST_ALLWEEKS`).
- **Pending-Friday hold**: if a just-finished week's latest data is pre-Friday and still fresh (≤6 days), HOLD the row (don't write, don't advance the week pointer) so the real Friday NAV can land next run.
- **Dormant guard**: newest email older than 45 days → likely redeemed, flag don't backfill.
- **Cross-checks** that produce anomalies instead of writes: parsed code must match expected product code; parsed NAV date must match subject date; `|weekly return| > 20%` or `unit ≤ 0` → flag.
- `citics_virtual` source: cumulative is reconstructed as `unit + offset` (offset = last recorded `cum - unit`).

### IMAP access (`navlib.py`)

All message access uses **IMAP UIDs (stable)**, never sequence numbers — this is a live inbox whose sequence numbers shift as mail arrives. `build_index` fetches only headers (`BODY.PEEK[HEADER.FIELDS ...]`, batched 200); bodies are fetched lazily only when validate/propose need them. Mailbox is opened `readonly=True`.

### Secret handling (intentionally minimal)

`get_password()` resolution order: `NAV_QQ_PW` env var → `%LOCALAPPDATA%/nav_tool/secret.json` → `config.json`. This matches the operator's real threat model — see the METHODOLOGY anti-pattern about over-engineering secrets ("放配置里就行" is an acceptable answer here). Do not "harden" this without the operator asking; a previous attempt to move secrets to a separate file broke the live system.

## Conventions

- Comments are bilingual: English docstrings, Chinese inline notes capturing real-world quirks ("坑"). When you hit a hand-maintained-data quirk, document it inline — they are unpredictable and only discoverable by encountering them.
- Excel dates are integer serials with epoch `datetime.date(1899, 12, 30)`; `s2d`/serial conversion helpers are duplicated across scripts.
- **One-time migrations** (renames, reordering, column adds) go in standalone `examples/one_time_migration.example.py`-style scripts that are run once and deleted — **never** in the scheduled pipeline. Don't add structural mutations to `write.py`.
- Failing softly is correct: a sheet name that doesn't match / a missing index / a stale product → skip and record, never crash the whole run.
