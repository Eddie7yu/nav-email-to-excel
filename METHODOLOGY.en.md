# Safely Writing "Data From Email" Into a "Living Spreadsheet" — A General Methodology

> Distilled from a real project: extracting fund NAV figures from custodian emails
> and writing them automatically into an Excel workbook maintained by hand by an
> operations person. All business specifics are stripped here; what remains is a
> reusable methodology for any **semi-structured input (email / PDF / attachment)
> → human-maintained system of record (Excel / Sheet / DB)** automation.

---

## 0. The Real Problem

The thing you're automating is usually a **living file**: someone uses it daily,
and mistakes are expensive.

The hard part is never "read the number." It's that:

- the input is **semi-structured** (every source has its own format, and they drift);
- the target is **human-maintained** (full of hand-entry quirks, edited at any time, must not be corrupted);
- a wrong write has **real consequences** (finance / compliance / trust).

So the whole methodology orbits one idea: **write trustably, reversibly, and without disrupting the human.**

---

## 1. Eight Design Principles

### ① Read first, write last, always preview before commit
Every stage that touches the source of truth has two gears: **preview** (writes to a
copy, original untouched) and **`--commit`** (writes for real, and **backs up first**).
The default is always preview. The operator confirms with their eyes, then you commit.

### ② Split the flow into independent stages
`fetch → parse → validate → write → notify`. Each is its own script, independently
runnable and debuggable. A failure pinpoints to one stage instead of one black box.

### ③ Validation is a first-class stage: prove "parsing is correct" before trusting it on new data
Before writing anything new, take the **last N rows already known to be correct**,
re-parse them, and confirm the parser **reproduces** what the human already recorded.
If it reproduces → the parser is trustworthy → only then let it touch new data. This
single step catches the vast majority of silent errors.

### ④ Never trust human-maintained data
People leave: trailing spaces in cells, hidden newlines (`\n`), pre-filled empty rows,
inconsistent fonts. Defenses:
- **Defensive parsing**: treat whitespace-only as empty; tolerate stray characters.
- **Sanitize on write**: `strip()` before writing, don't propagate dirty characters.
- Turn every pothole you hit into a code comment — they are **unpredictable; you only meet them by hitting them.**

### ⑤ Formatting follows the document, not the code
A new row copies the **style of the previous row** (font, color, border, number format
— all of it). **Never hardcode style in code.** This way the operator changes
formatting by editing the last few rows in the file, and the next run continues it —
**presentation belongs to the human, data belongs to the code.**

To actually pull this off on a **hand-formatted** file, two hard-won rules:

- **Write the real file with a "real program," not a pure library rewrite.** Libraries like
  openpyxl `load+save` silently **drop charts/pivots/some formatting** and re-emit hand-set
  colors as **theme colors** — which render as **different colors and broken layout on someone
  else's machine** (you won't see it on yours). For a human-maintained master, drive **real Excel
  (COM / `win32com`)** and act like a person: copy the row above → insert → overwrite only
  values/formulas, so formatting is inherited for free.
- **Separate compute from write:** use the pure library to compute new rows into a **throwaway
  preview copy** (master untouched), then use COM to transplant only the rows the preview has
  beyond the master back into the real master. You get the library's convenience without ever
  letting it touch the file you can't afford to corrupt.
- **Pin colors to explicit RGB, never "theme" colors**, for cross-machine consistency — and
  verify by opening it once **on the target machine** (theme drift is invisible on your own).

### ⑥ Idempotency + resilience: one bad sheet must not sink the whole run
- Target name doesn't match / temporarily missing → **skip and log**, don't throw and crash the whole batch.
- Any script must be safe to re-run (idempotent).
- An external dependency (network) failing → **defer and retry** that item, never write bad data.

### ⑦ Configuration-driven, not code-driven
Sheet names, identifiers, source mappings, processing scope — all in JSON config.
**Adding a new data source = a config edit, not a code change.**

### ⑧ Strictly separate one-time migrations from the recurring job
Structural changes (renames, reformatting, adding columns) go in "run-once-then-delete"
scripts; **never inside the scheduled job.** That way the pipeline that runs every day
only ever touches data, and will never silently reformat your file one morning.

---

## 2. Engineering Checklist

| Concern | Approach |
|---|---|
| **Portability** | Auto-locate the target file (by signature name / unique match); paths relative to the script; works as-is on a new machine |
| **Reversibility** | Auto-save a timestamped backup before every write |
| **Observability** | Log every run; append every write to a changelog; send a summary when done |
| **Self-diagnosis** | One command to check dependencies, config, connectivity, and target presence |
| **Guarded routing** | Match primarily by unique identifier; fall back to name; explicitly block ambiguous matches |
| **Minimal secrets** | Use the simplest scheme that **matches the operator's real threat model** — don't over-harden (see below) |

---

## 3. Process Methodology (how to build it, step by step)

1. **Inventory first, write nothing.** Map every input source, the target structure, and naming conventions — read-only throughout.
2. **Prove before you expand.** Get validation passing on known data before letting it touch new data.
3. **Smallest deployable change.** Preview → verify → commit, one step at a time.
4. **Don't do what wasn't asked.** Every extra layer of abstraction/hardening is future debt.
5. **Ask when blocked; don't guess-and-implement.** One clarifying question saves a whole round of rework.

---

## 4. Anti-Patterns (lessons paid for in real money)

- **Over-engineering security = breaking a working system with your own hands.**
  We once moved the secret from config to a "more secure" separate file — the operator's
  machine didn't have that file, so it couldn't connect and the whole thing went dark.
  **Lesson: secret handling must match the operator's real threat model, not the
  engineer's imagined best practice.** If the operator says "just put it in the config,
  I'm not worried about leaks," then put it in the config.

- **Debugging on a machine you can't see is brutally slow.**
  The environment that actually breaks is often one you can't touch; every
  "run it → report back → I'll fix" round-trip is high-latency.
  **Countermeasure: build self-diagnostics into the tool** so one command prints the
  environment's state.

- **Unstated dependencies and ordering get read as "you didn't do it."**
  Prerequisites like "migrate the structure first, then write data" must be baked into
  the docs and prompts from the start.

- **Few code changes ≠ a fast process.**
  In the end we changed very little code — which is precisely the sign the architecture
  was right. The time went into "undoing my own over-engineering" and "flushing out
  hand-entry quirks one by one." **Small diffs are an outcome, not the process.**

---

## 5. In One Sentence

> **Treat the human as the file's owner, and the code as a cautious assistant.**
> The assistant: reads first, backs up before changing, previews before acting, lets
> formatting follow the owner, asks when unsure, and never makes it complex when one
> command would do.
