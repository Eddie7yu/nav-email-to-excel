# Runtime configuration

## Contents

- [Top-level fields](#top-level-fields)
- [Route fields](#route-fields)
- [Column discovery](#column-discovery)
- [Benchmark mapping](#benchmark-mapping)
- [Scheduling](#scheduling)
- [Local sensitive files](#local-sensitive-files)

## Top-level fields

`config.json` is generated in the runtime and must never be committed.

| Field | Meaning |
| --- | --- |
| `schema_version` | Configuration schema; currently `1` |
| `runtime_id` | Random deployment identifier used to isolate secrets and tasks |
| `workbook_path` | Absolute `.xlsx` or `.xlsm` master path |
| `imap` | Host, port, user, mailbox, lookback and message-size/count limits; never a password |
| `routes` | Authorized sender/product-to-sheet mappings |
| `column_overrides` | Optional semantic column mappings per sheet |
| `style` | Preserve-first return formatting settings |
| `schedule` | Optional Windows task definitions |
| `validation` | Historical sample count and numeric tolerance |
| `retention` | Maximum local backup/preview counts and log age |

## Route fields

```json
{
  "sender": "sender@example.invalid",
  "subject_contains": "NAV",
  "sheet": "Demo Fund",
  "code": "DEMO01",
  "parser": "auto",
  "allow_sender_only": false,
  "cumulative_policy": "require",
  "cumulative_offset": null,
  "return_basis": "cumulative",
  "return_frequency": "weekly",
  "series_start": "2026-01-01",
  "max_staleness_days": 14,
  "benchmark": null
}
```

- Normalize codes case-insensitively but require exact equality after normalization.
- Store product codes as quoted JSON strings. Numeric JSON values are rejected so leading zeros cannot disappear.
- Use `subject_contains` when an authorized sender also sends non-NAV mail. Every in-scope message must parse; failures block preview instead of being skipped.
- Set `allow_sender_only` only when the sender is permanently dedicated to one product and messages contain no stable code.
- Use `cumulative_policy: require` when cumulative NAV must come from the email.
- Use `unit` only when historical samples prove unit NAV always equals cumulative NAV.
- Use `offset` only with an explicit `cumulative_offset` proven across historical samples.
- Use `series_start` to prevent observation, simulation, or pre-purchase history from entering a new held-position series.
- Choose `daily` or `weekly` return frequency explicitly.
- Set `max_staleness_days` to the product's actual publication cadence plus a conservative holiday buffer. Stale feeds block preview instead of succeeding as a no-op.
- Automatic catch-up appends only dates later than the workbook's latest NAV. A missing internal historical date blocks the run for supervised repair so rows and cross-sheet formulas cannot be silently reordered.

## Column discovery

The runtime scans early workbook rows for semantic headers such as date, product code, product name, unit NAV, cumulative NAV, return, benchmark return/level, and excess. Column order is not fixed.

When headers remain ambiguous, configure a 1-based number or Excel column letter:

```json
{
  "column_overrides": {
    "Demo Fund": {
      "header_row": 2,
      "date": "A",
      "return": "B",
      "name": "C",
      "unit": "D",
      "code": "E",
      "cumulative": "F",
      "benchmark_return": "H",
      "excess": "I"
    }
  }
}
```

Do not use an override to force an uncertain interpretation. Stop and inspect the workbook locally.

## Benchmark mapping

Map only to a workbook sheet whose historical dates and values have been verified:

```json
{
  "benchmark": {
    "source_sheet": "Demo Benchmark",
    "source_type": "aligned_return",
    "source_date": "A",
    "source_value": "B"
  }
}
```

Use `source_type: aligned_return` only when the source column is already aligned to the product's daily or weekly observation dates. A daily index-return column is not a weekly benchmark. Prefer `level` for index levels; the runtime then calculates between matching product-period anchors. Missing required source dates block formal commit.

## Scheduling

```json
{
  "schedule": [
    {"days": ["MON", "TUE", "WED"], "time": "09:30"}
  ]
}
```

Times use the target machine's local timezone. Scheduling is Windows-only, requires a local path and a logged-in user session, and generates previews only. It never writes the master or sends email.

## Validation and retention

`validation.minimum_history_dates` cannot be lower than `2`. `max_future_days` blocks future-dated values, and `max_period_change` blocks implausible unit-NAV jumps before a preview is created. Tune these only from documented product behavior, never merely to make a failing run pass.

`retention.backup_count`, `preview_count`, and `log_days` bound sensitive local artifacts. The runtime prunes only files under its own `backups/`, `previews/`, and `logs/` directories.

## Local sensitive files

Keep these runtime-only and Git-ignored: `config.json`, `route-report.json`, `validation-report.json`, `plan.json`, previews, workbooks, `logs/`, `backups/`, and `scheduled_tasks.json`. The Windows secret is stored under the current user's local application data and encrypted with DPAPI.

The exact parsed `From` address is checked after IMAP search, but this is routing validation, not cryptographic sender authentication. Where spoofing is a material threat, require provider-side DKIM/DMARC controls or a dedicated mailbox rule before enabling the workflow.
