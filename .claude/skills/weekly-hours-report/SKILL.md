---
name: weekly-hours-report
description: Generates a Weekly Hours report + Monthly Billing invoice (.xlsx, two tabs) from the "PTO Tracker Softtek" Google Sheet. Computes worked hours and bills using per-collaborator rates. Use when asked to generate the weekly hours report, monthly billing/invoice, or process PTO/team data.
---

# weekly-hours-report

Generates ONE `.xlsx` with two tabs (`Weekly Hours` + `Invoice`) from a single Google Sheet.

```
scripts/
  data_pipeline.py         — entry point (fetch, parse, cache, CLI)
  report_builder.py        — compute worked hours + backups (business logic)
  weekly_report_builder.py — Weekly Hours Excel tab (formatting)
  billing_report_builder.py — Invoice Excel tab (rates, template fill)
```

## Quick start

```bash
/usr/bin/python3 -c "import openpyxl" || /usr/bin/pip3 install --user openpyxl
/usr/bin/python3 scripts/data_pipeline.py --month Jul --year 2026
```

> Use `/usr/bin/python3` — NOT bare `python3`.

## Workflow

1. **Check rates file** — before anything, verify `assets/rates.json` exists.
   - If it **does NOT exist**: ask the user to provide the rates in **any format they
     want** — a list, a table, a screenshot, pasted text, etc. They are NOT technical.
     Do NOT ask for JSON. Accept whatever they give you (names + numbers) and YOU
     transform it into the correct JSON format and create `assets/rates.json`.
   - If it **exists**: run `--check-rates` to validate all collaborators have a rate > 0.
     If any are missing, list the names and ask "what's the rate for these people?"
     Accept the answer in any format and update the file. **STOP until resolved.**
2. **Ask month/year** — only after rates pass. Present these options:
   ```
   ¿Para qué mes genero el reporte?
   A) {previous month} {year}
   B) {current month} {year}
   C) Otro — escríbelo
   ```
   Use the actual month names in Spanish. Accept Spanish or English.
3. **Generate** — run for that month.

Month: Ene→`Jan`, Feb→`Feb`, Mar→`Mar`, Abr→`Apr`, May→`May`, Jun→`Jun`,
Jul→`Jul`, Ago→`Aug`, Sep→`Sep`, Oct→`Oct`, Nov→`Nov`, Dic→`Dec`.

## Flags

| Flag | Purpose |
|------|---------|
| `--month` / `--year` | Target month/year. Default: current. |
| `--check-rates` | Validate `rates.json` and exit. |
| `--refresh` | Force re-download (ignore 10-min cache). |
| `--rates <path>` | Alternate rates.json. |
| `--sheet-id <id>` | Override Google Sheet ID. |
| `--sheet-xlsx <path>` | Use local Sheet (skip download). |
| `--output <path>` | Override output path. |

## Authentication

```bash
meta google.sheets get --id 1Vae2OUAdYT3pMAQLLSYcNRBFklJ2WybctNia6OjNK_g
```

Auth error → `jf auth` via legacy:
1. https://www.internalfb.com/intern/jf/authenticate/ → **Legacy Options** → UID + NONCE.
2. `jf auth --skip-legacy-auth-upgrade <UID> <NONCE>`

## Rules

- **rates.json**: the user gives you names + rates in any format. YOU create/update the JSON file. Never ask for JSON — they are not technical.
- If rates check fails, **STOP completely**. Do NOT ask for month. Say "tell me when
  rates are ready" and wait.
- Never assume the month — always ask.
- Sheet data is cached 10 min. Use `--refresh` to force re-download.
- Output path controlled by `DRIVE_OUTPUTS_REL` in `data_pipeline.py`. Override with `--output`.
- `rates.json` is pay data — keep out of version control.
