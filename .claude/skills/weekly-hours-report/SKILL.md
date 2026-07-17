---
name: weekly-hours-report
description:
  Generate the Weekly Hours report AND the Monthly Billing invoice for QA teams
  from a single Google Sheet (collaborators + PTO), into ONE Excel file with two
  tabs. Reads PTO absences, team assignments and backup coverage; computes worked
  hours; and bills them using per-collaborator rates.
  Use when asked to generate the weekly hours report, the monthly billing/invoice,
  or process PTO/team data.
---

# Weekly Hours + Monthly Billing

One command produces **one `.xlsx` with two tabs** — `Weekly Hours` and
`Invoice` — from a single Google Sheet ("PTO Tracker Softtek") that holds both the
**collaborators** (tab "Team allocation 2026") and the **monthly PTO** (tabs like
"Jul 2026"). `scripts/main.py` orchestrates everything; `core.py` (shared
pipeline) and `billing.py` (invoice) are libraries.

## Workflow when the skill is invoked

Do these **in order**:

1. **Verify rates** — run `--check-rates` (no month needed). Every one of the ~61
   collaborators must have an hourly rate in `scripts/rates.json`. If any are
   missing, the script writes/updates a `rates.json` template (all collaborators,
   missing ones set to `0`) and lists who needs a rate. **Ask the user to fill in
   those QAs' rates in `scripts/rates.json` to continue**, then re-check.
2. **Ask which month** (and year) — only after rates are complete. Don't assume the
   current month. Confirm as `{Month} {Year}` (e.g. "Jul 2026").
3. **Generate** — run for that month; it creates the single two-tab Excel.

Accept Spanish or English month names, mapping to the English abbreviation the
script expects: Ene/Enero→`Jan`, Feb→`Feb`, Mar/Marzo→`Mar`, Abr/Abril→`Apr`,
May/Mayo→`May`, Jun/Junio→`Jun`, Jul/Julio→`Jul`, Ago/Agosto→`Aug`,
Sep/Septiembre→`Sep`, Oct/Octubre→`Oct`, Nov→`Nov`, Dic/Diciembre→`Dec`.

## Quick start

```bash
# 0. openpyxl available? (one-time)
/usr/bin/python3 -c "import openpyxl" || /usr/bin/pip3 install --user openpyxl

# 1. Check rates (creates/updates scripts/rates.json template if incomplete)
/usr/bin/python3 scripts/main.py --check-rates

# 2. Once every collaborator has a rate, generate for the chosen month
/usr/bin/python3 scripts/main.py --month Jul --year 2026
```

Output → one file `Weekly-Hours-Billing-{Month}-{Year}.xlsx` (tabs: Weekly Hours +
Invoice) in the Shared drive (see [Output location](#output-location)).

> Use `/usr/bin/python3` — NOT the bare `python3` (that resolves to the fbcode
> Python which lacks openpyxl; there is no `python`).

## Before running — authentication

Collaborators AND PTO come from a Google Sheet, so the `meta` CLI must be
authenticated. Quick check:

```bash
meta google.sheets get --id 1Vae2OUAdYT3pMAQLLSYcNRBFklJ2WybctNia6OjNK_g
```

If it lists tabs, you're good. If it fails with OAuth / 401 / auth errors, see
[Error: Auth expired](#error-auth-expired) below.

## Rates — `scripts/rates.json`

The Google Sheet has **no rate column**, so billing rates live in a
per-collaborator JSON file next to the scripts:

```json
{
  "Arath De la Cruz": 42.5,
  "Emmanuel Barrios": 42.5,
  "...": 0
}
```

- Keyed by the collaborator's **QA Analyst** short name (matched case/accent-insensitively).
- Value = hourly rate (USD). Must be **> 0** for every collaborator.
- A **backup** is billed at the rate of the person they cover (looked up by name).
- `--check-rates` validates completeness and regenerates the template (preserving
  existing values, missing → `0`) so the user only fills in numbers.
- `DISCOUNT_RATE` (2% volume discount) is set at the top of `billing.py`.
- `rates.json` holds pay data — keep it out of version control (gitignore it).

## How it runs (`scripts/main.py`)

`main.py` is the orchestrator; logic lives in `core.py` (shared) and `billing.py`.

1. **Fetch** — `core.fetch_sheet()` pulls the "Team allocation 2026" tab and the
   month's PTO tab via `meta google.sheets read` (concurrently), caching to
   `scripts/cache/sheet-{Month}-{Year}.xlsx` — see [Caching & performance](#caching--performance).
2. **Read** — `read_pto()` (absences/holidays/backups) and `read_team()` (roles, products).
3. **Validate rates** — every collaborator must have a rate (else template + stop).
4. **Weekly** — `build_report()` → Q1/Q2/Q3 structures with worked hours.
5. **Billing** — `build_billing(weekly, rates)` → line items, 2% discount, total.
6. **Export** — one `Workbook` gets `write_weekly_sheet()` + `write_invoice_sheet()`,
   saved as a single two-tab file.

> All **input data** (collaborators, absences, backups) comes from the single
> Google Sheet below — no other source is read (rates are the local `rates.json`).
> The finished workbook is then *written* to Google Drive.

### Flags

| Flag | Purpose |
|------|---------|
| `--month` | Month name (Jan–Dec). Default: current month. |
| `--year` | Year. Default: current year. |
| `--check-rates` | Validate `rates.json` against the roster and exit (no month needed). |
| `--rates <path>` | Use a specific rates.json (default: `scripts/rates.json`). |
| `--sheet-id <id>` | Override the Google Sheet ID (default baked into `core.py`). |
| `--sheet-xlsx <path>` | Reuse an already-downloaded Sheet copy instead of downloading. |
| `--no-cache` | Force a fresh download even if a valid cache exists. |
| `--output <path>` | Override the output `.xlsx` path. |

### Caching & performance

~95% of a run is network time on `meta google.sheets`. To keep reruns fast,
`fetch_sheet()` (in `core.py`):

- Reads the **team tab and month tab concurrently** (one round-trip's wall-clock).
- Caches the download at `scripts/cache/sheet-{Month}-{Year}.xlsx`:
  - **Within 2 minutes** → reuse the cache with **zero network** (~0.2s total).
  - **After 2 minutes** → check the Sheet's `modifiedTime` (via
    `google.sheets describe`); reuse if unchanged, else re-download.
- Use `--no-cache` to always re-download.

## Data source — one Google Sheet ("PTO Tracker Softtek")

ID `1Vae2OUAdYT3pMAQLLSYcNRBFklJ2WybctNia6OjNK_g`
https://docs.google.com/spreadsheets/d/1Vae2OUAdYT3pMAQLLSYcNRBFklJ2WybctNia6OjNK_g

The only source of input truth — collaborators, absences (PTO/ML), holidays,
backups. Monthly tabs are named in English abbreviations (`Jun 2026`, `Jul 2026`,
`Aug 2026`), **not** Spanish (there is no "Junio 2026").

### PTO — monthly tab
- Tab name format: `{Month} {Year}` (e.g. `Jul 2026`; September also matches `Sept 2026`).
- **Row 8** is the header: `QA3 | Collab | <day numbers…> | Backup | Notes`.
- **Row 9+**: one row per collaborator; the name lives in the **Collab** column
  and is the key used to match against the team tab.
- Absence codes: `PTO`, `PTO(PA)`, `ML`, `ML(PA)`, `H` (holiday), `Bench`.

### Collaborators — tab "Team allocation 2026"
- **Header on row 1.** Columns: `App | Feature | QA Lead | Role | QA Analyst | QA Analyst (Full Name) | New QA3`.
- Column mapping:

  | Sheet column | Used as |
  |--------------|---------|
  | `QA Analyst` | collaborator short name (match key vs PTO **Collab**, and rates key) |
  | `QA Analyst (Full Name)` | full name |
  | `App` | Product |
  | `Feature` | Pilar |
  | `Role` | QA grouping (`QA 1`→Q1, `QA 2`→Q2, `QA 3`→Q3) |
  | `QA Lead` / `New QA3` | reference only |

- **Tag** = `QA Analyst - App - Feature` (e.g. `Arath De la Cruz - Central Products - MAA`).
- A collaborator in the PTO tab but **not** in "Team allocation 2026" is skipped
  with a `WARN … not in team allocation` (expected).

## Business rules

- **Worked hours are computed**, not read: 8h × working days (Mon–Fri), minus
  absences, holidays and bench days. There is no worked-hours column in the Sheet.
- Standard day = **8 hours**; only **Mon–Fri** count as working days.
- **Holidays** (`H`) do not count as worked hours or absences.
- **Bench** days (`TEAM: Bench | Bench <range>` note) subtract worked hours but are not absences.
- **Backups** (`↳` rows) cover an absent collaborator. Credited only on the covered
  person's **actual PTO/ML days**; a backup day that is a weekend, holiday, or bench
  day is **skipped** with a `WARN` (never fails). Billed at the covered person's rate.
- **Emergency** highlight (red) when a collaborator's absence hours exceed 80.
- Collaborators are grouped into Q1/Q2/Q3 by their `Role`.

### Expected warnings (not errors)

- `WARN <name> not in team allocation, skipping` — in a PTO tab but not the team tab.
- `WARN <name>: backup <person> on day <n> without PTO/ML, skipping that day` — a
  backup range covers a non-PTO/ML day (e.g. bench); that day is ignored.
- `WARN no rate for <name>` — should not happen after `--check-rates` passes.

## Output location

One two-tab workbook saved into the team's **Shared drive** via the locally-synced
Google Drive for Desktop mount:

```
Shared drives/Meta - STK/Project Tracking/Automation/Automation Outputs/Weekly Hours Report/
  └─ Weekly-Hours-Billing-{Month}-{Year}.xlsx   (tabs: Weekly Hours + Invoice)
```

- The `GoogleDrive-<account>` mount is **auto-detected** from whoever runs the
  skill (`~/Library/CloudStorage/GoogleDrive-*`) — works on any team member's computer.
- We write to the **local synced folder** and let Drive for Desktop push it up.
  The `meta google.drive upload` API is **not** used (blocked by corpnet on laptops).
- If the Shared drive isn't mounted, it falls back to `scripts/output/` with a `WARN`.
- Override with `--output <path>`.

## Output format

**Weekly Hours tab:**
- Sections Q1/Q2/Q3 (dark-blue headers); columns Employee, Role, Wave, Product,
  Pilar, per-week day breakdown + Hrs, Abs Hrs, Work Hrs, Comments, Tag.
- Cell colors: PTO (blue), PTO(PA)/ML (yellow), Holiday (purple), Bench (gray),
  Emergency (red), Backup rows (cream), Totals (green).

**Invoice tab:**
- Softtek → Meta invoice: header (bill/ship to, PO, dates), line items grouped as
  QA Analyst I/II/III with QTY (hours) × RATE = AMOUNT, sub-total, 2% volume
  discount, and total (with amount in words).

## Error: Auth expired

If any step fails with "OAuth", "401", or "auth" errors:

1. Run `jf auth`. On x2p devservers this reports success but may not write a token —
   if reads still fail, use the legacy flow below.
2. Go to https://www.internalfb.com/intern/jf/authenticate/
3. Open the **"Legacy Options"** section and copy the **UID** and **NONCE**.
4. Run: `jf auth --skip-legacy-auth-upgrade <UID> <NONCE>`
5. Re-run the skill.
