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

One command → **one `.xlsx`, two tabs** (`Weekly Hours` + `Invoice`) from a single
Google Sheet ("PTO Tracker Softtek") holding both the **collaborators** (tab "Team
allocation 2026") and the **monthly PTO** (tabs like "Jul 2026"). `main.py`
orchestrates; `core.py` (pipeline) and `billing.py` (invoice) are libraries.

## Workflow (in order)

1. **Verify rates** — `--check-rates` (no month). Every collaborator (~61) needs an
   hourly rate in `assets/rates.json`. If any are missing the script rewrites the
   template (missing → `0`) and lists them → **ask the user to fill those rates**,
   then re-check.
2. **Ask which month/year** — only after rates pass. Don't assume the current month;
   confirm as `{Month} {Year}` (e.g. "Jul 2026").
3. **Generate** — run for that month.

Month names: accept Spanish or English, pass the **English abbreviation** to the
script — Ene→`Jan`, Feb→`Feb`, Mar→`Mar`, Abr→`Apr`, May→`May`, Jun→`Jun`,
Jul→`Jul`, Ago→`Aug`, Sep→`Sep`, Oct→`Oct`, Nov→`Nov`, Dic→`Dec`.

## Quick start

```bash
# openpyxl available? (one-time)
/usr/bin/python3 -c "import openpyxl" || /usr/bin/pip3 install --user openpyxl

# 1) check rates   2) generate once every rate is set
/usr/bin/python3 scripts/main.py --check-rates
/usr/bin/python3 scripts/main.py --month Jul --year 2026
```

Output → `Weekly-Hours-Billing-{Month}-{Year}.xlsx` in the Shared drive (see
[Output](#output)).

> Use `/usr/bin/python3` — NOT bare `python3` (fbcode Python, no openpyxl).

## Authentication

Data comes from a Google Sheet, so `meta` must be authenticated. Check:

```bash
meta google.sheets get --id 1Vae2OUAdYT3pMAQLLSYcNRBFklJ2WybctNia6OjNK_g
```

Lists tabs → good. OAuth / 401 / auth error → see [Auth expired](#auth-expired).

## Rates — `assets/rates.json`

The Sheet has **no rate column**, so rates live in a local JSON:

```json
{ "Arath De la Cruz": 42.5, "Emmanuel Barrios": 42.5, "...": 0 }
```

- Keyed by the **QA Analyst** short name (matched case/accent-insensitively).
- Value = hourly USD rate, must be **> 0** for everyone.
- A **backup** is billed at the rate of the person they cover.
- `--check-rates` validates and regenerates the template (keeps existing values,
  missing → `0`) so the user only fills numbers.
- `DISCOUNT_RATE` (2% volume discount) is at the top of `billing.py`.
- `rates.json` is pay data — keep it out of version control.

## How it runs

1. **Fetch** — `fetch_sheet()` reads the team tab + month PTO tab via
   `meta google.sheets read` (concurrently), to a fresh temp copy (**no cache**).
2. **Read** — `read_pto()` (absences/holidays/backups) + `read_team()` (roles/products).
3. **Validate rates** — every collaborator must have a rate (else template + stop).
4. **Weekly** — `build_report()` → Q1/Q2/Q3 with computed worked hours.
5. **Billing** — `build_billing()` → line items, 2% discount, total.
6. **Export** — the invoice **template** (`assets/Monthly Billing Report -
   Template.xlsx`) is loaded as the workbook; `write_weekly_sheet()` adds the
   Weekly Hours tab (returning each collaborator's Work Hrs **and Tag** cell) and
   `write_invoice_sheet()` fills the template's Invoice tab in place (line items +
   totals), using those maps for [live formulas](#live-links).

> All **input** comes from the Sheet below (rates from local `rates.json`); the
> finished workbook is *written* to Google Drive. Every run downloads fresh.

### Flags

| Flag | Purpose |
|------|---------|
| `--month` / `--year` | Target month (Jan–Dec) / year. Default: current. |
| `--check-rates` | Validate `rates.json` and exit (no month needed). |
| `--rates <path>` | Alternate rates.json (default: `assets/rates.json`). |
| `--sheet-id <id>` | Override the Google Sheet ID. |
| `--sheet-xlsx <path>` | Reuse an already-downloaded Sheet copy. |
| `--output <path>` | Override the output `.xlsx` path. |

## Data source — one Google Sheet ("PTO Tracker Softtek")

ID `1Vae2OUAdYT3pMAQLLSYcNRBFklJ2WybctNia6OjNK_g` — the only input truth
(collaborators, PTO/ML, holidays, backups). Monthly tabs use **English
abbreviations** (`Jul 2026`), never Spanish (no "Julio 2026").

**PTO — monthly tab** `{Month} {Year}` (Sep also matches `Sept 2026`):
- **Row 8** = header `QA3 | Collab | <day #s…> | Backup | Notes`.
- **Row 9+** = one row per collaborator; the **Collab** column is the match key.
- Codes: `PTO`, `PTO(PA)`, `ML`, `ML(PA)`, `H` (holiday), `Bench`, `O` (out of
  project — gray).

**Collaborators — "Team allocation 2026"** (header on **row 1**):

| Sheet column | Used as |
|--------------|---------|
| `QA Analyst` | short name — match key vs PTO **Collab** and rates key |
| `QA Analyst (Full Name)` | full name |
| `App` → Product · `Feature` → Pilar | |
| `Role` | Q grouping: `QA 1`→Q1, `QA 2`→Q2, `QA 3`→Q3 |

- **Tag** = `QA Analyst - App - Feature`.
- In PTO but not in team allocation → skipped with a `WARN` (expected).

## Business rules

- **Worked hours are computed** (not read): 8h × working days (Mon–Fri) − absences −
  holidays − bench. No worked-hours column exists in the Sheet.
- Standard day = **8h**; only **Mon–Fri** count.
- **Holidays** (`H`) count as neither worked nor absence.
- **Bench** (`TEAM: Bench | Bench <range>` note) subtracts worked hours, not absence.
- **Out of project** (`O`, gray) — the collaborator is off the project those days:
  it subtracts worked hours (not absence), and — like PTO/ML — is a **coverable**
  day, so a backup gets credited for covering it.
- **Backups** (`↳` rows) credited only on the covered person's **coverable days**
  (`PTO`/`ML`/`O`); a backup day that's weekend/holiday/bench is **skipped** with a
  `WARN`. Billed at the covered person's rate.
- **Emergency** red highlight when a collaborator's absence hours exceed 80.

### Expected warnings (not errors)

- `WARN <name> not in team allocation, skipping` — in PTO but not team tab.
- `WARN <name>: backup <person> on day <n> without PTO/ML/O, skipping that day`.
- `WARN no rate for <name>` — shouldn't happen after `--check-rates` passes.

## Output

Saved into the team **Shared drive** via the synced Google Drive for Desktop mount:

```
Shared drives/Meta - STK/Project Tracking/Automation/Automation Outputs/Weekly Hours Report/
  └─ Weekly-Hours-Billing-{Month}-{Year}.xlsx
```

- Mount auto-detected from `~/Library/CloudStorage/GoogleDrive-*` (any team member).
- Written to the local synced folder — Drive pushes it up. `meta google.drive
  upload` is **not** used (corpnet-blocked on laptops).
- No mount → falls back to your **home folder** with a `WARN`. Override: `--output`.

**Weekly Hours tab:** Q1/Q2/Q3 sections; columns Employee, Role, Wave, Product,
Pilar, then **per week only the Mon–Fri day columns + a weekly Hrs (sum) column**
(no per-week date-range/label column), then Abs Hrs, Work Hrs, Comments, Tag. The
`Week N` header on the top row is **merged across its whole week block**. Headers
are **color-coded by group** with white text — identity (1-5) **navy** | each Week
block **gold** | Abs+Work **green** | Comments+Tag **purple** — and **dark divider
lines** frame those same groups on top of the colors (light-gray gridlines
inside). Cell colors: PTO blue, PTO(PA)/ML yellow, Holiday purple, Bench gray,
Out-of-project (`O`) gray, Emergency red, Backup cream, Totals green.

**Invoice tab:** built by **filling the template** `assets/Monthly Billing Report -
Template.xlsx` (bill/ship to, PO, footer and styling all come from it). Only the
DATE (**centered** in `I8:L8`), the **Services Period** line in `C18`
(`Services Period: {m}/1/{year} to {m}/{lastDay}/{year}`), the QA Analyst I/II/III
line items (QTY × RATE = AMOUNT, one inserted row per collaborator under each
section marker, then a **blank spacer row after each QA section** so they aren't
cramped) and the totals — sub-total, 2% discount, total (+ amount in words) — are
written in. The line-item **DESCRIPTION** (column C) is a **live link** to the
collaborator's Tag cell in Weekly Hours (not static text). Everything written uses
**Times New Roman** and collaborator names are **not bold**.

### Live links

The Invoice uses **formulas**, so a manual edit to a collaborator's hours in
Weekly Hours flows into the Invoice **without re-running the skill**:

- **Work Hrs** (Weekly Hours) = `SUM(weekly Hrs cells)` → editing any weekly Hrs
  cell recalculates that collaborator's Work Hrs.
- **QTY** (Invoice) = `='Weekly Hours'!<cell>` → that same collaborator's **Work
  Hrs** cell, so the change lands on the matching Invoice line.
- **DESCRIPTION** (Invoice, column C) = `='Weekly Hours'!<cell>` → that
  collaborator's **Tag** cell, so editing the Tag in Weekly Hours updates the
  invoice name/label too. The column letter is computed per month (it shifts with
  the number of day columns), never hard-coded.
- **AMOUNT** = `QTY × RATE`; **SUB-TOTAL** = `SUM(amounts)`;
  **Discount** = `SUB-TOTAL × 2%`; **TOTAL** = `SUB-TOTAL − Discount`.

> Caveat: the spelled-out **amount in words** is fixed at generation time (Excel
> can't spell a number) — re-run the skill if a hand-edited total must match.

## Auth expired

On "OAuth" / "401" / "auth" errors:

1. `jf auth` (on x2p devservers may report success but not write a token → use legacy).
2. https://www.internalfb.com/intern/jf/authenticate/ → **Legacy Options** → copy UID + NONCE.
3. `jf auth --skip-legacy-auth-upgrade <UID> <NONCE>`
4. Re-run the skill.
