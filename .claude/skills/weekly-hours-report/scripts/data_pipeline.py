#!/usr/bin/env python3
"""Weekly Hours + Monthly Billing — entry point.

/usr/bin/python3 scripts/run.py --check-rates
/usr/bin/python3 scripts/run.py --month Jul --year 2026
"""
import argparse
import json
import subprocess
import sys
import tempfile
import unicodedata
from calendar import monthrange
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime
from pathlib import Path

from billing_report_builder import (
    build_billing,
    check_rates,
    load_rates,
    RATES_PATH_DEFAULT,
    TEMPLATE_PATH,
    write_invoice_sheet,
)
from openpyxl import load_workbook, Workbook
from report_builder import build_report
from weekly_report_builder import write_weekly_sheet


# ===========================================================================
# Helpers
# ===========================================================================
MONTHS = {
    "Jan": 1,
    "January": 1,
    "Feb": 2,
    "February": 2,
    "Mar": 3,
    "March": 3,
    "Apr": 4,
    "April": 4,
    "May": 5,
    "Jun": 6,
    "June": 6,
    "Jul": 7,
    "July": 7,
    "Aug": 8,
    "August": 8,
    "Sep": 9,
    "Sept": 9,
    "September": 9,
    "Oct": 10,
    "October": 10,
    "Nov": 11,
    "November": 11,
    "Dec": 12,
    "December": 12,
}
MONTH_NAMES_SHORT = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
]
MONTH_NAMES_FULL = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
]


def normalize_value(value):
    if not value:
        return ""
    s = unicodedata.normalize("NFD", str(value))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.lower().replace(" ", "").strip()


def get_month_number(month):
    n = MONTHS.get(month)
    if not n:
        raise ValueError(f"Invalid month: {month}")
    return n


def get_month_name(month_number, short=True):
    return (
        MONTH_NAMES_SHORT[month_number - 1]
        if short
        else MONTH_NAMES_FULL[month_number - 1]
    )


def month_tab_candidates(month, year):
    short = get_month_name(get_month_number(month), True)
    aliases = ["Sep", "Sept"] if short == "Sep" else [short]
    return [f"{a} {year}" for a in aliases]


# ===========================================================================
# Fetch — Google Sheet I/O
# ===========================================================================
DEFAULT_SHEET_ID = "1Vae2OUAdYT3pMAQLLSYcNRBFklJ2WybctNia6OjNK_g"
TEAM_TAB = "Team allocation 2026"
TEAM_RANGE = "A1:H500"
MONTH_RANGE = "A1:AC200"
SHEET_CACHE_MAX_AGE = 10 * 60  # 10 minutes


def _read_tab(sheet_id, tab, cell_range):
    res = subprocess.run(
        [
            "meta",
            "google.sheets",
            "read",
            "--id",
            sheet_id,
            "--range",
            f"'{tab}'!{cell_range}",
            "-o",
            "json",
        ],
        capture_output=True,
        text=True,
    )
    out = (res.stdout or "").strip()
    start = out.find("[")
    if start == -1:
        raise RuntimeError(f"Could not read tab '{tab}'. {res.stderr.strip() or out}")
    data = json.loads(out[start:])
    if isinstance(data, dict):
        raise RuntimeError(f"Could not read tab '{tab}': {data.get('message', data)}")
    return data


def _try_read_tab(sheet_id, tab, cell_range):
    try:
        return _read_tab(sheet_id, tab, cell_range) or None
    except RuntimeError:
        return None


def fetch_sheet(month, year, sheet_id=None, force_refresh=False):
    """Download team + monthly PTO tabs into a local .xlsx.
    Uses a 10-min cache to avoid re-downloading on repeated runs."""
    sheet_id = sheet_id or DEFAULT_SHEET_ID
    dest = Path(tempfile.gettempdir()) / f"weekly-hours-sheet-{month}-{year}.xlsx"

    # Cache: reuse if < 10 min old (unless forced)
    if not force_refresh and dest.exists():
        age = datetime.now().timestamp() - dest.stat().st_mtime
        if age < SHEET_CACHE_MAX_AGE:
            print(f"  Using cached sheet ({int(age)}s old, max {SHEET_CACHE_MAX_AGE}s)")
            return dest

    candidates = month_tab_candidates(month, year)

    with ThreadPoolExecutor(max_workers=len(candidates) + 1) as ex:
        f_team = ex.submit(_read_tab, sheet_id, TEAM_TAB, TEAM_RANGE)
        f_month = {
            c: ex.submit(_try_read_tab, sheet_id, c, MONTH_RANGE) for c in candidates
        }
        team_rows = f_team.result()
        month_title = month_rows = None
        for cand in candidates:
            rows = f_month[cand].result()
            if rows:
                month_title, month_rows = cand, rows
                break

    if month_rows is None:
        raise RuntimeError(
            f"No monthly tab found for {month} {year} (tried {candidates})."
        )

    wb = Workbook()
    wb.remove(wb.active)
    for title, rows in [(TEAM_TAB, team_rows), (month_title, month_rows)]:
        ws = wb.create_sheet(title)
        for r, row in enumerate(rows, 1):
            for c, val in enumerate(row, 1):
                if val not in ("", None):
                    ws.cell(r, c, val)
    dest.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dest)
    return dest


# ===========================================================================
# Team cache (1 week TTL)
# ===========================================================================


def fetch_team_roster(sheet_id=None):
    """Fetch team roster directly (no cache — use fetch_sheet cache instead)."""
    return _parse_team_rows(
        _read_tab(sheet_id or DEFAULT_SHEET_ID, TEAM_TAB, TEAM_RANGE)
    )


# ===========================================================================
# Parsers — PTO + Team from local xlsx
# ===========================================================================
PTO_HEADER_ROW = 8
DATA_START_ROW = 9
COLLAB_COL = 2
FIRST_DAY_COL = 3
HOLIDAY_PURPLE = "FFB4A7D6"

TEAM_COLS = {
    "team": "App",
    "module": "Feature",
    "job_role": "Role",
    "short_name": "QA Analyst",
    "full_name": "QA Analyst (Full Name)",
    "wave": "Wave",
}


def _is_purple(cell):
    fill = cell.fill
    if not fill or fill.patternType is None:
        return False
    fg = fill.fgColor
    if not fg or not fg.rgb:
        return False
    rgb = fg.rgb.upper()
    return HOLIDAY_PURPLE in rgb or rgb in ("FF00FF", "800080", "FFC000FF", "FF800080")


def _parse_team_rows(rows):
    if not rows:
        return []
    header = {str(h).strip(): i for i, h in enumerate(rows[0]) if h not in (None, "")}
    out = []
    for r in rows[1:]:

        def get(col_name):
            i = header.get(col_name)
            return (
                str(r[i]).strip()
                if (i is not None and i < len(r) and r[i] not in (None, ""))
                else ""
            )

        short = get(TEAM_COLS["short_name"])
        if not short:
            continue
        rec = {}
        for field, col_name in TEAM_COLS.items():
            val = get(col_name)
            rec[field] = val if val else ("" if field == "short_name" else "Missing")
        out.append(rec)
    return out


def read_team(team_path):
    p = Path(team_path)
    if not p.exists():
        raise FileNotFoundError(f"Team file not found at {p}")
    wb = load_workbook(p, data_only=True)
    ws = wb[TEAM_TAB]
    rows = [
        [c.value for c in row]
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column)
    ]
    return _parse_team_rows(rows)


def read_pto(month, year, pto_path):
    p = Path(pto_path)
    if not p.exists():
        raise FileNotFoundError(f"PTO file not found at {p}")
    wb = load_workbook(p, data_only=True)
    ws = None
    for name in month_tab_candidates(month, year):
        if name in wb.sheetnames:
            ws = wb[name]
            break
    if not ws:
        raise ValueError(f"Sheet not found for {month} {year}")

    header = ws[PTO_HEADER_ROW]
    backup_col = notes_col = None
    for idx, cell in enumerate(header, start=1):
        v = str(cell.value or "").strip().lower()
        if v == "backup":
            backup_col = idx
        if v == "notes":
            notes_col = idx
    if backup_col is None:
        backup_col = 24
    if notes_col is None:
        notes_col = backup_col + 1

    day_cols = []
    for col_idx in range(FIRST_DAY_COL, backup_col):
        cell = ws.cell(PTO_HEADER_ROW, col_idx)
        try:
            day_num = int(cell.value)
        except (TypeError, ValueError):
            continue
        if 1 <= day_num <= 31:
            day_cols.append((col_idx, day_num, _is_purple(cell)))

    result = {}
    for row_idx in range(DATA_START_ROW, ws.max_row + 1):
        name = str(ws.cell(row_idx, COLLAB_COL).value or "").strip()
        if not name:
            continue
        absences = []
        for col_idx, day_num, is_hol in day_cols:
            if is_hol:
                absences.append({"dayNumber": day_num, "type": "H"})
                continue
            v = ws.cell(row_idx, col_idx).value
            if v:
                absences.append({"dayNumber": day_num, "type": str(v).strip()})
        result[normalize_value(name)] = {
            "employeeName": name,
            "absences": absences,
            "backupsText": str(ws.cell(row_idx, backup_col).value or "").strip(),
            "notesText": str(ws.cell(row_idx, notes_col).value or "").strip(),
        }
    return result


# ===========================================================================
# Output path
# ===========================================================================
DRIVE_OUTPUTS_REL = "Shared drives/Softtek - Time Report & Invoice"


def resolve_output_path(month, year, override):
    month_num = get_month_number(month)
    month_full = get_month_name(month_num, short=False)
    fname = f"Softtek Time Report & Invoice - {month_full} {year}.xlsx"
    if override:
        return Path(override)
    home = Path.home()
    user = home.name
    cloud_root = home / "Library/CloudStorage"
    outputs = cloud_root / f"GoogleDrive-{user}@meta.com" / DRIVE_OUTPUTS_REL
    if not outputs.exists():
        for cloud in sorted(cloud_root.glob("GoogleDrive-*")):
            candidate = cloud / DRIVE_OUTPUTS_REL
            if candidate.exists():
                outputs = candidate
                break
    if not outputs.exists():
        raise FileNotFoundError(
            f"Google Drive not mounted or Shared drive missing.\n"
            f"Looked under: {cloud_root}/GoogleDrive-*/{DRIVE_OUTPUTS_REL}\n"
            f"Make sure Google Drive for Desktop is running and synced."
        )
    return outputs / fname


# ===========================================================================
# CLI
# ===========================================================================


def main():
    ap = argparse.ArgumentParser(description="Generate Weekly Hours + Monthly Billing")
    ap.add_argument("--month", help="Month name (Jan-Dec), default current")
    ap.add_argument("--year", type=int, help="Year, default current")
    ap.add_argument("--sheet-id", default=DEFAULT_SHEET_ID)
    ap.add_argument("--sheet-xlsx", help="Local Sheet xlsx (skip download)")
    ap.add_argument("--rates", help="Path to rates.json")
    ap.add_argument(
        "--check-rates", action="store_true", help="Validate rates and exit"
    )
    ap.add_argument(
        "--refresh", action="store_true", help="Force re-download (ignore cache)"
    )
    ap.add_argument("--output", help="Output path override")
    args = ap.parse_args()

    rates_path = Path(args.rates) if args.rates else RATES_PATH_DEFAULT

    if args.check_rates:
        print("Validating rates...")
        team = fetch_team_roster(sheet_id=args.sheet_id)
        print(f"  {len(team)} collaborators")
        rates = load_rates(rates_path)
        missing = check_rates(team, rates)
        if missing:
            print(f"\nMISSING RATES ({len(missing)}):")
            for n in missing:
                print(f"  - {n}")
            print(f"\nAdd their rates to {rates_path} and re-run.")
            sys.exit(2)
        print(f"\nAll {len(team)} rates OK. Ready.")
        sys.exit(0)

    now = datetime.now()
    month = args.month or MONTH_NAMES_SHORT[now.month - 1]
    year = args.year or now.year
    print(f"Generating for {month} {year}...")

    if args.sheet_xlsx:
        sheet_xlsx = Path(args.sheet_xlsx)
    else:
        print("Fetching Google Sheet...")
        sheet_xlsx = fetch_sheet(
            month, year, sheet_id=args.sheet_id, force_refresh=args.refresh
        )

    pto = read_pto(month, year, pto_path=sheet_xlsx)
    team = read_team(team_path=sheet_xlsx)
    print(f"  PTO: {len(pto)} | Team: {len(team)}")

    rates = load_rates(rates_path)
    missing = check_rates(team, rates)
    if missing:
        print(f"MISSING RATES ({len(missing)}):")
        for n in missing:
            print(f"  - {n}")
        print(f"\nAdd their rates to {rates_path} and re-run.")
        sys.exit(2)

    weekly = build_report(pto, team, month, year)
    billing = build_billing(weekly, rates)
    print(f"  Total: ${billing['total']:,.2f}")

    if not TEMPLATE_PATH.exists():
        print(f"ERROR: template not found at {TEMPLATE_PATH}")
        sys.exit(3)
    wb = load_workbook(TEMPLATE_PATH)
    invoice_ws = wb[wb.sheetnames[0]]
    invoice_ws.title = "Invoice"
    work_refs, tag_refs = write_weekly_sheet(wb, weekly)
    wb.move_sheet("Weekly Hours", offset=-1)
    write_invoice_sheet(
        invoice_ws, billing, month, year, work_refs=work_refs, tag_refs=tag_refs
    )

    out_path = resolve_output_path(month, year, args.output)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    print(f"\nDone! {out_path}")


if __name__ == "__main__":
    main()
