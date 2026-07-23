"""Worked hours computation — business logic.

Combines PTO absences + team data into Q1/Q2/Q3 report structures.
Pure logic: no I/O, no formatting.
"""

import re
import unicodedata
from calendar import monthrange
from collections import defaultdict
from datetime import date


# ===========================================================================
# Shared helpers (duplicated here to keep hours.py self-contained)
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


def _normalize(value):
    if not value:
        return ""
    s = unicodedata.normalize("NFD", str(value))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.lower().replace(" ", "").strip()


def _month_num(month):
    n = MONTHS.get(month)
    if not n:
        raise ValueError(f"Invalid month: {month}")
    return n


def _week_of_month(d):
    first = date(d.year, d.month, 1)
    return (d.day + first.weekday() - 1) // 7 + 1


def _workdays(year, month_num):
    days_in_month = monthrange(year, month_num)[1]
    return [
        date(year, month_num, day)
        for day in range(1, days_in_month + 1)
        if date(year, month_num, day).weekday() < 5
    ]


def _group_consecutive(days):
    if not days:
        return []
    days = sorted(days)
    ranges = []
    start = end = days[0]
    for d in days[1:]:
        if d == end + 1:
            end = d
        else:
            ranges.append((start, end))
            start = end = d
    ranges.append((start, end))
    return ranges


# ===========================================================================
# Constants
# ===========================================================================
ABSENCE_TYPES = {"pto", "pto(pa)", "ml", "ml(pa)"}
OUT_TYPES = {"o"}
COVERABLE_TYPES = ABSENCE_TYPES | OUT_TYPES
SILENT_SKIP = {"anacortez"}


# ===========================================================================
# Internal parsers (backup text, bench, team changes)
# ===========================================================================


def _parse_backups(text):
    if not text.strip():
        return []
    out = []
    for line in text.split("\n"):
        line = re.sub(r"\(.*?\)", "", line).strip()
        if not line:
            continue
        m = re.match(r"^[A-Za-z\u00C0-\u00FF\s]+", line)
        if not m:
            continue
        name = m.group(0).strip()
        rest = line[len(m.group(0)) :].strip()
        for rng in rest.split(","):
            rng = rng.strip()
            if not rng:
                continue
            parts = [p.strip() for p in rng.split("-")]
            try:
                start = int(parts[0])
                end = int(parts[1]) if len(parts) > 1 else start
            except (ValueError, IndexError):
                continue
            out.append(
                {
                    "backupName": _normalize(name),
                    "originalBackupName": name,
                    "startDay": start,
                    "endDay": end,
                }
            )
    return out


def _parse_bench_days(notes, year, month_num):
    bench = set()
    for m in re.finditer(r"TEAM:\s*(.+?)\s*\|\s*(.+?)\s+(\d+)-(\d+)", notes):
        if "bench" not in m.group(1).lower():
            continue
        for day in range(int(m.group(3)), int(m.group(4)) + 1):
            try:
                d = date(year, month_num, day)
            except ValueError:
                continue
            if d.weekday() < 5:
                bench.add(day)
    return bench


def _parse_team_changes(notes):
    return [
        {
            "product": m.group(1).strip(),
            "pilar": m.group(2).strip(),
            "fromDay": int(m.group(3)),
            "toDay": int(m.group(4)),
        }
        for m in re.finditer(r"TEAM:\s*(.+?)\s*\|\s*(.+?)\s+(\d+)-(\d+)", notes)
    ]


def _build_comments(absences_by_type, team_changes):
    parts = []
    if team_changes:
        parts.extend(
            f"\u2022 Team change: {c['product']} | {c['pilar']} (days {c['fromDay']}-{c['toDay']})"
            for c in team_changes
        )
    for typ, days in absences_by_type.items():
        if _normalize(typ) in ("h", "holiday"):
            continue
        ranges = _group_consecutive(sorted(days))
        rng_txt = ",".join(f"{s}" if s == e else f"{s}-{e}" for s, e in ranges)
        parts.append(f"\u2022 {typ}: {rng_txt}")
    return "\n".join(parts)


# ===========================================================================
# Per-employee processing
# ===========================================================================


def _process_employee(emp, data, workdays, month_num, year):
    weekly = {}
    for d in workdays:
        wn = _week_of_month(d)
        weekly.setdefault(wn, {"absHours": 0, "workedHours": 0})
        weekly[wn]["workedHours"] += 8

    absence_map = {}
    holiday_days = set()
    for a in data["absences"]:
        nt = _normalize(a["type"])
        absence_map[a["dayNumber"]] = a["type"]
        if nt in ("h", "holiday"):
            holiday_days.add(a["dayNumber"])

    absences_by_type = defaultdict(list)
    for a in data["absences"]:
        d = date(year, month_num, a["dayNumber"])
        wn = _week_of_month(d)
        if wn not in weekly:
            continue
        w = weekly[wn]
        nt = _normalize(a["type"])
        if nt in ("h", "holiday"):
            w["workedHours"] -= 8
        elif nt in ABSENCE_TYPES:
            w["absHours"] += 8
            w["workedHours"] -= 8
            absences_by_type[a["type"]].append(a["dayNumber"])
        else:
            w["workedHours"] -= 8
            if nt in OUT_TYPES:
                absences_by_type[a["type"]].append(a["dayNumber"])

    team_changes = _parse_team_changes(data["notesText"])
    bench_days = _parse_bench_days(data["notesText"], year, month_num)
    for day in bench_days:
        if day in absence_map:
            continue
        wn = _week_of_month(date(year, month_num, day))
        if wn in weekly:
            weekly[wn]["workedHours"] -= 8

    backup_ranges = _parse_backups(data["backupsText"])
    coverable_days = {
        a["dayNumber"]
        for a in data["absences"]
        if _normalize(a["type"]) in COVERABLE_TYPES
    }

    for br in backup_ranges:
        for day in range(br["startDay"], br["endDay"] + 1):
            try:
                d = date(year, month_num, day)
            except ValueError:
                continue
            if d.weekday() >= 5 or day in holiday_days:
                continue
            if day not in coverable_days:
                print(
                    f"WARN {emp['short_name']}: backup {br['originalBackupName']} on day {day} without PTO/ML/O, skipping"
                )

    all_weeks = sorted(weekly.keys())
    week_hours = [weekly[wn]["workedHours"] for wn in all_weeks]
    total_abs = sum(weekly[wn]["absHours"] for wn in all_weeks)
    total_work = sum(weekly[wn]["workedHours"] for wn in all_weeks)

    employee = {
        "fullName": emp["short_name"],
        "role": emp["job_role"],
        "wave": emp.get("wave", "Missing"),
        "product": emp["team"],
        "pilar": emp["module"],
        "tag": f"{emp['short_name']} - {emp['team']} - {emp['module']}",
        "weekHours": week_hours,
        "absHrs": total_abs,
        "workHrs": total_work,
        "comments": _build_comments(absences_by_type, team_changes),
        "isBackup": False,
        "isSeparator": False,
    }
    return employee, backup_ranges, coverable_days, all_weeks


# ===========================================================================
# Backup rows
# ===========================================================================


def _build_backup_rows(
    emp,
    backup_ranges,
    coverable_days,
    week_numbers,
    master,
    week_count,
    month_num,
    year,
):
    backup_map = {}
    for br in backup_ranges:
        be = master.get(br["backupName"])
        if not be:
            print(f"WARN backup {br['originalBackupName']} not found, skipping")
            continue
        for day in range(br["startDay"], br["endDay"] + 1):
            if day not in coverable_days:
                continue
            try:
                d = date(year, month_num, day)
            except ValueError:
                continue
            if d.weekday() >= 5:
                continue
            wn = _week_of_month(d)
            if wn not in week_numbers:
                continue
            wi = week_numbers.index(wn)
            key = br["backupName"]
            backup_map.setdefault(
                key,
                {
                    "name": be["short_name"],
                    "role": be["job_role"],
                    "wave": be.get("wave", "Missing"),
                    "product": be["team"],
                    "pilar": be["module"],
                    "weeklyHours": [0] * week_count,
                    "totalHours": 0,
                    "allWorkdays": [],
                },
            )
            bm = backup_map[key]
            bm["weeklyHours"][wi] += 8
            bm["totalHours"] += 8
            if day not in bm["allWorkdays"]:
                bm["allWorkdays"].append(day)

    rows = []
    for bm in backup_map.values():
        ranges = _group_consecutive(sorted(bm["allWorkdays"]))
        rng_txt = ",".join(
            f"{s}-{e}" if e - s + 1 >= 4 else ",".join(str(x) for x in range(s, e + 1))
            for s, e in ranges
        )
        day_word = "day" if len(bm["allWorkdays"]) == 1 else "days"
        rows.append(
            {
                "fullName": f"\u21b3 {bm['name']}",
                "role": bm["role"],
                "wave": bm.get("wave", "Missing"),
                "product": bm["product"],
                "pilar": bm["pilar"],
                "tag": f"{bm['name']} - {bm['product']} - {bm['pilar']}",
                "weekHours": bm["weeklyHours"],
                "absHrs": -1,
                "workHrs": bm["totalHours"],
                "comments": f"\u2022 Covered {emp['short_name']} {day_word} {rng_txt}",
                "isBackup": True,
                "isSeparator": False,
                "coveringFor": emp["short_name"],
            }
        )
    return rows


# ===========================================================================
# Main entry — build_report()
# ===========================================================================


def build_report(pto_data, team_data, month, year):
    """Combine PTO + team into Q1/Q2/Q3 with computed worked hours."""
    master = {_normalize(c["short_name"]): c for c in team_data if c["short_name"]}
    month_num = _month_num(month)
    workdays = _workdays(year, month_num)
    week_count = len(set(_week_of_month(d) for d in workdays))
    q1, q2, q3 = [], [], []

    for norm_name, data in pto_data.items():
        emp = master.get(norm_name)
        if not emp:
            if norm_name not in SILENT_SKIP:
                print(f"WARN {data['employeeName']} not in team allocation, skipping")
            continue
        if "bench" in emp["team"].lower():
            continue

        employee, backup_ranges, coverable_days, week_numbers = _process_employee(
            emp, data, workdays, month_num, year
        )
        role = emp["job_role"]
        target = (
            q1
            if "QA 1" in role
            else q2 if "QA 2" in role else q3 if "QA 3" in role else None
        )
        if target is None:
            continue
        target.append(employee)
        target.extend(
            _build_backup_rows(
                emp,
                backup_ranges,
                coverable_days,
                week_numbers,
                master,
                week_count,
                month_num,
                year,
            )
        )

    def sort_section(lst):
        principals = [e for e in lst if not e["isBackup"]]
        backups = [e for e in lst if e["isBackup"]]
        principals.sort(key=lambda x: (x["product"], x["fullName"]))
        res = []
        for p in principals:
            res.append(p)
            res.extend(b for b in backups if b.get("coveringFor") == p["fullName"])
        return res

    return {"q1": sort_section(q1), "q2": sort_section(q2), "q3": sort_section(q3)}
