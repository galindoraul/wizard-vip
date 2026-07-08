#!/usr/bin/env python3
"""Check who has done their C2C for the current week.
Reads PTO Sheet for team list + absence info (cached 5 min).
Reads Click2Sync/*.json for submission status.
Outputs JSON for the agent to format."""

import json
import os
import subprocess
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path

# Config is shared at .claude/skills/config.json
SCRIPTS_DIR = os.path.dirname(os.path.realpath(__file__))
SKILL_DIR = os.path.dirname(SCRIPTS_DIR)
SKILLS_DIR = os.path.dirname(SKILL_DIR)
CONFIG_PATH = os.path.join(SKILLS_DIR, "config.json")

PTO_SHEET_ID = "1Vae2OUAdYT3pMAQLLSYcNRBFklJ2WybctNia6OjNK_g"
HEADER_ROW = 8
ABSENCE_TYPES = ["pto", "ml", "pto(pa)", "ml(pa)"]

CACHE_DIR = "/tmp"
PTO_CACHE_TTL = 300  # 5 minutes


def get_month_name(month_number):
    months = [
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
    return months[month_number - 1]


def get_current_week_dates():
    today = datetime.now()
    monday = today - timedelta(days=today.weekday())
    monday = monday.replace(hour=0, minute=0, second=0, microsecond=0)
    friday = monday + timedelta(days=4)
    sunday = monday + timedelta(days=6)
    return monday, friday, sunday


def get_week_tab_name(monday, sunday):
    """Same format as json-writer.py uses for week keys."""
    return f"{monday.strftime('%b')} {monday.day}-{sunday.day}"


def read_sheet_data(sheet_name):
    cmd = [
        "meta",
        "google.sheets",
        "read",
        f"--id={PTO_SHEET_ID}",
        f"--range='{sheet_name}'",
        "--no-header",
        "-o",
        "json",
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        return None
    lines = result.stdout.strip().split("\n")
    json_start = next(
        (i for i, l in enumerate(lines) if l.strip().startswith("[")), None
    )
    if json_start is None:
        return None
    return json.loads("\n".join(lines[json_start:]))


def read_sheet_data_cached(sheet_name):
    """Read sheet data with 5-min cache."""
    cache_key = f"c2c_status_pto_{sheet_name.replace(' ', '_')}"
    cache_path = os.path.join(CACHE_DIR, f"{cache_key}.json")

    if os.path.exists(cache_path):
        age = time.time() - os.path.getmtime(cache_path)
        if age < PTO_CACHE_TTL:
            with open(cache_path) as f:
                return json.load(f)

    data = read_sheet_data(sheet_name)
    if data:
        with open(cache_path, "w") as f:
            json.dump(data, f)
    return data


def get_team_members_and_absences(monday, friday):
    """Get all team members from PTO sheet and their absence status for this week."""
    week_workdays = [(monday + timedelta(days=i)) for i in range(5)]

    # Determine which months to read
    months_needed = set()
    for day in week_workdays:
        months_needed.add((day.month, day.year))

    all_members = {}  # name -> {"absent_days": count, "working_days": count}

    for month, year in months_needed:
        sheet_name = f"{get_month_name(month)} {year}"
        rows = read_sheet_data_cached(sheet_name)
        if not rows or len(rows) < HEADER_ROW:
            continue

        header = rows[HEADER_ROW - 1]

        # Map day numbers to columns
        day_columns = {}
        for col_idx in range(2, len(header)):
            try:
                day_num = int(header[col_idx])
                day_columns[day_num] = col_idx
            except (ValueError, TypeError):
                continue

        # Holiday row (row 3, index 2)
        holiday_days = set()
        holiday_row = rows[2] if len(rows) > 2 else []
        for day_num, col_idx in day_columns.items():
            if col_idx < len(holiday_row):
                val = str(holiday_row[col_idx]).strip().lower()
                if val in ("h", "holiday"):
                    holiday_days.add(day_num)

        # Parse employees
        for row_idx in range(HEADER_ROW, len(rows)):
            row = rows[row_idx]
            if len(row) < 2:
                continue
            name = str(row[1]).strip()
            if not name:
                continue

            if name not in all_members:
                all_members[name] = {"absent_days": 0, "working_days": 0}

            # Check each workday in this month
            for workday in week_workdays:
                if workday.month != month:
                    continue
                day_num = workday.day

                # Holiday = not a working day
                if day_num in holiday_days:
                    continue

                all_members[name]["working_days"] += 1

                # Check absence
                col_idx = day_columns.get(day_num)
                if col_idx and col_idx < len(row):
                    val = str(row[col_idx]).strip().lower()
                    if val in ABSENCE_TYPES:
                        all_members[name]["absent_days"] += 1
                    elif val in ("h", "holiday"):
                        all_members[name]["working_days"] -= 1

    return all_members


def get_c2c_folder():
    """Find the Click2Sync folder via Google Drive Stream."""
    cloud_storage = Path.home() / "Library" / "CloudStorage"
    gdrive_dirs = list(cloud_storage.glob("GoogleDrive-*@meta.com"))
    if not gdrive_dirs:
        return None
    base = gdrive_dirs[0]
    c2c_folder = (
        base
        / "Shared drives"
        / "Meta - STK"
        / "Project Tracking"
        / "Automation"
        / "Automation Outputs"
        / "Click2Sync"
    )
    if not c2c_folder.exists():
        return None
    return c2c_folder


def get_all_persons(c2c_folder, week_key):
    """Read all person JSONs. Returns submitted and known persons with meta_unixname."""
    submitted = []
    known_persons = {}  # name.lower() -> {"name": ..., "meta_unixname": ...}
    json_files = list(c2c_folder.glob("*.json"))

    for json_path in json_files:
        try:
            with open(json_path) as f:
                data = json.load(f)
            person = data.get("person", json_path.stem)
            meta_unixname = data.get("meta_unixname", "")
            weeks = data.get("weeks", {})

            # Store person info regardless of submission
            known_persons[person.lower()] = {
                "name": person,
                "meta_unixname": meta_unixname,
            }

            # Check if current week exists
            if week_key in weeks and weeks[week_key]:
                submitted.append(
                    {
                        "name": person,
                        "meta_unixname": meta_unixname,
                        "rows": len(weeks[week_key]),
                    }
                )
        except (json.JSONDecodeError, OSError):
            continue

    return submitted, known_persons


def main():
    # Get week info
    monday, friday, sunday = get_current_week_dates()
    week_key = get_week_tab_name(monday, sunday)
    week_display = f"{monday.strftime('%b %d')} - {sunday.strftime('%b %d, %Y')}"

    # Get team members from PTO sheet
    members = get_team_members_and_absences(monday, friday)
    if not members:
        print(json.dumps({"status": "error", "message": "Could not read PTO sheet."}))
        sys.exit(1)

    # Get C2C folder
    c2c_folder = get_c2c_folder()
    if not c2c_folder:
        print(
            json.dumps({"status": "error", "message": "Click2Sync folder not found."})
        )
        sys.exit(1)

    # Get submissions and known persons
    submissions, known_persons = get_all_persons(c2c_folder, week_key)
    submitted_names = {s["name"].lower() for s in submissions}

    # Classify members
    pto_all_week = []
    missing = []

    for name, info in members.items():
        # If all working days are absent → PTO all week
        if info["working_days"] > 0 and info["absent_days"] >= info["working_days"]:
            pto_all_week.append(name)
        elif name.lower() not in submitted_names:
            # Get meta_unixname from known persons (previous weeks)
            person_info = known_persons.get(name.lower(), {})
            missing.append(
                {
                    "name": name,
                    "meta_unixname": person_info.get("meta_unixname", ""),
                }
            )

    # Sort
    submitted_sorted = sorted(submissions, key=lambda x: x["name"])
    missing.sort(key=lambda x: x["name"])
    pto_all_week.sort()

    report = {
        "status": "ok",
        "week": week_display,
        "weekKey": week_key,
        "submitted": submitted_sorted,
        "missing": missing,
        "ptoAllWeek": pto_all_week,
    }

    print(json.dumps(report))


if __name__ == "__main__":
    main()
