#!/usr/bin/env python3
"""Consolidates all person JSONs into one weekly xlsx.
Reads Click2Sync/*.json via Google Drive API in parallel (always fresh),
merges rows for current week, assigns requestNo,
and writes to Click2Sync/Weekly Reports/{week_name}.xlsx with formatting."""

import argparse
import concurrent.futures
import json
import os
import subprocess
import sys
from datetime import datetime, timedelta
from pathlib import Path

try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print(
        json.dumps(
            {
                "status": "error",
                "message": "openpyxl not installed. Run: pip3 install openpyxl",
            }
        )
    )
    sys.exit(1)

# Click2Sync folder in Shared Drive "Meta - STK"
C2C_FOLDER_ID = "1CsXnW7pq-l8E9A-ZpB1CvPtfS68END6W"

# Column definitions: (header, json_key, width, header_bg_color)
COLUMNS = [
    ("Project ID", "projectId", 14, "1A3A5C"),
    ("Created By", "createdBy", 17, "1A3A5C"),
    ("Request No", "requestNo", 15, "1A3A5C"),
    ("Assigned To", "assignTo", 17, "8B6D1A"),
    ("Peer Reviewer", "peerReviewer", 17, "1A3A5C"),
    ("Short Description", "shortDescription", 50, "1A3A5C"),
    ("Requirement Type", "requirementType", 27, "1A3A5C"),
    ("Requirement Subtype", "requirementSubtype", 27, "1A3A5C"),
    ("Module/Feature", "moduleFeature", 18, "8B6D1A"),
    ("Team", "team", 17, "8B6D1A"),
    ("Number of Defective Products", "numberOfDefectiveProducts", 32, "1A3A5C"),
    ("Number of Products (Generated or Modified)", "numberOfProducts", 46, "8B6D1A"),
    ("Total of Test Cases", "totalOfTestCases", 23, "1A3A5C"),
    ("Total", "total", 9, "1A3A5C"),
    ("Pass", "pass", 8, "1A3A5C"),
    ("Fail", "fail", 13, "1A3A5C"),
    ("Cant Test", "cantTest", 13, "1A3A5C"),
    ("Quantity of Bugs Closed", "qtyBugsClosed", 27, "1A3A5C"),
    ("Quantity of Bugs Re-opened", "qtyBugsReopened", 30, "1A3A5C"),
    ("Quantity of Bugs Verified", "qtyBugsVerified", 29, "1A3A5C"),
    ("Quantity of New Bugs Found", "qtyNewBugsFound", 30, "1A3A5C"),
    ("Release/Build", "releaseBuild", 17, "8B6D1A"),
    ("Complete %", "completePercent", 14, "1A3A5C"),
    ("Defects Advice Flag", "defectsAdviceFlag", 23, "1A3A5C"),
    ("Peer Review Checklist Reviewed", "peerReviewChecklistReviewed1", 34, "1A3A5C"),
    ("Peer Review Checklist Template", "peerReviewChecklistTemplate", 13, "1A3A5C"),
    ("Peer Review Checklist Reviewed", "peerReviewChecklistReviewed2", 13, "1A3A5C"),
    ("Scheduled Start Date", "scheduledStartDate", 24, "8B4513"),
    ("Scheduled Finish Date", "scheduledFinishDate", 25, "8B4513"),
    ("Scheduled Duration", "scheduledDuration", 22, "1A3A5C"),
    ("Scheduled Effort", "scheduledEffort", 20, "1A3A5C"),
    ("Peer Review Scheduled Effort (hrs)", "peerReviewScheduledEffort", 38, "2D5A3A"),
    ("Scheduled Rework Effort (hrs)", "scheduledReworkEffort", 33, "2D5A3A"),
    ("Scheduled UAT Rework Effort", "scheduledUATReworkEffort", 31, "2D5A3A"),
    ("Actual Start Date", "actualStartDate", 21, "1A3A5C"),
    ("Actual Finish Date", "actualFinishDate", 22, "1A3A5C"),
    ("Actual Duration", "actualDuration", 19, "1A3A5C"),
    ("Actual Effort", "actualEffort", 17, "1A3A5C"),
    ("Peer Review Actual Effort (hrs)", "peerReviewActualEffort", 35, "2D5A3A"),
    ("Actual Rework Effort (hrs)", "actualReworkEffort", 30, "2D5A3A"),
    ("Actual UAT Rework Effort", "actualUATReworkEffort", 28, "2D5A3A"),
    ("Closed On", "closedOn", 13, "1A3A5C"),
    ("ForClosing", "forClosing", 14, "1A3A5C"),
    ("Action", "action", 10, "1A3A5C"),
    ("Peer Review Checklist Reviewed", "peerReviewChecklist", 14, "1A4B6E"),
    ("Estimation", "estimationLink", 14, "1A4B6E"),
]

GRAY_FILL = PatternFill(start_color="D4D4D8", end_color="D4D4D8", fill_type="solid")
YELLOW_FILL = PatternFill(start_color="FFF9E6", end_color="FFF9E6", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=9)
DATA_FONT = Font(size=10)
DATA_ROW_HEIGHT = 15

THIN_BORDER = Border(
    left=Side(style="thin", color="808080"),
    right=Side(style="thin", color="808080"),
    top=Side(style="thin", color="808080"),
    bottom=Side(style="thin", color="808080"),
)


def get_week_tab_name():
    today = datetime.now()
    monday = today - timedelta(days=today.weekday())
    sunday = monday + timedelta(days=6)
    return f"{monday.strftime('%b')} {monday.day}-{sunday.day}"


def run_meta_cmd(args):
    """Run a meta CLI command and return stdout."""
    result = subprocess.run(
        ["meta"] + args,
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        return None
    return result.stdout.strip()


def get_c2c_files_from_drive():
    """List all JSON files in the Click2Sync folder via Drive API."""
    output = run_meta_cmd(
        [
            "google.drive",
            "list",
            f"--folder-id={C2C_FOLDER_ID}",
            "--limit=200",
            "--columns=name,id,mimeType",
            "-o",
            "json",
        ]
    )
    if not output:
        return []
    lines = output.split("\n")
    json_start = next(
        (i for i, l in enumerate(lines) if l.strip().startswith("[")), None
    )
    if json_start is None:
        return []
    files = json.loads("\n".join(lines[json_start:]))
    return [
        f
        for f in files
        if f.get("mimeType") == "application/json"
        or f.get("name", "").endswith(".json")
    ]


def read_drive_file(file_id):
    """Read a file's content from Google Drive via API."""
    output = run_meta_cmd(
        [
            "google.drive.file",
            "get",
            f"--id={file_id}",
        ]
    )
    if not output:
        return None
    try:
        return json.loads(output)
    except json.JSONDecodeError:
        return None


def get_all_rows(week_key):
    """Read all person JSONs from Drive API in parallel and collect rows."""
    files = get_c2c_files_from_drive()
    if not files:
        return []

    file_ids = [f["id"] for f in files if f.get("id")]
    all_data = []

    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
        future_to_id = {executor.submit(read_drive_file, fid): fid for fid in file_ids}
        for future in concurrent.futures.as_completed(future_to_id):
            data = future.result()
            if data:
                all_data.append(data)

    all_rows = []
    for data in all_data:
        person = data.get("person", "")
        weeks = data.get("weeks", {})
        if week_key in weeks:
            for row in weeks[week_key]:
                all_rows.append((person, row))

    all_rows.sort(key=lambda x: x[0].lower())
    return all_rows


def get_output_folder():
    """Find the Weekly Reports folder via Google Drive Stream (for writing)."""
    cloud_storage = Path.home() / "Library" / "CloudStorage"
    gdrive_dirs = list(cloud_storage.glob("GoogleDrive-*@meta.com"))
    if not gdrive_dirs:
        return None
    base = gdrive_dirs[0]
    weeks_folder = (
        base
        / "Shared drives"
        / "Meta - STK"
        / "Project Tracking"
        / "Automation"
        / "Automation Outputs"
        / "Click2Sync"
        / "Weekly Reports"
    )
    weeks_folder.mkdir(exist_ok=True)
    return weeks_folder


def get_last_request_no(weeks_folder):
    """Find the max requestNo from the most recent xlsx in weeks/."""
    if not weeks_folder or not weeks_folder.exists():
        return None
    xlsx_files = sorted(
        weeks_folder.glob("*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True
    )
    if not xlsx_files:
        return None
    try:
        wb = load_workbook(xlsx_files[0], read_only=True)
        ws = wb.active
        max_req = None
        for row in ws.iter_rows(min_row=2, max_col=3, values_only=True):
            req = row[2]
            if req and str(req).strip():
                max_req = str(req).strip()
        wb.close()
        return max_req
    except Exception:
        return None


def increment_request_no(request_no):
    """Increment: '29582-3-10001' -> '29582-3-10002'."""
    parts = request_no.rsplit("-", 1)
    if len(parts) == 2:
        prefix = parts[0]
        num = int(parts[1]) + 1
        return f"{prefix}-{num}"
    return request_no


def build_xlsx(all_rows, weeks_folder, week_key, first_request_no=None):
    """Build the xlsx file with template formatting."""
    if first_request_no:
        current_req = first_request_no
    else:
        last_req = get_last_request_no(weeks_folder)
        if last_req is None:
            return None, "need_request_no"
        current_req = increment_request_no(last_req)

    wb = Workbook()
    ws = wb.active
    ws.title = week_key

    # --- Headers ---
    for col_idx, (header, _, width, bg_color) in enumerate(COLUMNS):
        cell = ws.cell(row=1, column=col_idx + 1, value=header)
        cell.fill = PatternFill(
            start_color=bg_color, end_color=bg_color, fill_type="solid"
        )
        cell.font = HEADER_FONT
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        cell.border = THIN_BORDER
        ws.column_dimensions[get_column_letter(col_idx + 1)].width = width

    ws.freeze_panes = "A2"

    # --- Data rows ---
    current_person = None

    for row_idx, (person, row_data) in enumerate(all_rows):
        excel_row = row_idx + 2

        is_first_of_person = person != current_person
        if is_first_of_person:
            current_person = person

        row_data["requestNo"] = current_req
        current_req = increment_request_no(current_req)

        ws.row_dimensions[excel_row].height = DATA_ROW_HEIGHT

        for col_idx, (_, json_key, _, _) in enumerate(COLUMNS):
            value = row_data.get(json_key, "")
            cell = ws.cell(
                row=excel_row, column=col_idx + 1, value=value if value else None
            )
            cell.border = THIN_BORDER
            cell.font = DATA_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center")

            if not value:
                cell.fill = GRAY_FILL
            elif is_first_of_person:
                cell.fill = YELLOW_FILL

    # --- Auto-filter ---
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{len(all_rows) + 1}"

    # --- Save ---
    output_path = weeks_folder / f"{week_key}.xlsx"
    wb.save(str(output_path))
    wb.close()

    return output_path, "ok"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--first-request-no", help="First request number if no previous xlsx exists"
    )
    args = parser.parse_args()

    week_key = get_week_tab_name()

    # Read all rows from Drive API (parallel, always fresh)
    all_rows = get_all_rows(week_key)
    if not all_rows:
        print(
            json.dumps(
                {"status": "error", "message": f"No rows found for week '{week_key}'."}
            )
        )
        sys.exit(1)

    # Write xlsx to local Google Drive Stream path
    weeks_folder = get_output_folder()
    if not weeks_folder:
        print(
            json.dumps(
                {"status": "error", "message": "Weekly Reports folder not found."}
            )
        )
        sys.exit(1)

    output_path, status = build_xlsx(
        all_rows, weeks_folder, week_key, args.first_request_no
    )

    if status == "need_request_no":
        print(
            json.dumps(
                {
                    "status": "need_request_no",
                    "message": "No previous xlsx found. Provide --first-request-no.",
                }
            )
        )
        sys.exit(1)

    persons = set(person for person, _ in all_rows)
    print(
        json.dumps(
            {
                "status": "ok",
                "filename": output_path.name,
                "rows": len(all_rows),
                "persons": len(persons),
                "week": week_key,
            }
        )
    )


if __name__ == "__main__":
    main()
