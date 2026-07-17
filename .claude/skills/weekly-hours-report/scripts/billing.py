"""Monthly Billing — library used by main.py.

Turns each collaborator's worked hours (from the weekly report) into a
Softtek → Meta invoice. Reads the same data as the Weekly Hours report (via
core.py). Billing RATES come from a per-collaborator `rates.json` (the Google
Sheet has no rate column).
"""
import calendar
import json
from pathlib import Path

from openpyxl.styles import PatternFill, Border, Side, Font, Alignment
from openpyxl.utils import get_column_letter

from core import normalize_value, get_month_number, get_month_name

RATES_PATH_DEFAULT = Path(__file__).parent / "rates.json"
DISCOUNT_RATE = 0.02  # 2% volume discount


# ===========================================================================
# Rates — per-collaborator, loaded from rates.json
# ===========================================================================
def load_rates(path):
    """Return {normalized short name -> rate(float) or None}. Missing file -> {}."""
    path = Path(path)
    if not path.exists():
        return {}
    try:
        raw = json.loads(path.read_text())
    except Exception:
        return {}
    out = {}
    for name, val in raw.items():
        try:
            out[normalize_value(name)] = float(val)
        except (TypeError, ValueError):
            out[normalize_value(name)] = None
    return out

def check_rates(team_data, rates):
    """Return the list of collaborator short names that have no positive rate."""
    missing = []
    for c in team_data:
        name = c["short_name"]
        if not name:
            continue
        r = rates.get(normalize_value(name))
        if not r or r <= 0:
            missing.append(name)
    return missing

def write_rates_template(path, team_data, existing):
    """Write/refresh rates.json with every collaborator (existing values kept, missing -> 0)."""
    data = {}
    for c in team_data:
        name = c["short_name"]
        if not name:
            continue
        r = existing.get(normalize_value(name))
        data[name] = r if (r and r > 0) else 0
    Path(path).write_text(json.dumps(data, indent=2, ensure_ascii=False))


# ===========================================================================
# Build billing from the weekly report + rates
# ===========================================================================
def _rate_for(emp, rates):
    """Principal is billed at their own rate; a backup at the rate of whoever they cover."""
    key = normalize_value(emp.get("coveringFor", "") if emp.get("isBackup")
                          else emp["fullName"].lstrip("↳ ").strip())
    return rates.get(key) or 0.0

def build_billing(weekly_report, rates):
    def build_section(employees, role_key):
        rows = []
        for emp in employees:
            if emp.get("isSeparator"):
                continue
            rate = _rate_for(emp, rates)
            if rate == 0:
                print(f"WARN no rate for {emp['fullName']} — billed at $0")
            qty = emp["workHrs"]
            rows.append({
                "fullName": emp["fullName"], "tag": emp["tag"], "role": emp["role"],
                "product": emp["product"], "pilar": emp["pilar"],
                "qty": qty, "unit": "Hrs", "description": emp["tag"],
                "rate": rate, "amount": qty * rate,
                "isBackup": emp.get("isBackup", False), "coveringFor": emp.get("coveringFor"),
                "isSeparator": False, "roleKey": role_key,
            })
        return rows

    q1 = build_section(weekly_report["q1"], "q1")
    q2 = build_section(weekly_report["q2"], "q2")
    q3 = build_section(weekly_report["q3"], "q3")
    subtotal = sum(e["amount"] for e in q1 + q2 + q3)
    discount_amount = subtotal * DISCOUNT_RATE
    total = subtotal - discount_amount
    return {"q1": q1, "q2": q2, "q3": q3, "subtotal": subtotal,
            "discountRate": DISCOUNT_RATE, "discountAmount": discount_amount, "total": total}


# ===========================================================================
# Amount-in-words (invoice needs the total spelled out)
# ===========================================================================
_ONES = ["zero","one","two","three","four","five","six","seven","eight","nine","ten",
         "eleven","twelve","thirteen","fourteen","fifteen","sixteen","seventeen","eighteen","nineteen"]
_TENS = ["","","twenty","thirty","forty","fifty","sixty","seventy","eighty","ninety"]

def _under_1000(n):
    if n < 20: return _ONES[n]
    if n < 100: return _TENS[n // 10] + ("-" + _ONES[n % 10] if n % 10 else "")
    return _ONES[n // 100] + " hundred" + (" " + _under_1000(n % 100) if n % 100 else "")

def _int_words(n):
    if n == 0: return "zero"
    parts = []
    for div, name in [(1_000_000, "million"), (1000, "thousand"), (1, "")]:
        if n >= div:
            parts.append(_under_1000(n // div) + ((" " + name) if name else ""))
            n %= div
    return " ".join(parts)

def amount_to_words(total):
    dollars = int(total)
    cents = int(round((total - dollars) * 100))
    return f"{_int_words(dollars).upper()} DOLLARS AND {(_int_words(cents).upper() if cents else 'ZERO')} CENTS"


# ===========================================================================
# Invoice worksheet
# ===========================================================================
def fill(argb): return PatternFill(start_color=argb, end_color=argb, fill_type="solid")
def thin(): return Side(style="thin", color="D0D0D0")
def border_all(): return Border(left=thin(), right=thin(), top=thin(), bottom=thin())
def blue_border(): return Border(left=Side(style="thin", color="4472C4"), right=Side(style="thin", color="4472C4"), top=Side(style="thin", color="4472C4"), bottom=Side(style="thin", color="4472C4"))

BLUE = "2F5496"; BLACK = "000000"; GRAY_FILL = "D6E4F0"; HEADER_FILL = "BDD7EE"; BLUE_LINK = "0563C1"

def build_invoice_header(month, year):
    month_num = get_month_number(month)
    month_full = get_month_name(month_num, False)
    last_day = calendar.monthrange(year, month_num)[1]
    return {
        "date": f"{month_num}/{last_day}/{year}",
        "servicePeriod": f"{month_full} 1 - {last_day}, {year}",
        "poNumber": "70000801748",
        "billToName": "Meta Platforms, Inc.",
        "billToAddress": ["1601 Willow Rd", "Menlo Park, CA 94025", "United States"],
        "shipToName": "Meta Platforms, Inc.",
        "shipToAddress": ["7300 Gateway Blvd.", "Dock 15 - Ship Happens - FB1", "Newark, CA 94560 - USA"],
        "paymentTerms": "Within 30 days",
        "dueDate": "",
        "taxId": "91-1864740",
        "companyName": "SOFTTEK INTEGRATION SYSTEMS, INC",
    }

def write_invoice_sheet(wb, billing, month, year, title="Invoice"):
    """Add a Softtek invoice worksheet to an existing workbook."""
    header = build_invoice_header(month, year)
    ws = wb.create_sheet(title)
    widths = [10,8,12,12,12,12,12,12,12,12,12,16]
    for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width = w

    blue_font = Font(color=BLUE, size=11)
    blue_bold = Font(color=BLUE, bold=True, size=11)
    black11 = Font(color=BLACK, size=11)
    black_bold11 = Font(color=BLACK, bold=True, size=11)
    gray_fill = fill(GRAY_FILL)
    hdr_fill = fill(HEADER_FILL)
    b_all = border_all()

    r = 3
    ws.merge_cells(f"K{r}:L{r}"); ws.cell(r,11).value = header["companyName"]; ws.cell(r,11).font = blue_font; ws.cell(r,11).alignment = Alignment(horizontal="right"); r+=1
    ws.merge_cells(f"K{r}:L{r}"); ws.cell(r,11).value = f"Tax ID {header['taxId']}"; ws.cell(r,11).font = blue_font; ws.cell(r,11).alignment = Alignment(horizontal="right"); r+=2

    ws.merge_cells(f"H{r}:J{r}"); ws.cell(r,8).value="INVOICE #"; ws.cell(r,8).font=blue_font; ws.cell(r,8).alignment=Alignment(horizontal="center"); r+=1

    ws.cell(r,2).value="BILL TO"; ws.cell(r,2).font=blue_font
    ws.merge_cells(f"H{r}:J{r}"); ws.cell(r,8).value="DATE"; ws.cell(r,8).font=blue_font; ws.cell(r,8).alignment=Alignment(horizontal="center")
    ws.merge_cells(f"K{r}:L{r}"); ws.cell(r,11).value=header["date"]; ws.cell(r,11).alignment=Alignment(horizontal="right"); r+=1

    ws.merge_cells(f"B{r}:C{r}"); ws.cell(r,2).value=header["billToName"]; ws.cell(r,2).font=black_bold11; ws.cell(r,2).border=blue_border(); r+=1
    for line in header["billToAddress"]:
        ws.merge_cells(f"B{r}:C{r}"); ws.cell(r,2).value=line; ws.cell(r,2).font=black11; r+=1

    ws.cell(r,2).value="SHIP TO"; ws.cell(r,2).font=blue_font
    ws.merge_cells(f"F{r}:G{r}"); ws.cell(r,6).value="PAYMENT TERMS"; ws.cell(r,6).fill=gray_fill; ws.cell(r,6).font=blue_font; ws.cell(r,6).alignment=Alignment(horizontal="center")
    ws.cell(r,10).value="DUE DATE"; ws.cell(r,10).fill=gray_fill; ws.cell(r,10).font=blue_font; ws.cell(r,10).alignment=Alignment(horizontal="center")
    ws.cell(r,12).value="P.O. NO."; ws.cell(r,12).fill=gray_fill; ws.cell(r,12).font=blue_font; ws.cell(r,12).alignment=Alignment(horizontal="center"); r+=1

    ws.merge_cells(f"B{r}:C{r}"); ws.cell(r,2).value=header["shipToName"]; ws.cell(r,2).font=black_bold11
    ws.merge_cells(f"F{r}:G{r}"); ws.cell(r,6).value=header["paymentTerms"]; ws.cell(r,6).alignment=Alignment(horizontal="center"); ws.cell(r,6).font=black11
    ws.cell(r,10).value=header["dueDate"]; ws.cell(r,10).alignment=Alignment(horizontal="center")
    ws.cell(r,12).value=header["poNumber"]; ws.cell(r,12).alignment=Alignment(horizontal="center"); ws.cell(r,12).font=black_bold11; r+=1

    for line in header["shipToAddress"]:
        ws.merge_cells(f"B{r}:C{r}"); ws.cell(r,2).value=line; ws.cell(r,2).font=black11; r+=1

    while r < 17: r+=1

    # Table header
    ws.cell(r,1).value="QTY"; ws.cell(r,2).value="UN"; ws.merge_cells(f"C{r}:J{r}"); ws.cell(r,3).value="DESCRIPTION"; ws.cell(r,11).value="RATE"; ws.cell(r,12).value="AMOUNT"
    for col in [1,2,3,11,12]:
        c = ws.cell(r,col); c.fill=hdr_fill; c.font=blue_bold; c.alignment=Alignment(horizontal="center"); c.border=b_all
    for cc in range(3,11):
        c=ws.cell(r,cc); c.fill=hdr_fill; c.border=b_all
    r+=1

    ws.merge_cells(f"C{r}:J{r}"); ws.cell(r,3).value=f"Services Period: {header['servicePeriod'].replace(' - ',' to ').replace(',','')}"; ws.cell(r,3).font=black11; r+=1
    ws.merge_cells(f"C{r}:J{r}"); ws.cell(r,3).value="MSOW INB2830770. Order Form No. 1"; ws.cell(r,3).font=black11; r+=1
    ws.merge_cells(f"C{r}:J{r}"); ws.cell(r,3).value="Softtek nearshore program for QA analysts"; ws.cell(r,3).font=black11; r+=1
    r+=1

    def add_section(label, employees):
        nonlocal r
        if not employees: return
        ws.merge_cells(f"C{r}:J{r}"); ws.cell(r,3).value=label; ws.cell(r,3).font=black_bold11; r+=1
        for emp in employees:
            ws.cell(r,1).value=emp["qty"]; ws.cell(r,1).alignment=Alignment(horizontal="right"); ws.cell(r,1).font=black11
            ws.cell(r,2).value=emp["unit"]; ws.cell(r,2).alignment=Alignment(horizontal="center"); ws.cell(r,2).font=black11
            ws.merge_cells(f"C{r}:J{r}")
            name = emp["fullName"].lstrip("↳ ").strip()
            ws.cell(r,3).value=f"{name} - {emp['product']} - {emp['pilar']}"
            ws.cell(r,3).font = black11 if emp["isBackup"] else black_bold11
            if emp["isBackup"]: ws.cell(r,3).alignment = Alignment(indent=1)
            ws.cell(r,11).value=emp["rate"]; ws.cell(r,11).number_format='"$"#,##0.00'; ws.cell(r,11).alignment=Alignment(horizontal="right"); ws.cell(r,11).font=black11
            ws.cell(r,12).value=emp["amount"]; ws.cell(r,12).number_format='"$"#,##0.00'; ws.cell(r,12).alignment=Alignment(horizontal="right"); ws.cell(r,12).font=black11
            for col in [1,2,11,12]:
                ws.cell(r,col).border=b_all
            for cc in range(3,11):
                ws.cell(r,cc).border=b_all
            r+=1

    add_section("QA Analyst I", billing["q1"])
    add_section("QA Analyst II", billing["q2"])
    add_section("QA Analyst III", billing["q3"])
    r+=1

    # Totals
    ws.merge_cells(f"A{r}:B{r}"); ws.cell(r,1).value="Saxena, Sarthak"; ws.cell(r,1).font=black11; ws.cell(r,1).border=b_all
    ws.merge_cells(f"C{r}:H{r}"); ws.cell(r,3).border=b_all
    ws.merge_cells(f"I{r}:K{r}"); ws.cell(r,9).value="SUB-TOTAL"; ws.cell(r,9).font=blue_font; ws.cell(r,9).alignment=Alignment(horizontal="right"); ws.cell(r,9).border=b_all
    ws.cell(r,12).value=billing["subtotal"]; ws.cell(r,12).number_format="#,##0.00"; ws.cell(r,12).alignment=Alignment(horizontal="right"); ws.cell(r,12).font=black11; ws.cell(r,12).border=b_all; r+=1

    ws.merge_cells(f"A{r}:B{r}"); ws.cell(r,1).value=amount_to_words(billing["total"]); ws.cell(r,1).font=Font(size=9); ws.cell(r,1).border=b_all
    ws.merge_cells(f"C{r}:H{r}"); ws.cell(r,3).border=b_all
    ws.merge_cells(f"I{r}:K{r}"); ws.cell(r,9).value="Volume Discount"; ws.cell(r,9).font=blue_font; ws.cell(r,9).alignment=Alignment(horizontal="right"); ws.cell(r,9).border=b_all
    ws.cell(r,12).value=billing["discountAmount"]; ws.cell(r,12).number_format="#,##0.00"; ws.cell(r,12).alignment=Alignment(horizontal="right"); ws.cell(r,12).border=b_all; r+=1

    ws.merge_cells(f"A{r}:B{r}"); ws.cell(r,1).border=b_all
    ws.merge_cells(f"C{r}:H{r}"); ws.cell(r,3).value="PROJECT"; ws.cell(r,3).font=blue_font; ws.cell(r,3).border=b_all
    ws.merge_cells(f"I{r}:K{r}"); ws.cell(r,9).border=b_all; ws.cell(r,12).border=b_all; r+=1

    ws.merge_cells(f"A{r}:B{r}"); ws.cell(r,1).border=b_all
    ws.merge_cells(f"C{r}:H{r}"); ws.cell(r,3).value="1-0000029582-3"; ws.cell(r,3).font=black11; ws.cell(r,3).border=b_all
    ws.merge_cells(f"I{r}:J{r}"); ws.cell(r,9).value="TOTAL"; ws.cell(r,9).font=Font(color=BLUE,bold=True,size=11); ws.cell(r,9).alignment=Alignment(horizontal="right"); ws.cell(r,9).border=b_all
    ws.cell(r,11).value="USD"; ws.cell(r,11).alignment=Alignment(horizontal="center"); ws.cell(r,11).font=black_bold11; ws.cell(r,11).border=b_all
    ws.cell(r,12).value=billing["total"]; ws.cell(r,12).number_format="#,##0.00"; ws.cell(r,12).font=black_bold11; ws.cell(r,12).alignment=Alignment(horizontal="right"); ws.cell(r,12).border=b_all; r+=2

    # Footer
    ws.merge_cells(f"A{r}:L{r}"); ws.cell(r,1).value="Please wire your payments to:  Wells Fargo Bank,  250 East Ponce De Leon Avenue, GA9187"; ws.cell(r,1).font=Font(size=10); r+=1
    ws.merge_cells(f"A{r}:L{r}"); ws.cell(r,1).value="Decatur, Georgia 30030 United States Beneficiary: SOFTTEK INTEGRATION SYSTEMS, INC."; ws.cell(r,1).font=Font(size=10); r+=1
    ws.cell(r,1).value="softtek.com"; ws.cell(r,1).font=Font(color=BLUE_LINK, underline="single", size=10)
    ws.merge_cells(f"C{r}:L{r}"); ws.cell(r,3).value="15303 Dallas Pkwy, Suite 200, Addison TX 75001, US Account Number: 2000055970549"; ws.cell(r,3).font=Font(size=10); r+=1
    ws.cell(r,1).value="15303 Dallas Pkwy, Suite 200"; ws.cell(r,1).font=Font(color=BLUE_LINK, size=10)
    ws.merge_cells(f"C{r}:L{r}"); ws.cell(r,3).value="for ACH Payments 061000227 (ABA) / PNBPUS33 (SWIFT) / 0407 (CHIP); for international and"; ws.cell(r,3).font=Font(size=10); r+=1
    ws.cell(r,1).value="Addison, TX 75001"; ws.cell(r,1).font=Font(color=BLUE_LINK, size=10)
    ws.merge_cells(f"C{r}:L{r}"); ws.cell(r,3).value="Domestic Wire Transfers 121000248 (ABA) / WFBIUS6S (SWIFT) / 0407 (CHIP)."; ws.cell(r,3).font=Font(size=10); r+=1
    ws.merge_cells(f"A{r}:L{r}"); ws.cell(r,1).value="If you have any questions, please contact Accounts Receivable at:  Office 1(469) 283-2506"; ws.cell(r,1).alignment=Alignment(horizontal="center"); ws.cell(r,1).font=Font(size=10); r+=1
    ws.merge_cells(f"A{r}:L{r}"); ws.cell(r,1).value="Fax 1(214) 580-9778 Toll free from USA 1877-4723-029 Ext1031/1109 usbilling@softtek.com"; ws.cell(r,1).alignment=Alignment(horizontal="center"); ws.cell(r,1).font=Font(size=10)
    return ws
