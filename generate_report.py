from openpyxl import load_workbook
from jinja2 import Template
import os
import sys
import glob
import time
import re
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# -----------------------------
# Pick LATEST dated Excel file
# Expected name: uptime_latest_25th Dec_2025.xlsx
# -----------------------------
def extract_date_from_filename(filename):
    match = re.search(r'(\d{1,2})(st|nd|rd|th)\s+([A-Za-z]+)_([0-9]{4})', filename)
    if not match:
        return None

    day = match.group(1)
    month = match.group(3)
    year = match.group(4)

    try:
        return datetime.strptime(f"{day} {month} {year}", "%d %b %Y")
    except ValueError:
        return None


def find_excel():
    files = []
    for ext in ("*.xlsx", "*.xls"):
        files.extend(glob.glob(ext))

    if not files:
        raise Exception("❌ No Excel file found")

    dated_files = []
    for f in files:
        date = extract_date_from_filename(f)
        if date:
            dated_files.append((date, f))

    if not dated_files:
        raise Exception("❌ No Excel file with valid date format found")

    # Pick latest date
    dated_files.sort(reverse=True)
    return dated_files[0][1]


EXCEL_FILE = sys.argv[1] if len(sys.argv) > 1 else find_excel()
print(f"✅ Using Excel: {EXCEL_FILE}")

# -----------------------------
# Helper: show EXACT Excel display
# -----------------------------
def get_cell_display(cell):
    if cell.value is None:
        return ""

    # Percentage formatting (keep exactly like Excel)
    if isinstance(cell.value, (int, float)) and "%" in str(cell.number_format):
        decimals = 0
        match = re.search(r"\.(0+)%", cell.number_format)
        if match:
            decimals = len(match.group(1))
        return f"{cell.value * 100:.{decimals}f}%"

    return str(cell.value)

# -----------------------------
# Read sheet EXACTLY as Excel
# -----------------------------
def read_sheet_exact(sheet_name):
    wb = load_workbook(EXCEL_FILE, data_only=True)
    ws = wb[sheet_name]

    title = ws["A1"].value or ""

    headers = [cell.value or "" for cell in ws[2]]

    rows = []
    for row in ws.iter_rows(min_row=3):
        row_data = [get_cell_display(cell) for cell in row]
        if any(row_data):
            rows.append(row_data)

    html = '<table class="uptime-table">\n<thead><tr>'
    for h in headers:
        html += f"<th>{h}</th>"
    html += "</tr></thead>\n<tbody>\n"

    for row in rows:
        html += "<tr>"
        for val in row:
            html += f"<td>{val}</td>"
        html += "</tr>\n"

    html += "</tbody></table>"

    return title, headers, rows, html

# -----------------------------
# Downtime comparison helper
# -----------------------------
def downtime_to_minutes(text):
    if not text:
        return 0

    text = text.lower()
    mins = 0

    hrs = re.search(r'(\d+)\s*hr', text)
    mins_m = re.search(r'(\d+)\s*min', text)

    if hrs:
        mins += int(hrs.group(1)) * 60
    if mins_m:
        mins += int(mins_m.group(1))

    return mins

# -----------------------------
# Load sheets
# -----------------------------
wb = load_workbook(EXCEL_FILE, data_only=True)
sheets = wb.sheetnames

weekly_range, weekly_headers, weekly_rows, weekly_table = read_sheet_exact(sheets[0])

quarterly_range = ""
quarterly_table = ""
if len(sheets) > 1:
    quarterly_range, q_headers, q_rows, quarterly_table = read_sheet_exact(sheets[1])

# -----------------------------
# Major Incident (Weekly only)
# -----------------------------
major_incident = {"account": "", "outage": "", "rca": ""}
major_story = ""

if "Total Downtime(In Mins)" in weekly_headers and "Account Name" in weekly_headers:
    idx_downtime = weekly_headers.index("Total Downtime(In Mins)")
    idx_account = weekly_headers.index("Account Name")
    idx_rca = weekly_headers.index("RCA of Outage") if "RCA of Outage" in weekly_headers else None

    max_minutes = -1
    max_row = None

    for row in weekly_rows:
        minutes = downtime_to_minutes(row[idx_downtime])
        if minutes > max_minutes:
            max_minutes = minutes
            max_row = row

    if max_row and max_minutes > 0:
        major_incident = {
            "account": max_row[idx_account],
            "outage": max_row[idx_downtime],
            "rca": max_row[idx_rca] if idx_rca is not None else ""
        }

        major_story = (
            f"<b>{major_incident['account']}</b> experienced the highest outage "
            f"of <b>{major_incident['outage']}</b> during the week."
        )

# -----------------------------
# Render HTML
# -----------------------------
with open(os.path.join(BASE_DIR, "uptime_template.html"), encoding="utf-8") as f:
    template = Template(f.read())

html = template.render(
    weekly_range=weekly_range,
    quarterly_range=quarterly_range,
    weekly_table=weekly_table,
    quarterly_table=quarterly_table,
    major_incident=major_incident,
    major_story=major_story,
    generated_date=time.strftime("%d-%b-%Y %H:%M"),
    source_file=os.path.basename(EXCEL_FILE)
)

os.makedirs(OUTPUT_DIR, exist_ok=True)
out = os.path.join(OUTPUT_DIR, "uptime_report.html")
with open(out, "w", encoding="utf-8") as f:
    f.write(html)

print("✅ REPORT GENERATED:", out)
