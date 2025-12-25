from openpyxl import load_workbook
from jinja2 import Template
import os
import sys
import glob
import time
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# -----------------------------
# Pick Excel
# -----------------------------
def find_excel():
    files = []
    for ext in ("*.xlsx", "*.xls"):
        files.extend(glob.glob(ext))
    if not files:
        raise Exception("No Excel file found")
    files.sort()
    return files[0]

EXCEL_FILE = sys.argv[1] if len(sys.argv) > 1 else find_excel()
print(f"✅ Using Excel: {EXCEL_FILE}")

# -----------------------------
# Read sheet EXACTLY as displayed
# -----------------------------
def read_sheet_exact(sheet_name):
    wb = load_workbook(EXCEL_FILE, data_only=True)
    ws = wb[sheet_name]

    # Title (A1)
    title = ws["A1"].value or ""

    # Header row (row 2)
    headers = [cell.value for cell in ws[2]]

    data = []
    for row in ws.iter_rows(min_row=3, values_only=False):
        row_data = []
        for cell in row:
            # cell.text = EXACT Excel display value
            row_data.append(cell.text if cell.text is not None else "")
        if any(row_data):
            data.append(row_data)

    # Build HTML manually (no pandas at all)
    html = '<table class="uptime-table">\n<thead>\n<tr>'
    for h in headers:
        html += f"<th>{h}</th>"
    html += "</tr>\n</thead>\n<tbody>\n"

    for row in data:
        html += "<tr>"
        for val in row:
            html += f"<td>{val}</td>"
        html += "</tr>\n"

    html += "</tbody>\n</table>"

    return title, headers, data, html

# -----------------------------
# Downtime comparison helper (weekly only)
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

if "Total Downtime(In Mins)" in weekly_headers:
    idx_downtime = weekly_headers.index("Total Downtime(In Mins)")
    idx_account = weekly_headers.index("Account Name")
    idx_rca = weekly_headers.index("RCA of Outage")

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
            "rca": max_row[idx_rca]
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
