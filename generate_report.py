from openpyxl import load_workbook
from jinja2 import Template
import os, sys, glob, time, re
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# ----------------------------------
# Pick LATEST dated Excel file
# uptime_latest_25th Dec_2025.xlsx
# ----------------------------------
def extract_date_from_filename(filename):
    match = re.search(r'(\d{1,2})(st|nd|rd|th)\s+([A-Za-z]+)_([0-9]{4})', filename)
    if not match:
        return None
    try:
        return datetime.strptime(
            f"{match.group(1)} {match.group(3)} {match.group(4)}",
            "%d %b %Y"
        )
    except:
        return None

def find_excel():
    files = glob.glob("*.xlsx") + glob.glob("*.xls")
    dated = [(extract_date_from_filename(f), f) for f in files if extract_date_from_filename(f)]
    if not dated:
        raise Exception("❌ No valid dated Excel file found")
    dated.sort(reverse=True)
    return dated[0][1]

EXCEL_FILE = sys.argv[1] if len(sys.argv) > 1 else find_excel()
print(f"✅ Using Excel: {EXCEL_FILE}")

# ----------------------------------
# Helpers
# ----------------------------------
def get_cell_display(cell):
    if cell.value is None:
        return ""

    if isinstance(cell.value, (int, float)) and "%" in str(cell.number_format):
        return f"{cell.value * 100:.2f}%"

    return str(cell.value)

def wrap_uptime(val):
    try:
        num = float(val.replace("%", ""))
        cls = "uptime-bad" if num < 99.95 else "uptime-good"
        return f'<span class="{cls}">{val}</span>'
    except:
        return val

# ----------------------------------
# Read sheet EXACT (title from A1)
# ----------------------------------
def read_sheet_exact(sheet):
    wb = load_workbook(EXCEL_FILE, data_only=True)
    ws = wb[sheet]

    title = ws["A1"].value or ""
    headers = [c.value or "" for c in ws[2]]

    # ❌ Remove column after RCA of Outage (if exists)
    if "RCA of Outage" in headers:
        headers = headers[: headers.index("RCA of Outage") + 1]

    rows = []
    for r in ws.iter_rows(min_row=3, max_col=len(headers)):
        row = [get_cell_display(c) for c in r]
        if any(row):
            rows.append(row)

    # Apply uptime coloring
    for col in ["Total Uptime", "YTD uptime"]:
        if col in headers:
            idx = headers.index(col)
            for r in rows:
                r[idx] = wrap_uptime(r[idx])

    html = "<table class='uptime-table'><thead><tr>"
    for h in headers:
        html += f"<th>{h}</th>"
    html += "</tr></thead><tbody>"

    for r in rows:
        html += "<tr>"
        for v in r:
            html += f"<td>{v}</td>"
        html += "</tr>"

    html += "</tbody></table>"
    return title, html, headers, rows

# ----------------------------------
# Downtime helper
# ----------------------------------
def downtime_to_minutes(txt):
    if not txt:
        return 0
    mins = 0
    if m := re.search(r'(\d+)\s*hr', txt.lower()):
        mins += int(m.group(1)) * 60
    if m := re.search(r'(\d+)\s*min', txt.lower()):
        mins += int(m.group(1))
    return mins

# ----------------------------------
# Load sheets
# ----------------------------------
wb = load_workbook(EXCEL_FILE, data_only=True)
sheets = wb.sheetnames

weekly_title, weekly_table, weekly_headers, weekly_rows = read_sheet_exact(sheets[0])
quarterly_title = quarterly_table = ""

if len(sheets) > 1:
    quarterly_title, quarterly_table, _, _ = read_sheet_exact(sheets[1])

# ----------------------------------
# Major Incident (Weekly only)
# ----------------------------------
major_incident = {"account": "", "outage": ""}
major_story = ""

if "Total Downtime(In Mins)" in weekly_headers:
    idx_d = weekly_headers.index("Total Downtime(In Mins)")
    idx_a = weekly_headers.index("Account Name")

    max_row = max(weekly_rows, key=lambda r: downtime_to_minutes(r[idx_d]))
    if downtime_to_minutes(max_row[idx_d]) > 0:
        major_incident = {
            "account": max_row[idx_a],
            "outage": max_row[idx_d]
        }
        major_story = (
            f"<b>{major_incident['account']}</b> experienced the highest outage "
            f"of <b>{major_incident['outage']}</b> during the week."
        )

# ----------------------------------
# Render HTML
# ----------------------------------
with open(os.path.join(BASE_DIR, "uptime_template.html"), encoding="utf-8") as f:
    template = Template(f.read())

html = template.render(
    weekly_title=weekly_title,
    quarterly_title=quarterly_title,
    weekly_table=weekly_table,
    quarterly_table=quarterly_table,
    major_incident=major_incident,
    major_story=major_story,
    generated_date=time.strftime("%d-%b-%Y %H:%M")
)

os.makedirs(OUTPUT_DIR, exist_ok=True)
out = os.path.join(OUTPUT_DIR, "uptime_report.html")
with open(out, "w", encoding="utf-8") as f:
    f.write(html)

print("✅ REPORT GENERATED:", out)
