from openpyxl import load_workbook
from jinja2 import Template
import os, sys, glob, time, re
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# ----------------------------------
# Pick LATEST dated Excel file
# ----------------------------------
def extract_date_from_filename(name):
    m = re.search(r'(\d{1,2})(st|nd|rd|th)\s+([A-Za-z]+)_([0-9]{4})', name)
    if not m:
        return None
    try:
        return datetime.strptime(
            f"{m.group(1)} {m.group(3)} {m.group(4)}", "%d %b %Y"
        )
    except:
        return None

def find_excel():
    files = glob.glob("*.xlsx") + glob.glob("*.xls")
    dated = [(extract_date_from_filename(f), f)
             for f in files if extract_date_from_filename(f)]
    if not dated:
        raise Exception("❌ No valid dated Excel found")
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
        num = float(val.replace("%", "").strip())
        cls = "uptime-bad" if num < 99.95 else "uptime-good"
        return f'<span class="{cls}">{val}</span>'
    except:
        return val

def downtime_to_minutes(txt):
    if not txt:
        return 0
    txt = txt.lower()
    mins = 0
    if m := re.search(r'(\d+)\s*hr', txt):
        mins += int(m.group(1)) * 60
    if m := re.search(r'(\d+)\s*min', txt):
        mins += int(m.group(1))
    return mins

# ----------------------------------
# Read sheet (OUTLOOK SAFE)
# ----------------------------------
def read_sheet(sheet_name):
    wb = load_workbook(EXCEL_FILE, data_only=True)
    ws = wb[sheet_name]

    title = ws["A1"].value or ""
    raw_headers = [str(c.value).strip() if c.value else "" for c in ws[2]]

    end_idx = len(raw_headers)
    for i, h in enumerate(raw_headers):
        if h.lower() == "rca of outage":
            end_idx = i + 1
            break

    headers = raw_headers[:end_idx]
    rows = []

    for r in ws.iter_rows(min_row=3, max_col=end_idx):
        row = [get_cell_display(c) for c in r]
        if any(row):
            rows.append(row)

    for col in ["Total Uptime", "YTD uptime"]:
        if col in headers:
            idx = headers.index(col)
            for r in rows:
                r[idx] = wrap_uptime(r[idx])

    # ✅ OUTLOOK SAFE TABLE (COLGROUP CONTROLS WIDTH)
    html = """
    <table class="uptime-table" cellpadding="0" cellspacing="0">
      <colgroup>
        <col style="width:10%">
        <col style="width:9%">
        <col style="width:10%">
        <col style="width:9%">
        <col style="width:10%">
        <col style="width:22%">  <!-- Remarks -->
        <col style="width:30%">  <!-- RCA -->
      </colgroup>
      <thead><tr>
    """

    for h in headers:
        html += f"<th>{h}</th>"
    html += "</tr></thead><tbody>"

    for i, r in enumerate(rows):
        bg = "#ffffff" if i % 2 == 0 else "#f5f5f5"
        html += f"<tr style='background:{bg};'>"
        for v in r:
            html += f"<td>{v}</td>"
        html += "</tr>"

    html += "</tbody></table>"
    return title, html, headers, rows

# ----------------------------------
# Load sheets
# ----------------------------------
wb = load_workbook(EXCEL_FILE, data_only=True)
sheets = wb.sheetnames

weekly_title, weekly_table, weekly_headers, weekly_rows = read_sheet(sheets[0])

quarterly_title = ""
quarterly_table = ""
if len(sheets) > 1:
    quarterly_title, quarterly_table, _, _ = read_sheet(sheets[1])

# ----------------------------------
# Major Incident
# ----------------------------------
major_incident = {"account": "", "outage": "", "rca": ""}
major_story = ""

norm_headers = [h.lower() for h in weekly_headers]

def idx(name):
    return norm_headers.index(name.lower()) if name.lower() in norm_headers else None

idx_out = idx("outage downtime")
idx_acc = idx("account name")
idx_rca = idx("rca of outage")

if idx_out is not None and idx_acc is not None:
    max_row = max(weekly_rows, key=lambda r: downtime_to_minutes(r[idx_out]))
    if downtime_to_minutes(max_row[idx_out]) > 0:
        major_incident["account"] = max_row[idx_acc]
        major_incident["outage"] = max_row[idx_out]
        major_incident["rca"] = max_row[idx_rca] if idx_rca is not None else ""
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
