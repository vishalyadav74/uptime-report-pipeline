from openpyxl import load_workbook
from jinja2 import Template
import os, glob, re, base64
from datetime import datetime
import matplotlib.pyplot as plt
from io import BytesIO

# =================================================
# PATHS
# =================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# =================================================
# PICK LATEST EXCEL
# =================================================
def extract_date(name):
    m = re.search(r'(\d{1,2})(st|nd|rd|th)[_\-\s]*([A-Za-z]+)[_\-\s]*(\d{4})', name, re.I)
    if not m:
        return None
    return datetime.strptime(f"{m.group(1)} {m.group(3)} {m.group(4)}", "%d %b %Y")

def find_excel():
    files = glob.glob("*.xlsx")
    dated = [(extract_date(f), f) for f in files if extract_date(f)]
    dated.sort(reverse=True)
    return dated[0][1]

EXCEL_FILE = find_excel()

# =================================================
# HELPERS
# =================================================
def downtime_to_minutes(txt):
    if not txt:
        return 0
    txt = str(txt).lower()
    mins = 0
    if m := re.search(r'(\d+)\s*hr', txt):
        mins += int(m.group(1)) * 60
    if m := re.search(r'(\d+)\s*min', txt):
        mins += int(m.group(1))
    return mins

def normalize_pct(val):
    try:
        v = float(val)
        if v <= 1:
            v *= 100
        return f"{v:.2f}%"
    except:
        return val

# =================================================
# READ SHEET
# =================================================
def read_sheet(sheet):
    wb = load_workbook(EXCEL_FILE, data_only=True)
    ws = wb[sheet]
    title = ws["A1"].value or ""
    headers = [str(c.value).strip() for c in ws[2] if c.value]
    rows = []
    for r in ws.iter_rows(min_row=3, max_col=len(headers)):
        row = [str(c.value).strip() if c.value else "" for c in r]
        if any(row):
            rows.append(row)
    return title, headers, rows

wb = load_workbook(EXCEL_FILE, data_only=True)
weekly_title, weekly_headers, weekly_rows = read_sheet(wb.sheetnames[0])

quarterly_title, quarterly_headers, quarterly_rows = "", [], []
if len(wb.sheetnames) > 1:
    quarterly_title, quarterly_headers, quarterly_rows = read_sheet(wb.sheetnames[1])

# =================================================
# INDEX
# =================================================
def idx(headers, *names):
    h = [x.lower() for x in headers]
    for n in names:
        if n.lower() in h:
            return h.index(n.lower())
    return None

W_ACC = idx(weekly_headers, "account")
W_UP  = idx(weekly_headers, "uptime")
W_OUT = idx(weekly_headers, "outage downtime")

Q_ACC = idx(quarterly_headers, "account")
Q_YTD = idx(quarterly_headers, "ytd")
Q_OUT = idx(quarterly_headers, "outage downtime")

# =================================================
# WEEKLY KPI
# =================================================
weekly_uptimes = [float(normalize_pct(r[W_UP]).replace("%","")) for r in weekly_rows]
overall_uptime = f"{sum(weekly_uptimes)/len(weekly_uptimes):.2f}%"
total_downtime = sum(downtime_to_minutes(r[W_OUT]) for r in weekly_rows)
outage_count = sum(1 for r in weekly_rows if downtime_to_minutes(r[W_OUT]) > 0)

major_incident = max(
    weekly_rows,
    key=lambda r: downtime_to_minutes(r[W_OUT]),
    default=None
)
major_account = major_incident[W_ACC] if major_incident else "N/A"

# =================================================
# QUARTERLY KPI (ADD-ON)
# =================================================
quarterly_uptimes = [
    float(normalize_pct(r[Q_YTD]).replace("%",""))
    for r in quarterly_rows if Q_YTD is not None
]

quarterly_overall_uptime = (
    f"{sum(quarterly_uptimes)/len(quarterly_uptimes):.2f}%"
    if quarterly_uptimes else "N/A"
)

quarterly_total_downtime = sum(
    downtime_to_minutes(r[Q_OUT]) for r in quarterly_rows
) if Q_OUT is not None else 0

quarterly_outage_count = sum(
    1 for r in quarterly_rows if downtime_to_minutes(r[Q_OUT]) > 0
) if Q_OUT is not None else 0

quarterly_major = max(
    quarterly_rows,
    key=lambda r: downtime_to_minutes(r[Q_OUT]),
    default=None
)
quarterly_major_account = quarterly_major[Q_ACC] if quarterly_major else "N/A"

# =================================================
# RENDER
# =================================================
with open("uptime_template.html", encoding="utf-8") as f:
    template = Template(f.read())

html = template.render(
    overall_uptime=overall_uptime,
    outage_count=outage_count,
    total_downtime=total_downtime,
    major_account=major_account,

    quarterly_overall_uptime=quarterly_overall_uptime,
    quarterly_outage_count=quarterly_outage_count,
    quarterly_total_downtime=quarterly_total_downtime,
    quarterly_major_account=quarterly_major_account
)

with open(os.path.join(OUTPUT_DIR, "uptime_report.html"), "w", encoding="utf-8") as f:
    f.write(html)

print("✅ FINAL REPORT GENERATED")
