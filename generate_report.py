from openpyxl import load_workbook
from jinja2 import Template
import os, glob, re
from datetime import datetime
import matplotlib.pyplot as plt

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
    m = re.search(
        r'(\d{1,2})(st|nd|rd|th)[_\-\s]*([A-Za-z]+)[_\-\s]*(\d{4})',
        name,
        re.IGNORECASE
    )
    if not m:
        return None
    try:
        return datetime.strptime(
            f"{m.group(1)} {m.group(3)} {m.group(4)}",
            "%d %b %Y"
        )
    except:
        return None


def find_excel():
    files = glob.glob("*.xlsx")
    dated = [(extract_date(f), f) for f in files if extract_date(f)]
    if not dated:
        raise Exception("❌ No dated Excel file found")
    dated.sort(reverse=True)
    return dated[0][1]


EXCEL_FILE = find_excel()
print("📊 Using Excel:", EXCEL_FILE)

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
    raw_headers = [str(c.value).strip() if c.value else "" for c in ws[2]]
    headers = [h for h in raw_headers if h]

    rows = []
    for r in ws.iter_rows(min_row=3, max_col=len(raw_headers)):
        row = [str(c.value).strip() if c.value else "" for c in r]
        if any(row):
            rows.append(row[:len(headers)])

    return title, headers, rows

# =================================================
# LOAD DATA
# =================================================
wb = load_workbook(EXCEL_FILE, data_only=True)

weekly_title, weekly_headers, weekly_rows = read_sheet(wb.sheetnames[0])

quarterly_title, quarterly_headers, quarterly_rows = "", [], []
if len(wb.sheetnames) > 1:
    quarterly_title, quarterly_headers, quarterly_rows = read_sheet(
        wb.sheetnames[1]
    )

# =================================================
# INDEX HELPERS
# =================================================
def build_index(headers):
    norm = [h.lower() for h in headers]

    def idx(*names):
        for n in names:
            if n.lower() in norm:
                return norm.index(n.lower())
        return None

    return idx


w = build_index(weekly_headers)
W_ACC = w("account name", "account")
W_UP = w("total uptime", "uptime")
W_OUT = w("outage downtime")

# =================================================
# KPI CALCULATION
# =================================================
weekly_uptimes = []
downtime_map = {}

for r in weekly_rows:
    r[W_UP] = normalize_pct(r[W_UP])
    try:
        weekly_uptimes.append(float(r[W_UP].replace("%", "")))
    except:
        pass
    downtime_map[r[W_ACC]] = downtime_to_minutes(r[W_OUT])

overall_uptime = (
    f"{sum(weekly_uptimes)/len(weekly_uptimes):.2f}%"
    if weekly_uptimes else "N/A"
)

outage_count = sum(1 for v in downtime_map.values() if v > 0)
total_downtime = sum(downtime_map.values())
most_affected = max(downtime_map, key=downtime_map.get) if downtime_map else "N/A"

# =================================================
# BAR GRAPH (FIXED)
# =================================================
def make_bar_chart(rows, acc_idx, up_idx, out_file):
    accounts, uptimes = [], []

    for r in rows:
        try:
            accounts.append(r[acc_idx])
            uptimes.append(float(r[up_idx].replace("%", "")))
        except:
            pass

    if not uptimes:
        return

    plt.figure(figsize=(8, 4))
    plt.bar(accounts, uptimes)
    plt.ylim(90, 100)
    plt.ylabel("Uptime (%)")
    plt.xticks(rotation=25, ha="right")
    plt.title("Weekly Uptime by Account")
    plt.tight_layout()
    plt.savefig(os.path.join(OUTPUT_DIR, out_file))
    plt.close()


make_bar_chart(
    weekly_rows,
    W_ACC,
    W_UP,
    "uptime_bar.png"
)

# =================================================
# TABLE BUILDER
# =================================================
def build_table(headers, rows, uptime_idx):
    html = "<table class='uptime-table'><thead><tr>"
    for h in headers:
        html += f"<th>{h}</th>"
    html += "</tr></thead><tbody>"

    for r in rows:
        html += "<tr>"
        for i, v in enumerate(r):
            if i == uptime_idx:
                html += f"<td><span class='tick-circle'>✔</span>{v}</td>"
            else:
                html += f"<td>{v}</td>"
        html += "</tr>"

    html += "</tbody></table>"
    return html


weekly_table = build_table(weekly_headers, weekly_rows, W_UP)
quarterly_table = (
    build_table(quarterly_headers, quarterly_rows, None)
    if quarterly_rows else ""
)

# =================================================
# RENDER HTML
# =================================================
with open("uptime_template.html", encoding="utf-8") as f:
    template = Template(f.read())

html = template.render(
    weekly_title=weekly_title,
    quarterly_title=quarterly_title,
    weekly_table=weekly_table,
    quarterly_table=quarterly_table,
    overall_uptime=overall_uptime,
    outage_count=outage_count,
    total_downtime=total_downtime,
    most_affected=most_affected
)

out = os.path.join(OUTPUT_DIR, "uptime_report.html")
with open(out, "w", encoding="utf-8") as f:
    f.write(html)

print("✅ FINAL REPORT GENERATED:", out)
