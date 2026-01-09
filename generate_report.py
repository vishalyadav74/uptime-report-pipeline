from openpyxl import load_workbook
from jinja2 import Template
import os, glob, re, base64
from datetime import datetime
import matplotlib.pyplot as plt
from io import BytesIO

# =================================================
# PICK LATEST EXCEL
# =================================================
def extract_date(name):
    m = re.search(r'(\d{1,2})(st|nd|rd|th)[_\-\s]*([A-Za-z]+)[_\-\s]*(\d{4})', name)
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
def normalize_pct(val):
    try:
        v = float(val)
        if v <= 1:
            v *= 100
        return f"{v:.2f}%"
    except:
        return val

def downtime_to_minutes(txt):
    if not txt:
        return 0
    mins = 0
    if m := re.search(r'(\d+)\s*hr', txt.lower()):
        mins += int(m.group(1)) * 60
    if m := re.search(r'(\d+)\s*min', txt.lower()):
        mins += int(m.group(1))
    return mins

# =================================================
# READ SHEET
# =================================================
def read_sheet(sheet):
    wb = load_workbook(EXCEL_FILE, data_only=True)
    ws = wb[sheet]
    title = ws["A1"].value or ""
    headers = [c.value for c in ws[2] if c.value]
    rows = []

    for r in ws.iter_rows(min_row=3, max_col=len(headers)):
        row = [str(c.value).strip() if c.value else "" for c in r]
        if any(row):
            rows.append(row[:len(headers)])

    return title, headers, rows

wb = load_workbook(EXCEL_FILE, data_only=True)

weekly_title, weekly_headers, weekly_rows = read_sheet(wb.sheetnames[0])
quarterly_title, quarterly_headers, quarterly_rows = (
    read_sheet(wb.sheetnames[1]) if len(wb.sheetnames) > 1 else ("", [], [])
)

def idx(headers, *names):
    h = [x.lower() for x in headers]
    for n in names:
        if n.lower() in h:
            return h.index(n.lower())

W_ACC = idx(weekly_headers, "account", "account name")
W_UP  = idx(weekly_headers, "uptime", "total uptime")
W_OUT = idx(weekly_headers, "outage downtime")

# =================================================
# KPI
# =================================================
uptimes, downtime = [], {}

for r in weekly_rows:
    r[W_UP] = normalize_pct(r[W_UP])
    try:
        uptimes.append(float(r[W_UP].replace("%","")))
    except:
        pass
    downtime[r[W_ACC]] = downtime_to_minutes(r[W_OUT])

overall_uptime = f"{sum(uptimes)/len(uptimes):.2f}%"
outage_count = sum(1 for v in downtime.values() if v > 0)
total_downtime = sum(downtime.values())
most_affected = max(downtime, key=downtime.get)

# =================================================
# BAR GRAPH → BASE64 (EMAIL SAFE)
# =================================================
def build_bar_base64(rows):
    acc, up = [], []
    for r in rows:
        acc.append(r[W_ACC])
        up.append(float(r[W_UP].replace("%","")))

    fig, ax = plt.subplots(figsize=(7,3))
    ax.bar(acc, up)
    ax.set_ylim(90,100)
    ax.set_ylabel("Uptime (%)")
    ax.set_title("Weekly Uptime by Account")
    plt.xticks(rotation=25, ha="right")
    plt.tight_layout()

    buf = BytesIO()
    plt.savefig(buf, format="png")
    plt.close()
    return base64.b64encode(buf.getvalue()).decode()

bar_chart = build_bar_base64(weekly_rows)

# =================================================
# TABLE BUILDER
# =================================================
def build_table(headers, rows, uptime_idx):
    html = "<table class='uptime-table'><tr>"
    for h in headers:
        html += f"<th>{h}</th>"
    html += "</tr>"
    for r in rows:
        html += "<tr>"
        for i,v in enumerate(r):
            if i == uptime_idx:
                html += f"<td>✔ {v}</td>"
            else:
                html += f"<td>{v}</td>"
        html += "</tr>"
    return html + "</table>"

weekly_table = build_table(weekly_headers, weekly_rows, W_UP)
quarterly_table = build_table(quarterly_headers, quarterly_rows, None) if quarterly_rows else ""

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
    most_affected=most_affected,
    bar_chart=bar_chart
)

with open("uptime_report.html","w",encoding="utf-8") as f:
    f.write(html)

print("✅ EMAIL SAFE REPORT READY")
