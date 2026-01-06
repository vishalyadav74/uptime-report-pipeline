from openpyxl import load_workbook
from jinja2 import Template
import os, glob, re
from datetime import datetime
import matplotlib.pyplot as plt

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# -------------------------------------------------
# PICK LATEST DATED EXCEL
# -------------------------------------------------
def extract_date(name):
    m = re.search(r'(\d{1,2})(st|nd|rd|th)\s+([A-Za-z]+)\s*_?\s*([0-9]{4})', name)
    if not m:
        return None
    try:
        return datetime.strptime(
            f"{m.group(1)} {m.group(3)} {m.group(4)}", "%d %b %Y"
        )
    except:
        return None

def find_excel():
    files = glob.glob("*.xlsx")
    dated = [(extract_date(f), f) for f in files if extract_date(f)]
    if not dated:
        raise Exception("❌ No valid dated Excel found")
    dated.sort(reverse=True)
    return dated[0][1]

EXCEL_FILE = find_excel()
print("Using Excel:", EXCEL_FILE)

# -------------------------------------------------
# HELPERS
# -------------------------------------------------
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

# -------------------------------------------------
# READ SHEET (RAW)
# -------------------------------------------------
def read_sheet(sheet):
    wb = load_workbook(EXCEL_FILE, data_only=True)
    ws = wb[sheet]

    title = ws["A1"].value or ""
    headers = [str(c.value).strip() if c.value else "" for c in ws[2]]
    rows = []

    for r in ws.iter_rows(min_row=3, max_col=len(headers)):
        row = [str(c.value).strip() if c.value else "" for c in r]
        if any(row):
            rows.append(row)

    return title, headers, rows

# -------------------------------------------------
# LOAD DATA
# -------------------------------------------------
wb = load_workbook(EXCEL_FILE, data_only=True)
weekly_title, headers, weekly_rows = read_sheet(wb.sheetnames[0])

quarterly_title = ""
quarterly_rows = []
if len(wb.sheetnames) > 1:
    quarterly_title, _, quarterly_rows = read_sheet(wb.sheetnames[1])

# -------------------------------------------------
# DYNAMIC HEADER INDEX
# -------------------------------------------------
norm_headers = [h.lower() for h in headers]

def get_idx(*names):
    for n in names:
        if n.lower() in norm_headers:
            return norm_headers.index(n.lower())
    return None

idx_account = get_idx("account name", "account")
idx_uptime  = get_idx("total uptime", "uptime", "ytd uptime")
idx_outage  = get_idx("outage downtime", "downtime")

if idx_account is None or idx_uptime is None or idx_outage is None:
    raise Exception(f"❌ Required columns not found. Headers present: {headers}")

# -------------------------------------------------
# CONVERT DECIMAL UPTIME → PERCENT (IN PLACE)
# -------------------------------------------------
def normalize_uptime(rows):
    uptimes = []
    for r in rows:
        try:
            num = float(r[idx_uptime])
            if num <= 1:
                num = num * 100
            r[idx_uptime] = f"{num:.2f}%"
            uptimes.append(num)
        except:
            pass
    return uptimes

weekly_uptimes = normalize_uptime(weekly_rows)
quarterly_uptimes = normalize_uptime(quarterly_rows) if quarterly_rows else []

# -------------------------------------------------
# KPIs (WEEKLY)
# -------------------------------------------------
overall_uptime = (
    f"{sum(weekly_uptimes)/len(weekly_uptimes):.2f}%"
    if weekly_uptimes else "N/A"
)

total_downtime = sum(
    downtime_to_minutes(r[idx_outage]) for r in weekly_rows
)

outage_count = sum(
    1 for r in weekly_rows if downtime_to_minutes(r[idx_outage]) > 0
)

max_row = max(weekly_rows, key=lambda r: downtime_to_minutes(r[idx_outage]))
most_affected = max_row[idx_account]

# -------------------------------------------------
# BUILD TABLE (GREEN / RED LOGIC)
# -------------------------------------------------
def build_table(headers, rows):
    html = "<table class='uptime-table'><thead><tr>"
    for h in headers:
        html += f"<th>{h}</th>"
    html += "</tr></thead><tbody>"

    for i, r in enumerate(rows):
        bg = "#fafafa" if i % 2 else "#ffffff"
        html += f"<tr style='background:{bg}'>"

        for j, v in enumerate(r):
            if j == idx_uptime:
                if downtime_to_minutes(r[idx_outage]) > 0:
                    html += f"<td style='color:#dc2626;font-weight:600;'>{v}</td>"
                else:
                    html += f"<td style='color:#16a34a;font-weight:600;'>✔ {v}</td>"
            else:
                html += f"<td>{v}</td>"

        html += "</tr>"

    html += "</tbody></table>"
    return html

weekly_table = build_table(headers, weekly_rows)
quarterly_table = build_table(headers, quarterly_rows) if quarterly_rows else ""

# -------------------------------------------------
# CHART (WEEKLY DOWNTIME)
# -------------------------------------------------
labels, values = [], []
for r in weekly_rows:
    mins = downtime_to_minutes(r[idx_outage])
    if mins > 0:
        labels.append(r[idx_account])
        values.append(mins)

if values:
    plt.figure(figsize=(4,4))
    plt.pie(values, labels=labels, autopct="%1.1f%%")
    plt.title("Weekly Downtime Breakdown")
    plt.savefig(os.path.join(OUTPUT_DIR, "downtime_chart.png"))
    plt.close()

# -------------------------------------------------
# MAJOR INCIDENT
# -------------------------------------------------
major_incident = {
    "account": max_row[idx_account],
    "outage": max_row[idx_outage],
    "rca": ""
}

# -------------------------------------------------
# RENDER HTML
# -------------------------------------------------
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
    major_incident=major_incident
)

out = os.path.join(OUTPUT_DIR, "uptime_report.html")
with open(out, "w", encoding="utf-8") as f:
    f.write(html)

print("✅ Report generated:", out)
