from openpyxl import load_workbook
from jinja2 import Template
import os, sys, glob, time, re
from datetime import datetime
import matplotlib.pyplot as plt

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------------- PICK LATEST EXCEL ----------------
def extract_date(name):
    m = re.search(r'(\d{1,2})(st|nd|rd|th)\s+([A-Za-z]+)_([0-9]{4})', name)
    if not m:
        return None
    return datetime.strptime(f"{m.group(1)} {m.group(3)} {m.group(4)}", "%d %b %Y")

def find_excel():
    files = glob.glob("*.xlsx")
    dated = [(extract_date(f), f) for f in files if extract_date(f)]
    dated.sort(reverse=True)
    return dated[0][1]

EXCEL_FILE = find_excel()
print("Using Excel:", EXCEL_FILE)

# ---------------- HELPERS ----------------
def downtime_to_minutes(txt):
    if not txt:
        return 0
    mins = 0
    if m := re.search(r'(\d+)\s*hr', txt.lower()):
        mins += int(m.group(1)) * 60
    if m := re.search(r'(\d+)\s*min', txt.lower()):
        mins += int(m.group(1))
    return mins

# ---------------- READ SHEET ----------------
def read_sheet(sheet):
    wb = load_workbook(EXCEL_FILE, data_only=True)
    ws = wb[sheet]

    title = ws["A1"].value
    headers = [c.value for c in ws[2]]
    rows = []

    for r in ws.iter_rows(min_row=3, max_col=len(headers)):
        row = [str(c.value) if c.value else "" for c in r]
        if any(row):
            rows.append(row)

    html = "<table class='uptime-table'><thead><tr>"
    for h in headers:
        html += f"<th>{h}</th>"
    html += "</tr></thead><tbody>"

    for i, r in enumerate(rows):
        bg = "#fafafa" if i % 2 else "#ffffff"
        html += f"<tr style='background:{bg}'>"
        for v in r:
            html += f"<td>{v}</td>"
        html += "</tr>"

    html += "</tbody></table>"
    return title, html, headers, rows

# ---------------- LOAD DATA ----------------
wb = load_workbook(EXCEL_FILE, data_only=True)
weekly_title, weekly_table, headers, rows = read_sheet(wb.sheetnames[0])

quarterly_title = quarterly_table = ""
if len(wb.sheetnames) > 1:
    quarterly_title, quarterly_table, _, _ = read_sheet(wb.sheetnames[1])

# ---------------- KPIs ----------------
idx_outage = headers.index("Outage Downtime")
idx_account = headers.index("Account Name")
idx_uptime = headers.index("Total Uptime")

total_downtime = sum(downtime_to_minutes(r[idx_outage]) for r in rows)
outage_count = sum(1 for r in rows if downtime_to_minutes(r[idx_outage]) > 0)

uptimes = [float(r[idx_uptime].replace("%","")) for r in rows if "%" in r[idx_uptime]]
overall_uptime = f"{sum(uptimes)/len(uptimes):.2f}%"

max_row = max(rows, key=lambda r: downtime_to_minutes(r[idx_outage]))
most_affected = max_row[idx_account]

# ---------------- CHART ----------------
labels, values = [], []
for r in rows:
    mins = downtime_to_minutes(r[idx_outage])
    if mins > 0:
        labels.append(r[idx_account])
        values.append(mins)

plt.figure(figsize=(4,4))
plt.pie(values, labels=labels, autopct="%1.1f%%")
plt.title("Weekly Downtime Breakdown")
plt.savefig(os.path.join(OUTPUT_DIR, "downtime_chart.png"))
plt.close()

# ---------------- MAJOR INCIDENT ----------------
major_incident = {
    "account": most_affected,
    "outage": max_row[idx_outage],
    "rca": ""
}

# ---------------- RENDER HTML ----------------
with open("uptime_template.html") as f:
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
with open(out, "w") as f:
    f.write(html)

print("Report generated:", out)
