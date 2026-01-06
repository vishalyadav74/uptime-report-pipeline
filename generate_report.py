from openpyxl import load_workbook
from jinja2 import Template
import os, glob, re
from datetime import datetime
import matplotlib.pyplot as plt

# ---------------- PATHS ----------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------------- PICK LATEST EXCEL ----------------
def extract_date(name):
    m = re.search(r'(\d{1,2})(st|nd|rd|th)\s+([A-Za-z]+)\s*(\d{4})', name)
    if not m:
        return None
    return datetime.strptime(f"{m.group(1)} {m.group(3)} {m.group(4)}", "%d %b %Y")

def find_excel():
    files = glob.glob("*.xlsx")
    dated = [(extract_date(f), f) for f in files if extract_date(f)]
    if not dated:
        raise Exception("❌ No dated Excel found")
    dated.sort(reverse=True)
    return dated[0][1]

EXCEL_FILE = find_excel()
print("Using:", EXCEL_FILE)

# ---------------- HELPERS ----------------
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

# ---------------- READ SHEET ----------------
def read_sheet(sheet):
    wb = load_workbook(EXCEL_FILE, data_only=True)
    ws = wb[sheet]
    title = ws["A1"].value or ""
    headers = [str(c.value).strip() for c in ws[2]]
    rows = []
    for r in ws.iter_rows(min_row=3, max_col=len(headers)):
        row = [str(c.value).strip() if c.value else "" for c in r]
        if any(row):
            rows.append(row)
    return title, headers, rows

wb = load_workbook(EXCEL_FILE, data_only=True)
weekly_title, headers, weekly_rows = read_sheet(wb.sheetnames[0])
quarterly_title, quarterly_rows = "", []
if len(wb.sheetnames) > 1:
    quarterly_title, _, quarterly_rows = read_sheet(wb.sheetnames[1])

# ---------------- HEADER INDEX ----------------
norm = [h.lower() for h in headers]
def idx(*names):
    for n in names:
        if n.lower() in norm:
            return norm.index(n.lower())
    return None

IDX_ACCOUNT = idx("account name", "account")
IDX_UPTIME  = idx("total uptime", "uptime")
IDX_OUTAGE  = idx("outage downtime")
IDX_RCA     = idx("rca of outage")

# ---------------- NORMALIZE UPTIME ----------------
def normalize(rows):
    vals = []
    for r in rows:
        try:
            v = float(r[IDX_UPTIME])
            if v <= 1:
                v *= 100
            r[IDX_UPTIME] = f"{v:.2f}%"
            vals.append(v)
        except:
            pass
    return vals

weekly_uptimes = normalize(weekly_rows)
normalize(quarterly_rows)

# ---------------- KPIs ----------------
overall_uptime = f"{sum(weekly_uptimes)/len(weekly_uptimes):.2f}%"
total_downtime = sum(downtime_to_minutes(r[IDX_OUTAGE]) for r in weekly_rows)
outage_count   = sum(1 for r in weekly_rows if downtime_to_minutes(r[IDX_OUTAGE]) > 0)

major_row = max(weekly_rows, key=lambda r: downtime_to_minutes(r[IDX_OUTAGE]))
major_incident = {
    "account": major_row[IDX_ACCOUNT],
    "outage": major_row[IDX_OUTAGE],
    "rca": major_row[IDX_RCA] if IDX_RCA is not None else ""
}

# ---------------- WEEKLY BREAKDOWN ----------------
breakdown = []
total_outage = sum(downtime_to_minutes(r[IDX_OUTAGE]) for r in weekly_rows)

for r in weekly_rows:
    mins = downtime_to_minutes(r[IDX_OUTAGE])
    if mins > 0:
        pct = (mins / total_outage) * 100 if total_outage else 0
        breakdown.append({
            "account": r[IDX_ACCOUNT],
            "percent": f"{pct:.1f}%"
        })

affected_accounts = " | ".join(b["account"] for b in breakdown)

# ---------------- BUILD TABLE ----------------
def build_table(headers, rows):
    html = "<table class='uptime-table'><thead><tr>"
    for h in headers:
        html += f"<th>{h}</th>"
    html += "</tr></thead><tbody>"
    for r in rows:
        html += "<tr>"
        for i, v in enumerate(r):
            if i == IDX_UPTIME:
                html += f"<td><span class='tick-circle'>✔</span>{v}</td>"
            else:
                html += f"<td>{v}</td>"
        html += "</tr>"
    html += "</tbody></table>"
    return html

weekly_table = build_table(headers, weekly_rows)
quarterly_table = build_table(headers, quarterly_rows) if quarterly_rows else ""

# ---------------- DONUT ----------------
labels, values = [], []
for r in weekly_rows:
    mins = downtime_to_minutes(r[IDX_OUTAGE])
    if mins > 0:
        labels.append(r[IDX_ACCOUNT])
        values.append(mins)

plt.figure(figsize=(4,4))
plt.pie(values, startangle=90, wedgeprops=dict(width=0.32), autopct="%1.1f%%")
plt.text(0,0, overall_uptime, ha="center", va="center",
         fontsize=14, fontweight="bold")
plt.axis("equal")
plt.savefig(os.path.join(OUTPUT_DIR, "downtime_chart.png"),
            bbox_inches="tight", pad_inches=0.05)
plt.close()

# ---------------- RENDER ----------------
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
    major_incident=major_incident,
    breakdown=breakdown,
    affected_accounts=affected_accounts
)

out = os.path.join(OUTPUT_DIR, "uptime_report.html")
with open(out, "w", encoding="utf-8") as f:
    f.write(html)

print("FINAL REPORT:", out)
