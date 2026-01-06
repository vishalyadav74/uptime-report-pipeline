from openpyxl import load_workbook
from jinja2 import Template
import os, glob, re
from datetime import datetime
import matplotlib.pyplot as plt

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------------- PICK LATEST EXCEL ----------------
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

# ---------------- HELPERS ----------------
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

# ---------------- READ SHEET ----------------
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

# ---------------- LOAD DATA ----------------
wb = load_workbook(EXCEL_FILE, data_only=True)
weekly_title, headers, rows = read_sheet(wb.sheetnames[0])

quarterly_title = quarterly_rows = ""
if len(wb.sheetnames) > 1:
    quarterly_title, _, quarterly_rows = read_sheet(wb.sheetnames[1])

# ---------------- DYNAMIC HEADER INDEX ----------------
norm_headers = [h.lower() for h in headers]

def get_idx(*names):
    for n in names:
        if n.lower() in norm_headers:
            return norm_headers.index(n.lower())
    return None

idx_outage  = get_idx("outage downtime", "downtime")
idx_account = get_idx("account name", "account")
idx_uptime  = get_idx("total uptime", "uptime", "ytd uptime")

if idx_outage is None or idx_account is None or idx_uptime is None:
    raise Exception(f"❌ Required columns not found. Headers present: {headers}")

# ---------------- KPIs ----------------
total_downtime = sum(downtime_to_minutes(r[idx_outage]) for r in rows)
outage_count = sum(1 for r in rows if downtime_to_minutes(r[idx_outage]) > 0)

uptimes = []
for r in rows:
    val = r[idx_uptime]
    if "%" in val:
        try:
            uptimes.append(float(val.replace("%", "").strip()))
        except:
            pass

overall_uptime = f"{sum(uptimes)/len(uptimes):.2f}%" if uptimes else "N/A"

max_row = max(rows, key=lambda r: downtime_to_minutes(r[idx_outage]))
most_affected = max_row[idx_account]

# ---------------- BUILD WEEKLY TABLE (GREEN / RED LOGIC) ----------------
weekly_table = "<table class='uptime-table'><thead><tr>"
for h in headers:
    weekly_table += f"<th>{h}</th>"
weekly_table += "</tr></thead><tbody>"

for i, r in enumerate(rows):
    bg = "#fafafa" if i % 2 else "#ffffff"
    weekly_table += f"<tr style='background:{bg}'>"

    for j, v in enumerate(r):
        # uptime column
        if j == idx_uptime and "%" in v:
            if downtime_to_minutes(r[idx_outage]) > 0:
                weekly_table += f"<td style='color:#dc2626;font-weight:600;'>{v}</td>"
            else:
                weekly_table += f"<td style='color:#16a34a;font-weight:600;'>✔ {v}</td>"
        else:
            weekly_table += f"<td>{v}</td>"

    weekly_table += "</tr>"

weekly_table += "</tbody></table>"

# ---------------- CHART ----------------
labels, values = [], []
for r in rows:
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

# ---------------- MAJOR INCIDENT ----------------
major_incident = {
    "account": most_affected,
    "outage": max_row[idx_outage],
    "rca": ""
}

# ---------------- RENDER HTML ----------------
with open("uptime_template.html", encoding="utf-8") as f:
    template = Template(f.read())

html = template.render(
    weekly_title=weekly_title,
    quarterly_title=quarterly_title,
    weekly_table=weekly_table,
    quarterly_table="",   # quarterly simple for now
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
