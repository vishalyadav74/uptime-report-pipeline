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
    m = re.search(r'(\d{1,2})(st|nd|rd|th)\s+([A-Za-z]+).*?([0-9]{4})', name)
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
    dated.sort(reverse=True)
    return dated[0][1]

EXCEL_FILE = find_excel()
print("Using Excel:", EXCEL_FILE)

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
    raw_headers = [str(c.value).strip() if c.value else "" for c in ws[2]]

    # 🔥 REMOVE EMPTY TRAILING HEADERS
    valid_len = len(raw_headers)
    while valid_len > 0 and raw_headers[valid_len-1] == "":
        valid_len -= 1

    headers = raw_headers[:valid_len]
    rows = []

    for r in ws.iter_rows(min_row=3, max_col=valid_len):
        row = [str(c.value).strip() if c.value else "" for c in r]
        if any(row):
            rows.append(row)

    return title, headers, rows

# ---------------- LOAD DATA ----------------
wb = load_workbook(EXCEL_FILE, data_only=True)
weekly_title, headers, weekly_rows = read_sheet(wb.sheetnames[0])

quarterly_title = ""
quarterly_rows = []
if len(wb.sheetnames) > 1:
    quarterly_title, _, quarterly_rows = read_sheet(wb.sheetnames[1])

# ---------------- HEADER INDEX ----------------
norm_headers = [h.lower() for h in headers]

def idx(*names):
    for n in names:
        if n.lower() in norm_headers:
            return norm_headers.index(n.lower())
    return None

i_account = idx("account name")
i_uptime  = idx("total uptime")
i_outage  = idx("outage downtime")

# ---------------- UPTIME NORMALIZE (ONLY THIS COLUMN) ----------------
def normalize_uptime(rows):
    vals = []
    for r in rows:
        try:
            num = float(r[i_uptime])
            if num <= 1:
                num *= 100
            r[i_uptime] = f"{num:.2f}%"
            vals.append(num)
        except:
            pass
    return vals

weekly_uptimes = normalize_uptime(weekly_rows)
quarterly_uptimes = normalize_uptime(quarterly_rows)

# ---------------- KPIs ----------------
overall_uptime = f"{sum(weekly_uptimes)/len(weekly_uptimes):.2f}%"
total_downtime = sum(downtime_to_minutes(r[i_outage]) for r in weekly_rows)
outage_count = sum(1 for r in weekly_rows if downtime_to_minutes(r[i_outage]) > 0)

max_row = max(weekly_rows, key=lambda r: downtime_to_minutes(r[i_outage]))
most_affected = max_row[i_account]

# ---------------- TABLE BUILDER ----------------
def build_table(headers, rows):
    html = "<table class='uptime-table'><thead><tr>"
    for h in headers:
        html += f"<th>{h}</th>"
    html += "</tr></thead><tbody>"

    for r in rows:
        html += "<tr>"
        for j, v in enumerate(r):
            if j == i_uptime:
                if downtime_to_minutes(r[i_outage]) > 0:
                    html += f"<td style='color:#dc2626;font-weight:600;'>{v}</td>"
                else:
                    html += f"<td><span class='tick-circle'>✓</span><span style='color:#16a34a;font-weight:600;'>{v}</span></td>"
            else:
                html += f"<td>{v}</td>"
        html += "</tr>"

    html += "</tbody></table>"
    return html

weekly_table = build_table(headers, weekly_rows)
quarterly_table = build_table(headers, quarterly_rows)

# ---------------- DONUT CHART ----------------
labels, values = [], []
for r in weekly_rows:
    mins = downtime_to_minutes(r[i_outage])
    if mins > 0:
        labels.append(r[i_account])
        values.append(mins)

if values:
    plt.figure(figsize=(4,4))
    plt.pie(values, labels=labels, autopct="%1.1f%%", startangle=90)
    plt.savefig(os.path.join(OUTPUT_DIR, "downtime_chart.png"))
    plt.close()

# ---------------- MAJOR INCIDENT ----------------
major_incident = {
    "account": max_row[i_account],
    "outage": max_row[i_outage],
    "rca": ""
}

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
    most_affected=most_affected,
    major_incident=major_incident
)

with open(os.path.join(OUTPUT_DIR, "uptime_report.html"), "w", encoding="utf-8") as f:
    f.write(html)

print("✅ FINAL REPORT GENERATED")
