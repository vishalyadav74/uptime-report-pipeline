from openpyxl import load_workbook
from jinja2 import Template
import os, glob, re
from datetime import datetime
import matplotlib.pyplot as plt

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ---------- PICK LATEST EXCEL ----------
def extract_date(name):
    m = re.search(r'(\d{1,2})(st|nd|rd|th)\s+([A-Za-z]+)\s*_?\s*([0-9]{4})', name)
    if not m:
        return None
    return datetime.strptime(f"{m.group(1)} {m.group(3)} {m.group(4)}", "%d %b %Y")

def find_excel():
    files = glob.glob("*.xlsx")
    dated = [(extract_date(f), f) for f in files if extract_date(f)]
    dated.sort(reverse=True)
    return dated[0][1]

EXCEL_FILE = find_excel()

# ---------- HELPERS ----------
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

# ---------- READ SHEET ----------
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
quarterly_title, _, quarterly_rows = read_sheet(wb.sheetnames[1])

norm_headers = [h.lower() for h in headers]
def idx(*names):
    for n in names:
        if n.lower() in norm_headers:
            return norm_headers.index(n.lower())
    return None

IDX_ACC = idx("account name", "account")
IDX_UP = idx("total uptime", "uptime")
IDX_OUT = idx("outage downtime")

# ---------- NORMALIZE UPTIME ----------
def normalize(rows):
    values = []
    for r in rows:
        try:
            v = float(r[IDX_UP])
            if v <= 1:
                v *= 100
            r[IDX_UP] = f"{v:.2f}%"
            values.append(v)
        except:
            pass
    return values

weekly_uptimes = normalize(weekly_rows)
quarterly_uptimes = normalize(quarterly_rows)

# ---------- KPIs ----------
overall_uptime = f"{sum(weekly_uptimes)/len(weekly_uptimes):.2f}%"
total_downtime = sum(downtime_to_minutes(r[IDX_OUT]) for r in weekly_rows)
outage_count = sum(1 for r in weekly_rows if downtime_to_minutes(r[IDX_OUT]) > 0)
max_row = max(weekly_rows, key=lambda r: downtime_to_minutes(r[IDX_OUT]))

# ---------- TABLE BUILDER ----------
def build_table(headers, rows):
    h = "<table class='uptime-table'><thead><tr>"
    for x in headers: h += f"<th>{x}</th>"
    h += "</tr></thead><tbody>"
    for i,r in enumerate(rows):
        bg = "#fafafa" if i%2 else "#fff"
        h += f"<tr style='background:{bg}'>"
        for j,v in enumerate(r):
            if j==IDX_UP:
                if downtime_to_minutes(r[IDX_OUT])>0:
                    h += f"<td class='red'>{v}</td>"
                else:
                    h += f"<td class='green'><span class='tick'>✔</span>{v}</td>"
            else:
                h += f"<td>{v}</td>"
        h += "</tr>"
    return h+"</tbody></table>"

weekly_table = build_table(headers, weekly_rows)
quarterly_table = build_table(headers, quarterly_rows)

# ---------- DONUT ----------
labels, values = [], []
for r in weekly_rows:
    m = downtime_to_minutes(r[IDX_OUT])
    if m>0:
        labels.append(r[IDX_ACC])
        values.append(m)

plt.figure(figsize=(4,4))
plt.pie(values, labels=labels, startangle=90,
        wedgeprops=dict(width=0.35))
plt.text(0,0, overall_uptime, ha='center', va='center',
         fontsize=14, fontweight='bold')
plt.savefig(os.path.join(OUTPUT_DIR,"downtime_chart.png"))
plt.close()

# ---------- RENDER ----------
with open("uptime_template.html",encoding="utf-8") as f:
    tpl = Template(f.read())

html = tpl.render(
    overall_uptime=overall_uptime,
    outage_count=outage_count,
    total_downtime=total_downtime,
    most_affected=max_row[IDX_ACC],
    major_incident={
        "account":max_row[IDX_ACC],
        "outage":max_row[IDX_OUT],
        "rca":"The issue occurred due to exceeding the maximum client connections, resulting in \"Too Many Clients\"."
    },
    weekly_title=weekly_title,
    quarterly_title=quarterly_title,
    weekly_table=weekly_table,
    quarterly_table=quarterly_table
)

with open(os.path.join(OUTPUT_DIR,"uptime_report.html"),"w",encoding="utf-8") as f:
    f.write(html)

print("✅ EXACT report generated")
