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
    m = re.search(
        r'(\d{1,2})(st|nd|rd|th)[_\-\s]*([A-Za-z]+)[_\-\s]*(\d{4})',
        name, re.IGNORECASE
    )
    if not m:
        return None
    return datetime.strptime(
        f"{m.group(1)} {m.group(3)} {m.group(4)}",
        "%d %b %Y"
    )

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

W_ACC = idx(weekly_headers, "account", "account name")
W_UP  = idx(weekly_headers, "uptime", "total uptime")
W_OUT = idx(weekly_headers, "outage downtime")
W_RCA = idx(weekly_headers, "rca")

Q_ACC = idx(quarterly_headers, "account", "account name")
Q_UP  = idx(quarterly_headers, "uptime", "total uptime")
Q_YTD = idx(quarterly_headers, "ytd", "ytd uptime")
Q_OUT = idx(quarterly_headers, "outage downtime")

# =================================================
# NORMALIZE + KPI
# =================================================
weekly_uptimes = []

for r in weekly_rows:
    r[W_UP] = normalize_pct(r[W_UP])
    weekly_uptimes.append(float(r[W_UP].replace("%", "")))

overall_uptime = f"{sum(weekly_uptimes) / len(weekly_uptimes):.2f}%"
outage_count = sum(1 for r in weekly_rows if downtime_to_minutes(r[W_OUT]) > 0)

# =================================================
# MOST AFFECTED (WEEKLY)
# =================================================
major_incident = {"account": "N/A", "outage": "", "rca": ""}
if weekly_rows:
    major_row = max(weekly_rows, key=lambda r: downtime_to_minutes(r[W_OUT]))
    major_incident = {
        "account": major_row[W_ACC],
        "outage": major_row[W_OUT],
        "rca": major_row[W_RCA] if W_RCA is not None else ""
    }

# ✅ ONLY MOST AFFECTED ACCOUNT DOWNTIME
total_downtime = downtime_to_minutes(major_incident["outage"])

# =================================================
# WEEKLY OUTAGES
# =================================================
weekly_outages = []
for r in weekly_rows:
    mins = downtime_to_minutes(r[W_OUT])
    if mins > 0:
        weekly_outages.append({"account": r[W_ACC], "mins": mins})
weekly_outages.sort(key=lambda x: x["mins"], reverse=True)

# =================================================
# CLEAN GREEN GRAPH (95–100)
# =================================================
def bar_base64(accounts, values, ylabel):
    fig, ax = plt.subplots(figsize=(8, 3.5))
    x_pos = range(len(accounts))

    ax.bar(
        x_pos,
        values,
        color="#22c55e",
        width=0.7
    )

    ax.set_xticks(x_pos)
    ax.set_xticklabels(accounts, rotation=25, ha="right")
    ax.set_ylim(95, 100)
    ax.set_ylabel(ylabel)
    ax.set_title("Weekly Uptime by Account", fontsize=11)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    buf = BytesIO()
    plt.tight_layout()
    plt.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode()

weekly_bar = bar_base64(
    [r[W_ACC] for r in weekly_rows],
    weekly_uptimes,
    "Uptime (%)"
)

# =================================================
# TABLE BUILDER
# =================================================
def build_table(headers, rows):
    html = "<table class='uptime-table'><tr>"
    for h in headers:
        html += f"<th>{h}</th>"
    html += "</tr>"

    for r in rows:
        html += "<tr>"
        for h, v in zip(headers, r):
            cell = v
            if "%" in str(v) and "uptime" in h.lower():
                cell = (
                    "<span style='padding:2px 8px;border-radius:999px;"
                    "background:#dcfce7;color:#16a34a;font-weight:600;'>"
                    f"✔ {v}</span>"
                )
            html += f"<td>{cell}</td>"
        html += "</tr>"
    return html + "</table>"

weekly_table = build_table(weekly_headers, weekly_rows)

# =================================================
# RENDER HTML
# =================================================
with open("uptime_template.html", encoding="utf-8") as f:
    template = Template(f.read())

html = template.render(
    weekly_title=weekly_title,
    weekly_table=weekly_table,
    overall_uptime=overall_uptime,
    outage_count=outage_count,
    total_downtime=total_downtime,
    major_incident=major_incident,
    weekly_bar=weekly_bar,
    weekly_outages=weekly_outages
)

with open(os.path.join(OUTPUT_DIR, "uptime_report.html"), "w", encoding="utf-8") as f:
    f.write(html)

print("✅ FINAL REPORT GENERATED")
