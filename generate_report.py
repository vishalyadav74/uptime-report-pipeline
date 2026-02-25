# SaaS Uptime Report

from jinja2 import Template
import os, glob, re, base64
from datetime import datetime
import matplotlib.pyplot as plt
from io import BytesIO

# PATHS
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# PICK LATEST EXCEL
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

# HELPERS
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

# READ SHEET
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

# INDEX
def idx(headers, *names):
    h = [x.lower() for x in headers]
    for n in names:
        if n.lower() in h:
            return h.index(n.lower())
    return None

W_ACC = idx(weekly_headers, "account", "account name")
W_UP  = idx(weekly_headers, "uptime", "total uptime")
W_OUT = idx(weekly_headers, "outage downtime")

Q_ACC = idx(quarterly_headers, "account", "account name")
Q_UP  = idx(quarterly_headers, "uptime", "total uptime")
Q_YTD = idx(quarterly_headers, "ytd", "ytd uptime")
Q_OUT = idx(quarterly_headers, "outage downtime")

# KPI CALCULATIONS (UNCHANGED)
weekly_uptimes = []
for r in weekly_rows:
    r[W_UP] = normalize_pct(r[W_UP])
    weekly_uptimes.append(float(r[W_UP].replace("%", "")))

for r in quarterly_rows:
    r[Q_UP] = normalize_pct(r[Q_UP])
    if Q_YTD is not None:
        r[Q_YTD] = normalize_pct(r[Q_YTD])

overall_uptime = f"{sum(weekly_uptimes)/len(weekly_uptimes):.2f}%"
outage_count = sum(1 for r in weekly_rows if downtime_to_minutes(r[W_OUT]) > 0)

# OUTAGES LIST
weekly_outages = []
for r in weekly_rows:
    mins = downtime_to_minutes(r[W_OUT])
    if mins > 0:
        weekly_outages.append({"account": r[W_ACC], "mins": mins})

weekly_outages.sort(key=lambda x: x["mins"], reverse=True)

quarterly_outages = []
if quarterly_rows and Q_OUT is not None:
    for r in quarterly_rows:
        mins = downtime_to_minutes(r[Q_OUT])
        if mins > 0:
            quarterly_outages.append({"account": r[Q_ACC], "mins": mins})
    quarterly_outages.sort(key=lambda x: x["mins"], reverse=True)

# MOST AFFECTED ACCOUNT + DOWNTIME (FIX)
major_incident = {"account": weekly_outages[0]["account"] if weekly_outages else "N/A"}
affected_accounts = [o["account"] for o in weekly_outages]

most_affected_downtime = 0
if weekly_outages:
    most_affected_downtime = weekly_outages[0]["mins"]

# BAR GRAPH
def bar_base64(accounts, values, ylabel):
    fig, ax = plt.subplots(figsize=(6.8, 3.2))
    ax.bar(range(len(accounts)), values, color="#22c55e", width=0.55)
    ax.set_ylim(99, 100)
    ax.set_ylabel(ylabel, fontsize=10)
    ax.set_xticks(range(len(accounts)))
    ax.set_xticklabels(accounts, rotation=30, ha="right", fontsize=9)
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

quarterly_bar = None
if quarterly_rows and Q_YTD is not None:
    quarterly_bar = bar_base64(
        [r[Q_ACC] for r in quarterly_rows],
        [float(r[Q_YTD].replace("%", "")) for r in quarterly_rows],
        "YTD Uptime (%)"
    )

# TABLES
def build_table(headers, rows):
    col_count = len(headers)
    html = (
        "<table width='100%' cellpadding='6' cellspacing='0' "
        "style='border-collapse:separate;border-spacing:0;"
        "border-left:1px solid #e5e7eb;border-right:1px solid #e5e7eb;'>"
        "<tr>"
    )

    for i, h in enumerate(headers):
        rb = "border-right:1px solid #f1f5f9;" if i < col_count - 1 else ""
        html += (
            "<th style='background:#e01e7e;color:#ffffff;"
            "font-size:12px;font-weight:600;"
            "border-bottom:1px solid #e5e7eb;"
            f"{rb}'>"
            f"{h}</th>"
        )
    html += "</tr>"

    for r in rows:
        html += "<tr>"
        for i, v in enumerate(r):
            if "%" in str(v):
                v = (
                    "<span style='padding:2px 8px;border-radius:999px;"
                    "background:#dcfce7;color:#16a34a;font-weight:600;'>✔ "
                    f"{v}</span>"
                )
            rb = "border-right:1px solid #f1f5f9;" if i < col_count - 1 else ""
            html += (
                "<td style='font-size:12px;"
                "border-bottom:1px solid #e5e7eb;"
                f"{rb}'>"
                f"{v}</td>"
            )
        html += "</tr>"

    return html + "</table>"

weekly_table = build_table(weekly_headers, weekly_rows)
quarterly_table = build_table(quarterly_headers, quarterly_rows) if quarterly_rows else ""

# RENDER HTML
with open("uptime_template.html", encoding="utf-8") as f:
    template = Template(f.read())

html = template.render(
    weekly_title=weekly_title,
    quarterly_title=quarterly_title,
    weekly_table=weekly_table,
    quarterly_table=quarterly_table,
    overall_uptime=overall_uptime,
    outage_count=outage_count,
    total_downtime=most_affected_downtime,
    major_incident=major_incident,
    weekly_bar=weekly_bar,
    quarterly_bar=quarterly_bar,
    weekly_outages=weekly_outages,
    quarterly_outages=quarterly_outages,
    affected_accounts=affected_accounts
)

with open(os.path.join(OUTPUT_DIR, "uptime_report.html"), "w", encoding="utf-8") as f:
    f.write(html)

print("✅ FINAL REPORT GENERATED SUCCESSFULLY")
