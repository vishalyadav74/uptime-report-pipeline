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
    txt = str(txt).lower()
    mins = 0
    if m := re.search(r'(\d+)\s*hr', txt):
        mins += int(m.group(1)) * 60
    if m := re.search(r'(\d+)\s*min', txt):
        mins += int(m.group(1))
    return mins

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

def idx(headers, *names):
    h = [x.lower() for x in headers]
    for n in names:
        if n.lower() in h:
            return h.index(n.lower())

W_ACC = idx(weekly_headers, "account", "account name")
W_UP  = idx(weekly_headers, "uptime", "total uptime")
W_OUT = idx(weekly_headers, "outage downtime")
W_RCA = idx(weekly_headers, "rca")

Q_ACC = idx(quarterly_headers, "account", "account name")
Q_UP  = idx(quarterly_headers, "uptime", "total uptime")
Q_YTD = idx(quarterly_headers, "ytd", "ytd uptime")

weekly_vals = []
for r in weekly_rows:
    r[W_UP] = normalize_pct(r[W_UP])
    weekly_vals.append(float(r[W_UP].replace("%","")))

for r in quarterly_rows:
    r[Q_UP] = normalize_pct(r[Q_UP])
    if Q_YTD is not None:
        r[Q_YTD] = normalize_pct(r[Q_YTD])

overall_uptime = f"{sum(weekly_vals)/len(weekly_vals):.2f}%"
total_downtime = sum(downtime_to_minutes(r[W_OUT]) for r in weekly_rows)
outage_count = sum(1 for r in weekly_rows if downtime_to_minutes(r[W_OUT]) > 0)

major_row = max(weekly_rows, key=lambda r: downtime_to_minutes(r[W_OUT]))
major_incident = {
    "account": major_row[W_ACC],
    "outage": major_row[W_OUT],
    "rca": major_row[W_RCA] if W_RCA is not None else ""
}

# =================================================
# INFOGRAPHIC STYLE BAR (FINAL)
# =================================================
def infographic_bar(accounts, values, xlabel):
    fig, ax = plt.subplots(figsize=(7, 0.6 * len(accounts)))

    palette = ["#7c3aed","#ec4899","#0ea5e9","#2563eb",
               "#14b8a6","#22c55e","#f59e0b","#ef4444"]

    y = range(len(accounts))

    ax.barh(y, [100]*len(accounts), color="#e5e7eb", height=0.55)
    bars = ax.barh(y, values, color=[palette[i%len(palette)] for i in y], height=0.55)

    for i, bar in enumerate(bars):
        ax.text(values[i] + 0.4, bar.get_y()+bar.get_height()/2,
                f"{values[i]:.2f}%", va="center", fontsize=9, fontweight="bold")

    ax.set_yticks(y)
    ax.set_yticklabels(accounts)
    ax.set_xlim(0,100)
    ax.invert_yaxis()
    ax.set_xlabel(xlabel)

    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.tick_params(left=False, bottom=False)

    plt.tight_layout()
    buf = BytesIO()
    plt.savefig(buf, format="png", dpi=120)
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode()

weekly_bar = infographic_bar(
    [r[W_ACC] for r in weekly_rows],
    [float(r[W_UP].replace("%","")) for r in weekly_rows],
    "Uptime (%)"
)

quarterly_bar = ""
if quarterly_rows and Q_YTD is not None:
    quarterly_bar = infographic_bar(
        [r[Q_ACC] for r in quarterly_rows],
        [float(r[Q_YTD].replace("%","")) for r in quarterly_rows],
        "YTD Uptime (%)"
    )

def build_table(headers, rows):
    html = "<table class='uptime-table'><tr>"
    for h in headers:
        html += f"<th>{h}</th>"
    html += "</tr>"
    for r in rows:
        html += "<tr>"
        for v in r:
            html += f"<td>{v}</td>"
        html += "</tr>"
    return html + "</table>"

weekly_table = build_table(weekly_headers, weekly_rows)
quarterly_table = build_table(quarterly_headers, quarterly_rows) if quarterly_rows else ""

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
    weekly_bar=weekly_bar,
    quarterly_bar=quarterly_bar
)

with open(os.path.join(OUTPUT_DIR, "uptime_report.html"), "w", encoding="utf-8") as f:
    f.write(html)

print("✅ FINAL REPORT GENERATED")
