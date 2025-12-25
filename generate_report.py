import pandas as pd
from jinja2 import Template
import os
import sys
import glob
import time

# -----------------------------
# SMART EXCEL FILE DETECTOR
# -----------------------------
def find_best_excel_file():
    print("🔍 Scanning for Excel files...")
    files = []
    for ext in ("*.xlsx", "*.xls"):
        files.extend(glob.glob(ext))

    if not files:
        raise Exception("❌ No Excel file found")

    files.sort()
    print("📊 Found Excel files:", files)
    return files[0]

# -----------------------------
# CONFIG
# -----------------------------
SLA_THRESHOLD = 99.9
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# -----------------------------
# EXCEL FILE INPUT
# -----------------------------
if len(sys.argv) > 1:
    EXCEL_FILE = sys.argv[1]
elif os.getenv("UPTIME_EXCEL"):
    EXCEL_FILE = os.getenv("UPTIME_EXCEL")
else:
    EXCEL_FILE = find_best_excel_file()

if not os.path.exists(EXCEL_FILE):
    raise Exception(f"Excel file not found: {EXCEL_FILE}")

print(f"✅ Processing Excel: {EXCEL_FILE}")

# -----------------------------
# HELPER FUNCTIONS
# -----------------------------
def parse_downtime_to_minutes(val):
    if pd.isna(val):
        return 0

    val = str(val).lower()
    minutes = 0

    if "hr" in val:
        minutes += int(val.split("hr")[0].split()[-1]) * 60
    if "min" in val:
        minutes += int(val.split("min")[0].split()[-1])
    if "sec" in val:
        minutes += int(val.split("sec")[0].split()[-1]) / 60

    return round(minutes, 2)

def format_uptime(val):
    if pd.isna(val):
        return "N/A"

    val = str(val).replace("%", "").strip()
    uptime = float(val)

    css = "good" if uptime >= SLA_THRESHOLD else "bad"
    return f'<span class="{css}">{uptime:.2f}%</span>'

def read_uptime_sheet(sheet_name):
    """
    Handles:
    - Title in row 1 (A1)
    - Headers in row 2
    """
    raw = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=None)

    title = str(raw.iloc[0, 0]).strip()

    df = raw.iloc[1:].copy()
    df.columns = raw.iloc[1]
    df = df.iloc[1:].reset_index(drop=True)

    expected_cols = [
        "Account Name",
        "Total Uptime",
        "Planned Downtime",
        "Outage Downtime",
        "Total Downtime(In Mins)",
        "Remarks",
        "RCA of Outage"
    ]

    df = df[expected_cols]

    df["Total Uptime"] = df["Total Uptime"].apply(format_uptime)
    df["Outage Minutes"] = df["Total Downtime(In Mins)"].apply(parse_downtime_to_minutes)

    html_table = df.to_html(index=False, classes="uptime-table", escape=False)

    return title, df, html_table

# -----------------------------
# READ SHEETS
# -----------------------------
xls = pd.ExcelFile(EXCEL_FILE, engine="openpyxl")

weekly_sheet = xls.sheet_names[0]
quarterly_sheet = xls.sheet_names[1] if len(xls.sheet_names) > 1 else None

print("📑 Weekly Sheet:", weekly_sheet)
print("📑 Quarterly Sheet:", quarterly_sheet)

weekly_range, weekly_df, weekly_table = read_uptime_sheet(weekly_sheet)

quarterly_table = "<p>No quarterly data</p>"
quarterly_range = ""
if quarterly_sheet:
    quarterly_range, quarterly_df, quarterly_table = read_uptime_sheet(quarterly_sheet)

# -----------------------------
# MAJOR INCIDENT LOGIC
# -----------------------------
major_incident = {"account": "N/A", "outage": "0 mins", "rca": ""}
major_story = "No unplanned outages observed."

if weekly_df["Outage Minutes"].max() > 0:
    row = weekly_df.loc[weekly_df["Outage Minutes"].idxmax()]
    major_incident = {
        "account": row["Account Name"],
        "outage": f"{int(row['Outage Minutes'])} mins",
        "rca": row["RCA of Outage"]
    }
    major_story = (
        f"<b>{row['Account Name']}</b> had the highest outage of "
        f"<b>{int(row['Outage Minutes'])} minutes</b>.<br>"
        f"<b>Root Cause:</b> {row['RCA of Outage']}"
    )

# -----------------------------
# RENDER HTML
# -----------------------------
template_path = os.path.join(BASE_DIR, "uptime_template.html")
with open(template_path, encoding="utf-8") as f:
    template = Template(f.read())

html = template.render(
    weekly_range=weekly_range,
    quarterly_range=quarterly_range,
    weekly_table=weekly_table,
    quarterly_table=quarterly_table,
    major_incident=major_incident,
    major_story=major_story,
    generated_date=time.strftime("%d-%b-%Y %H:%M"),
    source_file=os.path.basename(EXCEL_FILE)
)

os.makedirs(OUTPUT_DIR, exist_ok=True)
output_file = os.path.join(OUTPUT_DIR, "uptime_report.html")

with open(output_file, "w", encoding="utf-8") as f:
    f.write(html)

print("✅ REPORT GENERATED")
print("📄 Output:", output_file)
print("📊 Weekly Records:", len(weekly_df))
