import pandas as pd
from jinja2 import Template
import os
import sys
import glob
import time
import numbers

# -----------------------------
# CONFIG
# -----------------------------
SLA_THRESHOLD = 99.9
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

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
# EXCEL INPUT
# -----------------------------
if len(sys.argv) > 1:
    EXCEL_FILE = sys.argv[1]
elif os.getenv("UPTIME_EXCEL"):
    EXCEL_FILE = os.getenv("UPTIME_EXCEL")
else:
    EXCEL_FILE = find_best_excel_file()

if not os.path.exists(EXCEL_FILE):
    raise Exception(f"❌ Excel file not found: {EXCEL_FILE}")

print(f"✅ Processing Excel: {EXCEL_FILE}")

# -----------------------------
# HELPERS
# -----------------------------
def parse_downtime_to_minutes(val):
    if val is None:
        return 0

    # 🔐 If accidentally Series → take first value
    if isinstance(val, pd.Series):
        val = val.iloc[0]

    if isinstance(val, numbers.Number):
        return float(val)

    if pd.isna(val):
        return 0

    val = str(val).lower()
    minutes = 0

    try:
        if "hr" in val:
            minutes += int(val.split("hr")[0].split()[-1]) * 60
        if "min" in val:
            minutes += int(val.split("min")[0].split()[-1])
        if "sec" in val:
            minutes += int(val.split("sec")[0].split()[-1]) / 60
    except:
        pass

    return round(minutes, 2)

def format_uptime(val):
    if pd.isna(val):
        return "N/A"

    val = str(val).replace("%", "").strip()
    uptime = float(val)

    css = "good" if uptime >= SLA_THRESHOLD else "bad"
    return f'<span class="{css}">{uptime:.2f}%</span>'

# -----------------------------
# READ & CLEAN SHEET
# -----------------------------
def read_uptime_sheet(sheet_name):
    raw = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=None)

    # Title from A1
    title = str(raw.iloc[0, 0]).strip()

    # Header + data
    df = raw.iloc[1:].copy()
    df.columns = raw.iloc[1]
    df = df.iloc[1:].reset_index(drop=True)

    # Clean column names
    df.columns = (
        df.columns
        .astype(str)
        .str.replace("\n", " ")
        .str.strip()
    )

    # 🔥 REMOVE DUPLICATE COLUMNS (CRITICAL FIX)
    df = df.loc[:, ~df.columns.duplicated()]

    # Flexible mapping
    COLUMN_MAP = {
        "account name": "Account Name",
        "total uptime": "Total Uptime",
        "planned downtime": "Planned Downtime",
        "outage downtime": "Outage Downtime",
        "total downtime(in mins)": "Total Downtime(In Mins)",
        "total downtime (in mins)": "Total Downtime(In Mins)",
        "remarks": "Remarks",
        "rca of outage": "RCA of Outage"
    }

    rename_cols = {}
    for col in df.columns:
        key = col.lower()
        if key in COLUMN_MAP:
            rename_cols[col] = COLUMN_MAP[key]

    df = df.rename(columns=rename_cols)

    required_cols = list(dict.fromkeys(COLUMN_MAP.values()))
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise Exception(f"❌ Missing required columns in Excel: {missing}")

    df = df[required_cols]

    # Formatting
    df["Total Uptime"] = df["Total Uptime"].apply(format_uptime)

    # 🔐 Always Series now
    downtime_series = df["Total Downtime(In Mins)"]
    df["Outage Minutes"] = downtime_series.apply(parse_downtime_to_minutes)

    html_table = df.to_html(index=False, classes="uptime-table", escape=False)

    return title, df, html_table

# -----------------------------
# LOAD SHEETS
# -----------------------------
xls = pd.ExcelFile(EXCEL_FILE, engine="openpyxl")
weekly_sheet = xls.sheet_names[0]
quarterly_sheet = xls.sheet_names[1] if len(xls.sheet_names) > 1 else None

print("📑 Weekly Sheet:", weekly_sheet)
print("📑 Quarterly Sheet:", quarterly_sheet)

weekly_range, weekly_df, weekly_table = read_uptime_sheet(weekly_sheet)

quarterly_range = ""
quarterly_table = "<p>No quarterly data available</p>"
if quarterly_sheet:
    quarterly_range, quarterly_df, quarterly_table = read_uptime_sheet(quarterly_sheet)

# -----------------------------
# MAJOR INCIDENT
# -----------------------------
major_incident = {"account": "N/A", "outage": "0 mins", "rca": ""}
major_story = "No unplanned outages observed during this period."

if weekly_df["Outage Minutes"].max() > 0:
    row = weekly_df.loc[weekly_df["Outage Minutes"].idxmax()]
    major_incident = {
        "account": row["Account Name"],
        "outage": f"{int(row['Outage Minutes'])} mins",
        "rca": row["RCA of Outage"]
    }
    major_story = (
        f"<b>{row['Account Name']}</b> experienced the highest outage of "
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

print("\n✅ REPORT GENERATED SUCCESSFULLY")
print("📄 Output:", output_file)
print("📊 Weekly records:", len(weekly_df))
