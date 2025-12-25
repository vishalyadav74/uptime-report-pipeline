import pandas as pd
from jinja2 import Template
import os
import sys
import glob
import time

# -----------------------------
# CONFIG
# -----------------------------
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
# READ & CLEAN SHEET
# -----------------------------
def read_uptime_sheet(sheet_name, is_quarterly=False):
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

    # Remove duplicate columns
    df = df.loc[:, ~df.columns.duplicated()]

    # Normalize column names
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

    # Required columns (NO Outage Minutes anywhere)
    required_cols = [
        "Account Name",
        "Total Uptime",
        "Planned Downtime",
        "Outage Downtime",
        "Total Downtime(In Mins)",
        "Remarks"
    ]

    if not is_quarterly:
        required_cols.append("RCA of Outage")

    for col in required_cols:
        if col not in df.columns:
            raise Exception(f"❌ Missing required column in {sheet_name}: {col}")

    # Quarterly: add empty RCA column if missing
    if is_quarterly and "RCA of Outage" not in df.columns:
        df["RCA of Outage"] = ""

    df = df[required_cols + (["RCA of Outage"] if is_quarterly else [])]

    # Total Uptime AS-IS (Excel value only)
    df["Total Uptime"] = df["Total Uptime"].astype(str)

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
    quarterly_range, quarterly_df, quarterly_table = read_uptime_sheet(
        quarterly_sheet, is_quarterly=True
    )

# -----------------------------
# DUMMY MAJOR INCIDENT (TEMPLATE SAFE)
# -----------------------------
major_incident = {
    "account": "",
    "outage": "",
    "rca": ""
}
major_story = ""

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
