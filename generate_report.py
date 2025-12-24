import pandas as pd
from jinja2 import Template
import os
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# -----------------------------
# EXCEL FILE (FROM JENKINS FILE PARAMETER)
# -----------------------------
EXCEL_FILE = os.getenv("UPTIME_EXCEL")

if not EXCEL_FILE or not os.path.exists(EXCEL_FILE):
    raise Exception("❌ UPTIME_EXCEL not provided. Upload Excel via Jenkins File Parameter.")

print(f"📄 Using Excel file: {EXCEL_FILE}")

SLA_THRESHOLD = 99.9


# -----------------------------
# Helper: Extract Date Range from Excel Title
# -----------------------------
def extract_date_range(title):
    if not title:
        return "N/A"
    match = re.search(r"\((.*?)\)", str(title))
    return match.group(1) if match else "N/A"


# -----------------------------
# Helper: Clean + Format Data
# -----------------------------
def clean_df(df):
    df = df.fillna("")
    for col in df.columns:
        if "uptime" in col.lower():
            try:
                df[col] = df[col].astype(float).apply(
                    lambda x: f'<span class="{"good" if x*100 >= SLA_THRESHOLD else "bad"}">{x*100:.2f}%</span>'
                )
            except:
                pass
    return df


# -----------------------------
# Helper: Convert downtime to minutes
# -----------------------------
def to_minutes(val):
    try:
        if isinstance(val, str):
            v = val.lower()
            hrs = mins = 0
            if "hr" in v:
                hrs = int(v.split("hr")[0].strip())
            if "min" in v:
                mins = int(v.split("min")[0].split()[-1])
            return hrs * 60 + mins
        return int(val)
    except:
        return 0


# -----------------------------
# Load Excel & Detect Sheets
# -----------------------------
xls = pd.ExcelFile(EXCEL_FILE, engine="openpyxl")
sheet_map = {s.strip().lower(): s for s in xls.sheet_names}

weekly_sheet = sheet_map.get("weekly")
quarterly_sheet = sheet_map.get("quarterly")

if not weekly_sheet or not quarterly_sheet:
    raise Exception("❌ Weekly or Quarterly sheet not found")

# -----------------------------
# Read Date Ranges from Excel A1
# -----------------------------
weekly_title = pd.read_excel(EXCEL_FILE, sheet_name=weekly_sheet, header=None, nrows=1).iloc[0, 0]
quarterly_title = pd.read_excel(EXCEL_FILE, sheet_name=quarterly_sheet, header=None, nrows=1).iloc[0, 0]

WEEKLY_RANGE = extract_date_range(weekly_title)
QUARTERLY_RANGE = extract_date_range(quarterly_title)

print(f"📅 Weekly Range    : {WEEKLY_RANGE}")
print(f"📅 Quarterly Range : {QUARTERLY_RANGE}")

# -----------------------------
# Load Actual Tables (skip title row)
# -----------------------------
weekly_df = pd.read_excel(EXCEL_FILE, sheet_name=weekly_sheet, skiprows=1)
quarterly_df = pd.read_excel(EXCEL_FILE, sheet_name=quarterly_sheet, skiprows=1)

weekly_df = clean_df(weekly_df)
quarterly_df = clean_df(quarterly_df)

weekly_table = weekly_df.to_html(index=False, classes="uptime-table", escape=False)
quarterly_table = quarterly_df.to_html(index=False, classes="uptime-table", escape=False)

# -----------------------------
# MAJOR INCIDENT (OUTAGE ONLY)
# -----------------------------
weekly_df["_outage_mins"] = weekly_df["Outage Downtime"].apply(to_minutes)
outage_df = weekly_df[weekly_df["_outage_mins"] > 0]

if outage_df.empty:
    major_incident = {"account": "N/A", "outage": "0 mins"}
    major_story = f"No unplanned outages observed during ({WEEKLY_RANGE})."
else:
    row = outage_df.loc[outage_df["_outage_mins"].idxmax()]
    account = row.get("Account Name", "N/A")
    outage = row["_outage_mins"]
    rca = str(row.get("RCA of Outage", "")).strip()

    major_incident = {"account": account, "outage": f"{outage} mins"}

    major_story = (
        f"<b>{account}</b> experienced the highest unplanned outage of "
        f"<b>{outage} mins</b> during ({WEEKLY_RANGE}).<br>"
        f"<b>Root Cause:</b> {rca if rca else 'RCA yet to be shared.'}"
    )

# -----------------------------
# Render HTML
# -----------------------------
with open(os.path.join(BASE_DIR, "uptime_template.html"), encoding="utf-8") as f:
    template = Template(f.read())

html = template.render(
    weekly_table=weekly_table,
    quarterly_table=quarterly_table,
    weekly_range=WEEKLY_RANGE,
    quarterly_range=QUARTERLY_RANGE,
    major_incident=major_incident,
    major_story=major_story
)

os.makedirs(OUTPUT_DIR, exist_ok=True)
output_file = os.path.join(OUTPUT_DIR, "uptime_report.html")

with open(output_file, "w", encoding="utf-8") as f:
    f.write(html)

print("🔥 EXECUTIVE UPTIME REPORT GENERATED")
print(f"📤 Output: {output_file}")
