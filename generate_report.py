import pandas as pd
from jinja2 import Template
import os
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# -------------------------------------------------
# Excel from Jenkins file parameter
# -------------------------------------------------
EXCEL_FILE = os.getenv("UPTIME_EXCEL")

if not EXCEL_FILE or not os.path.exists(EXCEL_FILE):
    raise Exception("❌ UPTIME_EXCEL not provided or file not found")

print(f"📄 Using Excel file: {EXCEL_FILE}")

SLA_THRESHOLD = 99.9


# -------------------------------------------------
# Helpers
# -------------------------------------------------
def extract_date_range(text):
    if not text:
        return "N/A"
    match = re.search(r"\((.*?)\)", str(text))
    return match.group(1) if match else str(text)


def clean_df(df):
    df = df.fillna("")
    for col in df.columns:
        col_str = str(col).lower()
        if "uptime" in col_str:
            try:
                df[col] = pd.to_numeric(df[col], errors="coerce").apply(
                    lambda x: f'<span class="{"good" if x >= SLA_THRESHOLD else "bad"}">{x:.2f}%</span>'
                    if pd.notna(x) else ""
                )
            except:
                pass
    return df


def to_minutes(val):
    try:
        if isinstance(val, str):
            v = val.lower()
            hrs = mins = 0
            if "hr" in v:
                hrs = int(re.search(r"(\d+)\s*hr", v).group(1))
            if "min" in v:
                mins = int(re.search(r"(\d+)\s*min", v).group(1))
            return hrs * 60 + mins
        return int(val)
    except:
        return 0


def find_column(df, keywords):
    for col in df.columns:
        col_l = str(col).lower()
        if all(k in col_l for k in keywords):
            return col
    return None


# -------------------------------------------------
# Load sheets
# -------------------------------------------------
xls = pd.ExcelFile(EXCEL_FILE, engine="openpyxl")
sheet_map = {s.strip().lower(): s for s in xls.sheet_names}

weekly_sheet = sheet_map.get("weekly")
quarterly_sheet = sheet_map.get("quarterly")

if not weekly_sheet or not quarterly_sheet:
    raise Exception("❌ Weekly / Quarterly sheet missing")

# -------------------------------------------------
# Read date ranges (first non-empty cell)
# -------------------------------------------------
def read_title(sheet):
    df = pd.read_excel(EXCEL_FILE, sheet_name=sheet, header=None)
    for val in df.iloc[0]:
        if pd.notna(val):
            return extract_date_range(val)
    return "N/A"

WEEKLY_RANGE = read_title(weekly_sheet)
QUARTERLY_RANGE = read_title(quarterly_sheet)

print(f"📅 Weekly Range    : {WEEKLY_RANGE}")
print(f"📅 Quarterly Range : {QUARTERLY_RANGE}")

# -------------------------------------------------
# Read actual data (skip title row)
# -------------------------------------------------
weekly_df = pd.read_excel(EXCEL_FILE, sheet_name=weekly_sheet, skiprows=1)
quarterly_df = pd.read_excel(EXCEL_FILE, sheet_name=quarterly_sheet, skiprows=1)

weekly_df = clean_df(weekly_df)
quarterly_df = clean_df(quarterly_df)

weekly_table = weekly_df.to_html(index=False, classes="uptime-table", escape=False)
quarterly_table = quarterly_df.to_html(index=False, classes="uptime-table", escape=False)

# -------------------------------------------------
# MAJOR INCIDENT LOGIC (robust)
# -------------------------------------------------
outage_col = find_column(weekly_df, ["outage"])
rca_col = find_column(weekly_df, ["rca"])
account_col = find_column(weekly_df, ["account"])

if outage_col:
    weekly_df["_outage_mins"] = weekly_df[outage_col].apply(to_minutes)
    outage_df = weekly_df[weekly_df["_outage_mins"] > 0]
else:
    outage_df = pd.DataFrame()

if outage_df.empty:
    major_incident = {"account": "N/A", "outage": "0 mins"}
    major_story = f"No unplanned outages observed during ({WEEKLY_RANGE})."
else:
    row = outage_df.loc[outage_df["_outage_mins"].idxmax()]
    account = row.get(account_col, "N/A")
    outage = row["_outage_mins"]
    rca = row.get(rca_col, "").strip() if rca_col else ""

    major_incident = {"account": account, "outage": f"{outage} mins"}
    major_story = (
        f"<b>{account}</b> experienced the highest unplanned outage of "
        f"<b>{outage} mins</b> during ({WEEKLY_RANGE}).<br>"
        f"<b>Root Cause:</b> {rca if rca else 'RCA yet to be shared.'}"
    )

# -------------------------------------------------
# Render HTML
# -------------------------------------------------
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
