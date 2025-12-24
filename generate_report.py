import pandas as pd
from jinja2 import Template
import os
import sys

# -----------------------------
# PATH CONFIG
# -----------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# ✅ FIRST CHECK COMMAND LINE ARGUMENT
if len(sys.argv) > 1:
    EXCEL_FILE = sys.argv[1]
    print(f"📄 Using Excel file from command line: {EXCEL_FILE}")
# ✅ THEN CHECK ENVIRONMENT VARIABLE
elif "UPTIME_EXCEL" in os.environ:
    EXCEL_FILE = os.environ["UPTIME_EXCEL"]
    print(f"📄 Using Excel file from environment variable: {EXCEL_FILE}")
# ✅ FINALLY USE DEFAULT
else:
    EXCEL_FILE = os.path.join(BASE_DIR, "uptime_latest1.xlsx")
    print(f"📄 Using default Excel file: {EXCEL_FILE}")

# ✅ CHECK IF FILE EXISTS
if not os.path.exists(EXCEL_FILE):
    raise Exception(f"❌ Excel file not found at path: {EXCEL_FILE}")

print(f"✅ Excel file found: {EXCEL_FILE}")

# -----------------------------
# CONFIG
# -----------------------------
SLA_THRESHOLD = 99.9

# -----------------------------
# Helper: Clean + Format Data
# -----------------------------
def clean_df(df):
    df = df.fillna("")

    for col in df.columns:
        col_str = str(col).lower()
        if "uptime" in col_str:
            try:
                df[col] = df[col].astype(float).apply(
                    lambda x: f'<span class="{"good" if x * 100 >= SLA_THRESHOLD else "bad"}">{x * 100:.2f}%</span>'
                )
            except Exception:
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
    except Exception:
        return 0


# -----------------------------
# Load Excel & Detect Sheets
# -----------------------------
try:
    xls = pd.ExcelFile(EXCEL_FILE, engine="openpyxl")
    sheet_map = {s.strip().lower(): s for s in xls.sheet_names}
    
    weekly_sheet = sheet_map.get("weekly")
    quarterly_sheet = sheet_map.get("quarterly")

    if not weekly_sheet:
        raise Exception("❌ Weekly sheet not found in Excel")
    if not quarterly_sheet:
        raise Exception("❌ Quarterly sheet not found in Excel")
        
except Exception as e:
    raise Exception(f"❌ Error reading Excel file: {str(e)}")

# -----------------------------
# Read Data
# -----------------------------
weekly_df = pd.read_excel(EXCEL_FILE, sheet_name=weekly_sheet)
quarterly_df = pd.read_excel(EXCEL_FILE, sheet_name=quarterly_sheet)

weekly_df = clean_df(weekly_df)
quarterly_df = clean_df(quarterly_df)

weekly_table = weekly_df.to_html(index=False, classes="uptime-table", escape=False)
quarterly_table = quarterly_df.to_html(index=False, classes="uptime-table", escape=False)

# -----------------------------
# DATE RANGE (OPTIONAL – STATIC FOR NOW)
# -----------------------------
WEEKLY_RANGE = "Weekly Uptime Summary"
QUARTERLY_RANGE = "Quarterly Uptime Summary"

# -----------------------------
# MAJOR INCIDENT (OUTAGE ONLY)
# -----------------------------
major_incident = {"account": "N/A", "outage": "0 mins", "rca": ""}
major_story = "No unplanned outages observed during this period."

if "Outage Downtime" in weekly_df.columns:
    weekly_df["_outage_mins"] = weekly_df["Outage Downtime"].apply(to_minutes)
    outage_df = weekly_df[weekly_df["_outage_mins"] > 0]

    if not outage_df.empty:
        row = outage_df.loc[outage_df["_outage_mins"].idxmax()]
        account = row.get("Account Name", "N/A")
        outage_mins = row["_outage_mins"]
        rca = str(row.get("RCA of Outage", "")).strip()

        major_incident = {
            "account": account,
            "outage": f"{outage_mins} mins",
            "rca": rca
        }

        major_story = (
            f"<b>{account}</b> experienced the highest unplanned outage of "
            f"<b>{outage_mins} mins</b>.<br>"
            f"<b>Root Cause:</b> {rca if rca else 'RCA yet to be shared.'}"
        )

# -----------------------------
# Render HTML
# -----------------------------
try:
    template_path = os.path.join(BASE_DIR, "uptime_template.html")
    with open(template_path, encoding="utf-8") as f:
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

    print("✅ EXECUTIVE UPTIME REPORT GENERATED SUCCESSFULLY")
    print(f"📤 Output file: {output_file}")
    
except Exception as e:
    raise Exception(f"❌ Error generating report: {str(e)}")
