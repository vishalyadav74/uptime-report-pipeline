import pandas as pd
from jinja2 import Template
import os
import sys
import glob
import time
import re

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# -----------------------------
# Pick Excel
# -----------------------------
def find_excel():
    files = []
    for ext in ("*.xlsx", "*.xls"):
        files.extend(glob.glob(ext))
    if not files:
        raise Exception("No Excel file found")
    files.sort()
    return files[0]

EXCEL_FILE = sys.argv[1] if len(sys.argv) > 1 else find_excel()
print(f"✅ Using Excel: {EXCEL_FILE}")

# -----------------------------
# Read sheet EXACTLY as text
# -----------------------------
def read_sheet_exact(sheet_name):
    raw = pd.read_excel(
        EXCEL_FILE,
        sheet_name=sheet_name,
        header=None,
        dtype=str,
        keep_default_na=False
    )

    title = raw.iloc[0, 0]

    df = raw.iloc[1:].copy()
    df.columns = raw.iloc[1]
    df = df.iloc[1:].reset_index(drop=True)

    # remove duplicate columns only
    df = df.loc[:, ~df.columns.duplicated()]

    html = df.to_html(index=False, classes="uptime-table", escape=False)
    return title, df, html

# -----------------------------
# Downtime parser ONLY for comparison
# -----------------------------
def downtime_to_minutes(text):
    if not text:
        return 0
    text = text.lower()
    mins = 0
    hrs = re.search(r'(\d+)\s*hr', text)
    mins_m = re.search(r'(\d+)\s*min', text)
    if hrs:
        mins += int(hrs.group(1)) * 60
    if mins_m:
        mins += int(mins_m.group(1))
    return mins

# -----------------------------
# Load sheets
# -----------------------------
xls = pd.ExcelFile(EXCEL_FILE)
weekly_sheet = xls.sheet_names[0]
quarterly_sheet = xls.sheet_names[1] if len(xls.sheet_names) > 1 else None

weekly_range, weekly_df, weekly_table = read_sheet_exact(weekly_sheet)

quarterly_range = ""
quarterly_table = ""
if quarterly_sheet:
    quarterly_range, quarterly_df, quarterly_table = read_sheet_exact(quarterly_sheet)

# -----------------------------
# Major Incident (Weekly only)
# -----------------------------
major_incident = {"account": "", "outage": "", "rca": ""}
major_story = ""

if "Total Downtime(In Mins)" in weekly_df.columns:
    weekly_df["_cmp"] = weekly_df["Total Downtime(In Mins)"].apply(downtime_to_minutes)
    idx = weekly_df["_cmp"].idxmax()

    if weekly_df.loc[idx, "_cmp"] > 0:
        row = weekly_df.loc[idx]
        major_incident = {
            "account": row.get("Account Name", ""),
            "outage": row.get("Total Downtime(In Mins)", ""),
            "rca": row.get("RCA of Outage", "")
        }
        major_story = (
            f"<b>{major_incident['account']}</b> experienced the highest outage "
            f"of <b>{major_incident['outage']}</b> during the week."
        )

# -----------------------------
# Render HTML
# -----------------------------
with open(os.path.join(BASE_DIR, "uptime_template.html"), encoding="utf-8") as f:
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
out = os.path.join(OUTPUT_DIR, "uptime_report.html")
with open(out, "w", encoding="utf-8") as f:
    f.write(html)

print("✅ REPORT GENERATED:", out)
