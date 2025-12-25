import pandas as pd
from jinja2 import Template
import os
import sys
import glob
import time

# -----------------------------
# PATHS
# -----------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# -----------------------------
# EXCEL PICKER
# -----------------------------
def find_best_excel_file():
    files = []
    for ext in ("*.xlsx", "*.xls"):
        files.extend(glob.glob(ext))
    if not files:
        raise Exception("No Excel file found")
    files.sort()
    return files[0]

# -----------------------------
# INPUT FILE
# -----------------------------
if len(sys.argv) > 1:
    EXCEL_FILE = sys.argv[1]
else:
    EXCEL_FILE = find_best_excel_file()

print(f"✅ Using Excel: {EXCEL_FILE}")

# -----------------------------
# READ SHEET EXACTLY AS IS
# -----------------------------
def read_sheet_exact(sheet_name):
    raw = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=None)

    title = str(raw.iloc[0, 0]).strip()

    df = raw.iloc[1:].copy()
    df.columns = raw.iloc[1]
    df = df.iloc[1:].reset_index(drop=True)

    # Clean headers only (values untouched)
    df.columns = (
        df.columns
        .astype(str)
        .str.replace("\n", " ")
        .str.strip()
    )

    df = df.loc[:, ~df.columns.duplicated()]

    html_table = df.to_html(
        index=False,
        classes="uptime-table",
        escape=False
    )

    return title, df, html_table

# -----------------------------
# LOAD SHEETS
# -----------------------------
xls = pd.ExcelFile(EXCEL_FILE, engine="openpyxl")

weekly_sheet = xls.sheet_names[0]
quarterly_sheet = xls.sheet_names[1] if len(xls.sheet_names) > 1 else None

weekly_range, weekly_df, weekly_table = read_sheet_exact(weekly_sheet)

quarterly_range = ""
quarterly_table = ""
if quarterly_sheet:
    quarterly_range, quarterly_df, quarterly_table = read_sheet_exact(quarterly_sheet)

# -----------------------------
# MAJOR INCIDENT (WEEKLY ONLY)
# -----------------------------
major_incident = {
    "account": "",
    "outage": "",
    "rca": ""
}
major_story = ""

if (
    "Total Downtime(In Mins)" in weekly_df.columns and
    "Account Name" in weekly_df.columns
):
    # Convert to numeric only for comparison (display stays original)
    downtime_numeric = pd.to_numeric(
        weekly_df["Total Downtime(In Mins)"],
        errors="coerce"
    )

    if downtime_numeric.notna().any():
        idx = downtime_numeric.idxmax()
        row = weekly_df.loc[idx]

        major_incident = {
            "account": row.get("Account Name", ""),
            "outage": str(row.get("Total Downtime(In Mins)", "")),
            "rca": str(row.get("RCA of Outage", ""))
        }

        major_story = (
            f"<b>{major_incident['account']}</b> experienced the highest outage "
            f"of <b>{major_incident['outage']}</b> during the week."
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
