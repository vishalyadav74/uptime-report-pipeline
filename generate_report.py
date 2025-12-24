import pandas as pd
from jinja2 import Template
import os
import sys
import glob
import time  # ✅ ADD THIS LINE

# -----------------------------
# SMART EXCEL FILE DETECTOR
# -----------------------------
def find_best_excel_file():
    """
    Find the best Excel file in current directory
    Priority: 1. Uptime related, 2. Most recent, 3. Any Excel
    """
    print("🔍 Scanning for Excel files...")
    
    # Get all Excel files
    all_excel_files = []
    for ext in ['.xlsx', '.xls']:
        all_excel_files.extend(glob.glob(f"*{ext}"))
    
    if not all_excel_files:
        print("❌ No Excel files found in directory")
        print("📁 Available files:")
        for file in sorted(os.listdir('.')):
            print(f"   • {file}")
        return None
    
    print(f"📊 Found {len(all_excel_files)} Excel file(s):")
    for file in sorted(all_excel_files):
        size = os.path.getsize(file) // 1024  # Size in KB
        mtime = os.path.getmtime(file)
        print(f"   • {file} ({size} KB)")
    
    # Score files based on relevance
    file_scores = []
    for file in all_excel_files:
        score = 0
        filename = file.lower()
        
        # Priority keywords (higher score = more relevant)
        if 'uptime' in filename:
            score += 10
        if 'latest' in filename:
            score += 5
        if 'report' in filename:
            score += 3
        if 'data' in filename:
            score += 2
        if 'weekly' in filename or 'monthly' in filename or 'quarterly' in filename:
            score += 3
        
        # File size score (reasonable size: 10KB - 10MB)
        size_kb = os.path.getsize(file) // 1024
        if 10 <= size_kb <= 10240:  # 10KB to 10MB
            score += 2
        
        # Modification time (newer files get higher score)
        days_old = (time.time() - os.path.getmtime(file)) / 86400
        if days_old < 7:  # Less than 7 days old
            score += 3
        elif days_old < 30:  # Less than 30 days old
            score += 1
        
        file_scores.append((file, score))
    
    # Sort by score (highest first)
    file_scores.sort(key=lambda x: x[1], reverse=True)
    
    print("\n🎯 File relevance scores:")
    for file, score in file_scores:
        print(f"   • {file}: {score} points")
    
    # Select best file
    best_file = file_scores[0][0]
    print(f"\n✅ Selected file: {best_file}")
    return best_file

# -----------------------------
# MAIN EXECUTION
# -----------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# ✅ GET EXCEL FILE (Priority order)
EXCEL_FILE = None

# 1. Command line argument (highest priority)
if len(sys.argv) > 1:
    EXCEL_FILE = sys.argv[1]
    print(f"📄 Using command line file: {EXCEL_FILE}")

# 2. Environment variable
if not EXCEL_FILE and "UPTIME_EXCEL" in os.environ:
    EXCEL_FILE = os.environ["UPTIME_EXCEL"]
    print(f"📄 Using environment variable file: {EXCEL_FILE}")

# 3. Auto-detect from repository
if not EXCEL_FILE:
    EXCEL_FILE = find_best_excel_file()
    if not EXCEL_FILE:
        raise Exception("""
❌ No suitable Excel file found!
Please either:
1. Upload a file when running Jenkins pipeline
2. Place an Excel file (.xlsx or .xls) in the repository
3. Set UPTIME_EXCEL environment variable
""")

# ✅ VALIDATE FILE
if not os.path.exists(EXCEL_FILE):
    raise Exception(f"❌ Excel file not found: {EXCEL_FILE}")

file_size = os.path.getsize(EXCEL_FILE) // 1024
print(f"✅ File selected: {EXCEL_FILE} ({file_size} KB)")

# -----------------------------
# REST OF YOUR CODE (UNCHANGED)
# -----------------------------
SLA_THRESHOLD = 99.9

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
# PROCESS EXCEL
# -----------------------------
try:
    xls = pd.ExcelFile(EXCEL_FILE, engine="openpyxl")
    sheet_map = {s.strip().lower(): s for s in xls.sheet_names}
    
    weekly_sheet = sheet_map.get("weekly")
    quarterly_sheet = sheet_map.get("quarterly")

    if not weekly_sheet:
        raise Exception("❌ 'Weekly' sheet not found in Excel")
    if not quarterly_sheet:
        raise Exception("❌ 'Quarterly' sheet not found in Excel")
        
except Exception as e:
    raise Exception(f"❌ Error reading Excel file: {str(e)}")

# Read data
weekly_df = pd.read_excel(EXCEL_FILE, sheet_name=weekly_sheet)
quarterly_df = pd.read_excel(EXCEL_FILE, sheet_name=quarterly_sheet)

weekly_df = clean_df(weekly_df)
quarterly_df = clean_df(quarterly_df)

weekly_table = weekly_df.to_html(index=False, classes="uptime-table", escape=False)
quarterly_table = quarterly_df.to_html(index=False, classes="uptime-table", escape=False)

# Major incident detection
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
# GENERATE HTML REPORT
# -----------------------------
try:
    template_path = os.path.join(BASE_DIR, "uptime_template.html")
    with open(template_path, encoding="utf-8") as f:
        template = Template(f.read())

    html = template.render(
        weekly_table=weekly_table,
        quarterly_table=quarterly_table,
        weekly_range="Weekly Uptime Summary",
        quarterly_range="Quarterly Uptime Summary",
        major_incident=major_incident,
        major_story=major_story
    )
    
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_file = os.path.join(OUTPUT_DIR, "uptime_report.html")
    
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\n✅ REPORT GENERATED SUCCESSFULLY!")
    print(f"📊 Source file: {os.path.basename(EXCEL_FILE)}")
    print(f"📁 Output: {output_file}")
    print(f"📈 Weekly rows: {len(weekly_df)}")
    print(f"📈 Quarterly rows: {len(quarterly_df)}")
    
except Exception as e:
    raise Exception(f"❌ Error generating report: {str(e)}")
