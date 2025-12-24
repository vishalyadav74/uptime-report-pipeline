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
    """
    Find the best Excel file in current directory
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
        print(f"   • {file} ({size} KB)")
    
    # Select first file (simplified)
    best_file = all_excel_files[0]
    print(f"\n✅ Selected file: {best_file}")
    return best_file

# -----------------------------
# MAIN EXECUTION
# -----------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# ✅ GET EXCEL FILE
EXCEL_FILE = None

if len(sys.argv) > 1:
    EXCEL_FILE = sys.argv[1]
    print(f"📄 Using command line file: {EXCEL_FILE}")
elif "UPTIME_EXCEL" in os.environ:
    EXCEL_FILE = os.environ["UPTIME_EXCEL"]
    print(f"📄 Using environment variable file: {EXCEL_FILE}")
else:
    EXCEL_FILE = find_best_excel_file()
    if not EXCEL_FILE:
        raise Exception("❌ No Excel file found!")

# ✅ VALIDATE FILE
if not os.path.exists(EXCEL_FILE):
    raise Exception(f"❌ Excel file not found: {EXCEL_FILE}")

print(f"✅ Processing: {os.path.basename(EXCEL_FILE)}")

# -----------------------------
# CONFIG
# -----------------------------
SLA_THRESHOLD = 99.9

# -----------------------------
# Helper: Convert downtime string to minutes
# -----------------------------
def parse_downtime_to_minutes(downtime_str):
    """
    Convert downtime string like "1 hrs 31 mins" or "34 mins 31 secs" to minutes
    """
    if pd.isna(downtime_str) or downtime_str == "":
        return 0
    
    downtime_str = str(downtime_str).lower()
    total_minutes = 0
    
    try:
        # Extract hours
        if "hr" in downtime_str:
            hrs_part = downtime_str.split("hr")[0]
            hours = 0
            for word in hrs_part.split():
                try:
                    hours = int(word)
                    break
                except ValueError:
                    continue
            total_minutes += hours * 60
        
        # Extract minutes
        if "min" in downtime_str:
            mins_part = downtime_str.split("min")[0]
            # Get the last number before "min"
            parts = mins_part.split()
            if parts:
                try:
                    mins = int(parts[-1])
                    total_minutes += mins
                except ValueError:
                    pass
        
        # Extract seconds and convert to minutes
        if "sec" in downtime_str:
            secs_part = downtime_str.split("sec")[0]
            parts = secs_part.split()
            if parts:
                try:
                    secs = int(parts[-1])
                    total_minutes += secs / 60
                except ValueError:
                    pass
        
        # If it's just a number (already in minutes)
        if total_minutes == 0:
            try:
                total_minutes = float(downtime_str.split()[0])
            except:
                pass
                
    except Exception as e:
        print(f"⚠️ Could not parse downtime: '{downtime_str}' - Error: {e}")
    
    return round(total_minutes, 2)

# -----------------------------
# Helper: Format percentages with colors
# -----------------------------
def format_uptime_percentage(value):
    """Format uptime percentage with color coding"""
    try:
        if isinstance(value, str):
            # Remove % sign if present
            value = value.replace('%', '').strip()
        uptime = float(value)
        
        if uptime >= SLA_THRESHOLD:
            return f'<span class="good">{uptime:.2f}%</span>'
        else:
            return f'<span class="bad">{uptime:.2f}%</span>'
    except:
        return f'<span>{value}</span>'

# -----------------------------
# Load Excel & Detect Sheets
# -----------------------------
try:
    xls = pd.ExcelFile(EXCEL_FILE, engine="openpyxl")
    print(f"📑 Sheets in Excel file: {xls.sheet_names}")
    
    # Find weekly and quarterly sheets
    weekly_sheet = None
    quarterly_sheet = None
    
    for sheet in xls.sheet_names:
        sheet_lower = sheet.lower()
        if 'weekly' in sheet_lower:
            weekly_sheet = sheet
        elif 'quarter' in sheet_lower:
            quarterly_sheet = sheet
    
    # If not found by name, use first two sheets
    if not weekly_sheet and len(xls.sheet_names) >= 1:
        weekly_sheet = xls.sheet_names[0]
    if not quarterly_sheet and len(xls.sheet_names) >= 2:
        quarterly_sheet = xls.sheet_names[1]
    
    print(f"📅 Weekly sheet: {weekly_sheet}")
    print(f"📅 Quarterly sheet: {quarterly_sheet}")
    
    if not weekly_sheet:
        raise Exception("❌ Weekly sheet not found in Excel")
        
except Exception as e:
    raise Exception(f"❌ Error reading Excel file: {str(e)}")

# -----------------------------
# Read and Process Weekly Data
# -----------------------------
try:
    # Read weekly data
    weekly_df = pd.read_excel(EXCEL_FILE, sheet_name=weekly_sheet)
    print(f"📊 Weekly data shape: {weekly_df.shape}")
    
    # Clean column names
    weekly_df.columns = [str(col).strip() for col in weekly_df.columns]
    print(f"📋 Weekly columns: {list(weekly_df.columns)}")
    
    # Display first few rows for debugging
    print("\n📝 Sample weekly data:")
    print(weekly_df.head(3).to_string())
    
    # Check for required columns
    required_cols = ['Account Name', 'Total Uptime']
    missing_cols = [col for col in required_cols if col not in weekly_df.columns]
    
    if missing_cols:
        print(f"⚠️ Missing columns in weekly sheet: {missing_cols}")
        print("🔍 Available columns:", list(weekly_df.columns))
    
    # Format uptime percentages
    if 'Total Uptime' in weekly_df.columns:
        weekly_df['Total Uptime'] = weekly_df['Total Uptime'].apply(format_uptime_percentage)
    
    # Calculate downtime in minutes
    if 'Outage Downtime' in weekly_df.columns:
        weekly_df['Outage Minutes'] = weekly_df['Outage Downtime'].apply(parse_downtime_to_minutes)
        weekly_df['Outage Minutes'] = weekly_df['Outage Minutes'].fillna(0)
    else:
        weekly_df['Outage Minutes'] = 0
    
    # Convert to HTML table
    weekly_table = weekly_df.to_html(index=False, classes="uptime-table", escape=False)
    
except Exception as e:
    raise Exception(f"❌ Error processing weekly data: {str(e)}")

# -----------------------------
# Read and Process Quarterly Data
# -----------------------------
quarterly_table = "<p>No quarterly data available</p>"
if quarterly_sheet:
    try:
        quarterly_df = pd.read_excel(EXCEL_FILE, sheet_name=quarterly_sheet)
        print(f"📊 Quarterly data shape: {quarterly_df.shape}")
        
        # Clean column names
        quarterly_df.columns = [str(col).strip() for col in quarterly_df.columns]
        print(f"📋 Quarterly columns: {list(quarterly_df.columns)}")
        
        # Format uptime percentages if column exists
        if 'Total Uptime' in quarterly_df.columns:
            quarterly_df['Total Uptime'] = quarterly_df['Total Uptime'].apply(format_uptime_percentage)
        
        # Convert to HTML table
        quarterly_table = quarterly_df.to_html(index=False, classes="uptime-table", escape=False)
        
    except Exception as e:
        print(f"⚠️ Could not process quarterly sheet: {e}")
        quarterly_table = f"<p>Error loading quarterly data: {str(e)}</p>"

# -----------------------------
# Major Incident Detection
# -----------------------------
major_incident = {"account": "N/A", "outage": "0 mins", "rca": ""}
major_story = "No unplanned outages observed during this period."

if 'Outage Minutes' in weekly_df.columns and 'Account Name' in weekly_df.columns:
    # Find account with maximum outage
    max_outage_idx = weekly_df['Outage Minutes'].idxmax()
    max_outage = weekly_df.loc[max_outage_idx]
    
    if max_outage['Outage Minutes'] > 0:
        account = max_outage.get('Account Name', 'N/A')
        outage_mins = int(max_outage['Outage Minutes'])
        
        # Get RCA if available
        rca = ""
        if 'RCA of Outage' in weekly_df.columns:
            rca = str(max_outage.get('RCA of Outage', '')).strip()
        elif 'Remarks' in weekly_df.columns:
            rca = str(max_outage.get('Remarks', '')).strip()
        
        major_incident = {
            "account": account,
            "outage": f"{outage_mins} mins",
            "rca": rca[:200] + "..." if len(rca) > 200 else rca  # Limit RCA length
        }
        
        major_story = (
            f"<b>{account}</b> experienced the highest unplanned outage of "
            f"<b>{outage_mins} minutes</b>."
        )
        if rca:
            major_story += f"<br><b>Root Cause:</b> {rca}"

# -----------------------------
# Date Ranges (Extract from sheet name or data)
# -----------------------------
# Try to extract date from sheet name
weekly_range = "Weekly Uptime Summary"
if weekly_sheet:
    weekly_range = weekly_sheet

quarterly_range = "Quarterly Uptime Summary"
if quarterly_sheet:
    quarterly_range = quarterly_sheet

# -----------------------------
# Render HTML Report
# -----------------------------
try:
    template_path = os.path.join(BASE_DIR, "uptime_template.html")
    with open(template_path, encoding="utf-8") as f:
        template = Template(f.read())

    html = template.render(
        weekly_table=weekly_table,
        quarterly_table=quarterly_table,
        weekly_range=weekly_range,
        quarterly_range=quarterly_range,
        major_incident=major_incident,
        major_story=major_story,
        generated_date=time.strftime("%d-%b-%Y %H:%M"),
        source_file=os.path.basename(EXCEL_FILE)
    )
    
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_file = os.path.join(OUTPUT_DIR, "uptime_report.html")
    
    with open(output_file, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\n✅ REPORT GENERATED SUCCESSFULLY!")
    print(f"📊 Source: {os.path.basename(EXCEL_FILE)}")
    print(f"📁 Output: {output_file}")
    print(f"📈 Weekly records: {len(weekly_df)}")
    if quarterly_sheet:
        print(f"📈 Quarterly records: {len(quarterly_df) if 'quarterly_df' in locals() else 0}")
    print(f"⚠️  Max outage: {major_incident['account']} - {major_incident['outage']}")
    
except Exception as e:
    raise Exception(f"❌ Error generating report: {str(e)}")
