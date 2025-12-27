from openpyxl import load_workbook
from jinja2 import Template
import os, sys, glob, time, re, argparse, smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# -------------------------------
# ARGUMENTS
# -------------------------------
parser = argparse.ArgumentParser()
parser.add_argument("--to", required=True, help="To recipients (comma separated)")
parser.add_argument("--cc", default="", help="CC recipients (comma separated)")
args = parser.parse_args()

MAIL_TO = args.to
MAIL_CC = args.cc

# -------------------------------
# SMTP CONFIG (ITSM)
# -------------------------------
smtp_server = 'smtp.office365.com'
smtp_port = 587
smtp_user = 'incident@businessnext.com'
smtp_password = 'btxnzsrnjgjfjpqf'
from_email = 'incident@businessnext.com'

# -------------------------------
# PATHS
# -------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")

# -------------------------------
# PICK LATEST EXCEL
# -------------------------------
def extract_date_from_filename(name):
    m = re.search(r'(\d{1,2})(st|nd|rd|th)\s+([A-Za-z]+)_([0-9]{4})', name)
    if not m:
        return None
    try:
        return datetime.strptime(f"{m.group(1)} {m.group(3)} {m.group(4)}", "%d %b %Y")
    except:
        return None

def find_excel():
    files = glob.glob("*.xlsx") + glob.glob("*.xls")
    dated = [(extract_date_from_filename(f), f) for f in files if extract_date_from_filename(f)]
    dated.sort(reverse=True)
    return dated[0][1]

EXCEL_FILE = find_excel()
print(f"✅ Using Excel: {EXCEL_FILE}")

# -------------------------------
# HELPERS
# -------------------------------
def get_cell_display(cell):
    if cell.value is None:
        return ""
    if isinstance(cell.value, (int, float)) and "%" in str(cell.number_format):
        return f"{cell.value * 100:.2f}%"
    return str(cell.value)

def wrap_uptime(val):
    try:
        num = float(val.replace("%", ""))
        cls = "uptime-bad" if num < 99.95 else "uptime-good"
        return f'<span class="{cls}">{val}</span>'
    except:
        return val

def downtime_to_minutes(txt):
    if not txt:
        return 0
    mins = 0
    if m := re.search(r'(\d+)\s*hr', txt.lower()):
        mins += int(m.group(1)) * 60
    if m := re.search(r'(\d+)\s*min', txt.lower()):
        mins += int(m.group(1))
    return mins

# -------------------------------
# READ SHEET
# -------------------------------
def read_sheet(sheet_name):
    wb = load_workbook(EXCEL_FILE, data_only=True)
    ws = wb[sheet_name]

    title = ws["A1"].value or ""
    headers = [str(c.value).strip() if c.value else "" for c in ws[2]]

    rows = []
    for r in ws.iter_rows(min_row=3, max_col=len(headers)):
        row = [get_cell_display(c) for c in r]
        if any(row):
            rows.append(row)

    html = "<table class='uptime-table'><thead><tr>"
    for h in headers:
        html += f"<th>{h}</th>"
    html += "</tr></thead><tbody>"

    for i, r in enumerate(rows):
        bg = "#ffffff" if i % 2 == 0 else "#f5f5f5"
        html += f"<tr style='background:{bg};'>"
        for v in r:
            html += f"<td>{wrap_uptime(v)}</td>"
        html += "</tr>"

    html += "</tbody></table>"
    return title, html

# -------------------------------
# LOAD DATA
# -------------------------------
wb = load_workbook(EXCEL_FILE, data_only=True)
sheets = wb.sheetnames

weekly_title, weekly_table = read_sheet(sheets[0])
quarterly_title, quarterly_table = ("", "")
if len(sheets) > 1:
    quarterly_title, quarterly_table = read_sheet(sheets[1])

# -------------------------------
# RENDER HTML
# -------------------------------
with open(os.path.join(BASE_DIR, "uptime_template.html"), encoding="utf-8") as f:
    template = Template(f.read())

html_body = template.render(
    weekly_title=weekly_title,
    quarterly_title=quarterly_title,
    weekly_table=weekly_table,
    quarterly_table=quarterly_table
)

os.makedirs(OUTPUT_DIR, exist_ok=True)
out = os.path.join(OUTPUT_DIR, "uptime_report.html")
with open(out, "w", encoding="utf-8") as f:
    f.write(html_body)

print("✅ REPORT GENERATED:", out)

# -------------------------------
# SEND EMAIL (TO + CC PROPER)
# -------------------------------
msg = MIMEMultipart()
msg["From"] = from_email
msg["To"] = MAIL_TO
msg["Cc"] = MAIL_CC
msg["Subject"] = "SAAS Accounts Weekly & Quarterly Application Uptime Report"

msg.attach(MIMEText(html_body, "html"))

to_list = [x.strip() for x in MAIL_TO.split(",") if x.strip()]
cc_list = [x.strip() for x in MAIL_CC.split(",") if x.strip()]
all_recipients = to_list + cc_list

server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(smtp_user, smtp_password)
server.sendmail(from_email, all_recipients, msg.as_string())
server.quit()

print(f"✅ Email sent | TO={to_list} | CC={cc_list}")
