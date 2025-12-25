📊 SaaS Accounts Weekly & Quarterly Uptime Report Automation

This project automates the generation and email delivery of Weekly and Quarterly Application Uptime Reports for SaaS accounts using Excel, Python, HTML, and Jenkins.

The solution reads uptime data exactly as present in Excel, generates a branded HTML report, highlights the Major Incident of the Week, and sends it automatically via email.

🚀 Features

✅ Reads Excel without altering values or formats

✅ Supports Weekly & Quarterly sheets

✅ Auto-detects the latest uptime Excel file

✅ Identifies Major Incident of the Week (highest downtime)

✅ Copies exact RCA from Excel

✅ Generates email-ready HTML report

✅ Fully automated via Jenkins

✅ Works with GitHub / Azure Repos

🧱 Tech Stack

Python 3.9+

openpyxl – Excel reading

Jinja2 – HTML templating

Jenkins – CI/CD automation

HTML/CSS – Email-friendly report

GitHub / Azure Repos

📁 Project Structure
uptime-report-pipeline/
│
├── generate_report.py        # Main Python script
├── uptime_template.html      # HTML email template
├── Businessnextlogo.png      # Company logo
├── requirements.txt          # Python dependencies
├── Jenkinsfile               # Jenkins pipeline
├── output/
│   └── uptime_report.html    # Generated report
└── README.md

📊 Excel File Requirements
🔹 File Naming Convention

The script auto-picks the latest file based on date in filename.

✅ Example:

uptime_latest_25th Dec_2025.xlsx

🔹 Sheet Structure
Sheet 1: Weekly Uptime

Account Name

Total Uptime

Planned Downtime

Outage Downtime

Total Downtime(In Mins)

Remarks

RCA of Outage

Sheet 2: Quarterly Uptime

Account Name

YTD Uptime

Total Uptime

Planned Downtime

Outage Downtime

Total Downtime(In Mins)

Remarks

⚠️ Important:
Values are read exactly as visible in Excel (no recalculation).

⚙️ Setup Instructions
1️⃣ Clone Repository
git clone https://github.com/your-org/uptime-report-pipeline.git
cd uptime-report-pipeline

2️⃣ Install Dependencies
pip install -r requirements.txt

3️⃣ Add Excel File

Place your uptime Excel file in the project root:

uptime_latest_25th Dec_2025.xlsx

▶️ Run Locally (Optional)
python generate_report.py


Output:

output/uptime_report.html


Open it in a browser or attach it to an email.

🔁 Jenkins Automation
Jenkins Pipeline Flow

Checkout code from repo

Create Python virtual environment

Install dependencies

Auto-detect latest Excel file

Generate HTML report

Email report

Archive artifact

Required Jenkins Plugins

Git

Pipeline

Email Extension Plugin

📧 Email Configuration

Email subject:

SAAS Accounts Weekly and Quarterly Application Uptime Report


Logo is loaded via public HTTPS URL to ensure visibility in Gmail/Outlook.

🚨 Major Incident Logic

Identifies row with highest “Total Downtime(In Mins)”

Picks:

Account Name

Outage Downtime

Exact RCA

Displays it in Major Incident of the Week section

🛡 Best Practices

✔ Keep Excel headers unchanged

✔ Ensure RCA column exists in Weekly sheet

✔ Use public HTTPS URL for images in emails

✔ Test locally before Jenkins run

**Note: Convert this to Azure DevOps pipeline if you want this for your organization.**
