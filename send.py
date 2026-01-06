import smtplib, argparse, os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage


def send_email(subject, body_html, to_list, cc_list):
    smtp_server = 'smtp.office365.com'
    smtp_port = 587
    smtp_user = 'incident@businessnext.com'
    smtp_password = 'btxnzsrnjgjfjpqf'  

    msg = MIMEMultipart('related')
    msg['From'] = smtp_user
    msg['To'] = ', '.join(to_list)
    msg['Cc'] = ', '.join(cc_list)
    msg['Subject'] = subject

    # HTML body
    msg.attach(MIMEText(body_html, 'html'))

    # --- Logo ---
    with open('logo-fixed.png', 'rb') as f:
        img = MIMEImage(f.read())
        img.add_header('Content-ID', '<businessnext_logo>')
        img.add_header('Content-Disposition', 'inline', filename='logo-fixed.png')
        msg.attach(img)

    # --- Downtime chart (if exists) ---
    chart_path = os.path.join('output', 'downtime_chart.png')
    if os.path.exists(chart_path):
        with open(chart_path, 'rb') as f:
            chart = MIMEImage(f.read())
            chart.add_header('Content-ID', '<downtime_chart>')
            chart.add_header('Content-Disposition', 'inline', filename='downtime_chart.png')
            msg.attach(chart)

    recipients = to_list + cc_list

    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(smtp_user, smtp_password)
    server.sendmail(smtp_user, recipients, msg.as_string())
    server.quit()

    print("✅ Email sent successfully")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('--subject', required=True)
    parser.add_argument('--to', required=True)
    parser.add_argument('--cc', default='')
    parser.add_argument('--body', required=True)
    args = parser.parse_args()

    with open(args.body, encoding='utf-8') as f:
        body = f.read()

    send_email(
        args.subject,
        body,
        [x.strip() for x in args.to.split(',') if x.strip()],
        [x.strip() for x in args.cc.split(',') if x.strip()]
    )
