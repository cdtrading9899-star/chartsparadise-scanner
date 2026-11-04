"""Email the latest-session watchlist as a plain table of stock names.

Reads scan_output/alerts_*.xlsx (written by run_scan.py) and sends it through Gmail SMTP.
Environment:
  GMAIL_USER          the Gmail address that sends the mail
  GMAIL_APP_PASSWORD  a Google App Password for that account (not the normal password)
  MAIL_TO             recipient(s), comma separated
  DRY_RUN=1           write email_preview.html instead of sending
  FORCE_EMAIL=1       send even if this session was already emailed
Without GMAIL_USER / GMAIL_APP_PASSWORD it prints a warning and exits 0, so the scan job still succeeds.
Each session is emailed once: the sent session is recorded in state/emailed_session.txt (committed by the workflow).
"""
import glob
import html
import os
import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd

xlsx = sorted(glob.glob("scan_output/alerts_*.xlsx"))
if not xlsx:
    sys.exit("No scan_output/alerts_*.xlsx found - run run_scan.py first.")
path = xlsx[-1]
session = pd.Timestamp(os.path.basename(path)[len("alerts_"):-len(".xlsx")])
watch = pd.read_excel(path, sheet_name="Latest_Watchlist")
names = [str(n).strip() for n in watch["Name"]] if len(watch) else []

day = f"{session:%d %b %Y}"
subject = f"NSE scan alerts {day}: {len(names)} stock{'s' if len(names) != 1 else ''}"

if names:
    width = max(len("Stock"), *(len(n) for n in names))
    text_rows = [f"{'No.':<4} {'Stock':<{width}}", f"{'-' * 4} {'-' * width}"]
    text_rows += [f"{i:<4} {n:<{width}}" for i, n in enumerate(names, 1)]
    text = f"Alerts for the {day} session\n\n" + "\n".join(text_rows) + "\n"
    rows = "".join(f"<tr><td>{i}</td><td>{html.escape(n)}</td></tr>" for i, n in enumerate(names, 1))
    body = (f"<p>Alerts for the {day} session</p>"
            '<table border="1" cellpadding="4" cellspacing="0">'
            f"<tr><th>No.</th><th>Stock</th></tr>{rows}</table>")
else:
    text = f"No alerts for the {day} session.\n"
    body = f"<p>No alerts for the {day} session.</p>"

msg = MIMEMultipart("alternative")
msg["Subject"] = subject
msg.attach(MIMEText(text, "plain", "utf-8"))
msg.attach(MIMEText(f"<html><body>{body}</body></html>", "html", "utf-8"))

if os.environ.get("DRY_RUN"):
    open("email_preview.html", "w", encoding="utf-8").write(f"<h3>{html.escape(subject)}</h3>{body}")
    print(subject)
    print(text)
    sys.exit(0)

MARKER = os.path.join("state", "emailed_session.txt")
done = open(MARKER, encoding="utf-8").read().strip() if os.path.exists(MARKER) else ""
if done == f"{session:%Y-%m-%d}" and not os.environ.get("FORCE_EMAIL"):
    print(f"Session {day} was already emailed; not sending again.")
    sys.exit(0)

user, password = os.environ.get("GMAIL_USER"), os.environ.get("GMAIL_APP_PASSWORD")
to = [a.strip() for a in os.environ.get("MAIL_TO", "").split(",") if a.strip()]
if not (user and password and to):
    print("::warning::Email not sent: set the GMAIL_USER and GMAIL_APP_PASSWORD secrets (and MAIL_TO).")
    sys.exit(0)

msg["From"] = user
msg["To"] = ", ".join(to)
with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=60) as s:
    s.login(user, password)
    s.sendmail(user, to, msg.as_string())
os.makedirs("state", exist_ok=True)
open(MARKER, "w", encoding="utf-8").write(f"{session:%Y-%m-%d}\n")
print(f"sent '{subject}' to {', '.join(to)}")
