"""Send the latest-session watchlist as a full table by email and (optionally) Telegram.

Reads scan_output/alerts_*.xlsx (written by run_scan.py).
Environment:
  GMAIL_USER / GMAIL_APP_PASSWORD   Gmail sender + Google App Password (email is skipped if missing)
  MAIL_TO                           recipient(s), comma separated
  TELEGRAM_TOKEN / TELEGRAM_CHAT_ID Telegram bot token + chat id (Telegram is skipped if missing)
  DRY_RUN=1                         write email_preview.html and print the Telegram text instead of sending
  FORCE_EMAIL=1                     send even if this session was already sent
Each session is sent once: the sent session is recorded in state/emailed_session.txt (committed by the workflow).
If one channel fails the other still goes out, and the script exits with an error so the run shows red.
"""
import glob
import html
import os
import smtplib
import sys
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
import requests

xlsx = sorted(glob.glob("scan_output/alerts_*.xlsx"))
if not xlsx:
    sys.exit("No scan_output/alerts_*.xlsx found - run run_scan.py first.")
path = xlsx[-1]
session = pd.Timestamp(os.path.basename(path)[len("alerts_"):-len(".xlsx")])
watch = pd.read_excel(path, sheet_name="Latest_Watchlist")
if len(watch) and "Score (0-6)" in watch:
    watch = watch.sort_values("Score (0-6)", ascending=False, kind="stable")

day = f"{session:%d %b %Y}"
n = len(watch)
subject = f"NSE scan alerts {day}: {n} stock{'s' if n != 1 else ''}"


def num(x, digits=2):
    return "" if pd.isna(x) else f"{x:,.{digits}f}"


def pct(x):
    return "" if pd.isna(x) else f"{x:.1%}"


rows = []
for i, (_, r) in enumerate(watch.iterrows(), 1):
    rows.append({
        "No.": i,
        "Symbol": str(r["Symbol"]).strip(),
        "Name": str(r["Name"]).strip(),
        "Close": num(r["Close"]),
        "Trigger": num(r["Trigger (alert-day high)"]),
        "Stop": num(r["Stop (5-bar low)"]),
        "Gap": pct(r["Trigger Gap (%)"]),
        "Risk": pct(r["Stop Risk (%)"]),
        "Score": int(r["Score (0-6)"]),
        "Status": str(r.get("Status", "")).strip(),
        "Flags": str(r.get("Flags", "")).strip(),
    })

# ---- email body ----
if rows:
    cols = ["No.", "Symbol", "Name", "Close", "Trigger", "Stop", "Gap", "Risk", "Score", "Status"]
    head = "".join(f"<th>{c}</th>" for c in cols)
    body_rows = "".join(
        "<tr>" + "".join(
            f'<td style="text-align:{"left" if c in ("Symbol", "Name", "Status") else "right"}">'
            f"{html.escape(str(r[c]))}</td>" for c in cols) + "</tr>"
        for r in rows)
    body = (f"<p>Alerts for the {day} session ({n}). Trigger = alert-day high; Stop = 5-bar low; "
            f"Gap = Trigger vs Close; Risk = Close to Stop.</p>"
            '<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;font-size:13px">'
            f"<tr>{head}</tr>{body_rows}</table>")
    widths = {c: max(len(c), *(len(str(r[c])) for r in rows)) for c in cols[:-1]}
    line = lambda r: " ".join(
        (f"{str(r[c]):<{widths[c]}}" if c in ("Symbol", "Name") else f"{str(r[c]):>{widths[c]}}") for c in cols[:-1])
    text = (f"Alerts for the {day} session ({n})\n\n"
            + line({c: c for c in cols[:-1]}) + "\n" + "\n".join(line(r) for r in rows) + "\n")
    # ---- compact Telegram text (monospace lines, phone friendly) ----
    tg_lines = [f"NSE alerts {day}: {n}"]
    for r in rows:
        tg_lines.append(f"{r['No.']}. {r['Symbol']}  s{r['Score']}  C {r['Close']}  T {r['Trigger']}  "
                        f"SL {r['Stop']}  ({r['Gap']}/{r['Risk']})")
else:
    body = f"<p>No alerts for the {day} session.</p>"
    text = f"No alerts for the {day} session.\n"
    tg_lines = [f"NSE alerts {day}: none"]


def tg_chunks(lines, limit=3500):
    chunk = ""
    for ln in lines:
        if len(chunk) + len(ln) + 1 > limit and chunk:
            yield chunk
            chunk = ""
        chunk += ln + "\n"
    if chunk:
        yield chunk


msg = MIMEMultipart("alternative")
msg["Subject"] = subject
msg.attach(MIMEText(text, "plain", "utf-8"))
msg.attach(MIMEText(f"<html><body>{body}</body></html>", "html", "utf-8"))

if os.environ.get("DRY_RUN"):
    open("email_preview.html", "w", encoding="utf-8").write(f"<h3>{html.escape(subject)}</h3>{body}")
    print(subject)
    print(text)
    print("--- Telegram ---")
    print("\n".join(tg_lines))
    sys.exit(0)

MARKER = os.path.join("state", "emailed_session.txt")
done = open(MARKER, encoding="utf-8").read().strip() if os.path.exists(MARKER) else ""
if done == f"{session:%Y-%m-%d}" and not os.environ.get("FORCE_EMAIL"):
    print(f"Session {day} was already sent; not sending again.")
    sys.exit(0)

sent, failed = [], []

# ---- email ----
user, password = os.environ.get("GMAIL_USER"), os.environ.get("GMAIL_APP_PASSWORD")
to = [a.strip() for a in os.environ.get("MAIL_TO", "").split(",") if a.strip()]
if user and password and to:
    try:
        msg["From"] = user
        msg["To"] = ", ".join(to)
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=60) as s:
            s.login(user, password)
            s.sendmail(user, to, msg.as_string())
        sent.append("email")
        print(f"sent '{subject}' to {', '.join(to)}")
    except Exception as e:  # keep going so Telegram can still go out
        failed.append("email")
        print(f"::error::Email failed: {e}")
else:
    print("::warning::Email skipped: set GMAIL_USER, GMAIL_APP_PASSWORD secrets and MAIL_TO.")

# ---- Telegram (optional) ----
tg_token, tg_chat = os.environ.get("TELEGRAM_TOKEN"), os.environ.get("TELEGRAM_CHAT_ID")
if tg_token and tg_chat:
    try:
        for chunk in tg_chunks(tg_lines):
            resp = requests.post(f"https://api.telegram.org/bot{tg_token}/sendMessage",
                                 data={"chat_id": tg_chat, "text": chunk}, timeout=30)
            resp.raise_for_status()
        sent.append("telegram")
        print("sent Telegram alert")
    except Exception as e:
        failed.append("telegram")
        print(f"::error::Telegram failed: {type(e).__name__}")  # never print the URL: it contains the token
else:
    print("Telegram skipped (TELEGRAM_TOKEN / TELEGRAM_CHAT_ID not set).")

if sent:
    os.makedirs("state", exist_ok=True)
    open(MARKER, "w", encoding="utf-8").write(f"{session:%Y-%m-%d}\n")
if failed:
    sys.exit(1)
