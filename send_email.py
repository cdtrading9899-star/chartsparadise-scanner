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
is_repeat = (watch["Alert Type"].astype(str).str.startswith("REPEAT") if "Alert Type" in watch
             else pd.Series(False, index=watch.index))
n_rep = int(is_repeat.sum())
n_new = n - n_rep
subject = (f"NSE scan alerts {day}: {n_new} new" + (f", {n_rep} repeat{'s' if n_rep != 1 else ''}" if n_rep else "")
           if "Alert Type" in watch else f"NSE scan alerts {day}: {n} stock{'s' if n != 1 else ''}")


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
        "Repeat": bool(is_repeat.loc[r.name]),
        "Days": int(r["Days On List"]) if "Days On List" in r and pd.notna(r["Days On List"]) else 1,
    })

# ---- email body ----
cols = ["No.", "Symbol", "Name", "Close", "Trigger", "Stop", "Gap", "Risk", "Score"]
LEFT = ("Symbol", "Name")


def html_table(rs, extra=None):
    head = "".join(f"<th>{c}</th>" for c in cols + ([extra[0]] if extra else []))
    out = []
    for r in rs:
        cells = "".join(
            f'<td style="text-align:{"left" if c in LEFT else "right"}">{html.escape(str(r[c]))}</td>' for c in cols)
        if extra:
            cells += f'<td style="text-align:right">{html.escape(str(extra[1](r)))}</td>'
        out.append(f"<tr>{cells}</tr>")
    return ('<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;font-size:13px">'
            f"<tr>{head}</tr>{''.join(out)}</table>")


def text_table(rs):
    ws = {c: max(len(c), *(len(str(r[c])) for r in rs)) for c in cols}
    fmt = lambda r: " ".join(
        (f"{str(r[c]):<{ws[c]}}" if c in LEFT else f"{str(r[c]):>{ws[c]}}") for c in cols)
    return fmt({c: c for c in cols}) + "\n" + "\n".join(fmt(r) for r in rs)


new_rows = [r for r in rows if not r["Repeat"]]
rep_rows = [r for r in rows if r["Repeat"]]
for grp in (new_rows, rep_rows):          # renumber each section from 1
    for i, r in enumerate(grp, 1):
        r["No."] = i

if rows:
    legend = ("Trigger = alert-day high; Stop = 5-bar low; Gap = Trigger vs Close; Risk = Close to Stop.")
    body = f"<p>Alerts for the {day} session. {legend}</p>"
    text = f"Alerts for the {day} session\n{legend}\n"
    body += f"<h3>New today ({len(new_rows)})</h3>" + (html_table(new_rows) if new_rows else "<p>None.</p>")
    text += f"\nNEW TODAY ({len(new_rows)})\n" + (text_table(new_rows) if new_rows else "None.") + "\n"
    if rep_rows:
        body += (f"<h3>Still in the setup, listed before ({len(rep_rows)})</h3>"
                 + html_table(rep_rows, ("Days on list", lambda r: r["Days"])))
        text += f"\nSTILL IN THE SETUP, LISTED BEFORE ({len(rep_rows)})\n" + text_table(rep_rows) + "\n"
    # ---- compact Telegram text (monospace lines, phone friendly) ----
    tg_lines = [f"NSE alerts {day}: {len(new_rows)} new" + (f", {len(rep_rows)} repeat" if rep_rows else "")]
    tg_line = lambda r: (f"{r['No.']}. {r['Symbol']}  s{r['Score']}  C {r['Close']}  T {r['Trigger']}  "
                         f"SL {r['Stop']}  ({r['Gap']}/{r['Risk']})")
    tg_lines += [tg_line(r) for r in new_rows] or ["(no new stocks)"]
    if rep_rows:
        tg_lines += ["", "Still in setup (listed before):"] + [
            f"{r['Symbol']} day {r['Days']}  T {r['Trigger']}" for r in rep_rows]
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
