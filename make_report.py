"""Turn scan_output/ into results/REPORT.md (tables + embedded charts) so it can be read on GitHub."""
import glob
import os
import shutil
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd

SRC, DST = "scan_output", "results"
xlsx = sorted(glob.glob(f"{SRC}/alerts_*.xlsx"))[-1]
sheets = pd.read_excel(xlsx, sheet_name=None)

shutil.rmtree(DST, ignore_errors=True)
os.makedirs(f"{DST}/charts")
for p in glob.glob(f"{SRC}/charts/*.png"):
    shutil.copy(p, f"{DST}/charts")
shutil.copy(xlsx, f"{DST}/latest.xlsx")


def pct(x):
    return "" if pd.isna(x) else f"{x:.1%}"


def md_table(df):
    if df.empty:
        return "_none_\n"
    df = df.fillna("").astype(str).apply(lambda s: s.str.replace("|", "/", regex=False))
    lines = ["| " + " | ".join(df.columns) + " |", "|" + "---|" * len(df.columns)]
    lines += ["| " + " | ".join(r) + " |" for r in df.values]
    return "\n".join(lines) + "\n"


def alert_table(df):
    d = df.copy()
    for c in ("Trigger Gap (%)", "Stop Risk (%)", "Max Gain After Entry (%)"):
        if c in d:
            d[c] = d[c].map(pct)
    if "Alert Date" in d:
        d["Alert Date"] = pd.to_datetime(d["Alert Date"]).dt.strftime("%d %b")
    d["Symbol"] = d["Symbol"].map(lambda s: f"[{s}](https://in.tradingview.com/chart/?symbol=NSE%3A{s})")
    cols = ["Alert Date", "Symbol", "Name", "Close", "Trigger (alert-day high)", "Stop (5-bar low)",
            "Trigger Gap (%)", "Stop Risk (%)", "Score (0-6)", "Flags", "Status", "Entry", "R Multiple"]
    return md_table(d[[c for c in cols if c in d]])


latest, allal = sheets["Latest_Watchlist"], sheets["All_Alerts"]
now = datetime.now(ZoneInfo("Asia/Kolkata")).strftime("%d %b %Y %H:%M IST")
out = [f"# NSE momentum-base scan\n\nGenerated {now} from `{os.path.basename(xlsx)}`. "
       f"Full workbook: [latest.xlsx](latest.xlsx).\n",
       f"## Latest session watchlist ({len(latest)})\n", alert_table(latest)]

for _, r in latest.iterrows():
    f = f"charts/{pd.Timestamp(r['Alert Date']):%Y%m%d}_{r['Symbol']}_s{r['Score (0-6)']}.png".replace("&", "and")
    if os.path.exists(f"{DST}/{f}"):
        out.append(f"\n### {r['Name']} ({r['Symbol']})\n![{r['Symbol']}]({f})\n")

out += [f"\n## All alerts, last 45 sessions ({len(allal)})\n", alert_table(allal)]
out += ["\n## Outcome summary (matured alerts)\n", md_table(sheets["Outcome_Summary"])]
if "Reference_Calibration" in sheets:
    cal = sheets["Reference_Calibration"][["Stock", "Alert (IST)", "Result"]]
    out += ["\n## Calibration vs 19 reference alerts\n", md_table(cal)]

report = "\n".join(out)
open(f"{DST}/REPORT.md", "w", encoding="utf-8").write(report)

summary = os.environ.get("GITHUB_STEP_SUMMARY")
if summary:   # job summary can't show repo images, so use the text-only part
    with open(summary, "a", encoding="utf-8") as fh:
        fh.write("\n".join(o for o in out if "![" not in o))
print(f"wrote {DST}/REPORT.md ({len(latest)} latest, {len(allal)} total alerts)")
