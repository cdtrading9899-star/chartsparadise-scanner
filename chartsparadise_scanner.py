#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Momentum -> Tight Consolidation -> Trigger scanner for NSE stocks (Yahoo Finance data)
======================================================================================
Reverse-engineered from 20 "ChartsParadise" TradingView alert screenshots (18 stocks,
Sep-Nov 2026).  Runs on Google Colab (also plain Python 3.10+).

WHAT IT DOES
  1. Downloads the NSE equity list and ~2 years of daily data from Yahoo Finance.
  2. Flags every stock/day that matches the setup:
        strong prior run-up (best 20-bar gain >= 30%, or >= 60% off the 120-bar low)
        -> close >= 3% above a rising 50 EMA with EMA10 >= EMA20 > EMA50
        -> price hugging the 10/20 EMA ribbon (<= 5% above EMA10) in a tight 10-bar base
        -> alert-day high <= 2.5% above the close (the "trigger" = dashed blue line).
     ENTRY RULE: default is "close_confirm" - a trade is entered only after a daily CLOSE above the
     alert-day high (CFG["entry_mode"]); the older "intraday_cross" rule is reported side by side.
     A 0-6 score (dry volume, inside/narrow bar, tight ribbon, volume spike HVQ/HVY/HVE,
     close near trigger, tight base) ranks the alerts; min_score=3 by default.
  3. Tracks what happened after each alert (trigger hit? stopped? 2R?) so you can verify.
  4. Shows chart snapshots inline, saves them as PNG, and writes an Excel workbook.
  5. Zips Excel + charts + rules into one file (auto-downloads in Colab).
  6. Re-checks the reference alerts against the rules ("Reference_Calibration" sheet)
     so you can see exactly which rule rejected a stock and tune it.

HOW TO RUN IN COLAB
  Paste the whole file into one cell (or upload it and run  %run chartsparadise_scanner.py).
  Edit CFG / RULES below first if you want.  A full-market run takes ~5-10 minutes.
  Tip: set CFG["universe_limit"] = 300 for a quick dry run.

NOTE: Research/screening aid only - not investment advice.  Rules were inferred from the
screenshots by eye; use the calibration sheet to tune them.
"""

# %% ======================================================================
# 0. SETUP
# ==========================================================================
import sys, os, io, json, time, zipfile, shutil, warnings, subprocess
from datetime import datetime, timedelta, time as dtime
from urllib.parse import quote
from zoneinfo import ZoneInfo


def _pip_install(*pkgs):
    subprocess.run([sys.executable, "-m", "pip", "install", "-q", "-U", *pkgs], check=False)


IN_COLAB = "google.colab" in sys.modules
if IN_COLAB:                      # Colab ships an old yfinance; Yahoo changes often
    _pip_install("yfinance", "openpyxl")

import numpy as np
import pandas as pd
import requests
import matplotlib
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter
import yfinance as yf

warnings.filterwarnings("ignore")
IST = ZoneInfo("Asia/Kolkata")

# %% ======================================================================
# 1. SETTINGS  (edit freely)
# ==========================================================================
CFG = {
    # ---- universe ----
    "universe": "auto",          # "auto" = NSE list from the web | "csv" = your file | "list" = your symbols
    "universe_csv": None,        # path to a CSV with a 'Symbol' column (when universe="csv")
    "universe_list": [],         # e.g. ["VARROC", "TDPOWERSYS"]       (when universe="list")
    "universe_limit": None,      # e.g. 300 for a quick test run
    "min_price": 40.0,           # ignore penny stocks
    "min_turnover": 2e7,         # median 20-day traded value >= Rs 2 crore
    # ---- data ----
    "history_period": "2y",
    "chunk_size": 100,
    "fill_missing": True,        # fill a missing latest-session bar from the NSE bhavcopy (else Yahoo 15m data)
    "fill_max_days": 7,          # look back at most this many calendar days for missing sessions
    # ---- scan window / alert logic ----
    "lookback_days": 45,         # list alerts from the last N trading days
    "min_score": 3,              # soft-score needed (0-6) on top of the hard rules
    "max_per_day": 25,           # keep only the best N alerts per day (None = no cap)
    "cooldown_bars": 5,          # don't re-alert the same stock within N bars ...
    "relist_latest": True,       # ... except on the latest session: every stock that qualifies today is listed
    "repeat_mode": "label",      # latest-session stocks that already qualified yesterday (still in the same base):
                                 #   "label" = keep them, tagged "REPEAT (day N)" / "Days On List";  "hide" = list only NEW ones
    # ---- forward tracking (to verify alerts) ----
    "entry_mode": "close_confirm",   # "close_confirm" = enter at the CLOSE of the first day that closes above the
                                     #   alert-day high (confirmation);  "intraday_cross" = enter as soon as price
                                     #   trades above it.  Both are always computed; this picks the main columns.
    "trigger_window": 10,        # bars allowed for price to confirm / cross the trigger
    "hold_bars": 10,             # bars to hold after entry
    "target_r": 2.0,             # target in R multiples
    "stop_lookback": 5,          # stop = lowest low of last N bars (incl. alert bar)
    # ---- output ----
    "chart_bars": 110,           # bars shown on each chart
    "chart_forward": 12,         # bars AFTER the alert shown in grey (what happened next)
    "max_display": 20,           # charts shown inline
    "max_save": 120,             # charts saved to PNG / zip
    "show_inline": True,
    "out_dir": "scan_output",
    "run_calibration": True,     # re-test the reference alerts
    "download_zip": True,        # auto-download in Colab
}

RULES = {
    # moving averages used for the ribbon (green/orange = fast/slow, red = long)
    "ma_type": "ema", "fast": 10, "slow": 20, "long": 50,
    "ma_tol": 0.005,              # fast EMA may sit up to 0.5% below slow EMA (ribbon twists in tight bases)
    "allow_ipo": True,            # new listings (<50 bars) are judged on 10/20 EMA only
    "min_bars": 25,
    # prior momentum
    "thrust_window": 20, "thrust_lookback": 120, "min_thrust": 0.30,   # best 20-bar run-up in last 120 bars >= 30% ...
    "min_runup_from_low": 0.35,                                         # ... AND close >= 35% above the 120-bar low
    "trend_runup": 0.50,                                                # ... OR a slower grind: close >= 50% above the 120-bar low
                                                                        #     (was 0.60; 0.50 catches staircase trends like NAZARA 3 Nov 2026)
    # location
    "high_lookback": 60, "max_off_high": 0.15,                          # within 15% of 60-bar high
    "max_above_fast": 0.05, "min_vs_fast": -0.02, "min_vs_slow": -0.01, # hugging the ribbon (close vs EMA10 / EMA20)
    "min_above_long": 0.03,                                             # close >= 3% above EMA50 (when it exists)
    # tightness
    "close_range_bars": 10, "max_close_range": 0.14,                    # 10-bar closing range <= 14%
    "max_ribbon": 0.065,                                                # |EMA10-EMA20|/EMA20 <= 6.5%
    "max_trigger_gap": 0.025,                                           # alert-day high <= 2.5% above close
    # soft score (each = 1 point, max 6)
    "dry_ratio": 0.90,            # median vol (5d) / 50d avg vol <= 0.9   -> "Dry volumes"
    "nr_atr": 0.75,               # alert-bar range <= 0.75 ATR14          -> narrow range (or inside bar)
    "tight_ribbon": 0.04,         # ribbon spread <= 4%
    "hv_mult": 3.0, "hv_lookback": 90,   # a >=3x volume spike in last 90 bars (HVQ/HVY/HVE)
    "score_trigger_gap": 0.015,   # trigger within 1.5% of close
    "score_close_range": 0.08,    # 10-bar closing range <= 8%
}

# %% ======================================================================
# 2. REFERENCE ALERTS (from the screenshots)  -  IST timestamps as printed on the charts
#    (name, yahoo symbol, alert time IST, trigger price on chart, last price, timeframe, family, tags)
# ==========================================================================
REFERENCE = [
    ("Shivalik Bimetal Controls", "SBCL",       "2026-09-02 16:49", 1060.90, 1050.90, "1D", "base",    "Thrust gap, tight base"),
    ("INDO-MIM",                  "INDOMIM",    "2026-09-09 14:52", 979.45,  966.00,  "1D", "base",    "IPO base, dry volume"),
    ("Avadh Sugar & Energy",      "AVADHSUGAR", "2026-09-08 18:57", 844.50,  836.25,  "1D", "base",    "Momentum, Tightness, Dry Volumes"),
    ("TD Power Systems",          "TDPOWERSYS", "2026-09-10 20:39", 772.00,  762.75,  "1D", "base",    "Up Momentum, Dry Volumes"),
    ("ASK Automotive",            "ASKAUTOLTD", "2026-09-11 10:50", 634.00,  627.55,  "1D", "base",    "Momentum + base on 10/20 EMA"),
    ("Talbros Automotive Comp.",  "TALBROAUTO", "2026-09-14 15:31", 428.00,  423.75,  "1D", "base",    "Momentum, Level to watch 428, invalidation 401.5"),
    ("ASK Automotive (1h)",       "ASKAUTOLTD", "2026-09-16 12:06", None,    621.75,  "1H", "support", "Triple-tested support 598.85-602.70 (NOT covered by daily scanner)"),
    ("MV Electrosystems",         "MVELECTRO",  "2026-09-18 11:22", 828.20,  812.00,  "1D", "ipo",     "IPO pullback to 10 EMA"),
    ("Varroc Engineering",        "VARROC",     "2026-09-21 18:06", 879.90,  872.55,  "1D", "base",    "Strong Up Momentum, HVE, Tightness"),
    ("Sona BLW Precision Forgings", "SONACOMS", "2026-09-22 15:09", 807.55,  803.20,  "1D", "base",    "Momentum, shake-out & rebalance at 50 EMA"),
    ("Electronics Mart India",    "EMIL",       "2026-09-25 14:06", None,    192.72,  "1D", "base",    "Strong run-up, HVE, sideways, dry volumes"),
    ("Venus Pipes & Tubes",       "VENUSPIPES", "2026-10-04 19:47", 2244.40, 2204.90, "1D", "base",    "Up Momentum, massive volumes"),
    ("Marine Electricals (India)", "MARINE",    "2026-10-04 19:51", 439.50,  436.55,  "1D", "base",    "Up Momentum, HVE, inside bar"),
    ("Aegis Logistics",           "AEGISLOG",   "2026-10-06 18:59", 1410.00, 1395.90, "1D", "base",    "Strong momentum, sideways, IB, dry volumes"),
    ("Wheels India (1h)",         "WHEELS",     "2026-10-08 11:08", None,    2216.50, "1H", "base",    "Contracting triangle / IB, dry volumes (1h view)"),
    ("Piramal Pharma",            "PPLPHARMA",  "2026-09-03 21:35", 216.90,  214.49,  "1D", "base",    "Strong momentum, HVY, tight range, dry volumes"),
    ("Knowledge Marine & Eng.",   "KMEW",       "2026-10-06 08:45", 3033.00, 2992.50, "1D", "base",    "Steady uptrend riding ribbon"),
    ("Wheels India",              "WHEELS",     "2026-10-07 21:31", 2315.80, 2283.90, "1D", "base",    "Up Momentum, IB, dry volumes"),
    ("Deep Industries",           "DEEPINDS",   "2026-10-13 18:44", 774.55,  770.00,  "1D", "base",    "Up Momentum, sideways"),
    # alert time not shown on the chart; daily bar of 3 Nov (high 408.40, close 399.20), broke out to 421.05 on 4 Nov
    ("Nazara Technologies",       "NAZARA",     "2026-11-03 15:30", None,    399.20,  "1D", "base",    "Staircase uptrend, tight base on 10/20 EMA, dry volumes"),
]

# %% ======================================================================
# 3. UNIVERSE
# ==========================================================================
NSE_URLS = [
    ("NSE all listed equities", "https://nsearchives.nseindia.com/content/equities/EQUITY_L.csv"),
    ("NSE all listed equities (alt)", "https://archives.nseindia.com/content/equities/EQUITY_L.csv"),
    ("Nifty Total Market", "https://niftyindices.com/IndexConstituent/ind_niftytotalmarket_list.csv"),
    ("Nifty 500", "https://nsearchives.nseindia.com/content/indices/ind_nifty500list.csv"),
]
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/124.0 Safari/537.36",
    "Accept": "text/csv,application/csv,*/*",
    "Referer": "https://www.nseindia.com/",
}


def parse_universe_csv(text):
    df = pd.read_csv(io.StringIO(text))
    df.columns = [str(c).strip() for c in df.columns]
    low = {c.lower(): c for c in df.columns}
    sym = low.get("symbol")
    if sym is None:
        raise ValueError("CSV has no 'Symbol' column")
    name = low.get("name of company") or low.get("company name") or low.get("name")
    ser = low.get("series")
    if ser:
        df = df[df[ser].astype(str).str.strip().str.upper() == "EQ"]
    out = pd.DataFrame({
        "symbol": df[sym].astype(str).str.strip().str.upper(),
        "name": df[name].astype(str).str.strip() if name else df[sym].astype(str).str.strip(),
    })
    return out.drop_duplicates("symbol").reset_index(drop=True)


def load_universe(cfg):
    mode = cfg["universe"]
    if mode == "list":
        syms = [s.strip().upper() for s in cfg["universe_list"] if s.strip()]
        return pd.DataFrame({"symbol": syms, "name": syms})
    if mode == "csv":
        with open(cfg["universe_csv"], "r", encoding="utf-8", errors="ignore") as fh:
            return parse_universe_csv(fh.read())
    for label, url in NSE_URLS:
        try:
            r = requests.get(url, headers=HEADERS, timeout=25)
            r.raise_for_status()
            uni = parse_universe_csv(r.text)
            if len(uni) > 100:
                print(f"Universe: {label} -> {len(uni)} symbols")
                return uni
        except Exception as e:                         # noqa: BLE001
            print(f"  [{label}] not reachable ({type(e).__name__}); trying next source...")
    print("!! Could not download an NSE list. Falling back to the 17 reference stocks only.\n"
          "   Fix: upload a CSV with a 'Symbol' column and set CFG['universe']='csv', CFG['universe_csv']='file.csv'")
    syms = sorted({r[1] for r in REFERENCE})
    return pd.DataFrame({"symbol": syms, "name": syms})


# %% ======================================================================
# 4. DATA (Yahoo Finance)
# ==========================================================================
OHLCV = ["Open", "High", "Low", "Close", "Volume"]


def clean_ohlcv(df):
    df = df[[c for c in OHLCV if c in df.columns]].copy()
    if len(df.columns) < 5:
        return None
    df.index = pd.to_datetime(df.index).tz_localize(None).normalize()
    df = df[~df.index.duplicated(keep="last")].sort_index()
    df = df.dropna(subset=["Close", "High", "Low", "Open"])
    df["Volume"] = df["Volume"].fillna(0)
    return df if len(df) else None


def split_download(raw):
    """yfinance -> {symbol_without_.NS: ohlcv_df}; handles both column orientations & single ticker."""
    out = {}
    if raw is None or len(raw) == 0:
        return out
    if isinstance(raw.columns, pd.MultiIndex):
        lv0 = set(raw.columns.get_level_values(0))
        if {"Close", "Open"} <= lv0:                      # (field, ticker) orientation
            raw = raw.swaplevel(axis=1)
        for t in raw.columns.get_level_values(0).unique():
            try:
                d = clean_ohlcv(raw.xs(t, axis=1, level=0))
            except Exception:                              # noqa: BLE001
                d = None
            if d is not None:
                out[str(t).replace(".NS", "")] = d
    else:
        d = clean_ohlcv(raw)
        if d is not None:
            out["__single__"] = d
    return out


def fetch_prices(symbols, period="2y", chunk=100, pause=1.0):
    tickers = [s + ".NS" for s in symbols]
    out = {}
    for k in range(0, len(tickers), chunk):
        part = tickers[k:k + chunk]
        raw = None
        for attempt in range(3):
            try:
                raw = yf.download(part, period=period, interval="1d", group_by="ticker",
                                  auto_adjust=True, threads=True, progress=False)
                if raw is not None and len(raw):
                    break
            except Exception as e:                         # noqa: BLE001
                print(f"  download error ({type(e).__name__}); retry {attempt + 1}/3")
            time.sleep(2 * (attempt + 1))
        got = split_download(raw)
        if "__single__" in got and len(part) == 1:
            got = {part[0].replace(".NS", ""): got["__single__"]}
        out.update({k2: v for k2, v in got.items() if k2 != "__single__"})
        print(f"  downloaded {min(k + chunk, len(tickers))}/{len(tickers)}  (usable so far: {len(out)})")
        time.sleep(pause)
    return out


def expected_session(now=None):
    """Latest weekday session that has closed (16:00 IST). Exchange holidays are not known here: a holiday simply
    has no bhavcopy / intraday data, so nothing gets filled for it."""
    now = now or datetime.now(IST)
    d = now.date()
    if not (now.weekday() < 5 and now.time() >= dtime(16, 0)):
        d -= timedelta(days=1)
    while d.weekday() >= 5:
        d -= timedelta(days=1)
    return pd.Timestamp(d)


def drop_incomplete_today(df):
    """Drop any bar newer than the latest closed session (today's unfinished bar during market hours)."""
    return df[df.index <= expected_session()]


def fetch_bhavcopy(day):
    """NSE equity bhavcopy (UDiFF) for one day -> DataFrame indexed by symbol, or None (holiday / blocked)."""
    url = f"https://nsearchives.nseindia.com/content/cm/BhavCopy_NSE_CM_0_0_0_{day:%Y%m%d}_F_0000.csv.zip"
    try:
        r = requests.get(url, headers={**HEADERS, "Accept": "*/*"}, timeout=30)
        if r.status_code != 200 or not r.content.startswith(b"PK"):
            return None
        z = zipfile.ZipFile(io.BytesIO(r.content))
        b = pd.read_csv(z.open(z.namelist()[0]))
    except Exception:                                      # noqa: BLE001
        return None
    b = b[b["SctySrs"].isin(["EQ", "BE", "BZ", "SM", "ST"])].copy()
    b["rank"] = b["SctySrs"].map({"EQ": 0, "BE": 1, "BZ": 2, "SM": 3, "ST": 4})
    b = b.sort_values("rank").drop_duplicates("TckrSymb").set_index("TckrSymb")
    return b.rename(columns={"OpnPric": "Open", "HghPric": "High", "LwPric": "Low", "ClsPric": "Close",
                             "TtlTradgVol": "Volume", "PrvsClsgPric": "PrevClose"})[
        ["Open", "High", "Low", "Close", "Volume", "PrevClose"]]


def intraday_daily(symbols, chunk=100):
    """Daily bars rebuilt from Yahoo 15-minute data (last 5 days) -> {symbol: df}. Fallback only: the close is the
    last 15-minute close, which can differ a little from NSE's official closing price."""
    out = {}
    for k in range(0, len(symbols), chunk):
        part = [s + ".NS" for s in symbols[k:k + chunk]]
        try:
            raw = yf.download(part, period="5d", interval="15m", group_by="ticker", auto_adjust=True,
                              threads=True, progress=False)
        except Exception:                                  # noqa: BLE001
            continue
        if raw is None or not len(raw):
            continue
        if not isinstance(raw.columns, pd.MultiIndex):
            raw = pd.concat({part[0]: raw}, axis=1)
        elif {"Close", "Open"} <= set(raw.columns.get_level_values(0)):
            raw = raw.swaplevel(axis=1)
        for t in raw.columns.get_level_values(0).unique():
            d = raw[t].dropna(subset=["Close"])
            if not len(d):
                continue
            idx = d.index.tz_convert(IST) if d.index.tz is not None else d.index
            g = d.groupby(pd.DatetimeIndex(idx).tz_localize(None).normalize()).agg(
                Open=("Open", "first"), High=("High", "max"), Low=("Low", "min"),
                Close=("Close", "last"), Volume=("Volume", "sum"))
            out[str(t).replace(".NS", "")] = g
    return out


def fill_missing_sessions(prices, cfg):
    """Yahoo sometimes has no daily bar yet for the latest session (it was missing for most stocks in every late-night
    GitHub run so far), so every alert of that session was lost.  Fill those bars from the NSE bhavcopy (exact
    official prices), else rebuild them from Yahoo 15-minute data.  Returns a data-check summary."""
    E = expected_session()
    last = {s: d.index[-1] for s, d in prices.items()}
    info = {"Expected session": str(E.date()), "Stocks": len(prices),
            "Had the session from Yahoo": sum(v >= E for v in last.values()),
            "Filled from NSE bhavcopy": 0, "Filled from Yahoo 15m": 0, "Bhavcopy days used": ""}
    stale = sorted(s for s, v in last.items() if v < E)
    if not stale or not cfg.get("fill_missing", True):
        info["Still without the session"] = len(stale)
        return info
    oldest = max(min(last[s] for s in stale), E - pd.Timedelta(days=cfg.get("fill_max_days", 7)))
    days = [d for d in pd.date_range(oldest + pd.Timedelta(days=1), E) if d.weekday() < 5]
    used = []
    for day in days:                                      # oldest first, so each fill continues the previous bar
        b = fetch_bhavcopy(day)
        if b is None:
            continue
        used.append(f"{day:%d %b}")
        for s in stale:
            df = prices[s]
            if df.index[-1] >= day or s not in b.index:
                continue
            row = b.loc[s]
            # continuity: NSE's previous close must be our last close, else a day is missing or a split happened
            if not (row["Volume"] > 0 and abs(row["PrevClose"] / df["Close"].iloc[-1] - 1) <= 0.02):
                continue
            prices[s] = pd.concat([df, pd.DataFrame([row[OHLCV].astype(float)], index=[day])])
            info["Filled from NSE bhavcopy"] += 1
    info["Bhavcopy days used"] = ", ".join(used) or "none (NSE not reachable or holiday)"
    stale = [s for s in stale if prices[s].index[-1] < E]
    have = info["Had the session from Yahoo"] + info["Filled from NSE bhavcopy"]
    # 15m fallback only when the session is broadly missing (NSE unreachable); thinly traded names are often absent
    if stale and have < 0.8 * len(prices):
        for s, g in intraday_daily(stale, cfg["chunk_size"]).items():
            df = prices[s]
            new = g[(g.index > df.index[-1]) & (g.index <= E) & (g["Volume"] > 0)]
            new = new[(new["Close"] / df["Close"].iloc[-1] - 1).abs() < 0.25]
            if len(new):
                prices[s] = pd.concat([df, new[OHLCV]])
                info["Filled from Yahoo 15m"] += 1
    info["Still without the session"] = sum(prices[s].index[-1] < E for s in prices)
    return info


# %% ======================================================================
# 5. PATTERN ENGINE
# ==========================================================================
GATES = ["trend", "thrust", "near_high", "hug", "tight", "trigger_gap", "liquidity"]


def compute_features(df, R, cfg):
    c, h, l, o, v = df["Close"], df["High"], df["Low"], df["Open"], df["Volume"]
    n = len(df)
    f = pd.DataFrame(index=df.index)
    if R["ma_type"] == "ema":
        ma = lambda s, k: s.ewm(span=k, adjust=False, min_periods=k).mean()
    else:
        ma = lambda s, k: s.rolling(k, min_periods=k).mean()
    f["ema_f"], f["ema_s"], f["ema_l"] = ma(c, R["fast"]), ma(c, R["slow"]), ma(c, R["long"])

    pc = c.shift(1)
    tr = pd.concat([h - l, (h - pc).abs(), (l - pc).abs()], axis=1).max(axis=1)
    atr14 = tr.rolling(14, min_periods=10).mean()
    vol_sma50 = v.rolling(50, min_periods=20).mean()

    # --- momentum / location metrics
    run20 = c / c.rolling(R["thrust_window"], min_periods=10).min() - 1
    f["thrust"] = run20.rolling(R["thrust_lookback"], min_periods=20).max()
    f["runup_low"] = c / l.rolling(R["thrust_lookback"], min_periods=20).min() - 1
    f["off_high"] = 1 - c / h.rolling(R["high_lookback"], min_periods=20).max()
    f["dist_fast"] = c / f["ema_f"] - 1
    f["ribbon"] = (f["ema_f"] - f["ema_s"]) / f["ema_s"]
    cr = R["close_range_bars"]
    f["close_range"] = (c.rolling(cr).max() - c.rolling(cr).min()) / c
    f["trigger_gap"] = h / c - 1
    f["vol_ratio"] = v.rolling(5).median() / vol_sma50
    f["turnover"] = (c * v).rolling(20, min_periods=10).median()
    spike = v / vol_sma50.shift(1)
    f["hv_event"] = spike.rolling(R["hv_lookback"], min_periods=10).max() >= R["hv_mult"]
    f["ib"] = (h <= h.shift(1)) & (l >= l.shift(1))
    f["nr"] = (h - l) <= R["nr_atr"] * atr14
    f["run20"] = run20

    # --- hard gates
    tol = R["ma_tol"]
    has_long = f["ema_l"].notna()
    full = ((c >= f["ema_l"] * (1 + R["min_above_long"])) & (f["ema_f"] >= f["ema_s"] * (1 - tol))
            & (f["ema_s"] > f["ema_l"]) & (f["ema_l"] > f["ema_l"].shift(10)))
    ipo = ((f["ema_f"] >= f["ema_s"] * (1 - tol)) & (c >= f["ema_s"] * 0.99)
           & (np.arange(1, n + 1) >= R["min_bars"]) & bool(R["allow_ipo"]))
    f["g_trend"] = np.where(has_long, full, ipo)
    f["g_thrust"] = (((f["thrust"] >= R["min_thrust"]) & (f["runup_low"] >= R["min_runup_from_low"]))
                     | (f["runup_low"] >= R["trend_runup"]))
    f["g_near_high"] = f["off_high"] <= R["max_off_high"]
    f["g_hug"] = ((f["dist_fast"] >= R["min_vs_fast"]) & (f["dist_fast"] <= R["max_above_fast"])
                  & (c >= f["ema_s"] * (1 + R["min_vs_slow"])))
    f["g_tight"] = (f["close_range"] <= R["max_close_range"]) & (f["ribbon"].abs() <= R["max_ribbon"])
    f["g_trigger_gap"] = f["trigger_gap"] <= R["max_trigger_gap"]
    f["g_liquidity"] = (c >= cfg["min_price"]) & (f["turnover"] >= cfg["min_turnover"])
    f["hard_pass"] = f[["g_" + g for g in GATES]].all(axis=1)

    # --- soft score (0-6)
    f["s_dry"] = f["vol_ratio"] <= R["dry_ratio"]
    f["s_narrow"] = f["ib"] | f["nr"]
    f["s_ribbon"] = f["ribbon"].abs() <= R["tight_ribbon"]
    f["s_hv"] = f["hv_event"]
    f["s_trig"] = f["trigger_gap"] <= R["score_trigger_gap"]
    f["s_base"] = f["close_range"] <= R["score_close_range"]
    f["score"] = f[["s_dry", "s_narrow", "s_ribbon", "s_hv", "s_trig", "s_base"]].sum(axis=1)
    f["signal"] = f["hard_pass"] & (f["score"] >= cfg["min_score"])
    return f


def hv_label(df, i, R):
    """HVE / HVY / HVQ tag for the biggest volume spike in the last hv_lookback bars."""
    v = df["Volume"].values
    lo = max(0, i - R["hv_lookback"] + 1)
    seg = v[lo:i + 1]
    if len(seg) == 0:
        return ""
    k = lo + int(np.argmax(seg))
    vm = v[k]
    prior = v[max(0, k - 50):k]
    base = prior.mean() if len(prior) else np.nan
    if not (base and base > 0 and vm >= R["hv_mult"] * base):
        return ""
    if vm >= v[:i + 1].max():
        return "HVE"            # highest volume in the downloaded history (~2y)
    if vm >= v[max(0, k - 249):k + 1].max():
        return "HVY"
    if vm >= v[max(0, k - 62):k + 1].max():
        return "HVQ"
    return "HV"


def track_outcome(df, i, cfg, mode=None):
    """What happened after the alert?  Trigger = alert-day high; stop = lowest low of last N bars.

    mode "close_confirm": entry = close of the first bar that CLOSES above the trigger; stop/target are
                          only checked from the NEXT bar (you can't be stopped out of a trade you haven't entered).
    mode "intraday_cross": entry = max(open, trigger) on the first bar whose HIGH exceeds the trigger.
    """
    mode = mode or cfg["entry_mode"]
    n = len(df)
    H, L, O, C = (df[k].values for k in ("High", "Low", "Open", "Close"))
    trig = float(H[i])
    stop = float(L[max(0, i - cfg["stop_lookback"] + 1): i + 1].min())
    out = dict(trigger=trig, stop=stop, status="", trigger_date=None, entry=np.nan,
               days_to_trigger=np.nan, risk_pct=np.nan, max_gain=np.nan, max_dd=np.nan, r_mult=np.nan)
    if i >= n - 1:
        out["status"] = "LIVE - watch next session"
        return out
    last = min(n - 1, i + cfg["trigger_window"])
    j = None
    for k in range(i + 1, last + 1):
        hit = (C[k] > trig) if mode == "close_confirm" else (H[k] > trig)
        if hit:
            j = k
            break
        if L[k] < stop:
            out["status"] = "Invalidated before trigger"
            return out
    if j is None:
        out["status"] = ("LIVE - awaiting trigger" if last < i + cfg["trigger_window"]
                         else f"No trigger in {cfg['trigger_window']}d")
        return out
    if mode == "close_confirm":
        entry, first_k = float(C[j]), j + 1
    else:
        entry, first_k = float(max(O[j], trig)), j
    risk = entry - stop
    end = min(n - 1, j + cfg["hold_bars"])
    target = entry + cfg["target_r"] * risk
    status, r, exit_k = None, np.nan, end
    for k in range(first_k, end + 1):
        if L[k] <= stop:                                   # stop checked first = conservative
            status, r, exit_k = "Triggered - stopped out", -1.0, k
            break
        if H[k] >= target:
            status, r, exit_k = f"Triggered - {cfg['target_r']:.0f}R target hit", cfg["target_r"], k
            break
    if status is None:
        r = (C[end] - entry) / risk
        in_progress = (end == n - 1) and (n - 1 - j) < cfg["hold_bars"]
        status = "Triggered - open" if in_progress else f"Triggered - timed exit ({cfg['hold_bars']}d)"
    seg_hi, seg_lo = H[first_k:exit_k + 1], L[first_k:exit_k + 1]
    out.update(status=status, trigger_date=df.index[j], entry=entry, days_to_trigger=j - i,
               risk_pct=risk / entry,
               max_gain=(seg_hi.max() / entry - 1) if len(seg_hi) else np.nan,
               max_dd=(seg_lo.min() / entry - 1) if len(seg_lo) else np.nan, r_mult=r)
    return out


def flags_text(f, i, df, R):
    row = f.iloc[i]
    fl = []
    if row["s_dry"]:
        fl.append("Dry volume")
    if row["ib"]:
        fl.append("IB")
    elif row["nr"]:
        fl.append("NR")
    if row["s_ribbon"]:
        fl.append("Tight ribbon")
    if row["s_hv"]:
        fl.append(hv_label(df, i, R) or "HV")
    if row["s_trig"]:
        fl.append("Close near trigger")
    if row["s_base"]:
        fl.append("Tight base")
    return ", ".join(fl)


def build_record(sym, name, df, f, i, cfg, R):
    row = f.iloc[i]
    oc = track_outcome(df, i, cfg)
    alt_mode = "intraday_cross" if cfg["entry_mode"] == "close_confirm" else "close_confirm"
    alt = track_outcome(df, i, cfg, alt_mode)
    return {
        "Alert Date": df.index[i].date(),
        "Alert Time (IST)": "EOD (after 15:30) - daily bar complete",
        "Symbol": sym, "Name": name, "Yahoo Ticker": sym + ".NS",
        "Close": round(float(df["Close"].iloc[i]), 2),
        "Trigger (alert-day high)": round(oc["trigger"], 2),
        "Stop (5-bar low)": round(oc["stop"], 2),
        "Trigger Gap (%)": float(row["trigger_gap"]),
        "Stop Risk (%)": float(1 - oc["stop"] / oc["trigger"]),
        "Score (0-6)": int(row["score"]),
        "Flags": flags_text(f, i, df, R),
        "Best 20d Run-up (%)": float(row["thrust"]),
        "Run-up from 120d Low (%)": float(row["runup_low"]),
        "Off 60d High (%)": float(row["off_high"]),
        "Dist to EMA10 (%)": float(row["dist_fast"]),
        "Ribbon Spread (%)": float(row["ribbon"]),
        "10d Close Range (%)": float(row["close_range"]),
        "Vol Ratio (5d med / 50d avg)": round(float(row["vol_ratio"]), 2) if pd.notna(row["vol_ratio"]) else np.nan,
        "Status": oc["status"],
        "Trigger Date": oc["trigger_date"].date() if oc["trigger_date"] is not None else None,
        "Entry": round(oc["entry"], 2) if pd.notna(oc["entry"]) else np.nan,
        "Days to Entry": oc["days_to_trigger"],
        "Entry Risk (%)": oc["risk_pct"],
        "Max Gain After Entry (%)": oc["max_gain"],
        "Max Drawdown After Entry (%)": oc["max_dd"],
        "R Multiple": round(oc["r_mult"], 2) if pd.notna(oc["r_mult"]) else np.nan,
        "Entry Mode": cfg["entry_mode"],
        "Alt Entry Mode": alt_mode,
        "Alt Status": alt["status"],
        "Alt Entry": round(alt["entry"], 2) if pd.notna(alt["entry"]) else np.nan,
        "Alt R Multiple": round(alt["r_mult"], 2) if pd.notna(alt["r_mult"]) else np.nan,
        "TradingView": f'=HYPERLINK("https://in.tradingview.com/chart/?symbol=NSE%3A{quote(sym)}","Open chart")',
        "Chart File": "",
        "_i": i,
        "_bars_after": len(df) - 1 - i,
    }


def scan(prices, names, cfg, R):
    """Run the engine over every symbol; returns (alerts_df, {symbol: (df, features)})."""
    records, keep = [], {}
    global_last = max(d.index[-1] for d in prices.values())
    for sym, df in prices.items():
        if len(df) < R["min_bars"] + 5:
            continue
        f = compute_features(df, R, cfg)
        start = max(0, len(df) - cfg["lookback_days"])
        idx = [i for i in range(start, len(df)) if bool(f["signal"].iloc[i])]
        sig = f["signal"].values

        def streak(i):
            """Consecutive sessions (ending at bar i) on which this stock met the setup."""
            k, j = 0, i
            while j >= 0 and bool(sig[j]):
                k += 1
                j -= 1
            return k

        last_i = -10 ** 9
        picked = []
        for i in idx:
            # the latest session is always listed when it qualifies, so a setup that is still building (e.g. NAZARA
            # 2 Nov -> 3 Nov 2026) stays on the evening watchlist; older bars keep the cooldown.
            # repeat_mode="hide" drops stocks that were already on yesterday's list (same base, no breakout yet).
            relist = cfg.get("relist_latest", True) and i == len(df) - 1 and df.index[i] == global_last
            if relist and cfg.get("repeat_mode") == "hide" and streak(i) > 1:
                relist = False
            if i - last_i > cfg["cooldown_bars"] or relist:
                picked.append(i)
                last_i = i
        if not picked:
            continue
        keep[sym] = (df, f)
        for i in picked:
            rec = build_record(sym, names.get(sym, sym), df, f, i, cfg, R)
            rec["Latest Session"] = bool(df.index[i] == global_last)
            days = streak(i)
            rec["Days On List"] = days
            rec["Alert Type"] = "NEW" if days <= 1 else f"REPEAT (day {days})"
            records.append(rec)
    if not records:
        return pd.DataFrame(), keep
    out = pd.DataFrame(records).sort_values(["Alert Date", "Score (0-6)", "10d Close Range (%)"],
                                            ascending=[False, False, True])
    if cfg.get("max_per_day"):
        out = out.groupby("Alert Date", sort=False).head(int(cfg["max_per_day"]))
    return out.reset_index(drop=True), keep


# %% ======================================================================
# 6. CHARTS
# ==========================================================================
UP, DOWN = "#26a69a", "#ef5350"


def _vol_fmt(x, _):
    if x >= 1e7:
        return f"{x / 1e7:.1f}Cr"
    if x >= 1e5:
        return f"{x / 1e5:.1f}L"
    if x >= 1e3:
        return f"{x / 1e3:.0f}K"
    return f"{x:.0f}"


def plot_alert(sym, name, df, f, i, rec, cfg, R, path=None, show=True):
    n = len(df)
    a, b = max(0, i - cfg["chart_bars"] + 1), min(n - 1, i + cfg["chart_forward"])
    seg, fs = df.iloc[a:b + 1], f.iloc[a:b + 1]
    xa, x = i - a, np.arange(len(seg))
    fig, (ax, axv) = plt.subplots(2, 1, figsize=(11.5, 6.6), dpi=110, sharex=True,
                                  gridspec_kw={"height_ratios": [3.4, 1], "hspace": 0.04})
    up = (seg["Close"] >= seg["Open"]).values
    col = np.where(up, UP, DOWN)
    yr = seg["High"].max() - seg["Low"].min()
    ax.vlines(x, seg["Low"], seg["High"], color=col, lw=0.9, zorder=3)
    body = np.maximum((seg["Close"] - seg["Open"]).abs().values, yr * 0.002)
    ax.bar(x, body, bottom=np.minimum(seg["Open"], seg["Close"]), width=0.7, color=col, zorder=4)
    ax.plot(x, fs["ema_f"], color="#2e9e48", lw=1.3, zorder=2, label=f"EMA{R['fast']}")
    ax.plot(x, fs["ema_s"], color="#f5a623", lw=1.3, zorder=2, label=f"EMA{R['slow']}")
    if fs["ema_l"].notna().any():
        ax.plot(x, fs["ema_l"], color="#e53935", lw=1.3, zorder=2, label=f"EMA{R['long']}")
    ax.fill_between(x, fs["ema_f"], fs["ema_s"], where=(fs["ema_f"] >= fs["ema_s"]).values,
                    color="#cfe3fb", alpha=0.85, zorder=1, interpolate=True)

    trig, stop = rec["Trigger (alert-day high)"], rec["Stop (5-bar low)"]
    ax.hlines(trig, xa, len(seg) + 1.5, colors="#2962ff", linestyles="--", lw=1.2, zorder=5)
    ax.text(len(seg) + 1.6, trig, f"Trigger {trig:,.2f}", color="white", fontsize=8, va="center", zorder=6,
            bbox=dict(boxstyle="square,pad=0.25", fc="#2962ff", ec="none"))
    ax.hlines(stop, xa, len(seg) + 1.5, colors="#ff9800", linestyles=":", lw=1.2, zorder=5)
    ax.text(len(seg) + 1.6, stop, f"Stop {stop:,.2f}", color="white", fontsize=8, va="center", zorder=6,
            bbox=dict(boxstyle="square,pad=0.25", fc="#ff9800", ec="none"))
    ax.axvline(xa, color="#777", ls=":", lw=1)
    if xa < len(seg) - 1:
        ax.axvspan(xa + 0.5, len(seg) - 0.5, color="black", alpha=0.045, zorder=0)
        ax.text(xa + 0.8, seg["Low"].min() + yr * 0.02, "after alert", color="#666", fontsize=8, style="italic")
    td = rec["Trigger Date"]
    if td is not None and pd.Timestamp(td) in seg.index:
        k = seg.index.get_loc(pd.Timestamp(td))
        ax.plot(k, seg["Low"].iloc[k] - yr * 0.02, marker="^", color="#2962ff", ms=9, zorder=7)

    # tags (like the author's hand-written labels)
    tags = ["Tight consolidation"]
    if rec["Flags"]:
        tags += [t.strip() for t in rec["Flags"].split(",") if t.strip() not in ("Close near trigger", "Tight base")]
    ax.text(0.01, 0.97, "  |  ".join(dict.fromkeys(tags)), transform=ax.transAxes, va="top", fontsize=9,
            color="#2962ff", style="italic", fontweight="bold")
    m = int(np.nanargmax(fs["run20"].values[:xa + 1])) if np.isfinite(fs["run20"].values[:xa + 1]).any() else None
    if m is not None and fs["run20"].iloc[m] >= R["min_thrust"]:
        ax.annotate("Momentum", xy=(m, seg["High"].iloc[m]), xytext=(max(m - 14, 0), seg["High"].iloc[m] + yr * 0.06),
                    fontsize=9, color="#2962ff", style="italic", fontweight="bold",
                    arrowprops=dict(arrowstyle="->", color="#2962ff", lw=0.9))

    # volume pane
    axv.bar(x, seg["Volume"], color=col, width=0.7)
    if rec["Flags"] and "Dry volume" in rec["Flags"]:
        axv.text(max(xa - 12, 0), axv.get_ylim()[1] * 0.8, "Dry Volumes", color="#2962ff", fontsize=9,
                 style="italic", fontweight="bold")
    hv = hv_label(df, i, R)
    if hv:
        lo = max(0, i - R["hv_lookback"] + 1)
        k = lo + int(np.argmax(df["Volume"].values[lo:i + 1])) - a
        if 0 <= k < len(seg):
            axv.text(k, seg["Volume"].iloc[k], hv, ha="center", va="bottom", fontsize=9, color="#2962ff",
                     style="italic", fontweight="bold")
    axv.yaxis.set_major_formatter(FuncFormatter(_vol_fmt))
    axv.yaxis.tick_right()
    ax.yaxis.tick_right()
    for a_ in (ax, axv):
        a_.grid(color="#e6e6e6", lw=0.6)
        a_.set_axisbelow(True)
        for s_ in ("top", "left"):
            a_.spines[s_].set_visible(False)
    ticks = np.linspace(0, len(seg) - 1, 9).astype(int)
    axv.set_xticks(ticks)
    axv.set_xticklabels([seg.index[t].strftime("%d %b") for t in ticks], fontsize=8)
    ax.set_xlim(-1, len(seg) + 9)

    sub = rec["Status"]
    if rec["R Multiple"] == rec["R Multiple"] and rec["Trigger Date"]:
        mg = rec["Max Gain After Entry (%)"]
        sub += (f"  |  entry {rec['Entry']:,.2f} ({rec['Entry Mode']})"
                + (f"  |  max gain {mg:+.1%}" if mg == mg else "") + f"  |  R {rec['R Multiple']:+.1f}")
    fig.suptitle(f"{name} ({sym})  NSE 1D   Alert {df.index[i]:%d %b %Y}   Score {rec['Score (0-6)']}/6   "
                 f"Close {rec['Close']:,.2f}\n{sub}", fontsize=10.5, x=0.01, ha="left", y=0.995)
    ax.legend(loc="lower left", fontsize=8, frameon=False, ncol=3)
    fig.subplots_adjust(left=0.04, right=0.90, top=0.91, bottom=0.07)
    if path:
        fig.savefig(path, bbox_inches="tight")
    if show:
        try:
            from IPython.display import display
            display(fig)
        except Exception:                                  # noqa: BLE001
            plt.show()
    plt.close(fig)


# %% ======================================================================
# 7. EXCEL
# ==========================================================================
def _style_sheet(ws, pct_cols=(), widths=None, freeze="A2"):
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    hdr_fill = PatternFill("solid", fgColor="1F3864")
    for c in ws[1]:
        c.font, c.fill = Font(bold=True, color="FFFFFF"), hdr_fill
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 34
    ws.freeze_panes = freeze
    for ci, cell in enumerate(ws[1], start=1):
        name = str(cell.value or "")
        letter = get_column_letter(ci)
        longest = max((len(str(r[ci - 1].value)) if r[ci - 1].value is not None else 0)
                      for r in ws.iter_rows(min_row=2, max_row=min(ws.max_row, 200))) if ws.max_row > 1 else 0
        ws.column_dimensions[letter].width = min(max(len(name) * 0.9, longest + 2, 9), 60)
        if "(%)" in name or name in pct_cols:
            for r in ws.iter_rows(min_row=2, min_col=ci, max_col=ci):
                r[0].number_format = "0.0%"
    ws.auto_filter.ref = ws.dimensions


def summarize_outcomes(alerts, cfg):
    """Headline stats on MATURED alerts only (>= trigger_window + hold_bars sessions of follow-up).
    Younger alerts are excluded because fast stop-outs resolve first and would bias the numbers."""
    if alerts.empty:
        return pd.DataFrame({"Metric": ["No alerts"], "Value": [""]})
    need = cfg["trigger_window"] + cfg["hold_bars"]
    mat = alerts[alerts["_bars_after"] >= need]
    prim = cfg["entry_mode"]
    alt = "intraday_cross" if prim == "close_confirm" else "close_confirm"

    def col(status, r, entry, risk_entry):
        trig = status.str.startswith("Triggered")
        closed = trig & ~status.str.contains("open")
        rr = r[closed]
        risk = (1 - mat["Stop (5-bar low)"] / entry)[trig]
        pct = lambda x: f"{x:.0%}"
        n_closed = int(closed.sum())
        return [
            f"{int(trig.sum())}  ({pct(trig.mean())})" if len(mat) else "n/a",
            pct((rr > 0).mean()) if n_closed else "n/a",
            pct(status[closed].str.contains("target").mean()) if n_closed else "n/a",
            pct(status[closed].str.contains("stopped").mean()) if n_closed else "n/a",
            pct(status[closed].str.contains("timed").mean()) if n_closed else "n/a",
            round(float(rr.mean()), 2) if n_closed else "n/a",
            round(float(rr.sum() / len(mat) * 100), 1) if len(mat) else "n/a",
            f"{risk.median():.1%}" if trig.any() else "n/a",
        ]

    labels = ["Entered within window (of matured alerts)", "Win rate (R > 0)", f"{cfg['target_r']:.0f}R target hit",
              "Stopped out", f"Timed exit ({cfg['hold_bars']}d)", "Average R per trade",
              "Total R per 100 matured alerts (fair head-to-head)", "Median entry risk (entry to stop)"]
    a_col = col(mat["Status"], mat["R Multiple"], mat["Entry"], None)
    b_col = col(mat["Alt Status"], mat["Alt R Multiple"], mat["Alt Entry"], None)
    rows = [("Alerts in window (all)", len(alerts), len(alerts)),
            (f"Matured alerts (>= {need} sessions old) - used below", len(mat), len(mat)),
            ("Excluded: too recent to judge", len(alerts) - len(mat), len(alerts) - len(mat))]
    rows += [(lab, x, y) for lab, x, y in zip(labels, a_col, b_col)]
    rows.append(("Caveat", "Daily bars only, no slippage/fees; stop is checked before target on the same bar. "
                           "close_confirm enters at that day's close and starts stop/target checks the next bar.", ""))
    return pd.DataFrame(rows, columns=["Metric", f"{prim} (selected)", f"{alt} (comparison)"])


def timing_table():
    rows = []
    for name, sym, ts, trig, last, tf, fam, tags in REFERENCE:
        t = pd.Timestamp(ts)
        tm = t.time()
        if tm < dtime(9, 15):
            phase = "Pre-market"
        elif tm < dtime(15, 30):
            phase = "Intraday (market open)"
        else:
            phase = "Post-market (EOD)"
        rows.append({"Stock": name, "Yahoo Ticker": sym + ".NS", "Alert Date": t.date(),
                     "Weekday": t.day_name(), "Alert Time (IST)": t.strftime("%H:%M"), "Session": phase,
                     "Timeframe": tf, "Family": fam, "Trigger on chart": trig, "Price on chart": last,
                     "Chart tags": tags})
    return pd.DataFrame(rows)


def write_excel(path, alerts, latest, summary, calib, timing, cfg, R, data_check=None):
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        if data_check:
            pd.DataFrame(list(data_check.items()), columns=["Check", "Value"]).to_excel(
                xw, sheet_name="Data_Check", index=False)
        cols = [c for c in alerts.columns if not c.startswith("_") and c != "Latest Session"]
        (latest[cols] if len(latest) else pd.DataFrame(columns=cols)).to_excel(xw, sheet_name="Latest_Watchlist", index=False)
        (alerts[cols] if len(alerts) else pd.DataFrame(columns=cols)).to_excel(xw, sheet_name="All_Alerts", index=False)
        summary.to_excel(xw, sheet_name="Outcome_Summary", index=False)
        if calib is not None:
            calib.to_excel(xw, sheet_name="Reference_Calibration", index=False)
        timing.to_excel(xw, sheet_name="Reference_Alert_Timing", index=False)
        rules = pd.DataFrame([(k, str(v)) for k, v in {**{"CFG." + k: v for k, v in cfg.items()},
                                                       **{"RULES." + k: v for k, v in R.items()}}.items()],
                             columns=["Parameter", "Value"])
        rules.to_excel(xw, sheet_name="Rules_Used", index=False)
        for name, ws in xw.sheets.items():
            _style_sheet(ws, freeze="C2" if name in ("Latest_Watchlist", "All_Alerts") else "A2")


# %% ======================================================================
# 8. CALIBRATION  (do the rules catch the reference alerts?)
# ==========================================================================
def calibrate(prices, cfg, R):
    rows = []
    for name, sym, ts, ref_trig, last, tf, fam, tags in REFERENCE:
        t = pd.Timestamp(ts)
        d = t.normalize() if t.time() >= dtime(15, 30) else t.normalize() - pd.Timedelta(days=1)
        row = {"Stock": name, "Yahoo Ticker": sym + ".NS", "Alert (IST)": ts, "Timeframe": tf, "Family": fam,
               "Chart tags": tags, "Chart trigger": ref_trig}
        df = prices.get(sym)
        if df is None or len(df) < 30:
            row.update(Result="NO DATA from Yahoo (check ticker)")
            rows.append(row)
            continue
        f = compute_features(df, R, cfg)
        pos = int(df.index.searchsorted(d, side="right")) - 1
        if pos < 0:
            row.update(Result="Alert date before first bar")
            rows.append(row)
            continue
        win = range(max(0, pos - 2), min(len(df) - 1, pos + 2) + 1)
        exact = bool(f["signal"].iloc[pos])
        near = any(bool(f["signal"].iloc[k]) for k in win)
        fr = f.iloc[pos]
        failed = [g for g in GATES if not bool(fr["g_" + g])]
        our_trig = float(df["High"].iloc[pos])
        if fam == "support":
            res = "Not covered (1h support-retest setup)"
        elif exact:
            res = "MATCH on alert day"
        elif near:
            res = "MATCH within +/-2 bars"
        elif not failed:
            res = f"Hard rules pass, score {int(fr['score'])} < min_score {cfg['min_score']}"
        else:
            res = "MISSED - failed: " + ", ".join(failed)
        row.update({
            "Evaluated bar": df.index[pos].date(), "Result": res,
            "Our trigger (day high)": round(our_trig, 2),
            "Trigger matches chart?": (abs(our_trig / ref_trig - 1) <= 0.003) if ref_trig else None,
            "Score (0-6)": int(fr["score"]), "Failed gates": ", ".join(failed),
            "Best 20d Run-up (%)": float(fr["thrust"]), "Run-up from 120d Low (%)": float(fr["runup_low"]),
            "Off 60d High (%)": float(fr["off_high"]), "Dist to EMA10 (%)": float(fr["dist_fast"]),
            "Ribbon Spread (%)": float(fr["ribbon"]), "10d Close Range (%)": float(fr["close_range"]),
            "Trigger Gap (%)": float(fr["trigger_gap"]),
            "Vol Ratio (5d med / 50d avg)": round(float(fr["vol_ratio"]), 2) if pd.notna(fr["vol_ratio"]) else np.nan,
            "Inside bar": bool(fr["ib"]), "Narrow range": bool(fr["nr"]),
        })
        rows.append(row)
    return pd.DataFrame(rows)


# %% ======================================================================
# 9. MAIN
# ==========================================================================
def main(cfg=None, R=None):
    cfg = {**CFG, **(cfg or {})}
    R = {**RULES, **(R or {})}
    t0 = time.time()
    out_dir = cfg["out_dir"]
    shutil.rmtree(out_dir, ignore_errors=True)
    os.makedirs(os.path.join(out_dir, "charts"), exist_ok=True)

    print("=" * 78, "\n1/5  Loading NSE universe")
    uni = load_universe(cfg)
    if cfg["universe_limit"]:
        uni = uni.head(int(cfg["universe_limit"]))
    names = dict(zip(uni["symbol"], uni["name"]))
    print(f"     {len(uni)} symbols")

    print("2/5  Downloading daily prices from Yahoo Finance")
    prices = fetch_prices(list(uni["symbol"]), cfg["history_period"], cfg["chunk_size"])
    prices = {k: drop_incomplete_today(v) for k, v in prices.items()}
    prices = {k: v for k, v in prices.items() if len(v) > 0}
    if not prices:
        raise SystemExit("No price data downloaded. Check your internet / Yahoo access and try again.")
    data_check = fill_missing_sessions(prices, cfg)
    print("     data check: " + "; ".join(f"{k} {v}" for k, v in data_check.items()))
    last_date = max(d.index[-1] for d in prices.values()).date()
    print(f"     {len(prices)} stocks with data, latest completed session: {last_date}")

    print("3/5  Scanning for the pattern")
    alerts, keep = scan(prices, names, cfg, R)
    latest = alerts[alerts["Latest Session"]] if len(alerts) else alerts
    print(f"     {len(alerts)} alerts in the last {cfg['lookback_days']} sessions; "
          f"{len(latest)} on the latest session ({last_date})")

    print("4/5  Rendering charts")
    if len(alerts):
        order = alerts.assign(_latest=alerts["Latest Session"].astype(int)).sort_values(
            ["_latest", "Alert Date", "Score (0-6)"], ascending=[False, False, False])
        for n_, (ridx, rec) in enumerate(order.head(cfg["max_save"]).iterrows()):
            sym = rec["Symbol"]
            df, f = keep[sym]
            fn = f"{pd.Timestamp(rec['Alert Date']):%Y%m%d}_{sym}_s{rec['Score (0-6)']}.png".replace("&", "and")
            alerts.at[ridx, "Chart File"] = f"charts/{fn}"
            if n_ == 0:
                print(f"     (newest alerts first; showing up to {cfg['max_display']} inline)")
            plot_alert(sym, names.get(sym, sym), df, f, int(rec["_i"]), rec, cfg, R,
                       path=os.path.join(out_dir, "charts", fn),
                       show=cfg["show_inline"] and n_ < cfg["max_display"])
        latest = alerts[alerts["Latest Session"]]

    calib = None
    if cfg["run_calibration"]:
        print(f"5/5  Calibrating against the {len(REFERENCE)} reference alerts")
        refs = sorted({r[1] for r in REFERENCE})
        ref_prices = {k: prices[k] for k in refs if k in prices}
        missing = [s for s in refs if s not in ref_prices]
        if missing:
            more = fetch_prices(missing, cfg["history_period"], cfg["chunk_size"], pause=0.2)
            ref_prices.update({k: drop_incomplete_today(v) for k, v in more.items()})
        calib = calibrate(ref_prices, cfg, R)
        scored = calib[calib["Family"] != "support"]
        hits = scored["Result"].str.startswith("MATCH").sum()
        print(f"     scanner matches {hits}/{len(scored)} comparable reference alerts")
        with pd.option_context("display.max_colwidth", 70, "display.width", 200):
            print(calib[["Stock", "Alert (IST)", "Result"]].to_string(index=False))

    xlsx = os.path.join(out_dir, f"alerts_{last_date}.xlsx")
    write_excel(xlsx, alerts, latest, summarize_outcomes(alerts, cfg), calib, timing_table(), cfg, R, data_check)
    with open(os.path.join(out_dir, "rules_used.json"), "w") as fh:
        json.dump({"CFG": {k: (v if not isinstance(v, set) else list(v)) for k, v in cfg.items()}, "RULES": R},
                  fh, indent=2, default=str)

    zip_path = f"ChartsParadise_Scan_{last_date}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(out_dir):
            for fn in files:
                full = os.path.join(root, fn)
                z.write(full, os.path.relpath(full, out_dir))
    print(f"\nDone in {time.time() - t0:.0f}s  ->  {zip_path}  ({os.path.getsize(zip_path) / 1e6:.1f} MB)")
    print("Tip: 11 of the first 19 reference alerts were posted AFTER the close (15:30 IST), mostly 18:00-21:35. "
          "Run this after ~16:30 IST so Yahoo's daily bar is final; entry needs a CLOSE above the trigger.")
    if cfg["download_zip"] and IN_COLAB:
        from google.colab import files as colab_files
        colab_files.download(zip_path)
    return alerts, calib, zip_path


if __name__ == "__main__":
    alerts_df, calib_df, zip_file = main()
