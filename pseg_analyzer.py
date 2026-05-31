#!/usr/bin/env python3
"""
PSEG Long Island consumption analyzer.

Reads a downloaded interval-consumption CSV (consumption meter + optional
generation meter whose number ends in 'g'), splits it into billing cycles, and
writes an .xlsx dashboard estimating the full bill - every line item that
appears on a real PSEG bill - under four rate plans: 580, 180, 194 and 195.

Rates are DATE-MATCHED: a monthly rate schedule (delivery, supply and riders)
is kept in a local cache and shown on a "Rates" tab. For each billing cycle the
applicable rate is computed by day-weighting the schedule across the cycle's
days, so a cycle that straddles a rate change (e.g. the Jan 1 increase) or a
season boundary gets the correct blended rate - the way PSEG actually bills.

Historical rates are available on PSEG LI back to June 2025:
  https://www.psegliny.com/aboutpseglongisland/ratesandtariffs/rateinformation
The scraper refreshes/extends the cache from that page when the data covers
months not yet downloaded. Dates before June 2025 fall back to the earliest
known rates and a warning is printed.

Usage:
    python pseg_analyzer.py --cons_file Usage.csv
    python pseg_analyzer.py --cons_file Usage.csv --split 2025-12-15 2026-01-16 ...
    python pseg_analyzer.py --cons_file Usage.csv --demand 16 --out report.xlsx [--refresh]
"""

import argparse
import copy
import datetime as dt
import json
import math
import os
import re
import sys
from typing import Any, cast

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

CACHE_FILE = "pseg_rates_cache.json"
RATE_INFO_URL = "https://www.psegliny.com/aboutpseglongisland/ratesandtariffs/rateinformation"
HIST_FLOOR = (2025, 6)   # PSEG publishes historical rates back to June 2025

# --------------------------------------------------------------------------- #
# Rate epochs. PSEG raised residential delivery/rider rates on 2026-01-01.     #
# Values confirmed from residential bills + the PSEG rate guide. Where a value #
# is not independently confirmed it is flagged (see month flags) and the       #
# scraper is expected to correct it. Seasons: Summer = Jun-Sep, Winter = rest. #
# --------------------------------------------------------------------------- #
EPOCH_PRE: dict[str, Any] = {            # in effect through 2025-12-31
    "basic_day": 0.5400, "der_rate": 0.005554, "mfc_rate": 0.001886,
    "cbc_rate": 0.0330, "dsa_rate": -0.0006, "rda_rate": -0.0003,
    "nysa_pct": 0.0027, "pilots_pct": 0.0215, "suffolk_pct": 0.0231, "salestax_pct": 0.025,
    "delivery": {
        "580": {"t1_S": 0.1064, "t1_W": 0.1021, "t2_S": 0.1348, "t2_W": 0.1021,
                "exc_S": 0.1348, "exc_W": 0.0561, "psc_mult": 1.00},
        "180": {"t1_S": 0.1064, "t1_W": 0.1021, "exc_S": 0.1348, "exc_W": 0.1021, "psc_mult": 1.00},
        "194": {"off_S": 0.1093, "off_W": 0.0929, "peak_S": 0.2217, "peak_W": 0.1885,
                "mult_off": 0.83, "mult_peak_S": 1.9505, "mult_peak_W": 2.1340},
        "195": {"so_S": 0.0452, "so_W": 0.0450, "off_S": 0.1388, "off_W": 0.0929,
                "peak_S": 0.2979, "peak_W": 0.2440, "mult_so": 0.60, "mult_off": 1.00,
                "mult_peak_S": 1.7419, "mult_peak_W": 1.9688},
    },
}
EPOCH_2026: dict[str, Any] = {           # in effect from 2026-01-01
    "basic_day": 0.5600, "der_rate": 0.006014, "mfc_rate": 0.001763,
    "cbc_rate": 0.0372, "dsa_rate": -0.0006, "rda_rate": -0.0003,
    "nysa_pct": 0.0027, "pilots_pct": 0.0215, "suffolk_pct": 0.0231, "salestax_pct": 0.025,
    "delivery": {
        "580": {"t1_S": 0.1064, "t1_W": 0.1064, "t2_S": 0.1348, "t2_W": 0.1064,
                "exc_S": 0.1348, "exc_W": 0.0585, "psc_mult": 1.00},
        "180": {"t1_S": 0.1064, "t1_W": 0.1064, "exc_S": 0.1348, "exc_W": 0.1064, "psc_mult": 1.00},
        "194": {"off_S": 0.1093, "off_W": 0.0929, "peak_S": 0.2217, "peak_W": 0.1885,
                "mult_off": 0.83, "mult_peak_S": 1.9505, "mult_peak_W": 2.1340},
        "195": {"so_S": 0.0452, "so_W": 0.0450, "off_S": 0.1388, "off_W": 0.0929,
                "peak_S": 0.2979, "peak_W": 0.2440, "mult_so": 0.60, "mult_off": 1.00,
                "mult_peak_S": 1.7419, "mult_peak_W": 1.9688},
    },
}
# Monthly Power Supply Charge ($/kWh) - the time-varying supply base rate.
SEED_PSC = {
    "2025-05": 0.122603, "2025-06": 0.130053, "2025-07": 0.118428, "2025-08": 0.127798,
    "2025-09": 0.130177, "2025-10": 0.129685, "2025-11": 0.127804, "2025-12": 0.125526,
    "2026-01": 0.129777, "2026-02": 0.146948, "2026-03": 0.165666, "2026-04": 0.165927,
    "2026-05": 0.160182,
}
# Delivery/rider values are independently CONFIRMED for these months (from bills).
CONFIRMED = {"2025-10", "2025-11", "2025-12", "2026-01", "2026-02", "2026-03", "2026-04", "2026-05"}

RATE_TITLES = {
    "580": "Rate 580 - Residential, Home Heating",
    "180": "Rate 180 - Residential (standard)",
    "194": "Rate 194 - Time-of-Day, Off-Peak",
    "195": "Rate 195 - Time-of-Day, Super Off-Peak",
}


# --------------------------------------------------------------------------- #
# Rate schedule (monthly) construction + cache + scraping                      #
# --------------------------------------------------------------------------- #
def month_key(y, mo):
    return f"{y}-{mo:02d}"


def build_seed_schedule():
    sched = {}
    for mk, psc in SEED_PSC.items():
        y, mo = int(mk[:4]), int(mk[5:7])
        epoch = EPOCH_2026 if (y, mo) >= (2026, 1) else EPOCH_PRE
        rec = copy.deepcopy(epoch)
        rec["psc"] = psc
        if (y, mo) < HIST_FLOOR:
            rec["flag"] = "stale"            # before published history
        elif mk in CONFIRMED:
            rec["flag"] = "confirmed"
        else:
            rec["flag"] = "estimate"         # PSC exact; delivery/riders estimated
        sched[mk] = rec
    return sched


def load_cache():
    seed = build_seed_schedule()
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE) as f:
                disk = json.load(f)
            sched = disk.get("schedule", {})
            for mk, rec in seed.items():        # backfill seed months not on disk
                sched.setdefault(mk, rec)
            return {"scraped_at": disk.get("scraped_at"), "schedule": sched}
        except Exception as e:
            print(f"  (cache unreadable: {e}; using seed schedule)")
    return {"scraped_at": None, "schedule": seed}


def save_cache(cache):
    with open(CACHE_FILE, "w") as f:
        json.dump(cache, f, indent=2)


def months_in_span(a, b):
    out, y, mo = [], a.year, a.month
    while (y, mo) <= (b.year, b.month):
        out.append(month_key(y, mo))
        mo += 1
        if mo > 12:
            mo, y = 1, y + 1
    return out


def scrape_rate_information(month_keys):
    """Best-effort scrape of the PSEG LI rate-information page. Returns
    {month_key: {'psc':..., optionally delivery/rider fields}}. Network or
    parse failures degrade gracefully to {} (callers keep the seed)."""
    found = {}
    try:
        import urllib.request
        req = urllib.request.Request(RATE_INFO_URL, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=25) as r:
            html = r.read().decode("utf-8", "ignore")
        names = ["january", "february", "march", "april", "may", "june", "july",
                 "august", "september", "october", "november", "december"]
        for i, nm in enumerate(names, 1):
            for mt in re.finditer(rf"{nm}\s+(\d{{4}}).{{0,120}}?\$?\s*(0\.\d{{5,6}})",
                                  html, re.IGNORECASE | re.DOTALL):
                mk = month_key(int(mt.group(1)), i)
                if mk in month_keys:
                    found.setdefault(mk, {})["psc"] = float(mt.group(2))
    except Exception as e:
        print(f"  (scrape failed: {e})")
    return found


def refresh_schedule(cache, start, end, force=False):
    sched = cache["schedule"]
    needed = months_in_span(start, end)
    missing = [m for m in needed if m not in sched]
    target = set(missing) | ({m for m in needed if sched.get(m, {}).get("flag") == "estimate"}
                             if force else set())
    if not target:
        if missing:
            pass
        print("  Rate schedule already covers the data period; no scrape needed.")
        return cache
    print(f"  Fetching rates for {len(target)} month(s) from PSEG LI rate-information...")
    got = scrape_rate_information(set(target))
    for mk, fields in got.items():
        base = sched.get(mk) or copy.deepcopy(EPOCH_2026)
        base.update({k: v for k, v in fields.items()})
        base["flag"] = "scraped"
        sched[mk] = base
    still = [m for m in missing if m not in sched]
    if still:
        last = sorted(sched)[-1]
        for m in still:
            rec = copy.deepcopy(sched[last]); rec["flag"] = "carried"
            sched[m] = rec
        print(f"  Could not obtain {', '.join(still)}; carried forward {last}. Verify on psegliny.com.")
    cache["scraped_at"] = dt.datetime.now().isoformat(timespec="seconds")
    save_cache(cache)
    return cache


def get_month(sched, y, mo, warnings):
    """Return the monthly rate record for (y,mo), with fallback + warning."""
    mk = month_key(y, mo)
    if mk in sched:
        rec = sched[mk]
        if rec.get("flag") == "stale":
            warnings.add(f"Dates in/before {mk} predate published history (back to "
                         f"{HIST_FLOOR[0]}-{HIST_FLOOR[1]:02d}); using earliest known rates.")
        elif rec.get("flag") in ("estimate", "carried"):
            warnings.add(f"Delivery/rider rates for {mk} are estimated (PSC may be exact); "
                         f"run with --refresh once online to scrape exact values.")
        return rec
    keys = sorted(sched)
    fb = keys[0] if (y, mo) < tuple(map(int, keys[0].split("-"))) else keys[-1]
    warnings.add(f"No rate data for {mk}; used nearest available ({fb}).")
    return sched[fb]


# --------------------------------------------------------------------------- #
# Consumption loading + usage primitives + per-cycle effective rates           #
# --------------------------------------------------------------------------- #
def load_consumption(path):
    df = pd.read_csv(path)
    cols = {c.lower(): c for c in df.columns}
    if not all(n in cols for n in ("start", "meter", "kwh")):
        sys.exit(f"CSV must have Start, Meter, kWh columns; found {list(df.columns)}")
    df = df.rename(columns={cols["start"]: "Start", cols["meter"]: "Meter", cols["kwh"]: "kWh"})
    try:
        df["ts"] = pd.to_datetime(df["Start"], format="%m/%d/%Y %I:%M:%S %p")
    except (ValueError, TypeError):
        df["ts"] = pd.to_datetime(df["Start"])
    df["kWh"] = cast("pd.Series", pd.to_numeric(df["kWh"], errors="coerce")).fillna(0.0)
    df["is_gen"] = df["Meter"].astype(str).str.contains(r"#\s*\d+\s*g", case=True, regex=True)
    con = cast("pd.DataFrame", df[~df["is_gen"]])
    gen = cast("pd.DataFrame", df[df["is_gen"]]).copy()
    gen["kWh"] = gen["kWh"].abs()
    imp = con.groupby("ts")["kWh"].sum()
    exp = gen.groupby("ts")["kWh"].sum()
    m = pd.DataFrame({"imp": imp, "exp": exp}).fillna(0.0).sort_index()
    if m.empty:
        sys.exit("No usable rows parsed from the CSV.")
    m["net"] = m["imp"] - m["exp"]
    return m


def federal_holidays(years):
    hol = set()
    def nth(y, mo, wd, n):
        d = dt.date(y, mo, 1) + dt.timedelta((wd - dt.date(y, mo, 1).weekday()) % 7)
        return d + dt.timedelta(weeks=n - 1)
    def last(y, mo, wd):
        d = dt.date(y, mo, 28)
        while d.month == mo:
            d += dt.timedelta(days=1)
        d -= dt.timedelta(days=1)
        return d - dt.timedelta((d.weekday() - wd) % 7)
    for y in years:
        for mo, da in [(1, 1), (6, 19), (7, 4), (11, 11), (12, 25)]:
            hol.add(dt.date(y, mo, da))
        hol.add(nth(y, 1, 0, 3)); hol.add(nth(y, 2, 0, 3)); hol.add(last(y, 5, 0))
        hol.add(nth(y, 9, 0, 1)); hol.add(nth(y, 10, 0, 2)); hol.add(nth(y, 11, 3, 4))
    return hol


def build_cycles(start, end, splits):
    if not splits:
        return [(start, end)]
    sp = [d for d in sorted(set(splits)) if start <= d < end]
    cycles, cur = [], start
    for d in sp:
        cycles.append((cur, d)); cur = d + dt.timedelta(days=1)
    cycles.append((cur, end))
    return cycles


def season_of(month):
    return "S" if month in (6, 7, 8, 9) else "W"


def effective_rates(a, b, sched, warnings):
    """Day-weight every rate parameter across the cycle [a,b], picking each
    day's month and season. Returns a flat dict of blended values."""
    keys_shared = ["basic_day", "der_rate", "mfc_rate", "cbc_rate", "dsa_rate", "rda_rate",
                   "nysa_pct", "pilots_pct", "suffolk_pct", "salestax_pct", "psc"]
    delivery_spec = {
        "580": ["t1", "t2", "exc"], "180": ["t1", "exc"],
        "194": ["off", "peak"], "195": ["so", "off", "peak"],
    }
    mult_spec = {  # (output key, plan, field, seasonal?)
        "m194_off": ("194", "mult_off", False), "m194_peak": ("194", "mult_peak", True),
        "m195_so": ("195", "mult_so", False), "m195_off": ("195", "mult_off", False),
        "m195_peak": ("195", "mult_peak", True),
    }
    acc = {k: 0.0 for k in keys_shared}
    for plan, fields in delivery_spec.items():
        for f in fields:
            acc[f"d{plan}_{f}"] = 0.0
    for k in mult_spec:
        acc[k] = 0.0
    n = 0
    d = a
    while d <= b:
        rec = get_month(sched, d.year, d.month, warnings)
        s = season_of(d.month)
        for k in keys_shared:
            acc[k] += rec[k]
        for plan, fields in delivery_spec.items():
            for f in fields:
                acc[f"d{plan}_{f}"] += rec["delivery"][plan][f"{f}_{s}"]
        for ok, (plan, field, seasonal) in mult_spec.items():
            key = f"{field}_{s}" if seasonal else field
            acc[ok] += rec["delivery"][plan][key]
        n += 1
        d += dt.timedelta(days=1)
    return {k: v / n for k, v in acc.items()}


def cycle_primitives(m, cycles, holidays, sched, warnings):
    idx = cast("pd.DatetimeIndex", m.index)
    dtacc = idx.to_series().dt
    hour = dtacc.hour.to_numpy()
    dow = dtacc.dayofweek.to_numpy()
    dates = dtacc.date.to_numpy()
    is_hol = np.array([d in holidays for d in dates])
    is_peak = (dow < 5) & (~is_hol) & (hour >= 15) & (hour < 19)
    is_so = (hour >= 22) | (hour < 6)
    m = m.copy(); m["peak"] = is_peak; m["so"] = is_so
    rows = []
    for (a, b) in cycles:
        sel = (dates >= a) & (dates <= b)
        g = cast("pd.DataFrame", m[sel])
        if g.empty:
            continue
        net = float(g["net"].to_numpy().sum())
        peak = float(g.loc[g["peak"], "net"].to_numpy().sum())
        so = float(g.loc[g["so"], "net"].to_numpy().sum())
        eff = effective_rates(a, b, sched, warnings)
        rows.append(dict(start=a, end=b, days=(b - a).days + 1, net=net, peak=peak, so=so,
                         off195=net - peak - so, off194=net - peak,
                         demand=math.ceil(float(g["imp"].to_numpy().max()) * 4), eff=eff))
    return rows


# --------------------------------------------------------------------------- #
# Workbook                                                                     #
# --------------------------------------------------------------------------- #
FONT = "Arial"
BLUE, BLACK, GREEN, WHITE = "0000FF", "000000", "008000", "FFFFFF"
HEAD = PatternFill("solid", fgColor="0F6E56")
SUB = PatternFill("solid", fgColor="E1F5EE")
TOT = PatternFill("solid", fgColor="F1EFE8")
YEL = PatternFill("solid", fgColor="FFF6CC")
MONEY = '$#,##0.00;($#,##0.00)'
KWH = '#,##0'
RFMT = '0.000000'
EFF_KEYS = ["basic_day", "der_rate", "mfc_rate", "cbc_rate", "dsa_rate", "rda_rate",
            "nysa_pct", "pilots_pct", "suffolk_pct", "salestax_pct", "psc",
            "d580_t1", "d580_t2", "d580_exc", "d180_t1", "d180_exc",
            "d194_off", "d194_peak", "m194_off", "m194_peak",
            "d195_so", "d195_off", "d195_peak", "m195_so", "m195_off", "m195_peak"]
EFF_LABELS = {
    "basic_day": "Basic service ($/day)", "der_rate": "DER ($/kWh)", "mfc_rate": "MFC ($/kWh)",
    "cbc_rate": "CBC ($/kW/day)", "dsa_rate": "Delivery Svc Adj ($/kWh)",
    "rda_rate": "Rev Decoupling Adj ($/kWh)", "nysa_pct": "NY State Assmt (frac)",
    "pilots_pct": "PILOTs (frac)", "suffolk_pct": "Suffolk Prop Tax (frac)",
    "salestax_pct": "Sales tax (frac)", "psc": "Supply rate ($/kWh)",
    "d580_t1": "580 tier1 ($/kWh)", "d580_t2": "580 tier2 ($/kWh)", "d580_exc": "580 excess ($/kWh)",
    "d180_t1": "180 first250 ($/kWh)", "d180_exc": "180 excess ($/kWh)",
    "d194_off": "194 off-peak ($/kWh)", "d194_peak": "194 peak ($/kWh)",
    "m194_off": "194 off PSC x", "m194_peak": "194 peak PSC x",
    "d195_so": "195 super-off ($/kWh)", "d195_off": "195 off-peak ($/kWh)",
    "d195_peak": "195 peak ($/kWh)", "m195_so": "195 super-off PSC x",
    "m195_off": "195 off PSC x", "m195_peak": "195 peak PSC x",
}


def write_rates_tab(wb, sched):
    ws = wb.create_sheet("Rates")
    ws.cell(1, 1, "PSEG LI MONTHLY RATE SCHEDULE (source)").font = Font(FONT, bold=True, size=13)
    ws.cell(2, 1, "Blue = inputs. 'flag' shows data provenance. Dashboard uses date-weighted "
                  "blends of these per billing cycle.").font = Font(FONT, italic=True, size=9, color="5F5E5A")
    hdr = ["Month", "flag", "Supply PSC", "Basic/day", "DER", "MFC", "CBC",
           "580 t1", "580 t2", "580 exc", "180 first", "180 exc",
           "194 off", "194 peak", "195 so", "195 off", "195 peak"]
    for j, h in enumerate(hdr, 1):
        c = ws.cell(4, j, h); c.font = Font(FONT, bold=True, color=WHITE, size=9); c.fill = HEAD
    r = 5
    for mk in sorted(sched):
        rec = sched[mk]; s = season_of(int(mk[5:7]))
        d = rec["delivery"]
        vals = [mk, rec.get("flag", ""), rec["psc"], rec["basic_day"], rec["der_rate"],
                rec["mfc_rate"], rec["cbc_rate"],
                d["580"][f"t1_{s}"], d["580"][f"t2_{s}"], d["580"][f"exc_{s}"],
                d["180"][f"t1_{s}"], d["180"][f"exc_{s}"],
                d["194"][f"off_{s}"], d["194"][f"peak_{s}"],
                d["195"][f"so_{s}"], d["195"][f"off_{s}"], d["195"][f"peak_{s}"]]
        for j, v in enumerate(vals, 1):
            c = ws.cell(r, j, v)
            c.font = Font(FONT, size=9, color=BLUE if j >= 3 else BLACK)
            if j >= 3:
                c.number_format = RFMT
            if rec.get("flag") in ("estimate", "stale", "carried") and j >= 3:
                c.fill = YEL
        r += 1
    ws.column_dimensions["A"].width = 9
    ws.column_dimensions["B"].width = 10
    for col in range(3, 18):
        ws.column_dimensions[get_column_letter(col)].width = 9
    ws.freeze_panes = "C5"


def write_effective_tab(wb, prims):
    ws = wb.create_sheet("EffectiveRates")
    ws.cell(1, 1, "EFFECTIVE RATES PER CYCLE (date-weighted from Rates schedule)").font = \
        Font(FONT, bold=True, size=12)
    ws.cell(2, 1, "Each value is the day-weighted average of the monthly schedule across the "
                  "cycle - blends rate changes & seasons.").font = Font(FONT, italic=True, size=9, color="5F5E5A")
    ws.cell(4, 1, "Parameter").font = Font(FONT, bold=True, color=WHITE, size=9)
    ws.cell(4, 1).fill = HEAD
    for i, p in enumerate(prims):
        c = ws.cell(4, 2 + i, f"Cycle {i+1}"); c.font = Font(FONT, bold=True, color=WHITE, size=9)
        c.fill = HEAD; c.alignment = Alignment(horizontal="center")
    eref = {}
    for ri, key in enumerate(EFF_KEYS):
        rr = 5 + ri
        ws.cell(rr, 1, EFF_LABELS[key]).font = Font(FONT, size=9)
        for i, p in enumerate(prims):
            c = ws.cell(rr, 2 + i, round(p["eff"][key], 6))
            c.font = Font(FONT, color=BLUE, size=9); c.number_format = RFMT
            eref[(i, key)] = f"'EffectiveRates'!${get_column_letter(2 + i)}${rr}"
    ws.column_dimensions["A"].width = 24
    for i in range(len(prims)):
        ws.column_dimensions[get_column_letter(2 + i)].width = 11
    ws.freeze_panes = "B5"
    return eref


def build_dashboard(wb, prims, eref, meta):
    ws = wb.active; ws.title = "Dashboard"
    ncyc = len(prims)
    cyc_cols = [2 + i for i in range(ncyc)]
    tot_col = 2 + ncyc
    L = get_column_letter
    cl = [L(c) for c in cyc_cols]
    tl = L(tot_col)

    def H(rr, cc, val, fill=None, color=BLACK, bold=False, align="right"):
        c = ws.cell(rr, cc, val); c.font = Font(FONT, bold=bold, color=color, size=10)
        if fill:
            c.fill = fill
        c.alignment = Alignment(horizontal=align)
        return c

    ws.cell(1, 1, "PSEG Long Island - Billing Analysis").font = Font(FONT, bold=True, size=15)
    ws.cell(2, 1, meta).font = Font(FONT, italic=True, size=9, color="5F5E5A")
    row = 4
    H(row, 1, "BILLING CYCLE", HEAD, WHITE, True, "left")
    for i, c in enumerate(cyc_cols):
        H(row, c, f"Cycle {i+1}", HEAD, WHITE, True)
    H(row, tot_col, "TOTAL", HEAD, WHITE, True)
    row += 1
    H(row, 1, "Service period", SUB, align="left")
    for i, c in enumerate(cyc_cols):
        H(row, c, f"{prims[i]['start']:%m/%d/%y}-{prims[i]['end']:%m/%d/%y}", SUB, align="center")
    H(row, tot_col, "All cycles", SUB, align="center")
    row += 1

    ws.cell(row, 1, "USAGE").font = Font(FONT, bold=True, color="0F6E56", size=11)
    row += 1
    urow = {}

    def usage(label, key, agg="sum"):
        nonlocal row
        ws.cell(row, 1, label).font = Font(FONT, size=10)
        for i, c in enumerate(cyc_cols):
            cc = ws.cell(row, c, prims[i][key]); cc.font = Font(FONT, color=BLUE, size=10)
            cc.number_format = KWH; cc.alignment = Alignment(horizontal="right")
        rng = f"{cl[0]}{row}:{cl[-1]}{row}"
        t = ws.cell(row, tot_col, f"=SUM({rng})" if agg == "sum" else f"=MAX({rng})")
        t.font = Font(FONT, bold=True, size=10); t.number_format = KWH
        urow[key] = row; row += 1

    usage("Days", "days"); usage("Net consumption (kWh)", "net")
    usage("  Super-off-peak 10p-6a (kWh)", "so"); usage("  Off-peak (195 basis) (kWh)", "off195")
    usage("  Off-peak (194 basis) (kWh)", "off194"); usage("  On-peak wkdy 3-7p (kWh)", "peak")
    usage("Peak demand (kW)", "demand", "max")

    def u(i, key):
        return f"{cl[i]}{urow[key]}"

    rate_total_rows = {}

    def line(label, fs, bold=False, fill=None):
        nonlocal row
        lc = ws.cell(row, 1, label); lc.font = Font(FONT, size=10, bold=bold)
        if fill:
            lc.fill = fill
        for i, c in enumerate(cyc_cols):
            cc = ws.cell(row, c, fs[i]); cc.font = Font(FONT, color=BLACK, size=10, bold=bold)
            cc.number_format = MONEY
            if fill:
                cc.fill = fill
        t = ws.cell(row, tot_col, f"=SUM({cl[0]}{row}:{cl[-1]}{row})")
        t.font = Font(FONT, bold=True, size=10); t.number_format = MONEY
        if fill:
            t.fill = fill
        this = row; row += 1; return this

    def block(rate):
        nonlocal row
        row += 1
        t = ws.cell(row, 1, RATE_TITLES[rate]); t.font = Font(FONT, bold=True, color=WHITE, size=10)
        for c in cyc_cols + [tot_col]:
            ws.cell(row, c).fill = HEAD
        t.fill = HEAD; row += 1

        def e(i, key):
            return eref[(i, key)]

        deliv, supply = [], []
        for i in range(ncyc):
            net, days = u(i, "net"), u(i, "days")
            if rate == "580":
                t1, t2 = f"(250*{days}/30)", f"(150*{days}/30)"
                deliv.append(f"=MIN({net},{t1})*{e(i,'d580_t1')}"
                             f"+MIN(MAX({net}-{t1},0),{t2})*{e(i,'d580_t2')}"
                             f"+MAX({net}-{t1}-{t2},0)*{e(i,'d580_exc')}")
                supply.append(f"={net}*{e(i,'psc')}")
            elif rate == "180":
                t1 = f"(250*{days}/30)"
                deliv.append(f"=MIN({net},{t1})*{e(i,'d180_t1')}+MAX({net}-{t1},0)*{e(i,'d180_exc')}")
                supply.append(f"={net}*{e(i,'psc')}")
            elif rate == "194":
                deliv.append(f"={u(i,'peak')}*{e(i,'d194_peak')}+{u(i,'off194')}*{e(i,'d194_off')}")
                supply.append(f"={e(i,'psc')}*({u(i,'off194')}*{e(i,'m194_off')}"
                              f"+{u(i,'peak')}*{e(i,'m194_peak')})")
            else:
                deliv.append(f"={u(i,'so')}*{e(i,'d195_so')}+{u(i,'off195')}*{e(i,'d195_off')}"
                             f"+{u(i,'peak')}*{e(i,'d195_peak')}")
                supply.append(f"={e(i,'psc')}*({u(i,'so')}*{e(i,'m195_so')}"
                              f"+{u(i,'off195')}*{e(i,'m195_off')}+{u(i,'peak')}*{e(i,'m195_peak')})")

        br = line("Basic Service", [f"={e(i,'basic_day')}*{u(i,'days')}" for i in range(ncyc)])
        cr = line("Customer Benefit Contribution",
                  [f"={e(i,'cbc_rate')}*{u(i,'demand')}*{u(i,'days')}" for i in range(ncyc)])
        mr = line("Merchant Function Charge", [f"={e(i,'mfc_rate')}*{u(i,'net')}" for i in range(ncyc)])
        dr = line("Delivery (energy)", deliv)
        dsr = line("Delivery & System subtotal",
                   [f"={cl[i]}{br}+{cl[i]}{cr}+{cl[i]}{mr}+{cl[i]}{dr}" for i in range(ncyc)],
                   bold=True, fill=SUB)
        sr = line("Power Supply", supply)
        der = line("DER Charge", [f"={e(i,'der_rate')}*{u(i,'net')}" for i in range(ncyc)])
        dsa = line("Delivery Service Adjustment", [f"={e(i,'dsa_rate')}*{u(i,'net')}" for i in range(ncyc)])
        rda = line("Revenue Decoupling Adjustment", [f"={e(i,'rda_rate')}*{u(i,'net')}" for i in range(ncyc)])
        base = [f"({cl[i]}{dsr}+{cl[i]}{sr})" for i in range(ncyc)]
        nr = line("NY State Assessment", [f"={e(i,'nysa_pct')}*{base[i]}" for i in range(ncyc)])
        pr = line("Revenue-Based PILOTs", [f"={e(i,'pilots_pct')}*{base[i]}" for i in range(ncyc)])
        sf = line("Suffolk Property Tax Adj", [f"={e(i,'suffolk_pct')}*{base[i]}" for i in range(ncyc)])
        pretax = [f"({cl[i]}{dsr}+{cl[i]}{sr}+{cl[i]}{der}+{cl[i]}{dsa}+{cl[i]}{rda}"
                  f"+{cl[i]}{nr}+{cl[i]}{pr}+{cl[i]}{sf})" for i in range(ncyc)]
        sx = line("Sales Tax", [f"={e(i,'salestax_pct')}*{pretax[i]}" for i in range(ncyc)])
        tr = line("TOTAL CHARGES", [f"={pretax[i]}+{cl[i]}{sx}" for i in range(ncyc)], bold=True, fill=TOT)
        rate_total_rows[rate] = tr

    for rate in ["580", "180", "194", "195"]:
        block(rate)

    row += 2
    ws.cell(row, 1, "RATE COMPARISON - total bill ($)").font = Font(FONT, bold=True, color="0F6E56", size=11)
    row += 1
    H(row, 1, "Rate plan", HEAD, WHITE, True, "left")
    for i, c in enumerate(cyc_cols):
        H(row, c, f"Cycle {i+1}", HEAD, WHITE, True)
    H(row, tot_col, "TOTAL", HEAD, WHITE, True)
    row += 1
    cmp = {}
    for rate in ["580", "180", "194", "195"]:
        ws.cell(row, 1, RATE_TITLES[rate]).font = Font(FONT, size=10)
        tr = rate_total_rows[rate]
        for i, c in enumerate(cyc_cols):
            cc = ws.cell(row, c, f"={cl[i]}{tr}"); cc.font = Font(FONT, color=GREEN, size=10)
            cc.number_format = MONEY
        t = ws.cell(row, tot_col, f"={tl}{tr}"); t.font = Font(FONT, bold=True, size=10); t.number_format = MONEY
        cmp[rate] = row; row += 1
    ws.cell(row, 1, "Lowest-cost plan").font = Font(FONT, bold=True, size=10)
    f0, f1 = cmp["580"], cmp["195"]
    for i, c in enumerate(cyc_cols):
        rng = f"{cl[i]}{f0}:{cl[i]}{f1}"
        cc = ws.cell(row, c, f'=INDEX({{"580";"180";"194";"195"}},MATCH(MIN({rng}),{rng},0))')
        cc.font = Font(FONT, bold=True, color="0F6E56", size=10)
        cc.alignment = Alignment(horizontal="center"); cc.fill = SUB
    rng = f"{tl}{f0}:{tl}{f1}"
    t = ws.cell(row, tot_col, f'=INDEX({{"580";"180";"194";"195"}},MATCH(MIN({rng}),{rng},0))')
    t.font = Font(FONT, bold=True, color="0F6E56", size=10)
    t.alignment = Alignment(horizontal="center"); t.fill = SUB
    row += 2
    ws.cell(row, 1, "Date-matched estimate: each cycle uses day-weighted rates from the Rates schedule "
                    "(blends rate changes & seasons). Suffolk Prop Tax applies to Suffolk County only. "
                    "Yellow cells on Rates = estimated/stale; refresh with --refresh while online.").font = \
        Font(FONT, italic=True, size=8, color="5F5E5A")

    ws.column_dimensions["A"].width = 32
    for c in cyc_cols + [tot_col]:
        ws.column_dimensions[L(c)].width = 13
    ws.freeze_panes = "B5"


# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description="Analyze PSEG LI consumption into an xlsx dashboard.")
    ap.add_argument("--cons_file", required=True)
    ap.add_argument("--split", nargs="*", default=[], help="Billing-cycle end dates YYYY-MM-DD ...")
    ap.add_argument("--out", default=None)
    ap.add_argument("--demand", type=float, default=None, help="Override peak demand (kW) for CBC")
    ap.add_argument("--refresh", action="store_true", help="Force re-scrape of estimated months")
    args = ap.parse_args()

    if not os.path.exists(args.cons_file):
        sys.exit(f"File not found: {args.cons_file}")
    try:
        splits = [dt.datetime.strptime(s, "%Y-%m-%d").date() for s in args.split]
    except ValueError:
        sys.exit("Split dates must be YYYY-MM-DD.")

    print("Loading consumption data...")
    m = load_consumption(args.cons_file)
    didx = cast("pd.DatetimeIndex", m.index)
    start = cast("pd.Timestamp", didx.min()).date()
    end = cast("pd.Timestamp", didx.max()).date()
    print(f"  {len(m):,} intervals, {start} to {end}")

    print("Loading & refreshing rate schedule...")
    cache = load_cache()
    cache = refresh_schedule(cache, start, end, force=args.refresh)
    sched = cache["schedule"]

    warnings = set()
    cycles = build_cycles(start, end, splits)
    prims = cycle_primitives(m, cycles, federal_holidays(range(start.year, end.year + 1)), sched, warnings)
    if args.demand is not None:
        for p in prims:
            p["demand"] = args.demand
    print(f"  {len(prims)} billing cycle(s).")
    for w in sorted(warnings):
        print(f"  WARNING: {w}")

    wb = Workbook()
    write_rates_tab(wb, sched)
    eref = write_effective_tab(wb, prims)
    meta = (f"Source: {os.path.basename(args.cons_file)}  |  Data {start} to {end}  |  "
            f"{len(prims)} cycle(s)  |  Generated {dt.date.today()}")
    build_dashboard(wb, prims, eref, meta)
    wb.move_sheet("Dashboard", -(len(wb.sheetnames) - 1))

    out = args.out or os.path.splitext(os.path.basename(args.cons_file))[0] + "_analysis.xlsx"
    wb.save(out)
    print(f"Wrote {out}")


if __name__ == "__main__":
    main()
