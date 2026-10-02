#!/usr/bin/env python3
"""
EVI Dashboard build engine — PMS AIF LEAPS(TM)
Reads the EVI 2025 daily dataset (Google Sheet CSV export or local CSV),
computes the 11-factor Equity Valuation Index, EPS growth analytics,
and renders a self-contained docs/index.html for GitHub Pages.

Data source resolution order:
  1. env SHEET_CSV_URL  (published Google Sheet CSV export URL)
  2. data/evi_data.csv  (local fallback / seed)
"""
import os, sys, json, csv, io, math, urllib.request
from datetime import datetime, timedelta

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_CSV = os.path.join(ROOT, "data", "evi_data.csv")
TEMPLATE = os.path.join(ROOT, "scripts", "template.html")
OUT_HTML = os.path.join(ROOT, "docs", "index.html")

# Canonical column order = columns A..AB of the "EVI 2025" sheet
COLS = ["sno","date","mcap_inr_cr","mcap_usd","usdinr","mcap_usd_tn","mcap_gdp",
        "beer","n50_ey","n50_eps","in10y","nifty50","pb","yield_gap","n50_pe",
        "us10y","in_us_spread","dxy","preity","tbill91","mid150","mid_pe",
        "mid_eps","mid_ey","small250","small_pe","small_eps","small_ey"]

NUMERIC = COLS[2:]

# ---- 11 prime factors: key, label, unit, direction (+1: high = expensive, -1: high = cheap)
# NOTE: Market Cap to GDP (both ₹ and $ rows) reads the sheet's own `mcap_gdp` column (col 7)
# directly — no recomputation. Because a MCap/GDP ratio is currency-neutral, both rows show
# the same maintained figure.
FACTORS = [
    ("n50_pe",        "P/E Ratio — Nifty 50",                "x",   +1),
    ("pb",            "P/B Ratio — Nifty 50",                "x",   +1),
    ("mcapgdp_inr",   "Market Cap to GDP — ₹ terms",         "%",   +1),
    ("mcapgdp_usd",   "Market Cap to GDP — $ terms",         "%",   +1),
    ("n50_ey",        "Earnings Yield — Nifty 50",           "%",   -1),
    ("in10y",         "India 10-Year G-Sec Yield",           "%",   +1),
    ("beer",          "BEER Ratio (10Y ÷ Earnings Yield)",   "x",   +1),
    ("tbill91",       "91-Day T-Bill Yield",                 "%",   +1),
    ("preity",        "PREITY Ratio (P/E × 91-Day)",         "",    +1),
    ("yield_gap",     "Yield Gap (10Y − Earnings Yield)",    "pp",  +1),
    ("in_us_spread",  "India 10Y − US 10Y Spread",           "pp",  -1),
]

BANDS = [
    (0, 20,  "Deep Value",        "#1F6B4E"),
    (20, 40, "Value",             "#5C9A6F"),
    (40, 60, "Fair Value",        "#C8A24B"),
    (60, 80, "Expensive",         "#C2622E"),
    (80, 101,"Extreme Expensive", "#962B25"),
]

def parse_date(s):
    s = str(s).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%m-%Y", "%d-%b-%Y",
                "%Y-%m-%d %H:%M:%S", "%d/%m/%y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    raise ValueError(f"Unparseable date: {s!r}")

def to_float(v):
    if v is None: return None
    s = str(v).replace(",", "").replace("₹", "").strip()
    if s in ("", "-", "#N/A", "N/A", "NA", "#DIV/0!", "#REF!", "#VALUE!"): return None
    try: return float(s)
    except ValueError: return None

def load_rows():
    url = os.environ.get("SHEET_CSV_URL", "").strip()
    if url:
        print(f"Fetching Google Sheet: {url[:80]}...")
        with urllib.request.urlopen(url, timeout=60) as r:
            text = r.read().decode("utf-8-sig")
    else:
        print(f"Reading local CSV: {DATA_CSV}")
        with open(DATA_CSV, encoding="utf-8-sig") as f:
            text = f.read()
    reader = csv.reader(io.StringIO(text))
    raw = list(reader)
    start = 0
    for i, row in enumerate(raw[:5]):
        if any("date" in str(c).lower() for c in row):
            start = i + 1
            break
    rows = []
    for row in raw[start:]:
        if len(row) < 20 or not str(row[1]).strip():
            continue
        try:
            d = parse_date(row[1])
        except ValueError:
            continue
        rec = {"date": d}
        for j, key in enumerate(COLS[2:], start=2):
            rec[key] = to_float(row[j]) if j < len(row) else None
        rows.append(rec)
    rows.sort(key=lambda r: r["date"])
    dedup = {}
    for r in rows: dedup[r["date"]] = r
    rows = [dedup[k] for k in sorted(dedup)]
    ffill_cols = [c for c in NUMERIC if c not in ("yield_gap", "in_us_spread")]
    prev = {}
    for r in rows:
        for c in NUMERIC:
            v = r[c]
            bad = v is None or (c in ffill_cols and v == 0 and c not in
                  ("mid150","mid_pe","mid_eps","mid_ey","small250","small_pe","small_eps","small_ey"))
            if bad and c in prev:
                r[c] = prev[c]
            elif v is not None:
                prev[c] = v
    return rows

def percentile_ranks(values):
    pairs = sorted((v, i) for i, v in enumerate(values))
    n = len(values)
    out = [0.0] * n
    k = 0
    while k < n:
        j = k
        while j + 1 < n and pairs[j + 1][0] == pairs[k][0]:
            j += 1
        avg_rank = (k + j) / 2.0
        pct = 100.0 * avg_rank / (n - 1) if n > 1 else 50.0
        for m in range(k, j + 1):
            out[pairs[m][1]] = pct
        k = j + 1
    return out

def band_of(score):
    for lo, hi, name, color in BANDS:
        if lo <= score < hi:
            return name, color
    return BANDS[-1][2], BANDS[-1][3]

def cagr(latest, past, years):
    if not latest or not past or past <= 0 or latest <= 0:
        return None
    return (latest / past) ** (1.0 / years) - 1.0

def nearest_on_or_before(rows, target):
    lo, hi, best = 0, len(rows) - 1, None
    while lo <= hi:
        mid = (lo + hi) // 2
        if rows[mid]["date"] <= target:
            best = mid; lo = mid + 1
        else:
            hi = mid - 1
    return best

def main():
    rows = load_rows()
    n = len(rows)
    print(f"{n} daily observations | {rows[0]['date']:%d-%b-%Y} → {rows[-1]['date']:%d-%b-%Y}")

    # ---- Market Cap to GDP: read the sheet's own `mcap_gdp` column (col 7) directly.
    #      Both ₹ and $ rows use this single maintained figure (ratio is currency-neutral).
    for r in rows:
        r["mcapgdp_inr"] = r["mcap_gdp"]
        r["mcapgdp_usd"] = r["mcap_gdp"]

    # ---- Composite EVI: equal-weighted direction-adjusted percentile ranks
    pct = {}
    for key, label, unit, direction in FACTORS:
        vals = [r[key] for r in rows]
        ranks = percentile_ranks(vals)
        if direction < 0:
            ranks = [100.0 - x for x in ranks]
        pct[key] = ranks
    evi = [sum(pct[k][i] for k, *_ in FACTORS) / len(FACTORS) for i in range(n)]
    evi_pct = percentile_ranks(evi)

    evi_smooth = []
    for i in range(n):
        w = evi[max(0, i - 29):i + 1]
        evi_smooth.append(sum(w) / len(w))

    latest = rows[-1]
    cur_evi = evi[-1]
    cur_band, cur_color = band_of(cur_evi)
    evi_30d_ago = evi[max(0, n - 31)]

    band_share = {b[2]: 0 for b in BANDS}
    for v in evi:
        band_share[band_of(v)[0]] += 1
    band_share = {k: round(100.0 * v / n, 1) for k, v in band_share.items()}

    factor_meta = []
    for key, label, unit, direction in FACTORS:
        vals = [r[key] for r in rows if r[key] is not None]
        cur = latest[key]
        sv = sorted(vals)
        med = sv[len(sv)//2]
        factor_meta.append({
            "key": key, "label": label, "unit": unit,
            "direction": "high = expensive" if direction > 0 else "high = cheap",
            "current": round(cur, 2),
            "median": round(med, 2),
            "min": round(sv[0], 2), "max": round(sv[-1], 2),
            "pctl": round(pct[key][-1], 1),
        })

    eps_defs = [("n50_eps", "Nifty 50"), ("mid_eps", "Nifty Midcap 150"), ("small_eps", "Nifty Smallcap 250")]
    eps_growth = []
    for key, name in eps_defs:
        latest_eps = latest[key]
        g = {"index": name, "latest": round(latest_eps, 1) if latest_eps else None}
        for yrs in (1, 3, 5, 7):
            idx = nearest_on_or_before(rows, latest["date"] - timedelta(days=int(365.25 * yrs)))
            past = rows[idx][key] if idx is not None else None
            c = cagr(latest_eps, past, yrs)
            g[f"y{yrs}"] = round(c * 100, 2) if c is not None else None
        eps_growth.append(g)

    def series(key, dec=5, start=0):
        pts = []
        for i in range(start, n):
            if (i - start) % dec == 0 or i == n - 1:
                v = rows[i][key]
                if v is not None:
                    pts.append([rows[i]["date"].strftime("%Y-%m-%d"), round(v, 3)])
        return pts

    mid_start = next((i for i, r in enumerate(rows) if (r["mid_eps"] or 0) > 0), 0)

    payload = {
        "asof": latest["date"].strftime("%d %b %Y"),
        "built": datetime.now().strftime("%d %b %Y, %H:%M UTC"),
        "n_obs": n,
        "span": f"{rows[0]['date']:%b %Y} – {rows[-1]['date']:%b %Y}",
        "evi": {
            "score": round(cur_evi, 1),
            "band": cur_band, "color": cur_color,
            "delta30": round(cur_evi - evi_30d_ago, 1),
            "pctl": round(evi_pct[-1], 1),
            "band_share": band_share,
            "series": [[rows[i]["date"].strftime("%Y-%m-%d"), round(evi[i], 2)]
                       for i in range(n) if i % 5 == 0 or i == n - 1],
            "smooth": [[rows[i]["date"].strftime("%Y-%m-%d"), round(evi_smooth[i], 2)]
                       for i in range(n) if i % 5 == 0 or i == n - 1],
        },
        "ribbon": [round(evi[i], 1) for i in range(0, n, max(1, n // 900))],
        "snapshot": {
            "nifty": latest["nifty50"], "pe": latest["n50_pe"], "pb": latest["pb"],
            "mcap_tn": round(latest["mcap_usd_tn"], 2), "in10y": latest["in10y"],
            "mid_pe": latest["mid_pe"], "small_pe": latest["small_pe"],
        },
        "factors": factor_meta,
        "factor_series": {key: series(key) for key, *_ in FACTORS},
        "eps_series": {
            "n50_eps": series("n50_eps"),
            "mid_eps": series("mid_eps", start=mid_start),
            "small_eps": series("small_eps", start=mid_start),
        },
        "eps_growth": eps_growth,
        "bands": [{"lo": b[0], "hi": min(b[1], 100), "name": b[2], "color": b[3]} for b in BANDS],
    }

    with open(TEMPLATE, encoding="utf-8") as f:
        html = f.read()
    vendor = os.path.join(ROOT, "scripts", "vendor")
    with open(os.path.join(vendor, "chart.umd.js"), encoding="utf-8") as f:
        html = html.replace("/*__CHARTJS__*/", f.read())
    with open(os.path.join(vendor, "chartjs-adapter.min.js"), encoding="utf-8") as f:
        html = html.replace("/*__ADAPTER__*/", f.read())
    html = html.replace("/*__DATA__*/null", json.dumps(payload, separators=(",", ":")))
    os.makedirs(os.path.dirname(OUT_HTML), exist_ok=True)
    with open(OUT_HTML, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"EVI = {cur_evi:.1f}  →  {cur_band}")
    for fm in factor_meta:
        print(f"  {fm['label']:<42} {fm['current']:>10}   pctl {fm['pctl']:>5}")
    print(f"Wrote {OUT_HTML} ({os.path.getsize(OUT_HTML)//1024} KB)")

if __name__ == "__main__":
    main()
