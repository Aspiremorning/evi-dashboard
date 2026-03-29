#!/usr/bin/env python3
"""EVI Dashboard Builder v2 — Run: python3 scripts/build.py"""
import json, datetime, statistics, zipfile, re, sys, io
from pathlib import Path

ROOT   = Path(__file__).parent.parent
EXCEL  = ROOT / "data" / "EVI_2025-26.xlsx"
OUTPUT = ROOT / "docs" / "index.html"
WINDOW = 252  # ~1 trading year for YoY EPS growth

def load_workbook_safe(path):
    from openpyxl import load_workbook
    with zipfile.ZipFile(path, "r") as z:
        files = {n: z.read(n) for n in z.namelist()}
    styles = files.get("xl/styles.xml", b"")
    styles = re.sub(
        rb'(<family val=")(\d+)(")',
        lambda m: m.group(0) if int(m.group(2)) <= 14 else m.group(1) + b"2" + m.group(3),
        styles,
    )
    files["xl/styles.xml"] = styles
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in files.items():
            z.writestr(name, data)
    buf.seek(0)
    return load_workbook(buf, data_only=True)

def extract_data(wb):
    ws = wb["EVI 2025"]
    rows = []
    for r in range(2, ws.max_row + 1):
        def v(col):
            val = ws.cell(row=r, column=col).value
            if val is None: return 0.0
            if isinstance(val, (datetime.datetime, datetime.date)): return val
            try: return float(val)
            except: return 0.0
        date_raw = ws.cell(row=r, column=2).value
        if date_raw is None: continue
        if isinstance(date_raw, (datetime.datetime, datetime.date)):
            dt = date_raw.date() if isinstance(date_raw, datetime.datetime) else date_raw
        else:
            try: dt = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(float(date_raw)))
            except: continue
        nifty, pe = v(12), v(15)
        if nifty == 0 or pe == 0: continue
        mc_inr  = v(3)   # Market cap INR crores
        usdinr  = v(5)
        mid_pe  = v(22)
        sc_pe   = v(26)
        rows.append({
            "date":               dt.isoformat(),
            "nifty50":            nifty,
            "pe":                 pe,
            "pb":                 v(13),
            "eps":                v(10),
            "earning_yield":      v(9),
            "india_10yr":         v(11),
            "us_10yr":            v(16),
            "yield_gap":          v(14),
            "usdinr":             usdinr,
            "dollar_index":       v(18),
            "marketcap_inr":      mc_inr,
            "marketcap_trillion": round(mc_inr * 1e7 / usdinr / 1e12, 2) if usdinr > 0 else 0,
            "marketcap_gdp":      v(7),
            "beer":               v(8),
            "preity":             v(19),
            "midcap150":          v(21),
            "midcap_pe":          mid_pe,
            "midcap_eps":         round(v(21) / mid_pe, 2) if mid_pe > 0 else 0,
            "midcap_earn_yield":  v(24),
            "smallcap250":        v(25),
            "smallcap_pe":        sc_pe,
            "smallcap_eps":       round(v(25) / sc_pe, 2) if sc_pe > 0 else 0,
            "smallcap_earn_yield":v(28),
        })

    # Calculate YoY EPS growth (~252 trading days = 1 year)
    for i, r in enumerate(rows):
        if i >= WINDOW:
            prev = rows[i - WINDOW]
            r["nifty_eps_growth"]   = round((r["eps"]         - prev["eps"])         / prev["eps"]         * 100, 2) if prev["eps"]         > 0 else 0
            r["midcap_eps_growth"]  = round((r["midcap_eps"]  - prev["midcap_eps"])  / prev["midcap_eps"]  * 100, 2) if prev["midcap_eps"]  > 0 else 0
            r["smallcap_eps_growth"]= round((r["smallcap_eps"]- prev["smallcap_eps"])/ prev["smallcap_eps"]* 100, 2) if prev["smallcap_eps"]> 0 else 0
        else:
            r["nifty_eps_growth"] = r["midcap_eps_growth"] = r["smallcap_eps_growth"] = 0
    return rows

def build_chart_data(rows):
    n = len(rows)
    thin_idx = list(range(0, max(0, n - 90), 3)) + list(range(max(0, n - 90), n))
    def pick(key):
        return [round(rows[i][key], 4) if isinstance(rows[i][key], float) else rows[i][key] for i in thin_idx]
    keys = ["date","nifty50","pe","pb","earning_yield","india_10yr","us_10yr","yield_gap",
            "usdinr","dollar_index","marketcap_gdp","marketcap_trillion","beer","preity",
            "midcap_earn_yield","smallcap_earn_yield",
            "nifty_eps_growth","midcap_eps_growth","smallcap_eps_growth"]
    return {k: pick(k) for k in keys}

def compute_stats(rows):
    def med(key):
        vals = [r[key] for r in rows if r[key] > 0]
        return round(statistics.median(vals), 4) if vals else 0
    latest = rows[-1]
    return {
        "last_date": latest["date"], "total_rows": len(rows), "date_from": rows[0]["date"],
        "nifty": latest["nifty50"], "pe": latest["pe"], "pb": latest["pb"],
        "earning_yield": latest["earning_yield"], "india_10yr": latest["india_10yr"],
        "us_10yr": latest["us_10yr"], "yield_gap": latest["yield_gap"],
        "usdinr": latest["usdinr"], "dollar_index": latest["dollar_index"],
        "beer": latest["beer"], "marketcap_gdp": latest["marketcap_gdp"],
        "marketcap_trillion": latest["marketcap_trillion"],
        "preity": latest["preity"],
        "midcap_earn_yield": latest["midcap_earn_yield"],
        "smallcap_earn_yield": latest["smallcap_earn_yield"],
        "nifty_eps_growth": latest["nifty_eps_growth"],
        "midcap_eps_growth": latest["midcap_eps_growth"],
        "smallcap_eps_growth": latest["smallcap_eps_growth"],
        "pe_median": med("pe"), "beer_median": med("beer"),
        "mcgdp_median": med("marketcap_gdp"), "yg_median": med("yield_gap"),
    }

HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>EVI Dashboard — __LAST_DATE__</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@400;500;600&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{--bg:#0b0e14;--bg2:#111620;--bg3:#181e2c;--border:#1f2a3c;--accent:#00c8ff;--gold:#f0b429;--green:#00d68f;--red:#ff5c6a;--purple:#c084fc;--orange:#ff8c42;--text:#c8d6e5;--muted:#5a7089;--white:#eef4fb;}
body{font-family:'IBM Plex Sans',sans-serif;background:var(--bg);color:var(--text)}
.hdr{background:var(--bg2);border-bottom:1px solid var(--border);padding:10px 20px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100}
.hdr-t{font-family:'IBM Plex Mono',monospace;font-size:13px;font-weight:600;color:var(--accent);letter-spacing:.1em}
.hdr-s{font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--muted);margin-left:10px}
.hdr-r{font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--muted)}
.hdr-r strong{color:var(--white)}
.wrap{max-width:1400px;margin:0 auto;padding:10px 16px}

/* EVI */
.evi-row{display:flex;gap:16px;align-items:center;background:var(--bg2);border:1px solid var(--border);border-radius:6px;padding:10px 16px;margin-bottom:10px}
.evi-num{font-family:'IBM Plex Mono',monospace;font-size:40px;font-weight:600;line-height:1;background:linear-gradient(135deg,var(--accent),var(--green));-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;min-width:80px}
.evi-txt{flex:1}
.evi-hl{font-size:13px;font-weight:600;color:var(--white);margin-bottom:3px}
.evi-desc{font-size:11px;color:var(--muted);line-height:1.4}
.evi-bar-wrap{min-width:220px}
.evi-bar-bg{height:8px;background:var(--border);border-radius:4px;margin-bottom:4px;position:relative}
.evi-bar-fill{height:100%;border-radius:4px;background:linear-gradient(90deg,var(--green),var(--gold),var(--red))}
.evi-dot{position:absolute;top:-4px;width:16px;height:16px;background:var(--white);border-radius:50%;border:2px solid var(--bg);transform:translateX(-50%);box-shadow:0 0 6px rgba(0,200,255,.5)}
.evi-bar-lbl{display:flex;justify-content:space-between;font-family:'IBM Plex Mono',monospace;font-size:8px;color:var(--muted)}

/* KPI */
.kpi-strip{display:grid;grid-template-columns:repeat(8,1fr);gap:1px;background:var(--border);border:1px solid var(--border);border-radius:6px;overflow:hidden;margin-bottom:10px}
.kpi{background:var(--bg2);padding:8px 10px}
.kpi-l{font-family:'IBM Plex Mono',monospace;font-size:8px;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);margin-bottom:3px}
.kpi-v{font-family:'IBM Plex Mono',monospace;font-size:14px;font-weight:600;color:var(--white);line-height:1}
.kpi-v.g{color:var(--green)}.kpi-v.r{color:var(--red)}.kpi-v.gld{color:var(--gold)}.kpi-v.ac{color:var(--accent)}.kpi-v.pu{color:var(--purple)}
.kpi-c{font-family:'IBM Plex Mono',monospace;font-size:9px;color:var(--muted);margin-top:2px}
.kpi-c.up{color:var(--green)}.kpi-c.dn{color:var(--red)}

/* SEC */
.sec{font-family:'IBM Plex Mono',monospace;font-size:9px;font-weight:600;letter-spacing:.15em;text-transform:uppercase;color:var(--muted);margin:8px 0 6px;padding-bottom:5px;border-bottom:1px solid var(--border)}
.sec span{color:var(--accent);margin-right:6px}

/* GAUGES */
.gauge-row{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:10px}
.gauge-card{background:var(--bg2);border:1px solid var(--border);border-radius:6px;padding:10px 12px}
.gauge-title{font-family:'IBM Plex Mono',monospace;font-size:8px;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);margin-bottom:8px}
.gw{display:flex;align-items:center;gap:10px}
.g-arc{position:relative;width:56px;height:56px;flex-shrink:0}
.g-arc svg{width:100%;height:100%}
.g-num{position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;font-family:'IBM Plex Mono',monospace;font-size:11px;font-weight:600;color:var(--white);line-height:1}
.g-pct{font-size:8px;color:var(--muted);margin-top:1px}
.g-val{font-family:'IBM Plex Mono',monospace;font-size:18px;font-weight:600;color:var(--white)}
.g-med{font-family:'IBM Plex Mono',monospace;font-size:9px;color:var(--muted);margin-top:2px}
.g-zone{display:inline-block;margin-top:5px;padding:2px 6px;border-radius:3px;font-family:'IBM Plex Mono',monospace;font-size:8px;font-weight:600;text-transform:uppercase}
.zc{background:rgba(0,214,143,.15);color:var(--green)}.zf{background:rgba(240,180,41,.15);color:var(--gold)}.zr{background:rgba(255,92,106,.15);color:var(--red)}

/* FILTER */
.filter-row{display:flex;gap:5px;margin-bottom:8px}
.fb{font-family:'IBM Plex Mono',monospace;font-size:9px;padding:4px 10px;border-radius:3px;border:1px solid var(--border);background:transparent;color:var(--muted);cursor:pointer;transition:all .15s}
.fb:hover{border-color:var(--accent);color:var(--accent)}.fb.on{background:var(--accent);border-color:var(--accent);color:var(--bg)}

/* CHART GRIDS */
.g2{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:8px}
.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:8px}
.g4{display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:8px;margin-bottom:8px}
.cc{background:var(--bg2);border:1px solid var(--border);border-radius:6px;padding:10px 12px}
.cc.sp2{grid-column:span 2}.cc.sp3{grid-column:span 3}
.ch{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:6px}
.ct{font-family:'IBM Plex Mono',monospace;font-size:9px;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:var(--text)}
.cv{font-family:'IBM Plex Mono',monospace;font-size:12px;font-weight:600;color:var(--accent)}
.cw{height:130px;position:relative}
.cw canvas{position:absolute;top:0;left:0;width:100%;height:100%}
.lgd{display:flex;gap:12px;margin-top:5px;flex-wrap:wrap}
.li{display:flex;align-items:center;gap:4px;font-family:'IBM Plex Mono',monospace;font-size:8px;color:var(--muted)}
.ld{width:7px;height:7px;border-radius:50%}

/* Growth badge */
.gbadge{display:inline-flex;align-items:center;gap:4px;font-family:'IBM Plex Mono',monospace;font-size:10px;font-weight:600;padding:1px 5px;border-radius:3px}
.gbadge.pos{background:rgba(0,214,143,.15);color:var(--green)}.gbadge.neg{background:rgba(255,92,106,.15);color:var(--red)}

@media(max-width:900px){.kpi-strip{grid-template-columns:repeat(4,1fr)}.gauge-row,.g3,.g4{grid-template-columns:repeat(2,1fr)}.g2{grid-template-columns:1fr}.cc.sp2,.cc.sp3{grid-column:span 1}}
</style>
</head>
<body>
<div class="hdr">
  <div style="display:flex;align-items:baseline">
    <span class="hdr-t">EVI // Equity Valuation Index</span>
    <span class="hdr-s">India Market Monitor</span>
  </div>
  <div class="hdr-r"><strong>__LAST_DATE__</strong> &nbsp;|&nbsp; __DATE_FROM__ → __LAST_DATE__ &nbsp;|&nbsp; __TOTAL_ROWS__ trading days</div>
</div>

<div class="wrap">

  <!-- EVI COMPOSITE -->
  <div class="evi-row">
    <div class="evi-num" id="eviNum">—</div>
    <div class="evi-txt"><div class="evi-hl" id="eviHl">—</div><div class="evi-desc" id="eviDesc">—</div></div>
    <div class="evi-bar-wrap">
      <div class="evi-bar-bg"><div class="evi-bar-fill" style="width:100%"></div><div class="evi-dot" id="eviDot"></div></div>
      <div class="evi-bar-lbl"><span>Cheap</span><span>Fair</span><span>Expensive</span></div>
    </div>
  </div>

  <!-- KPI ROW 1: Core -->
  <div class="kpi-strip">
    <div class="kpi"><div class="kpi-l">Nifty 50</div><div class="kpi-v" id="kN">—</div><div class="kpi-c" id="kNc">—</div></div>
    <div class="kpi"><div class="kpi-l">P/E</div><div class="kpi-v" id="kPE">—</div><div class="kpi-c" id="kPEm">—</div></div>
    <div class="kpi"><div class="kpi-l">P/B</div><div class="kpi-v" id="kPB">—</div><div class="kpi-c">price/book</div></div>
    <div class="kpi"><div class="kpi-l">Earning Yield</div><div class="kpi-v g" id="kEY">—</div><div class="kpi-c">1/PE×100</div></div>
    <div class="kpi"><div class="kpi-l">India 10yr</div><div class="kpi-v gld" id="kI10">—</div><div class="kpi-c" id="kYG">—</div></div>
    <div class="kpi"><div class="kpi-l">US 10yr</div><div class="kpi-v" id="kU10">—</div><div class="kpi-c" id="kSp">—</div></div>
    <div class="kpi"><div class="kpi-l">USD/INR</div><div class="kpi-v" id="kFX">—</div><div class="kpi-c" id="kDXY">—</div></div>
    <div class="kpi"><div class="kpi-l">MC/GDP</div><div class="kpi-v" id="kMC">—</div><div class="kpi-c">Buffett</div></div>
  </div>

  <!-- KPI ROW 2: Extended -->
  <div class="kpi-strip" style="margin-bottom:10px">
    <div class="kpi"><div class="kpi-l">Market Cap</div><div class="kpi-v ac" id="kMCT">—</div><div class="kpi-c">USD Trillion</div></div>
    <div class="kpi"><div class="kpi-l">PREITY Ratio</div><div class="kpi-v pu" id="kPREITY">—</div><div class="kpi-c">Nifty/US10yr</div></div>
    <div class="kpi"><div class="kpi-l">Midcap EY</div><div class="kpi-v g" id="kMEY">—</div><div class="kpi-c">Midcap 150</div></div>
    <div class="kpi"><div class="kpi-l">Smallcap EY</div><div class="kpi-v g" id="kSEY">—</div><div class="kpi-c">SC 250</div></div>
    <div class="kpi"><div class="kpi-l">Nifty EPS Growth</div><div class="kpi-v" id="kNEG">—</div><div class="kpi-c">YoY 1yr</div></div>
    <div class="kpi"><div class="kpi-l">Midcap EPS Growth</div><div class="kpi-v" id="kMEG">—</div><div class="kpi-c">YoY 1yr</div></div>
    <div class="kpi"><div class="kpi-l">Smallcap EPS Growth</div><div class="kpi-v" id="kSEG">—</div><div class="kpi-c">YoY 1yr</div></div>
    <div class="kpi"><div class="kpi-l">BEER Ratio</div><div class="kpi-v gld" id="kBEER">—</div><div class="kpi-c">EY/Bond Yield</div></div>
  </div>

  <!-- FILTER -->
  <div class="filter-row" id="fr">
    <button class="fb" data-r="90">3M</button>
    <button class="fb" data-r="180">6M</button>
    <button class="fb on" data-r="365">1Y</button>
    <button class="fb" data-r="730">2Y</button>
    <button class="fb" data-r="9999">All</button>
  </div>

  <!-- GAUGES -->
  <div class="sec"><span>◈</span>Valuation Gauges — Percentile Rank vs Full History</div>
  <div class="gauge-row">
    <div class="gauge-card"><div class="gauge-title">P/E Ratio</div><div class="gw"><div class="g-arc"><svg viewBox="0 0 56 56"><circle cx="28" cy="28" r="22" fill="none" stroke="#1f2a3c" stroke-width="6" stroke-dasharray="138.2" stroke-linecap="round" transform="rotate(-90 28 28)"/><circle cx="28" cy="28" r="22" fill="none" id="arcPE" stroke="#00c8ff" stroke-width="6" stroke-dasharray="138.2" stroke-dashoffset="138.2" stroke-linecap="round" transform="rotate(-90 28 28)"/></svg><div class="g-num"><span id="pPE">—</span><span class="g-pct">%ile</span></div></div><div><div class="g-val" id="vPE">—</div><div class="g-med">Med: __PE_MED__</div><div class="g-zone" id="zPE">—</div></div></div></div>
    <div class="gauge-card"><div class="gauge-title">BEER Ratio</div><div class="gw"><div class="g-arc"><svg viewBox="0 0 56 56"><circle cx="28" cy="28" r="22" fill="none" stroke="#1f2a3c" stroke-width="6" stroke-dasharray="138.2" stroke-linecap="round" transform="rotate(-90 28 28)"/><circle cx="28" cy="28" r="22" fill="none" id="arcBEER" stroke="#f0b429" stroke-width="6" stroke-dasharray="138.2" stroke-dashoffset="138.2" stroke-linecap="round" transform="rotate(-90 28 28)"/></svg><div class="g-num"><span id="pBEER">—</span><span class="g-pct">%ile</span></div></div><div><div class="g-val" id="vBEER">—</div><div class="g-med">Med: __BEER_MED__</div><div class="g-zone" id="zBEER">—</div></div></div></div>
    <div class="gauge-card"><div class="gauge-title">MC / GDP %</div><div class="gw"><div class="g-arc"><svg viewBox="0 0 56 56"><circle cx="28" cy="28" r="22" fill="none" stroke="#1f2a3c" stroke-width="6" stroke-dasharray="138.2" stroke-linecap="round" transform="rotate(-90 28 28)"/><circle cx="28" cy="28" r="22" fill="none" id="arcMC" stroke="#ff5c6a" stroke-width="6" stroke-dasharray="138.2" stroke-dashoffset="138.2" stroke-linecap="round" transform="rotate(-90 28 28)"/></svg><div class="g-num"><span id="pMC">—</span><span class="g-pct">%ile</span></div></div><div><div class="g-val" id="vMC">—</div><div class="g-med">Med: __MCGDP_MED__%</div><div class="g-zone" id="zMC">—</div></div></div></div>
    <div class="gauge-card"><div class="gauge-title">Yield Gap (EY−Bond)</div><div class="gw"><div class="g-arc"><svg viewBox="0 0 56 56"><circle cx="28" cy="28" r="22" fill="none" stroke="#1f2a3c" stroke-width="6" stroke-dasharray="138.2" stroke-linecap="round" transform="rotate(-90 28 28)"/><circle cx="28" cy="28" r="22" fill="none" id="arcYG" stroke="#00d68f" stroke-width="6" stroke-dasharray="138.2" stroke-dashoffset="138.2" stroke-linecap="round" transform="rotate(-90 28 28)"/></svg><div class="g-num"><span id="pYG">—</span><span class="g-pct">%ile</span></div></div><div><div class="g-val" id="vYG">—</div><div class="g-med">Med: __YG_MED__%</div><div class="g-zone" id="zYG">—</div></div></div></div>
  </div>

  <!-- SECTION 1: NIFTY -->
  <div class="sec"><span>◈</span>Nifty 50 Valuation</div>
  <div class="g3">
    <div class="cc sp2"><div class="ch"><span class="ct">Nifty 50 Index</span><span class="cv" id="cN">—</span></div><div class="cw"><canvas id="cNifty"></canvas></div></div>
    <div class="cc"><div class="ch"><span class="ct">P/E Ratio</span><span class="cv" id="cPE">—</span></div><div class="cw"><canvas id="cPEchart"></canvas></div></div>
  </div>
  <div class="g4">
    <div class="cc"><div class="ch"><span class="ct">Earning Yield %</span><span class="cv" id="cEY">—</span></div><div class="cw"><canvas id="cEYchart"></canvas></div></div>
    <div class="cc"><div class="ch"><span class="ct">BEER Ratio</span><span class="cv" id="cBEER">—</span></div><div class="cw"><canvas id="cBEERchart"></canvas></div></div>
    <div class="cc"><div class="ch"><span class="ct">MC/GDP % (Buffett)</span><span class="cv" id="cMC">—</span></div><div class="cw"><canvas id="cMCchart"></canvas></div></div>
    <div class="cc"><div class="ch"><span class="ct">Market Cap (USD T)</span><span class="cv" id="cMCT">—</span></div><div class="cw"><canvas id="cMCTchart"></canvas></div></div>
  </div>

  <!-- SECTION 2: MIDCAP & SMALLCAP -->
  <div class="sec"><span>◈</span>Midcap 150 &amp; Smallcap 250 — Earning Yield &amp; PREITY</div>
  <div class="g3">
    <div class="cc"><div class="ch"><span class="ct">Midcap 150 Earning Yield %</span><span class="cv" id="cMEY">—</span></div><div class="cw"><canvas id="cMEYchart"></canvas></div></div>
    <div class="cc"><div class="ch"><span class="ct">Smallcap 250 Earning Yield %</span><span class="cv" id="cSEY">—</span></div><div class="cw"><canvas id="cSEYchart"></canvas></div></div>
    <div class="cc"><div class="ch"><span class="ct">PREITY Ratio (Nifty/US10yr)</span><span class="cv" id="cPREITY">—</span></div><div class="cw"><canvas id="cPREITYchart"></canvas></div></div>
  </div>

  <!-- SECTION 3: EPS GROWTH -->
  <div class="sec"><span>◈</span>Earnings Growth Rate — YoY (trailing 252 trading days ≈ 1 year)</div>
  <div class="g3">
    <div class="cc"><div class="ch"><span class="ct">Nifty 50 EPS Growth %</span><span class="cv" id="cNEG">—</span></div><div class="cw"><canvas id="cNEGchart"></canvas></div></div>
    <div class="cc"><div class="ch"><span class="ct">Midcap 150 EPS Growth %</span><span class="cv" id="cMEG">—</span></div><div class="cw"><canvas id="cMEGchart"></canvas></div></div>
    <div class="cc"><div class="ch"><span class="ct">Smallcap 250 EPS Growth %</span><span class="cv" id="cSEG">—</span></div><div class="cw"><canvas id="cSEGchart"></canvas></div></div>
  </div>

  <!-- SECTION 4: BONDS & FX -->
  <div class="sec"><span>◈</span>Bond Yields &amp; Dollar Index</div>
  <div class="g3">
    <div class="cc sp2"><div class="ch"><span class="ct">India 10yr vs US 10yr &amp; Yield Gap</span><span class="cv" id="cBond">—</span></div><div class="cw"><canvas id="cBondchart"></canvas></div><div class="lgd"><div class="li"><div class="ld" style="background:#f0b429"></div>India 10yr</div><div class="li"><div class="ld" style="background:#00c8ff"></div>US 10yr</div><div class="li"><div class="ld" style="background:#00d68f"></div>Yield Gap</div></div></div>
    <div class="cc"><div class="ch"><span class="ct">USD / INR</span><span class="cv" id="cFX">—</span></div><div class="cw"><canvas id="cFXchart"></canvas></div></div>
  </div>

</div>

<script>
const RAW=__CHART_DATA__;
const STATS=__STATS_DATA__;

Chart.defaults.color='#5a7089';
Chart.defaults.font.family="'IBM Plex Mono',monospace";
Chart.defaults.font.size=9;
Chart.defaults.animation={duration:200};

const GRID={color:'rgba(31,42,60,0.9)',lineWidth:1};
const TICK={color:'#5a7089',maxTicksLimit:5};

function bo(){
  return{responsive:true,maintainAspectRatio:false,
    interaction:{mode:'index',intersect:false},
    plugins:{legend:{display:false},tooltip:{backgroundColor:'#111620',borderColor:'#1f2a3c',borderWidth:1,titleColor:'#c8d6e5',bodyColor:'#5a7089',padding:8}},
    scales:{x:{grid:GRID,ticks:{...TICK,maxRotation:0,maxTicksLimit:6}},y:{grid:GRID,ticks:TICK}}};
}

function grad(id,c){
  const cv=document.getElementById(id);if(!cv)return c+'33';
  const g=cv.getContext('2d').createLinearGradient(0,0,0,130);
  g.addColorStop(0,c+'55');g.addColorStop(1,c+'00');return g;
}

function lds(data,color,fill,id){
  return{data,borderColor:color,backgroundColor:fill?grad(id,color):'transparent',borderWidth:1.5,pointRadius:0,tension:0.3,fill:!!fill};
}

function barDs(data,color){
  return{type:'bar',data,backgroundColor:data.map(v=>v>=0?color+'99':v<0?'#ff5c6a99':color+'99'),borderColor:data.map(v=>v>=0?color:'#ff5c6a'),borderWidth:1,borderRadius:2};
}

function lbl(dates){
  return dates.map(d=>new Date(d).toLocaleDateString('en-IN',{day:'2-digit',month:'short',year:'2-digit'}));
}

function pct(arr,v){return Math.round(arr.filter(x=>x<=v).length/arr.length*100);}

function zone(p,flip){
  const cheap=flip?p>60:p<30,rich=flip?p<30:p>70;
  if(cheap)return['Attractive','zc'];if(rich)return['Stretched','zr'];return['Fair Value','zf'];
}

function set(id,v){const el=document.getElementById(id);if(el)el.textContent=v;}

function setGauge(arcId,pId,vId,zId,p,val,flip,color){
  const arc=document.getElementById(arcId);
  arc.style.strokeDashoffset=138.2*(1-p/100);arc.style.stroke=color;
  document.getElementById(pId).textContent=p+'%';
  document.getElementById(vId).textContent=typeof val==='number'?val.toFixed(2):val;
  const[zt,zc]=zone(p,flip);const el=document.getElementById(zId);el.textContent=zt;el.className='g-zone '+zc;
}

function growthColor(v){return v>=0?'var(--green)':'var(--red)';}

const charts={};

function filterData(n){
  const len=RAW.date.length,start=n>=9999?0:Math.max(0,len-n);
  const sl=k=>RAW[k].slice(start);
  return{date:sl('date'),nifty50:sl('nifty50'),pe:sl('pe'),pb:sl('pb'),
    earning_yield:sl('earning_yield'),india_10yr:sl('india_10yr'),us_10yr:sl('us_10yr'),
    yield_gap:sl('yield_gap'),usdinr:sl('usdinr'),dollar_index:sl('dollar_index'),
    marketcap_gdp:sl('marketcap_gdp'),marketcap_trillion:sl('marketcap_trillion'),
    beer:sl('beer'),preity:sl('preity'),
    midcap_earn_yield:sl('midcap_earn_yield'),smallcap_earn_yield:sl('smallcap_earn_yield'),
    nifty_eps_growth:sl('nifty_eps_growth'),midcap_eps_growth:sl('midcap_eps_growth'),
    smallcap_eps_growth:sl('smallcap_eps_growth')};
}

function updateKPIs(d){
  const n=d.date.length-1,p=n>0?n-1:0;
  const nifty=d.nifty50[n],pe=d.pe[n],pb=d.pb[n],ey=d.earning_yield[n];
  const i10=d.india_10yr[n],u10=d.us_10yr[n],yg=d.yield_gap[n];
  const fx=d.usdinr[n],mc=d.marketcap_gdp[n],beer=d.beer[n];
  const mct=d.marketcap_trillion[n],preity=d.preity[n];
  const mey=d.midcap_earn_yield[n],sey=d.smallcap_earn_yield[n];
  const neg=d.nifty_eps_growth[n],meg=d.midcap_eps_growth[n],seg=d.smallcap_eps_growth[n];
  const chg=((nifty-d.nifty50[p])/d.nifty50[p]*100).toFixed(2);

  set('kN',nifty.toLocaleString('en-IN',{maximumFractionDigits:0}));
  const kNc=document.getElementById('kNc');kNc.textContent=(chg>0?'+':'')+chg+'%';kNc.className='kpi-c '+(chg>=0?'up':'dn');
  set('kPE',pe.toFixed(2));set('kPEm','med __PE_MED__');set('kPB',pb.toFixed(2));set('kEY',ey.toFixed(2)+'%');
  set('kI10',i10.toFixed(2)+'%');
  const kYG=document.getElementById('kYG');kYG.textContent='gap '+(yg>=0?'+':'')+yg.toFixed(2)+'%';kYG.className='kpi-c '+(yg>=0?'up':'dn');
  set('kU10',u10.toFixed(2)+'%');set('kSp','spr '+(i10-u10).toFixed(2)+'%');
  set('kFX',fx.toFixed(2));set('kDXY','DXY '+d.dollar_index[n].toFixed(2));
  const mcEl=document.getElementById('kMC');mcEl.textContent=mc.toFixed(1)+'%';mcEl.className='kpi-v '+(mc>150?'r':mc>120?'gld':'g');
  set('kMCT','$'+mct.toFixed(2)+'T');
  set('kPREITY',preity.toFixed(0));
  set('kMEY',mey.toFixed(2)+'%');set('kSEY',sey.toFixed(2)+'%');
  set('kBEER',beer.toFixed(3));

  const negEl=document.getElementById('kNEG');negEl.textContent=(neg>=0?'+':'')+neg.toFixed(1)+'%';negEl.className='kpi-v '+(neg>=0?'g':'r');
  const megEl=document.getElementById('kMEG');megEl.textContent=(meg>=0?'+':'')+meg.toFixed(1)+'%';megEl.className='kpi-v '+(meg>=0?'g':'r');
  const segEl=document.getElementById('kSEG');segEl.textContent=(seg>=0?'+':'')+seg.toFixed(1)+'%';segEl.className='kpi-v '+(seg>=0?'g':'r');

  set('cN',nifty.toLocaleString('en-IN',{maximumFractionDigits:0}));
  set('cPE',pe.toFixed(2)+'x');set('cEY',ey.toFixed(2)+'%');set('cBEER',beer.toFixed(3));
  set('cMC',mc.toFixed(1)+'%');set('cMCT','$'+mct.toFixed(2)+'T');
  set('cMEY',mey.toFixed(2)+'%');set('cSEY',sey.toFixed(2)+'%');
  set('cPREITY',preity.toFixed(0));
  set('cBond',i10.toFixed(2)+'% / '+u10.toFixed(2)+'%');set('cFX',fx.toFixed(2));
  set('cNEG',(neg>=0?'+':'')+neg.toFixed(1)+'%');
  set('cMEG',(meg>=0?'+':'')+meg.toFixed(1)+'%');
  set('cSEG',(seg>=0?'+':'')+seg.toFixed(1)+'%');

  const peArr=RAW.pe,beerArr=RAW.beer.filter(x=>x>0),mcArr=RAW.marketcap_gdp.filter(x=>x>0),ygArr=RAW.yield_gap;
  setGauge('arcPE','pPE','vPE','zPE',pct(peArr,pe),pe,false,'#00c8ff');
  setGauge('arcBEER','pBEER','vBEER','zBEER',pct(beerArr,beer),beer,true,'#f0b429');
  setGauge('arcMC','pMC','vMC','zMC',pct(mcArr,mc),mc,false,'#ff5c6a');
  setGauge('arcYG','pYG','vYG','zYG',pct(ygArr,yg),yg,true,'#00d68f');

  const score=Math.round((pct(peArr,pe)+(100-pct(beerArr,beer))+pct(mcArr,mc)+(100-pct(ygArr,yg)))/4);
  set('eviNum',score);document.getElementById('eviDot').style.left=score+'%';
  const[hl,desc]=score<35?['Market Attractive','PE, BEER and yield metrics below historical medians. Strong long-term entry zone.']:score<55?['Fairly Valued','Valuations near historical medians. Balanced risk-reward.']:score<75?['Mildly Stretched','Several indicators above median. Prefer quality stocks.']:['Expensive','Valuations in top quartile. Consider reducing equity allocation.'];
  set('eviHl',hl);set('eviDesc',desc);
}

function buildCharts(d){
  const labels=lbl(d.date);
  const ref=(arr,v)=>arr.map(()=>v);
  const opt=bo();

  function mk(key,id,datasets,yOpts){
    if(charts[key]){charts[key].destroy();delete charts[key];}
    const canvas=document.getElementById(id);if(!canvas)return;
    const opts=JSON.parse(JSON.stringify(opt));
    if(yOpts)Object.assign(opts.scales.y,yOpts);
    charts[key]=new Chart(canvas.getContext('2d'),{type:'line',data:{labels,datasets},options:opts});
  }

  function mkBar(key,id,datasets,yOpts){
    if(charts[key]){charts[key].destroy();delete charts[key];}
    const canvas=document.getElementById(id);if(!canvas)return;
    const opts=JSON.parse(JSON.stringify(opt));
    opts.scales.x.grid={display:false};
    if(yOpts)Object.assign(opts.scales.y,yOpts);
    charts[key]=new Chart(canvas.getContext('2d'),{type:'bar',data:{labels,datasets},options:opts});
  }

  mk('nifty','cNifty',[lds(d.nifty50,'#00c8ff',true,'cNifty')]);
  mk('pe','cPEchart',[lds(d.pe,'#00c8ff',false),{data:ref(d.pe,STATS.pe_median),borderColor:'#f0b42966',borderDash:[4,3],borderWidth:1,pointRadius:0,tension:0}],{min:15,max:30});
  mk('ey','cEYchart',[lds(d.earning_yield,'#00d68f',true,'cEYchart')]);
  mk('beer','cBEERchart',[lds(d.beer,'#f0b429',false),{data:ref(d.beer,1.0),borderColor:'#ff5c6a66',borderDash:[4,3],borderWidth:1,pointRadius:0,tension:0},{data:ref(d.beer,STATS.beer_median),borderColor:'#00d68f66',borderDash:[4,3],borderWidth:1,pointRadius:0,tension:0}],{min:0.4,max:2.0});
  mk('mc','cMCchart',[lds(d.marketcap_gdp,'#ff5c6a',false),{data:ref(d.marketcap_gdp,STATS.mcgdp_median),borderColor:'#f0b42966',borderDash:[4,3],borderWidth:1,pointRadius:0,tension:0}],{min:80,max:170});
  mk('mct','cMCTchart',[lds(d.marketcap_trillion,'#00c8ff',true,'cMCTchart')]);
  mk('mey','cMEYchart',[lds(d.midcap_earn_yield,'#00d68f',true,'cMEYchart')]);
  mk('sey','cSEYchart',[lds(d.smallcap_earn_yield,'#c084fc',true,'cSEYchart')]);
  mk('preity','cPREITYchart',[lds(d.preity,'#ff8c42',true,'cPREITYchart')]);
  mk('bond','cBondchart',[lds(d.india_10yr,'#f0b429',false),lds(d.us_10yr,'#00c8ff',false),lds(d.yield_gap,'#00d68f',true,'cBondchart')]);
  mk('fx','cFXchart',[lds(d.usdinr,'#c084fc',true,'cFXchart')]);

  // EPS Growth bars — skip zero values (first year has no data)
  const negData=d.nifty_eps_growth.map(v=>v===0?null:v);
  const megData=d.midcap_eps_growth.map(v=>v===0?null:v);
  const segData=d.smallcap_eps_growth.map(v=>v===0?null:v);

  function growthBar(key,id,data,color){
    if(charts[key]){charts[key].destroy();delete charts[key];}
    const canvas=document.getElementById(id);if(!canvas)return;
    const opts=JSON.parse(JSON.stringify(opt));
    opts.scales.x.grid={display:false};
    opts.plugins.tooltip.callbacks={label:ctx=>(ctx.parsed.y>=0?'+':'')+ctx.parsed.y.toFixed(1)+'%'};
    charts[key]=new Chart(canvas.getContext('2d'),{
      type:'bar',
      data:{labels,datasets:[{data,backgroundColor:data.map(v=>v===null?'transparent':v>=0?color+'99':'#ff5c6a99'),borderColor:data.map(v=>v===null?'transparent':v>=0?color:'#ff5c6a'),borderWidth:1,borderRadius:2}]},
      options:opts
    });
  }
  growthBar('neg','cNEGchart',negData,'#00c8ff');
  growthBar('meg','cMEGchart',megData,'#00d68f');
  growthBar('seg','cSEGchart',segData,'#c084fc');
}

function refresh(range){
  const d=filterData(range);
  updateKPIs(d);
  buildCharts(d);
}

document.getElementById('fr').addEventListener('click',e=>{
  const btn=e.target.closest('[data-r]');if(!btn)return;
  document.querySelectorAll('.fb').forEach(b=>b.classList.remove('on'));
  btn.classList.add('on');refresh(parseInt(btn.dataset.r));
});

refresh(365);
</script>
</body>
</html>
"""

def main():
    if not EXCEL.exists():
        print(f"ERROR: {EXCEL} not found"); sys.exit(1)
    print(f"📖  Reading {EXCEL.name} …")
    wb   = load_workbook_safe(EXCEL)
    rows = extract_data(wb)
    if not rows:
        print("ERROR: No data rows found"); sys.exit(1)
    print(f"    {len(rows)} rows  |  {rows[0]['date']} → {rows[-1]['date']}")
    cd    = build_chart_data(rows)
    stats = compute_stats(rows)
    html = HTML
    html = html.replace("__LAST_DATE__",   stats["last_date"])
    html = html.replace("__DATE_FROM__",   stats["date_from"])
    html = html.replace("__TOTAL_ROWS__",  str(stats["total_rows"]))
    html = html.replace("__PE_MED__",      str(stats["pe_median"]))
    html = html.replace("__BEER_MED__",    str(stats["beer_median"]))
    html = html.replace("__MCGDP_MED__",   str(stats["mcgdp_median"]))
    html = html.replace("__YG_MED__",      str(stats["yg_median"]))
    html = html.replace("__CHART_DATA__",  json.dumps(cd))
    html = html.replace("__STATS_DATA__",  json.dumps(stats))
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(html, encoding="utf-8")
    print(f"✅  Dashboard → {OUTPUT}  ({OUTPUT.stat().st_size//1024} KB)")
    print(f"    Nifty {stats['nifty']:,.0f}  |  PE {stats['pe']:.2f}  |  MC ${stats['marketcap_trillion']:.2f}T")
    print(f"    Midcap EY {stats['midcap_earn_yield']:.2f}%  |  SC EY {stats['smallcap_earn_yield']:.2f}%")
    print(f"    EPS Growth → Nifty {stats['nifty_eps_growth']:.1f}%  Midcap {stats['midcap_eps_growth']:.1f}%  SC {stats['smallcap_eps_growth']:.1f}%")

if __name__ == "__main__":
    main()
