#!/usr/bin/env python3
"""
EVI Dashboard Builder
---------------------
Reads EVI_2025-26.xlsx  →  generates docs/index.html

Run:  python scripts/build.py
"""

import json
import datetime
import statistics
import zipfile
import re
import sys
from pathlib import Path

# ── Paths ─────────────────────────────────────────────────────────────────────
ROOT   = Path(__file__).parent.parent
EXCEL  = ROOT / "data" / "EVI_2025-26.xlsx"
OUTPUT = ROOT / "docs" / "index.html"

# ── 1. Read Excel (handles the NSE stylesheet quirk) ──────────────────────────
def load_workbook_safe(path):
    """Patch the font-family value issue before loading with openpyxl."""
    from openpyxl import load_workbook
    import io

    with zipfile.ZipFile(path, "r") as z:
        files = {n: z.read(n) for n in z.namelist()}

    styles = files.get("xl/styles.xml", b"")
    styles = re.sub(
        rb'(<family val=")(\d+)(")',
        lambda m: m.group(0) if int(m.group(2)) <= 14
                  else m.group(1) + b"2" + m.group(3),
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
            if val is None:
                return 0.0
            if isinstance(val, (datetime.datetime, datetime.date)):
                return val
            try:
                return float(val)
            except (TypeError, ValueError):
                return 0.0

        date_raw = ws.cell(row=r, column=2).value
        if date_raw is None:
            continue
        if isinstance(date_raw, (datetime.datetime, datetime.date)):
            dt = date_raw.date() if isinstance(date_raw, datetime.datetime) else date_raw
        else:
            try:
                dt = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(float(date_raw)))
            except Exception:
                continue

        nifty  = v(12)
        pe     = v(15)
        if nifty == 0 or pe == 0:
            continue

        rows.append({
            "date":             dt.isoformat(),
            "nifty50":          nifty,
            "pe":               pe,
            "pb":               v(13),
            "eps":              v(10),
            "earning_yield":    v(9),
            "india_10yr":       v(11),
            "us_10yr":          v(16),
            "yield_gap":        v(14),
            "usdinr":           v(5),
            "dollar_index":     v(18),
            "marketcap_inr":    v(3),
            "marketcap_gdp":    v(7),
            "beer":             v(8),
            "preity":           v(19),
            "t91":              v(20),
            "midcap150":        v(21),
            "midcap_pe":        v(22),
            "midcap_earn_yield":v(24),
            "smallcap250":      v(25),
            "smallcap_pe":      v(26),
            "smallcap_earn_yield": v(28),
        })
    return rows


# ── 2. Build chart data (thinned for performance) ─────────────────────────────
def build_chart_data(rows):
    n = len(rows)
    # Keep every 3rd point for older data, full resolution for last 90 days
    thin_idx = list(range(0, max(0, n - 90), 3)) + list(range(max(0, n - 90), n))

    def pick(key):
        return [round(rows[i][key], 4) if isinstance(rows[i][key], float)
                else rows[i][key]
                for i in thin_idx]

    return {
        "dates":          pick("date"),
        "nifty50":        pick("nifty50"),
        "pe":             pick("pe"),
        "pb":             pick("pb"),
        "earning_yield":  pick("earning_yield"),
        "india_10yr":     pick("india_10yr"),
        "us_10yr":        pick("us_10yr"),
        "yield_gap":      pick("yield_gap"),
        "usdinr":         pick("usdinr"),
        "dollar_index":   pick("dollar_index"),
        "beer":           pick("beer"),
        "marketcap_gdp":  pick("marketcap_gdp"),
        "midcap_pe":      pick("midcap_pe"),
        "smallcap_pe":    pick("smallcap_pe"),
    }


# ── 3. Compute summary stats ───────────────────────────────────────────────────
def compute_stats(rows):
    def med(key):
        vals = [r[key] for r in rows if r[key] > 0]
        return round(statistics.median(vals), 4) if vals else 0

    latest  = rows[-1]
    first   = rows[0]
    nifty_ytd_start = next(
        (r["nifty50"] for r in rows
         if r["date"].startswith(str(datetime.date.today().year))),
        first["nifty50"],
    )

    return {
        "last_date":       latest["date"],
        "total_rows":      len(rows),
        "date_from":       first["date"],
        "nifty":           latest["nifty50"],
        "pe":              latest["pe"],
        "pb":              latest["pb"],
        "earning_yield":   latest["earning_yield"],
        "india_10yr":      latest["india_10yr"],
        "us_10yr":         latest["us_10yr"],
        "yield_gap":       latest["yield_gap"],
        "usdinr":          latest["usdinr"],
        "dollar_index":    latest["dollar_index"],
        "beer":            latest["beer"],
        "marketcap_gdp":   latest["marketcap_gdp"],
        "pe_median":       med("pe"),
        "beer_median":     med("beer"),
        "mcgdp_median":    med("marketcap_gdp"),
        "yg_median":       med("yield_gap"),
        "nifty_ytd_start": nifty_ytd_start,
    }


# ── 4. HTML template ──────────────────────────────────────────────────────────
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>EVI Dashboard — {last_date}</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');
:root{{
  --bg:#0b0e14;--bg2:#111620;--bg3:#181e2c;--border:#1f2a3c;
  --accent:#00c8ff;--gold:#f0b429;--green:#00d68f;--red:#ff5c6a;
  --orange:#ff8c42;--text:#c8d6e5;--muted:#5a7089;--white:#eef4fb;
  --mono:'IBM Plex Mono',monospace;--sans:'IBM Plex Sans',sans-serif;
}}
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:var(--sans);background:var(--bg);color:var(--text);min-height:100vh;overflow-x:hidden}}
.header{{background:var(--bg2);border-bottom:1px solid var(--border);padding:18px 32px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:100}}
.header h1{{font-family:var(--mono);font-size:15px;font-weight:600;color:var(--accent);letter-spacing:.12em;text-transform:uppercase}}
.header .sub{{font-family:var(--mono);font-size:11px;color:var(--muted);letter-spacing:.08em;margin-left:12px}}
.header-right{{font-family:var(--mono);font-size:11px;color:var(--muted);text-align:right;line-height:1.6}}
.header-right strong{{color:var(--text);font-weight:500}}
.wrapper{{max-width:1440px;margin:0 auto;padding:24px 32px}}
.data-note{{background:rgba(0,200,255,.05);border:1px solid rgba(0,200,255,.15);border-radius:6px;padding:12px 16px;font-size:12px;color:var(--muted);margin-bottom:20px;line-height:1.5}}
.data-note strong{{color:var(--accent)}}
.kpi-strip{{display:grid;grid-template-columns:repeat(8,1fr);gap:1px;background:var(--border);border:1px solid var(--border);border-radius:8px;overflow:hidden;margin-bottom:20px}}
.kpi-cell{{background:var(--bg2);padding:16px 14px;display:flex;flex-direction:column;gap:4px;transition:background .15s}}
.kpi-cell:hover{{background:var(--bg3)}}
.kpi-label{{font-family:var(--mono);font-size:9px;font-weight:500;letter-spacing:.12em;text-transform:uppercase;color:var(--muted)}}
.kpi-value{{font-family:var(--mono);font-size:20px;font-weight:600;color:var(--white);line-height:1}}
.kpi-value.green{{color:var(--green)}}.kpi-value.red{{color:var(--red)}}.kpi-value.gold{{color:var(--gold)}}.kpi-value.accent{{color:var(--accent)}}
.kpi-change{{font-family:var(--mono);font-size:10px;color:var(--muted)}}
.kpi-change.up{{color:var(--green)}}.kpi-change.down{{color:var(--red)}}
.section-title{{font-family:var(--mono);font-size:10px;font-weight:600;letter-spacing:.18em;text-transform:uppercase;color:var(--muted);margin-bottom:12px;padding-bottom:8px;border-bottom:1px solid var(--border)}}
.section-title span{{color:var(--accent);margin-right:8px}}
.gauge-row{{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:20px}}
.gauge-card{{background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:20px}}
.gauge-title{{font-family:var(--mono);font-size:9px;font-weight:600;letter-spacing:.14em;text-transform:uppercase;color:var(--muted);margin-bottom:12px}}
.gauge-wrap{{display:flex;align-items:center;gap:14px}}
.gauge-arc{{position:relative;width:72px;height:72px;flex-shrink:0}}
.gauge-arc svg{{width:100%;height:100%}}
.gauge-arc .gauge-num{{position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;font-family:var(--mono);font-size:13px;font-weight:600;line-height:1;color:var(--white)}}
.gauge-arc .gauge-pct{{font-size:9px;color:var(--muted);margin-top:2px}}
.gauge-info{{flex:1}}
.gauge-val{{font-family:var(--mono);font-size:22px;font-weight:600;color:var(--white);line-height:1}}
.gauge-median{{font-family:var(--mono);font-size:10px;color:var(--muted);margin-top:4px}}
.gauge-zone{{display:inline-block;margin-top:8px;padding:2px 8px;border-radius:3px;font-family:var(--mono);font-size:9px;font-weight:600;letter-spacing:.08em;text-transform:uppercase}}
.zone-cheap{{background:rgba(0,214,143,.15);color:var(--green)}}
.zone-fair{{background:rgba(240,180,41,.15);color:var(--gold)}}
.zone-rich{{background:rgba(255,92,106,.15);color:var(--red)}}
.evi-score-card{{background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:24px 28px;margin-bottom:20px;display:flex;align-items:center;gap:32px}}
.evi-big{{font-family:var(--mono);font-size:64px;font-weight:600;line-height:1;background:linear-gradient(135deg,var(--accent),var(--green));-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text;min-width:160px}}
.evi-details{{flex:1}}
.evi-headline{{font-size:18px;font-weight:500;color:var(--white);margin-bottom:6px}}
.evi-desc{{font-size:13px;color:var(--muted);line-height:1.6;max-width:600px}}
.evi-bar-wrap{{flex:1;min-width:280px}}
.evi-bar-bg{{height:10px;background:var(--border);border-radius:5px;overflow:hidden;margin-bottom:6px}}
.evi-bar-fill{{height:100%;border-radius:5px;background:linear-gradient(90deg,var(--green),var(--gold),var(--red));position:relative}}
.evi-marker{{position:absolute;top:-4px;width:18px;height:18px;background:var(--white);border-radius:50%;border:2px solid var(--bg);transform:translateX(-50%);box-shadow:0 0 8px rgba(0,200,255,.6)}}
.evi-bar-labels{{display:flex;justify-content:space-between;font-family:var(--mono);font-size:9px;color:var(--muted)}}
.filter-row{{display:flex;gap:6px;margin-bottom:20px;flex-wrap:wrap}}
.filter-btn{{font-family:var(--mono);font-size:10px;font-weight:500;letter-spacing:.08em;padding:5px 12px;border-radius:4px;border:1px solid var(--border);background:transparent;color:var(--muted);cursor:pointer;transition:all .15s}}
.filter-btn:hover{{border-color:var(--accent);color:var(--accent)}}
.filter-btn.active{{background:var(--accent);border-color:var(--accent);color:var(--bg)}}
.charts-grid{{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px}}
.chart-card{{background:var(--bg2);border:1px solid var(--border);border-radius:8px;padding:20px}}
.chart-card.wide{{grid-column:span 2}}
.chart-header{{display:flex;align-items:baseline;justify-content:space-between;margin-bottom:16px}}
.chart-title{{font-family:var(--mono);font-size:10px;font-weight:600;letter-spacing:.14em;text-transform:uppercase;color:var(--text)}}
.chart-curr{{font-family:var(--mono);font-size:13px;font-weight:600;color:var(--accent)}}
.chart-canvas{{height:200px}}.chart-canvas.tall{{height:260px}}
.legend-row{{display:flex;gap:16px;flex-wrap:wrap;margin-top:8px}}
.legend-item{{display:flex;align-items:center;gap:6px;font-family:var(--mono);font-size:9px;color:var(--muted)}}
.legend-dot{{width:8px;height:8px;border-radius:50%;flex-shrink:0}}
@media(max-width:1100px){{.kpi-strip{{grid-template-columns:repeat(4,1fr)}}.gauge-row{{grid-template-columns:repeat(2,1fr)}}}}
@media(max-width:720px){{.wrapper{{padding:16px}}.kpi-strip{{grid-template-columns:repeat(2,1fr)}}.charts-grid{{grid-template-columns:1fr}}.chart-card.wide{{grid-column:span 1}}.gauge-row{{grid-template-columns:1fr}}.evi-score-card{{flex-direction:column}}.header{{flex-direction:column;gap:8px;align-items:flex-start}}}}
</style>
</head>
<body>

<div class="header">
  <div style="display:flex;align-items:baseline;gap:12px">
    <h1>EVI // Equity Valuation Index</h1>
    <span class="sub">India Market Monitor</span>
  </div>
  <div class="header-right">
    <strong>{last_date}</strong><br>
    {date_from} → {last_date} &nbsp;|&nbsp; {total_rows} trading days
  </div>
</div>

<div class="wrapper">
  <div class="data-note">
    <strong>Auto-generated</strong> from <code>EVI_2025-26.xlsx</code> on <strong>{built_at}</strong>.
    Run <code>python scripts/build.py</code> after each daily update to refresh this dashboard.
  </div>

  <div class="section-title"><span>◈</span>Composite Valuation Signal</div>
  <div class="evi-score-card">
    <div class="evi-big" id="eviScore">—</div>
    <div class="evi-details">
      <div class="evi-headline" id="eviHeadline">—</div>
      <div class="evi-desc" id="eviDesc">—</div>
    </div>
    <div class="evi-bar-wrap">
      <div class="evi-bar-bg"><div class="evi-bar-fill" style="width:100%"><div class="evi-marker" id="eviMarker"></div></div></div>
      <div class="evi-bar-labels"><span>Cheap</span><span>Fair Value</span><span>Expensive</span></div>
    </div>
  </div>

  <div class="kpi-strip">
    <div class="kpi-cell"><div class="kpi-label">Nifty 50</div><div class="kpi-value" id="kNifty">—</div><div class="kpi-change" id="kNiftyChg">—</div></div>
    <div class="kpi-cell"><div class="kpi-label">P/E Ratio</div><div class="kpi-value" id="kPE">—</div><div class="kpi-change" id="kPEmd">—</div></div>
    <div class="kpi-cell"><div class="kpi-label">P/B Ratio</div><div class="kpi-value" id="kPB">—</div><div class="kpi-change">price to book</div></div>
    <div class="kpi-cell"><div class="kpi-label">Earning Yield</div><div class="kpi-value green" id="kEY">—</div><div class="kpi-change">1/PE × 100</div></div>
    <div class="kpi-cell"><div class="kpi-label">India 10yr</div><div class="kpi-value gold" id="kIN10">—</div><div class="kpi-change" id="kYG">—</div></div>
    <div class="kpi-cell"><div class="kpi-label">US 10yr</div><div class="kpi-value" id="kUS10">—</div><div class="kpi-change" id="kSpread">—</div></div>
    <div class="kpi-cell"><div class="kpi-label">USD / INR</div><div class="kpi-value" id="kUSDINR">—</div><div class="kpi-change" id="kDXY">—</div></div>
    <div class="kpi-cell"><div class="kpi-label">MC / GDP</div><div class="kpi-value" id="kMCGDP">—</div><div class="kpi-change">Buffett indicator</div></div>
  </div>

  <div class="filter-row" id="filterRow">
    <button class="filter-btn" data-range="90">3M</button>
    <button class="filter-btn" data-range="180">6M</button>
    <button class="filter-btn active" data-range="365">1Y</button>
    <button class="filter-btn" data-range="730">2Y</button>
    <button class="filter-btn" data-range="9999">All</button>
  </div>

  <div class="section-title"><span>◈</span>Valuation Gauges — Percentile vs Full History</div>
  <div class="gauge-row">
    <div class="gauge-card">
      <div class="gauge-title">Nifty 50 P/E</div>
      <div class="gauge-wrap">
        <div class="gauge-arc"><svg viewBox="0 0 72 72"><circle cx="36" cy="36" r="28" fill="none" stroke="#1f2a3c" stroke-width="8" stroke-dasharray="175.9" stroke-linecap="round" transform="rotate(-90 36 36)"/><circle cx="36" cy="36" r="28" fill="none" id="gaugePEArc" stroke="#00c8ff" stroke-width="8" stroke-dasharray="175.9" stroke-dashoffset="175.9" stroke-linecap="round" transform="rotate(-90 36 36)"/></svg><div class="gauge-num"><span id="gaugePEPct">—</span><span class="gauge-pct">pctile</span></div></div>
        <div class="gauge-info"><div class="gauge-val" id="gaugePEVal">—</div><div class="gauge-median" id="gaugePEMd">—</div><div class="gauge-zone" id="gaugePEZone">—</div></div>
      </div>
    </div>
    <div class="gauge-card">
      <div class="gauge-title">BEER Ratio</div>
      <div class="gauge-wrap">
        <div class="gauge-arc"><svg viewBox="0 0 72 72"><circle cx="36" cy="36" r="28" fill="none" stroke="#1f2a3c" stroke-width="8" stroke-dasharray="175.9" stroke-linecap="round" transform="rotate(-90 36 36)"/><circle cx="36" cy="36" r="28" fill="none" id="gaugeBEERArc" stroke="#f0b429" stroke-width="8" stroke-dasharray="175.9" stroke-dashoffset="175.9" stroke-linecap="round" transform="rotate(-90 36 36)"/></svg><div class="gauge-num"><span id="gaugeBEERPct">—</span><span class="gauge-pct">pctile</span></div></div>
        <div class="gauge-info"><div class="gauge-val" id="gaugeBEERVal">—</div><div class="gauge-median" id="gaugeBEERMd">—</div><div class="gauge-zone" id="gaugeBEERZone">—</div></div>
      </div>
    </div>
    <div class="gauge-card">
      <div class="gauge-title">Market Cap / GDP %</div>
      <div class="gauge-wrap">
        <div class="gauge-arc"><svg viewBox="0 0 72 72"><circle cx="36" cy="36" r="28" fill="none" stroke="#1f2a3c" stroke-width="8" stroke-dasharray="175.9" stroke-linecap="round" transform="rotate(-90 36 36)"/><circle cx="36" cy="36" r="28" fill="none" id="gaugeMCGDPArc" stroke="#ff5c6a" stroke-width="8" stroke-dasharray="175.9" stroke-dashoffset="175.9" stroke-linecap="round" transform="rotate(-90 36 36)"/></svg><div class="gauge-num"><span id="gaugeMCGDPPct">—</span><span class="gauge-pct">pctile</span></div></div>
        <div class="gauge-info"><div class="gauge-val" id="gaugeMCGDPVal">—</div><div class="gauge-median" id="gaugeMCGDPMd">—</div><div class="gauge-zone" id="gaugeMCGDPZone">—</div></div>
      </div>
    </div>
    <div class="gauge-card">
      <div class="gauge-title">Yield Gap (EY − Bond)</div>
      <div class="gauge-wrap">
        <div class="gauge-arc"><svg viewBox="0 0 72 72"><circle cx="36" cy="36" r="28" fill="none" stroke="#1f2a3c" stroke-width="8" stroke-dasharray="175.9" stroke-linecap="round" transform="rotate(-90 36 36)"/><circle cx="36" cy="36" r="28" fill="none" id="gaugeYGArc" stroke="#00d68f" stroke-width="8" stroke-dasharray="175.9" stroke-dashoffset="175.9" stroke-linecap="round" transform="rotate(-90 36 36)"/></svg><div class="gauge-num"><span id="gaugeYGPct">—</span><span class="gauge-pct">pctile</span></div></div>
        <div class="gauge-info"><div class="gauge-val" id="gaugeYGVal">—</div><div class="gauge-median" id="gaugeYGMd">—</div><div class="gauge-zone" id="gaugeYGZone">—</div></div>
      </div>
    </div>
  </div>

  <div class="section-title"><span>◈</span>Nifty 50 Valuation</div>
  <div class="charts-grid">
    <div class="chart-card wide"><div class="chart-header"><div class="chart-title">Nifty 50 Index</div><div class="chart-curr" id="cNifty">—</div></div><canvas id="chartNifty" class="chart-canvas tall"></canvas></div>
    <div class="chart-card"><div class="chart-header"><div class="chart-title">P/E Ratio</div><div class="chart-curr" id="cPE">—</div></div><canvas id="chartPE" class="chart-canvas"></canvas></div>
    <div class="chart-card"><div class="chart-header"><div class="chart-title">Earning Yield %</div><div class="chart-curr" id="cEY">—</div></div><canvas id="chartEY" class="chart-canvas"></canvas></div>
  </div>

  <div class="section-title"><span>◈</span>Market Cap to GDP &amp; BEER</div>
  <div class="charts-grid">
    <div class="chart-card"><div class="chart-header"><div class="chart-title">Market Cap / GDP % (Buffett Indicator)</div><div class="chart-curr" id="cMCGDP">—</div></div><canvas id="chartMCGDP" class="chart-canvas"></canvas></div>
    <div class="chart-card"><div class="chart-header"><div class="chart-title">BEER Ratio — Equity Yield / Bond Yield</div><div class="chart-curr" id="cBEER">—</div></div><canvas id="chartBEER" class="chart-canvas"></canvas></div>
  </div>

  <div class="section-title"><span>◈</span>Bond Yields &amp; Dollar Index</div>
  <div class="charts-grid">
    <div class="chart-card wide"><div class="chart-header"><div class="chart-title">India 10yr vs US 10yr &amp; Yield Gap</div><div class="chart-curr" id="cBond">—</div></div><canvas id="chartBond" class="chart-canvas"></canvas><div class="legend-row"><div class="legend-item"><div class="legend-dot" style="background:#f0b429"></div>India 10yr</div><div class="legend-item"><div class="legend-dot" style="background:#00c8ff"></div>US 10yr</div><div class="legend-item"><div class="legend-dot" style="background:#00d68f"></div>Yield Gap</div></div></div>
    <div class="chart-card"><div class="chart-header"><div class="chart-title">Dollar Index (DXY)</div><div class="chart-curr" id="cDXY">—</div></div><canvas id="chartDXY" class="chart-canvas"></canvas></div>
    <div class="chart-card"><div class="chart-header"><div class="chart-title">USD / INR</div><div class="chart-curr" id="cUSDINR">—</div></div><canvas id="chartUSDINR" class="chart-canvas"></canvas></div>
  </div>
</div>

<script>
const RAW = __CHART_DATA__;
const STATS = __STATS_DATA__;

Chart.defaults.color='#5a7089';
Chart.defaults.font.family="'IBM Plex Mono',monospace";
Chart.defaults.font.size=10;

const GRID={{color:'rgba(31,42,60,0.8)',lineWidth:1}};
const TICK={{color:'#5a7089',maxTicksLimit:6}};

function baseOpts(){{
  return{{responsive:true,maintainAspectRatio:false,interaction:{{mode:'index',intersect:false}},
    plugins:{{legend:{{display:false}},tooltip:{{backgroundColor:'#111620',borderColor:'#1f2a3c',borderWidth:1,titleColor:'#c8d6e5',bodyColor:'#5a7089',padding:10}}}},
    scales:{{x:{{grid:GRID,ticks:{{...TICK,maxRotation:0,maxTicksLimit:8}}}},y:{{grid:GRID,ticks:TICK}}}}
  }};
}}

function makeGrad(ctx,c1){{
  const g=ctx.createLinearGradient(0,0,0,300);
  g.addColorStop(0,c1+'44');g.addColorStop(1,c1+'00');return g;
}}

const charts={{}};
let currentRange=365;

function filterData(n){{
  const len=RAW.dates.length,start=n>=9999?0:Math.max(0,len-n);
  const sl=k=>RAW[k].slice(start);
  return{{dates:sl('dates'),nifty50:sl('nifty50'),pe:sl('pe'),pb:sl('pb'),
    earning_yield:sl('earning_yield'),india_10yr:sl('india_10yr'),us_10yr:sl('us_10yr'),
    yield_gap:sl('yield_gap'),usdinr:sl('usdinr'),dollar_index:sl('dollar_index'),
    beer:sl('beer'),marketcap_gdp:sl('marketcap_gdp'),midcap_pe:sl('midcap_pe'),smallcap_pe:sl('smallcap_pe')}};
}}

function fmtLabels(dates){{
  return dates.map(d=>new Date(d).toLocaleDateString('en-IN',{{day:'2-digit',month:'short',year:'2-digit'}}));
}}

function pctRank(arr,val){{
  return Math.round(arr.filter(x=>x<=val).length/arr.length*100);
}}

function zoneInfo(pct,flip){{
  const cheap=flip?pct>60:pct<30,rich=flip?pct<30:pct>70;
  if(cheap)return['Attractive','zone-cheap'];
  if(rich) return['Stretched','zone-rich'];
  return['Fair Value','zone-fair'];
}}

function setGauge(arcId,pctId,valId,mdId,zoneId,pct,val,mdVal,flip,color){{
  const arc=document.getElementById(arcId);
  arc.style.strokeDashoffset=175.9*(1-pct/100);
  arc.style.stroke=color;
  document.getElementById(pctId).textContent=pct+'%';
  document.getElementById(valId).textContent=typeof val==='number'?val.toFixed(2):val;
  document.getElementById(mdId).textContent='Median: '+mdVal;
  const[zt,zc]=zoneInfo(pct,flip);
  const el=document.getElementById(zoneId);
  el.textContent=zt;el.className='gauge-zone '+zc;
}}

function updateAll(d){{
  const n=d.dates.length-1,p=n>0?n-1:0;
  const nifty=d.nifty50[n],pe=d.pe[n],pb=d.pb[n];
  const ey=d.earning_yield[n],i10=d.india_10yr[n],u10=d.us_10yr[n];
  const yg=d.yield_gap[n],usdinr=d.usdinr[n],dxy=d.dollar_index[n];
  const beer=d.beer[n],mcgdp=d.marketcap_gdp[n];

  const chg=((nifty-d.nifty50[p])/d.nifty50[p]*100).toFixed(2);
  const set=(id,v)=>{{document.getElementById(id).textContent=v}};
  const cls=(id,c)=>{{document.getElementById(id).className='kpi-change '+c}};

  set('kNifty',nifty.toLocaleString('en-IN',{{maximumFractionDigits:0}}));
  set('kNiftyChg',(chg>0?'+':'')+chg+'%'); cls('kNiftyChg',chg>=0?'up':'down');
  set('kPE',pe.toFixed(2)); set('kPEmd','median '+STATS.pe_median);
  set('kPB',pb.toFixed(2)); set('kEY',ey.toFixed(2)+'%');
  set('kIN10',i10.toFixed(2)+'%');
  set('kYG','Gap: '+yg.toFixed(2)+'%'); cls('kYG',yg>0?'up':'down');
  set('kUS10',u10.toFixed(2)+'%');
  set('kSpread','Spread: '+(i10-u10).toFixed(2)+'%');
  set('kUSDINR',usdinr.toFixed(2));
  set('kDXY','DXY '+dxy.toFixed(2));
  set('kMCGDP',mcgdp.toFixed(1)+'%');
  const mcCls=mcgdp>150?'red':mcgdp>120?'gold':'green';
  document.getElementById('kMCGDP').className='kpi-value '+mcCls;

  set('cNifty',nifty.toLocaleString('en-IN',{{maximumFractionDigits:0}}));
  set('cPE',pe.toFixed(2)+'x'); set('cEY',ey.toFixed(2)+'%');
  set('cMCGDP',mcgdp.toFixed(1)+'%'); set('cBEER',beer.toFixed(3));
  set('cBond',i10.toFixed(2)+'% / '+u10.toFixed(2)+'%');
  set('cDXY',dxy.toFixed(2)); set('cUSDINR',usdinr.toFixed(2));

  const peArr=RAW.pe,beerArr=RAW.beer.filter(x=>x>0),mcArr=RAW.marketcap_gdp.filter(x=>x>0),ygArr=RAW.yield_gap;
  setGauge('gaugePEArc','gaugePEPct','gaugePEVal','gaugePEMd','gaugePEZone',pctRank(peArr,pe),pe,'Median: '+STATS.pe_median,false,'#00c8ff');
  setGauge('gaugeBEERArc','gaugeBEERPct','gaugeBEERVal','gaugeBEERMd','gaugeBEERZone',pctRank(beerArr,beer),beer,'Median: '+STATS.beer_median,true,'#f0b429');
  setGauge('gaugeMCGDPArc','gaugeMCGDPPct','gaugeMCGDPVal','gaugeMCGDPMd','gaugeMCGDPZone',pctRank(mcArr,mcgdp),mcgdp,'Median: '+STATS.mcgdp_median+'%',false,'#ff5c6a');
  setGauge('gaugeYGArc','gaugeYGPct','gaugeYGVal','gaugeYGMd','gaugeYGZone',pctRank(ygArr,yg),yg,'Median: '+STATS.yg_median+'%',true,'#00d68f');

  const pePct=pctRank(peArr,pe),beerPct=100-pctRank(beerArr,beer);
  const mcPct=pctRank(mcArr,mcgdp),ygPct=100-pctRank(ygArr,yg);
  const score=Math.round((pePct+beerPct+mcPct+ygPct)/4);
  set('eviScore',score);
  document.getElementById('eviMarker').style.left=score+'%';
  const[hl,desc]=score<35
    ?['Market Attractive — Valuations Compelling','PE, BEER, and yield gap metrics are below historical medians. Historically a strong long-term entry zone.']
    :score<55
    ?['Fairly Valued — Balanced Risk-Reward','Valuations near historical medians. Selective stock-picking remains rewarding.']
    :score<75
    ?['Mildly Stretched — Caution Warranted','Several indicators above median. Prefer quality, limit speculative positions.']
    :['Expensive — High Caution Zone','Valuations in top quartile. Risk-reward unfavourable. Consider reducing equity allocation.'];
  set('eviHeadline',hl); set('eviDesc',desc);
}}

function lineDS(data,color,fill,ctx){{
  return{{data,borderColor:color,backgroundColor:fill&&ctx?makeGrad(ctx,color):'transparent',
    borderWidth:1.5,pointRadius:0,tension:0.3,fill:!!fill}};
}}

function buildCharts(d){{
  const lbl=fmtLabels(d.dates);
  const rebuild=(key,id,datasets,opts)=>{{
    if(charts[key])charts[key].destroy();
    charts[key]=new Chart(document.getElementById(id).getContext('2d'),{{type:'line',data:{{labels:lbl,datasets}},options:opts}});
  }};
  const ctx=id=>document.getElementById(id).getContext('2d');
  const ref=(arr,v)=>arr.map(()=>v);
  const bo=baseOpts();

  rebuild('nifty','chartNifty',[lineDS(d.nifty50,'#00c8ff',true,ctx('chartNifty'))],{{...bo,scales:{{...bo.scales,y:{{...bo.scales.y}}}}}});
  const peOpts={{...bo}};peOpts.scales={{...bo.scales,y:{{...bo.scales.y,min:15,max:30}}}};
  rebuild('pe','chartPE',[lineDS(d.pe,'#00c8ff'),{{data:ref(d.pe,STATS.pe_median),borderColor:'#f0b42944',borderDash:[4,3],borderWidth:1,pointRadius:0,tension:0}}],peOpts);
  rebuild('ey','chartEY',[lineDS(d.earning_yield,'#00d68f',true,ctx('chartEY'))],bo);
  const mcOpts={{...bo}};mcOpts.scales={{...bo.scales,y:{{...bo.scales.y,min:80,max:170}}}};
  rebuild('mcgdp','chartMCGDP',[lineDS(d.marketcap_gdp,'#ff5c6a'),{{data:ref(d.marketcap_gdp,STATS.mcgdp_median),borderColor:'#f0b42944',borderDash:[4,3],borderWidth:1,pointRadius:0,tension:0}}],mcOpts);
  const beerOpts={{...bo}};beerOpts.scales={{...bo.scales,y:{{...bo.scales.y,min:0.5,max:2.0}}}};
  rebuild('beer','chartBEER',[lineDS(d.beer,'#f0b429'),{{data:ref(d.beer,1.0),borderColor:'#ff5c6a44',borderDash:[4,3],borderWidth:1,pointRadius:0,tension:0}},{{data:ref(d.beer,STATS.beer_median),borderColor:'#00d68f44',borderDash:[4,3],borderWidth:1,pointRadius:0,tension:0}}],beerOpts);
  rebuild('bond','chartBond',[lineDS(d.india_10yr,'#f0b429'),lineDS(d.us_10yr,'#00c8ff'),lineDS(d.yield_gap,'#00d68f',true,ctx('chartBond'))],bo);
  rebuild('dxy','chartDXY',[lineDS(d.dollar_index,'#ff8c42',true,ctx('chartDXY'))],bo);
  rebuild('usdinr','chartUSDINR',[lineDS(d.usdinr,'#c084fc',true,ctx('chartUSDINR'))],bo);
}}

document.getElementById('filterRow').addEventListener('click',e=>{{
  const btn=e.target.closest('[data-range]');
  if(!btn)return;
  document.querySelectorAll('.filter-btn').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  const d=filterData(parseInt(btn.dataset.range));
  updateAll(d);buildCharts(d);
}});

(function(){{const d=filterData(365);updateAll(d);buildCharts(d);}})();
</script>
</body>
</html>
"""


# ── 5. Main ────────────────────────────────────────────────────────────────────
def main():
    if not EXCEL.exists():
        print(f"ERROR: Excel file not found at {EXCEL}")
        sys.exit(1)

    print(f"📖  Reading {EXCEL.name} …")
    wb   = load_workbook_safe(EXCEL)
    rows = extract_data(wb)

    if not rows:
        print("ERROR: No data rows found in 'EVI 2025' sheet.")
        sys.exit(1)

    print(f"    {len(rows)} rows  |  {rows[0]['date']} → {rows[-1]['date']}")

    cd    = build_chart_data(rows)
    stats = compute_stats(rows)

    html = HTML_TEMPLATE.format(
        last_date = stats["last_date"],
        date_from = stats["date_from"],
        total_rows= stats["total_rows"],
        built_at  = datetime.datetime.now().strftime("%d %b %Y %H:%M"),
    )
    html = html.replace("__CHART_DATA__", json.dumps(cd))
    html = html.replace("__STATS_DATA__", json.dumps(stats))

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT.write_text(html, encoding="utf-8")

    size = OUTPUT.stat().st_size // 1024
    print(f"✅  Dashboard written → {OUTPUT}  ({size} KB)")
    print(f"    Latest: Nifty {stats['nifty']:,.0f}  |  PE {stats['pe']:.2f}  |  India 10yr {stats['india_10yr']:.2f}%")


if __name__ == "__main__":
    main()
