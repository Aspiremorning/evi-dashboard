# EVI Dashboard — Equity Valuation Index

Auto-regenerating India market valuation dashboard.  
You update the Excel file → run one command → live dashboard refreshes on GitHub Pages.

---

## Daily Workflow (once set up)

```
1. Open data/EVI_2025-26.xlsx
2. Add today's row in the "EVI 2025" sheet
3. Save and close
4. Run:  ./update.sh
```

That's it. Dashboard is live in ~30 seconds.

---

## One-Time Setup

### Step 1 — Clone or create this repo on GitHub

1. Go to https://github.com/new
2. Create a repo named `evi-dashboard` (keep it public for free GitHub Pages)
3. Copy your repo URL, e.g. `https://github.com/YOUR_USERNAME/evi-dashboard`

### Step 2 — Set up the project on your Mac/Linux

```bash
# Move into this folder (wherever you saved it)
cd /path/to/evi-dashboard

# Initialise git
git init
git add .
git commit -m "Initial commit"

# Connect to GitHub and push
git remote add origin https://github.com/YOUR_USERNAME/evi-dashboard.git
git branch -M main
git push -u origin main
```

### Step 3 — Enable GitHub Pages

1. Go to your repo on GitHub → **Settings** → **Pages**
2. Source: **Deploy from a branch**
3. Branch: `main` / folder: `/docs`
4. Click **Save**

Your dashboard will be live at:
```
https://YOUR_USERNAME.github.io/evi-dashboard/
```

### Step 4 — Make the update script executable (one time)

```bash
chmod +x update.sh
```

### Step 5 — Install Python dependency (one time)

```bash
pip install openpyxl
```

---

## Project Structure

```
evi-dashboard/
├── data/
│   └── EVI_2025-26.xlsx       ← Edit this daily
├── scripts/
│   └── build.py               ← Reads Excel → generates HTML
├── docs/
│   └── index.html             ← Auto-generated dashboard (GitHub Pages serves this)
├── .github/
│   └── workflows/
│       └── build.yml          ← Auto-rebuilds when you push Excel to GitHub
├── update.sh                  ← One-command: build + commit + push
├── requirements.txt
└── README.md
```

---

## How It Works

```
You edit Excel
      │
      ▼
./update.sh
      │
      ├─ python scripts/build.py   ← reads EVI 2025 sheet, extracts all rows
      │                              generates docs/index.html with data embedded
      │
      └─ git add → git commit → git push
                                    │
                                    ▼
                             GitHub Actions
                             (runs build.py again server-side as backup)
                                    │
                                    ▼
                             GitHub Pages serves
                             docs/index.html live
```

---

## Troubleshooting

**"openpyxl not found"**
```bash
pip install openpyxl
# or
pip3 install openpyxl
```

**"Permission denied: ./update.sh"**
```bash
chmod +x update.sh
```

**Dashboard not updating after push**
- Check GitHub Actions tab in your repo for build errors
- Try: `python scripts/build.py` locally first to confirm it works

**Build script can't find Excel**
- Make sure the file is at `data/EVI_2025-26.xlsx` (exact name matters)

---

## Columns Read from "EVI 2025" Sheet

| Col | Field             |
|-----|-------------------|
| B   | Date              |
| C   | Market Cap (INR)  |
| E   | USD/INR           |
| G   | Market Cap / GDP  |
| H   | BEER Ratio        |
| I   | Earning Yield     |
| J   | EPS               |
| K   | India 10yr Yield  |
| L   | Nifty 50          |
| M   | Price to Book     |
| N   | Yield Gap         |
| O   | P/E Ratio         |
| P   | US 10yr Bond      |
| R   | Dollar Index      |
| S   | PREITY Ratio      |
| T   | 91-day T-Bill     |
| U   | Nifty Midcap 150  |
| V   | Midcap 150 PE     |
| X   | Midcap Earn Yield |
| Y   | Nifty Smallcap 250|
| Z   | Smallcap 250 PE   |
| AB  | Smallcap EY       |
