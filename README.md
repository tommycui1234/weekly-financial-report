# Weekly Financial Report

A Claude Code skill that produces a comprehensive weekly cross-asset financial report as a Word document.

**Coverage:** Commodity futures · Equity indices · China/US bond yields · FX  
**Data sources:** yfinance · akshare  
**Charts:** matplotlib (headless PNG, embedded in docx)  
**Macro news:** Tavily real-time search (8 items per week)

---

## Quick Start

### 1. Install Python dependencies

```bash
pip install yfinance akshare python-docx matplotlib numpy
```

### 2. Clone this repo

```bash
git clone https://github.com/tommycui1234/weekly-financial-report.git ~/weekly-financial-report
```

### 3. Run directly

```bash
python ~/weekly-financial-report/scripts/build_full_report.py
# Dates are auto-calculated. Output goes to ~/weekly-report-output/
```

Or pass dates explicitly:

```bash
python ~/weekly-financial-report/scripts/build_full_report.py \
  2026-04-20 2026-04-24 2026-04-17 2025-12-31
```

### 4. Use as a Claude Code skill

Install the skill:

```bash
npx skills add tommycui1234/weekly-financial-report
```

Then in Claude Code:

```
/weekly-financial-report
```

---

## Output

```
~/weekly-report-output/
├── docs/
│   └── 2026-04-24_综合周度报告.docx   ← Word report
└── charts/
    ├── commodity_weekly_cumulative.png
    ├── index_weekly_cumulative.png
    ├── china_bond_yield_change.png
    ├── us_bond_yield_change.png
    └── fx_combined.png
```

Override the output location:

```bash
# Via env var (persistent across runs)
export WEEKLY_REPORT_DIR=~/my/reports
python ~/weekly-financial-report/scripts/build_full_report.py

# Via argv (one-off override)
python ~/weekly-financial-report/scripts/build_full_report.py \
  2026-04-20 2026-04-24 2026-04-17 2025-12-31 \
  ~/my/charts ~/my/docs
```

---

## Script Arguments

| Position | Arg | Description | Default |
|----------|-----|-------------|---------|
| argv[1] | `WEEK_START` | Report period start (Monday) | Auto-calculated |
| argv[2] | `WEEK_END` | Report period end | Today or last Friday |
| argv[3] | `PREV_FRIDAY` | Prior Friday (weekly % base) | WEEK_START − 3 days |
| argv[4] | `YTD_START` | YTD base date | Dec 31 of prior year |
| argv[5] | `OUTPUT_CHARTS_DIR` | Chart PNG output dir | `~/weekly-report-output/charts` |
| argv[6] | `OUTPUT_DOCS_DIR` | Word document output dir | `~/weekly-report-output/docs` |

---

## Report Structure

```
Title: 商品期货、股指、债券与汇率 周度报告
Subtitle: 报告期间 | 数据截至各市场最新收盘

一、本周宏观事件   (8 items — China 5, International 3)
二、商品期货       (table + chart + analysis)
三、股指           (table + chart + analysis)
四、中债利率       (table + chart + analysis)
五、美债利率       (table + chart + analysis)
六、汇率           (table + chart + analysis)

Disclaimer
```

---

## Mid-Week Behaviour

The report can be run any day of the week:

- **Column label:** `WTD涨跌幅` (Mon–Thu) · `周涨跌幅` (Fri–Sun)
- **Each asset** uses its own last available close (Option A: cross-market dates may differ)
- **Subtitle** notes: "数据截至各市场最新收盘，不同市场收盘日期可能不同"

---

## Requirements

- Python 3.9+
- `yfinance >= 0.2`
- `akshare >= 1.12`
- `python-docx >= 1.1`
- `matplotlib >= 3.8`
- `numpy`
- Claude Code with Tavily MCP (for macro news section)

---

## Font Note

Charts use `Arial Unicode MS` for CJK character support. If not available on your system, install `noto-fonts-cjk` or the script will fall back to the system default (Chinese characters may not render).

---

## License

MIT
