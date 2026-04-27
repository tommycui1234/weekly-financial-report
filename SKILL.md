---
name: weekly-financial-report
description: Produces a comprehensive weekly cross-asset financial report (Word document) covering commodities, equity indices, China/US bond yields, and FX. Data fetched live from yfinance and akshare; charts generated with matplotlib; 8-item macro news section from Tavily search.
version: 1.0.0
author: tommycui1234
tools:
  - bash
  - tavily
---

# Weekly Financial Report

You are a **senior cross-asset financial analyst** producing a comprehensive weekly report covering commodities, equity indices, China/US bond yields, and FX. The report is delivered as a Word document with embedded charts.

## First-Time Setup

Before running for the first time, ensure the following Python packages are installed:

```bash
pip install yfinance akshare python-docx matplotlib numpy
```

Clone the scripts from the repo:

```bash
git clone https://github.com/tommycui1234/weekly-financial-report.git ~/weekly-financial-report
```

The main script is at `~/weekly-financial-report/scripts/build_full_report.py`.

Output defaults to `~/weekly-report-output/` (created automatically on first run).  
Override via env var: `export WEEKLY_REPORT_DIR=~/my/custom/path`

---

## How to Start

**Auto-calculate all dates from today — no manual input needed.** Run this Python snippet first to derive the report window, then proceed immediately without asking the user:

```python
from datetime import date, timedelta

today = date.today()
# Roll back to the most recent Friday (weekday 4)
days_since_friday = (today.weekday() - 4) % 7
WEEK_END    = today - timedelta(days=days_since_friday)
WEEK_START  = WEEK_END - timedelta(days=4)       # Monday
PREV_FRIDAY = WEEK_END - timedelta(days=7)       # Friday of prior week
YTD_START   = date(WEEK_END.year - 1, 12, 31)   # Dec 31 of prior year

print(WEEK_START, WEEK_END, PREV_FRIDAY, YTD_START)
```

The user may optionally override dates by passing them explicitly after `/weekly-financial-report`.

**Cross-market data note (Option A):** Each asset uses its own last available close from yfinance/akshare. On days when Asian and US markets have different last-close dates (e.g. Monday evening HKT), the change column will reflect each market's own latest data — this is by design. The subtitle says "数据截至各市场最新收盘" to make this clear. The column label is `WTD涨跌幅` (Mon–Thu) or `周涨跌幅` (Fri–Sun).

### Parameters (all optional — auto-calculated if omitted)

| Parameter | Description | Auto value |
|-----------|-------------|------------|
| `WEEK_START` | Report period start (Monday) | Most recent Mon |
| `WEEK_END` | Report period end (Friday) | Most recent Fri |
| `PREV_FRIDAY` | Prior Friday (weekly % base) | WEEK_END − 7 days |
| `YTD_START` | YTD base date | Dec 31 of prior year |
| `OUTPUT_CHARTS_DIR` | Where to save chart PNGs | `~/weekly-report-output/charts/` |
| `OUTPUT_DOCS_DIR` | Where to save the Word document | `~/weekly-report-output/docs/` |
| `LANGUAGE` | Report language | Chinese |

---

## Invocation

```bash
python ~/weekly-financial-report/scripts/build_full_report.py \
  {WEEK_START} {WEEK_END} {PREV_FRIDAY} {YTD_START}
```

Optional path overrides (argv[5] and argv[6]):

```bash
python ~/weekly-financial-report/scripts/build_full_report.py \
  {WEEK_START} {WEEK_END} {PREV_FRIDAY} {YTD_START} \
  ~/my/charts/dir ~/my/docs/dir
```

Or set env var before running:

```bash
export WEEKLY_REPORT_DIR=~/my/report/base
python ~/weekly-financial-report/scripts/build_full_report.py \
  {WEEK_START} {WEEK_END} {PREV_FRIDAY} {YTD_START}
```

---

## Execution Pipeline (run strictly in this order)

### [0/6] Collect Macro News (Tavily MCP)

Search for **8 news items from `WEEK_START`–`WEEK_END`**, strictly split as:
- **China 5 items** (priority order):
  1. China economy & finance: PBOC, NFRA, CSRC, SAFE, SASAC, MoF policies
  2. Real estate: MOHURD, natural resources ministry, major developer news
  3. Tech & industry: semiconductors, AI, new energy
  4. Construction & infrastructure: SOE/state enterprise headlines
  5. Hong Kong & Macau
- **International 3 items** (priority order):
  1. US economy: Fed, key data releases, tariff policy
  2. Geopolitics: Middle East / Russia-Ukraine / Taiwan Strait
  3. Other major economies

**Writing rules for each item:**
- Title ≤ 16 Chinese characters
- Body ~200 characters, objective and neutral
- Cite authoritative source (e.g. "据央行公告", "据统计局数据")
- No sensationalist words: 重磅/震撼/炸裂 are banned
- All 8 items must be from the current report week — no stale news
- Cover distinct topics across the 8 items, no overlap

Run 3 targeted Tavily searches:
1. `"中国经济金融政策 {WEEK_START} {WEEK_END}"`
2. `"Hong Kong Macau economy news {WEEK_START} {WEEK_END}"`
3. `"US Fed economy geopolitics {WEEK_START} {WEEK_END}"`

After collecting the 8 news items, insert them into the `macro_events` list in `build_full_report.py` before running the script, OR pass them via a JSON file as `sys.argv[7]` if the script supports it. (Current version: edit the list directly in the script before running.)

---

### [1/6] Fetch International Data (yfinance)

Handled automatically by the script. Tickers:

```
BZ=F, GC=F, ^IXIC, ^DJI, ^GSPC, ^HSI, 000001.SS, CNY=X, DX-Y.NYB, EURUSD=X, JPY=X
```

Weekly % = (WEEK_END close − PREV_FRIDAY close) / PREV_FRIDAY close × 100.
YTD % = (WEEK_END close − YTD_START close) / YTD_START close × 100.

---

### [2/6] Fetch China Bond Yields (akshare)

Handled automatically. `close` values are already in % (1.764 = 1.764%) — **never multiply by 100**.

---

### [3/6] Fetch US Bond Yields (akshare)

Handled automatically. Same caveat on % values.

---

### [4/6] Fetch Domestic Futures (akshare)

Handled automatically, unconditionally (not inside any conditional branch).

Symbols: `cu0` (SHFE Copper), `al0` (SHFE Aluminium), `rb0` (SHFE Rebar), `jm0` (DCE Coking Coal).

---

### [5/6] Generate Charts (matplotlib)

5 charts generated automatically:

| Chart | Series | Y-axis |
|-------|--------|--------|
| 商品期货 | Brent, Gold, SHFE Copper, SHFE Aluminium, SHFE Rebar, DCE Coking Coal | Cumulative % |
| 股指 | Nasdaq, Dow Jones, S&P 500, HSI, Shanghai Composite | Cumulative % |
| 中债利率 | China 1Y/5Y/10Y/30Y | Absolute yield (%) |
| 美债利率 | US 2Y/5Y/10Y/30Y | Absolute yield (%) |
| 汇率 | USD/CNY, DXY, EUR/USD, USD/JPY | Cumulative % |

---

### [6/6] Generate Word Document (python-docx)

Output filename: `{WEEK_END}_综合周度报告.docx`

Document sections (strict order):

```
Title: 商品期货、股指、债券与汇率 周度报告
Subtitle: 报告期间 | 数据来源 | 数据截至各市场最新收盘

Section 0: 本周宏观事件       ← 8 news items
Section 1: 商品期货           ← table + chart + analysis
Section 2: 股指               ← table + chart + analysis
Section 3: 中债利率           ← table + chart + analysis
Section 4: 美债利率           ← table + chart + analysis
Section 5: 汇率               ← table + chart + analysis

Disclaimer
```

**Table format (7 columns):**

| 品种 | 代码 | 英文名 | 周涨跌幅 | 最新收盘价 | 单位 | 数据来源 |
|------|------|--------|----------|-----------|------|---------|

---

## Known Pitfalls Checklist

- [ ] akshare bond `close` values are already % — never ×100
- [ ] Domestic futures fetched unconditionally
- [ ] All 4 FX pairs in chart and table: USD/CNY, DXY, EUR/USD, USD/JPY
- [ ] Bond/FX Y-axes do NOT start at 0 (auto-focus)
- [ ] Commodity/equity Y-axes start at 0 (cumulative %)
- [ ] All 8 macro news items are from the current report week
- [ ] Table 英文名 column present and matches chart legend labels
- [ ] All table values sourced from variables, not hardcoded
- [ ] matplotlib font: `Arial Unicode MS` (fallback: `Noto CJK`)
- [ ] `BZ=F`: the `=F` suffix is Yahoo Finance's futures identifier, not a math operator
- [ ] Chart titles are empty strings (captions go in Word, not in the image)
- [ ] Dual colour+linestyle encoding applied to all multi-series charts
- [ ] Total runtime estimate: ~1–2 minutes
