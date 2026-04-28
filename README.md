# Global Macro and Cross-Asset Weekly
## 环球宏观与多资产周度观察

A cross-platform agent skill that produces a comprehensive weekly cross-asset financial report as a Word document.

**Coverage:** Commodity futures · Equity indices · China/US bond yields · FX  
**Data sources:** yfinance · akshare  
**Charts:** matplotlib (headless PNG, embedded in .docx)  
**Macro news:** Tavily real-time search (8 items per week)  
**Output language:** Chinese (report body and section headings)

Compatible with **Claude Code**, **Hermes**, and **OpenClaw**.

---

## Installation

### Python dependencies (required on all platforms)

```bash
pip install yfinance akshare python-docx matplotlib numpy
```

Clone the scripts:

```bash
git clone https://github.com/tommycui1234/weekly-financial-report.git ~/weekly-financial-report
```

---

### Claude Code

```bash
npx skills add tommycui1234/weekly-financial-report
```

Then invoke:

```
/weekly-financial-report
```

---

### Hermes

Install from GitHub:

```bash
hermes skills install tommycui1234/weekly-financial-report
```

Or install from raw URL:

```bash
hermes skills install \
  https://raw.githubusercontent.com/tommycui1234/weekly-financial-report/main/SKILL.md \
  --name weekly-financial-report
```

Or add as a tap for browsing alongside other skills:

```bash
hermes skills tap add tommycui1234/weekly-financial-report
hermes skills install tommycui1234/weekly-financial-report/weekly-financial-report
```

Invoke in Hermes with the skill name: `weekly-financial-report`

---

### OpenClaw

OpenClaw does not have a GitHub-direct install command. Copy the skill manually:

```bash
mkdir -p ~/.openclaw/workspace/skills/weekly-financial-report
curl -sSL https://raw.githubusercontent.com/tommycui1234/weekly-financial-report/main/SKILL.md \
  -o ~/.openclaw/workspace/skills/weekly-financial-report/SKILL.md
```

Then reload the gateway:

```bash
openclaw gateway restart
# or type /new in chat
```

Verify:

```bash
openclaw skills list | grep weekly
```

---

## Output

```
~/weekly-report-output/          # default; override with WEEKLY_REPORT_DIR env var
├── docs/
│   └── 2026-04-24_综合周度报告.docx
└── charts/
    ├── commodity_weekly_cumulative.png
    ├── index_weekly_cumulative.png
    ├── china_bond_yield_change.png
    ├── us_bond_yield_change.png
    └── fx_combined.png
```

Override the output location:

```bash
# Persistent: set env var before running
export WEEKLY_REPORT_DIR=~/my/reports

# One-off: pass argv[5] and argv[6]
python ~/weekly-financial-report/scripts/build_full_report.py \
  2026-04-20 2026-04-24 2026-04-17 2025-12-31 \
  ~/my/charts ~/my/docs
```

---

## Report Structure

```
Title: 商品期货、股指、债券与汇率 周度报告
Subtitle: 报告期间 | 数据截至各市场最新收盘，不同市场收盘日期可能不同

一、本周宏观事件   (8 items — China 5, International 3)
二、商品期货       (7-column table + chart + analysis)
三、股指           (7-column table + chart + analysis)
四、中债利率       (table + chart + analysis)
五、美债利率       (table + chart + analysis)
六、汇率           (7-column table + chart + analysis)

Disclaimer
```

---

## Running on Any Day of the Week

The report can be run Mon–Sun. Dates are auto-calculated:

| Run day | WEEK_END | Change label |
|---------|----------|--------------|
| Mon–Thu | Today | `WTD涨跌幅` |
| Fri | Today (Friday) | `周涨跌幅` |
| Sat–Sun | Last Friday | `周涨跌幅` |

Each asset uses its own last available close (Option A). Asian and US markets may have different last-close dates on mid-week runs — the subtitle notes this.

---

## Script Arguments

```bash
python ~/weekly-financial-report/scripts/build_full_report.py \
  [WEEK_START] [WEEK_END] [PREV_FRIDAY] [YTD_START] \
  [OUTPUT_CHARTS_DIR] [OUTPUT_DOCS_DIR]
```

| Position | Arg | Default |
|----------|-----|---------|
| argv[1] | WEEK_START | Auto (Monday of current week) |
| argv[2] | WEEK_END | Auto (today or last Friday) |
| argv[3] | PREV_FRIDAY | Auto (WEEK_START − 3 days) |
| argv[4] | YTD_START | Auto (Dec 31 of prior year) |
| argv[5] | OUTPUT_CHARTS_DIR | `~/weekly-report-output/charts` |
| argv[6] | OUTPUT_DOCS_DIR | `~/weekly-report-output/docs` |

---

## Requirements

- Python 3.9+
- `yfinance >= 0.2`
- `akshare >= 1.12`
- `python-docx >= 1.1`
- `matplotlib >= 3.8`
- `numpy`
- Tavily MCP server (for macro news; required by Claude Code, Hermes, and OpenClaw via their Tavily integrations)

---

## Font Note

Charts use `Arial Unicode MS` for CJK character support. On Linux systems without it, install `noto-fonts-cjk` or equivalent — otherwise Chinese labels in charts will not render.

---

## License

MIT
