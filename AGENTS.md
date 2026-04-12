# AGENTS.md

## Cursor Cloud specific instructions

### Overview

This repository contains a single-file Python script (`joinquant_strategy_monitor.py`) — a **Strategy Environment Monitoring System** for the [JoinQuant (聚宽)](https://www.joinquant.com/) quantitative trading platform. It scores 7 Chinese private-fund strategy types (主观多头, 量化多头, CTA, ETF套利, 股指套利, 期权套利, 市场中性) based on market indicators and generates an Excel report.

### Platform dependency

The script is designed for JoinQuant's cloud research environment and imports `from jqdata import *`. A local **mock `jqdata` package** (`jqdata/__init__.py`) is provided so the script can be imported, linted, and run locally with synthetic data. The mock stubs `get_price`, `get_extras`, `get_dominant_future`, `get_index_stocks`, `get_fundamentals`, `query`, and `valuation`.

### Running locally

```bash
python3 -c "from joinquant_strategy_monitor import run_strategy_monitor; run_strategy_monitor(output_excel=True)"
```

This exercises all 4 classes (`DataFetcher`, `IndicatorCalculator`, `StrategyScorer`, `ExcelReportGenerator`) end-to-end with synthetic data and produces an `.xlsx` report in the current directory.

### Linting

```bash
flake8 --max-line-length=120 joinquant_strategy_monitor.py
python3 -m py_compile joinquant_strategy_monitor.py
```

Most existing flake8 warnings are style issues (trailing whitespace, bare excepts) and `F403`/`F405` from the `jqdata` star import — these are inherent to JoinQuant's platform conventions.

### Key caveats

- **No automated tests exist.** There is no test framework or CI configuration.
- The `jqdata` mock returns deterministic data via `np.random.RandomState(42)`. Scores will vary from real JoinQuant output but the full pipeline runs.
- The Excel report file is written to the current working directory. Clean up after demo runs.
