# fund-portfolio-backtester
A Python engine for historical Dollar Cost Averaging (DCA) backtesting and performance comparison of TEFAS and BEFAS (Pension) funds.

# TEFAS & BEFAS (Pension) DCA Backtest Engine

A flexible Python engine that fetches historical data for Turkish mutual funds (TEFAS) and pension funds (BEFAS) via API, designed to accurately backtest Dollar Cost Averaging (DCA) strategies and portfolio performances.

## Why This Engine?
Standard fund scraping tools usually default to the 1st day of the month or use static prices. However, in reality, investors make purchases on specific days (e.g., when they receive their salary on the 15th). 
This engine:
* Fetches real historical prices directly from the official API.
* Allows targeting **specific days of the month** (e.g., 15th).
* Automatically detects holidays/weekends and shifts to the **next available trading day**.
* Exports the aggregated data into highly readable, formatted Excel (`.xlsx`) pivot matrices using `openpyxl`.

## Installation

To set up the environment, simply run the following command in your terminal:

```bash
pip install -r requirements.txt
