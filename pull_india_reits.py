#!/usr/bin/env python3
"""
pull_india_reits.py
-------------------
Download full price history + distribution (dividend) history for all six
listed Indian REITs (Embassy, Mindspace, Brookfield, Nexus, Knowledge Realty,
Bagmane) from Yahoo Finance and write everything to a single Excel workbook.

Note: Bagmane Prime Office REIT listed only in May 2026, so it has a short
history and few (possibly zero) distributions so far.

WHAT YOU GET (one .xlsx file):
  - "Dividends_All"   : every distribution event across all REITs
                        (REIT, Ticker, Ex-Date, Distribution/unit)
  - "Close_Combined"  : raw daily Close of all five REITs aligned by date
                        (handy for your ex-date recovery backtests)
  - one sheet per REIT: full daily OHLCV + Adj Close + Dividends since listing

NOTES / CAVEATS:
  * Yahoo stores each distribution as ONE total per-unit number. It does NOT
    break out the interest / SPV-debt-repayment / dividend components. For the
    tax split you still need each REIT's official distribution notice.
  * 'Close' is the raw (unadjusted) traded price -- correct for ex-date
    recovery tests. 'Adj Close' is back-adjusted for distributions.
  * Yahoo lists these REITs under a few symbol forms. The script auto-resolves
    by trying candidates in order and keeping whichever returns the most data.

USAGE:
  pip install yfinance pandas openpyxl
  python pull_india_reits.py
  # optional: python pull_india_reits.py --out my_reits.xlsx --start 2019-01-01
"""

import argparse
import sys
import time
from datetime import datetime

import pandas as pd

try:
    import yfinance as yf
except ImportError:
    sys.exit("Missing dependency. Run:  pip install yfinance pandas openpyxl")


# REIT name -> candidate Yahoo tickers, tried in order.
# The -RR.NS form is the tradeable NSE unit; .NS and .BO are fallbacks.
REITS = {
    "Embassy Office Parks": ["EMBASSY-RR.NS", "EMBASSY.NS", "EMBASSY.BO"],
    "Mindspace Business Parks": ["MINDSPACE-RR.NS", "MINDSPACE.NS", "MINDSPACE.BO"],
    "Brookfield India": ["BIRET-RR.NS", "BIRET.NS", "BIRET.BO"],
    "Nexus Select Trust": ["NXST-RR.NS", "NXST.NS", "NXST.BO"],
    "Knowledge Realty Trust": ["KRT-RR.NS", "KRT.NS", "KRT.BO"],
    "Bagmane Prime Office": ["BAGMANE-RR.NS", "BAGMANE.NS", "BAGMANE.BO"],
}

# Excel sheet names max 31 chars and can't contain : \ / ? * [ ]
SHEET_NAMES = {
    "Embassy Office Parks": "Embassy",
    "Mindspace Business Parks": "Mindspace",
    "Brookfield India": "Brookfield",
    "Nexus Select Trust": "Nexus",
    "Knowledge Realty Trust": "KnowledgeRealty",
    "Bagmane Prime Office": "Bagmane",
}


def fetch_one(name, candidates, start=None):
    """Try each candidate ticker; return (ticker, DataFrame) for the best hit."""
    best_ticker, best_df = None, None
    for tkr in candidates:
        try:
            df = yf.Ticker(tkr).history(
                period="max" if start is None else None,
                start=start,
                auto_adjust=False,   # keep raw Close + separate Adj Close + Dividends
                actions=True,        # include Dividends / Stock Splits columns
            )
        except Exception as e:
            print(f"    {tkr}: error ({e})")
            continue
        if df is None or df.empty:
            print(f"    {tkr}: no data")
            continue
        rows = len(df)
        print(f"    {tkr}: {rows} rows, {df.index.min().date()} -> {df.index.max().date()}")
        if best_df is None or rows > len(best_df):
            best_ticker, best_df = tkr, df
        time.sleep(0.5)  # be gentle with Yahoo
    return best_ticker, best_df


def clean_prices(df):
    """Tidy a raw yfinance frame into the columns we want to store."""
    df = df.copy()
    df.index = df.index.tz_localize(None)          # drop tz for clean Excel dates
    df.index.name = "Date"
    keep = [c for c in ["Open", "High", "Low", "Close", "Adj Close",
                        "Volume", "Dividends"] if c in df.columns]
    return df[keep].round(4)


def extract_dividends(name, ticker, df):
    """Pull the non-zero distribution rows into a flat table."""
    if "Dividends" not in df.columns:
        return pd.DataFrame(columns=["REIT", "Ticker", "Ex_Date", "Distribution_per_unit"])
    d = df[df["Dividends"] > 0]["Dividends"]
    return pd.DataFrame({
        "REIT": name,
        "Ticker": ticker,
        "Ex_Date": d.index.tz_localize(None).date,
        "Distribution_per_unit": d.values.round(4),
    })


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=f"india_reits_{datetime.now():%Y%m%d}.xlsx")
    ap.add_argument("--start", default=None,
                    help="optional start date YYYY-MM-DD (default = since listing)")
    args = ap.parse_args()

    price_frames, div_frames, close_series = {}, [], {}

    for name, candidates in REITS.items():
        print(f"\n{name}:")
        ticker, raw = fetch_one(name, candidates, start=args.start)
        if raw is None:
            print(f"  !! could not resolve any ticker for {name} -- skipping")
            continue
        prices = clean_prices(raw)
        price_frames[name] = (ticker, prices)
        div_frames.append(extract_dividends(name, ticker, raw))
        if "Close" in prices.columns:
            close_series[name] = prices["Close"]
        print(f"  -> using {ticker}: {len(prices)} rows, "
              f"{int((raw['Dividends'] > 0).sum())} distributions")

    if not price_frames:
        sys.exit("No data pulled for any REIT. Check network / tickers.")

    # Combined close-price matrix (aligned by date, outer join)
    close_combined = pd.DataFrame(close_series).sort_index()
    close_combined.index.name = "Date"

    # All distributions in one table
    all_divs = (pd.concat(div_frames, ignore_index=True)
                .sort_values(["Ex_Date", "REIT"]) if div_frames else pd.DataFrame())

    with pd.ExcelWriter(args.out, engine="openpyxl", datetime_format="yyyy-mm-dd") as xw:
        if not all_divs.empty:
            all_divs.to_excel(xw, sheet_name="Dividends_All", index=False)
        close_combined.to_excel(xw, sheet_name="Close_Combined")
        for name, (ticker, prices) in price_frames.items():
            prices.to_excel(xw, sheet_name=SHEET_NAMES.get(name, name)[:31])

    print(f"\nDone. Wrote {args.out}")
    print(f"  REITs pulled : {len(price_frames)}")
    if not all_divs.empty:
        print(f"  Distributions: {len(all_divs)} events "
              f"({all_divs['Ex_Date'].min()} -> {all_divs['Ex_Date'].max()})")


if __name__ == "__main__":
    main()
