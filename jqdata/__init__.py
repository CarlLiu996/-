"""
Mock jqdata module for local development and testing.
Provides stub implementations of JoinQuant API functions used by
joinquant_strategy_monitor.py, returning synthetic data so the script
can be imported, linted, and exercised offline.
"""

import numpy as np
import pandas as pd
import datetime as dt

_RNG = np.random.RandomState(42)


class _Column:
    """Mimics an ORM column used in jqdata query()."""

    def __init__(self, name):
        self._name = name

    def in_(self, values):
        return self

    def __ne__(self, other):
        return self

    def __eq__(self, other):
        return self


class _Valuation:
    """Mimics jqdata's valuation ORM table."""
    code = _Column("code")
    pe_ratio = _Column("pe_ratio")
    pb_ratio = _Column("pb_ratio")
    market_cap = _Column("market_cap")


valuation = _Valuation()


class _Query:
    """Mimics a SQLAlchemy-like query object with filter()."""

    def __init__(self, *columns):
        self._columns = columns

    def filter(self, *conditions):
        return self


def query(*args, **kwargs):
    """Stub for jqdata.query()."""
    return _Query(*args, **kwargs)


def get_fundamentals(q, date=None):
    """Return a DataFrame of synthetic fundamentals for ~50 stocks."""
    n = 50
    codes = [f"60{i:04d}.XSHG" for i in range(n)]
    return pd.DataFrame({
        "code": codes,
        "pe_ratio": _RNG.uniform(8, 40, n).round(2),
        "pb_ratio": _RNG.uniform(0.5, 5, n).round(2),
        "market_cap": _RNG.uniform(50, 5000, n).round(2),
    })


def get_price(symbol, count=120, end_date=None, frequency="daily", fields=None):
    """Return synthetic OHLCV / IV data, aligned to a business-day index."""
    if end_date is None:
        end_date = dt.datetime.now().strftime("%Y-%m-%d")
    end = pd.Timestamp(end_date)
    dates = pd.bdate_range(end=end, periods=count)
    n = len(dates)

    base = _RNG.uniform(3000, 5000)
    returns = _RNG.normal(0.0003, 0.012, n)
    closes = base * np.cumprod(1 + returns)

    if fields is None:
        fields = ["open", "close", "high", "low", "volume", "money"]

    data = {}
    for f in fields:
        if f == "close":
            data[f] = closes
        elif f == "open":
            data[f] = closes * _RNG.uniform(0.995, 1.005, n)
        elif f == "high":
            data[f] = closes * _RNG.uniform(1.0, 1.02, n)
        elif f == "low":
            data[f] = closes * _RNG.uniform(0.98, 1.0, n)
        elif f == "volume":
            data[f] = _RNG.uniform(1e8, 5e8, n)
        elif f == "money":
            data[f] = _RNG.uniform(1e10, 5e10, n)
        elif f == "implied_volatility":
            data[f] = _RNG.uniform(0.15, 0.35, n)

    return pd.DataFrame(data, index=dates)


def get_extras(info_type, securities, start_date=None, end_date=None, **kwargs):
    """Return synthetic extras (unit_net_value, futures_positions, etc.)."""
    if end_date is None:
        end_date = dt.datetime.now().strftime("%Y-%m-%d")

    sec = securities[0] if isinstance(securities, list) else securities

    if info_type == "unit_net_value":
        if start_date is None:
            start_date = (pd.Timestamp(end_date) - pd.Timedelta(days=365)).strftime("%Y-%m-%d")
        dates = pd.bdate_range(start=start_date, end=end_date)
        n = len(dates)
        nav = 1.0 + np.cumsum(_RNG.normal(0.0002, 0.005, n))
        return pd.DataFrame({sec: nav}, index=dates)

    elif info_type == "futures_positions":
        dates = pd.bdate_range(end=end_date, periods=252)
        n = len(dates)
        return pd.DataFrame({sec: _RNG.uniform(50000, 200000, n)}, index=dates)

    return pd.DataFrame()


def get_dominant_future(underlying_symbol):
    """Return a plausible dominant futures contract code."""
    mapping = {
        "AU": "AU2506.XSGE",
        "CU": "CU2506.XSGE",
        "RB": "RB2506.XSGE",
        "I": "I2506.XDCE",
        "M": "M2506.XDCE",
        "MA": "MA2506.XZCE",
        "TA": "TA2506.XZCE",
        "SC": "SC2506.XINE",
        "IF": "IF2506.CCFX",
        "IC": "IC2506.CCFX",
        "IM": "IM2506.CCFX",
        "IH": "IH2506.CCFX",
    }
    return mapping.get(underlying_symbol, f"{underlying_symbol}2506.XSGE")


def get_index_stocks(index_code, date=None):
    """Return a list of synthetic stock codes."""
    n = 50
    return [f"60{i:04d}.XSHG" for i in range(n)]
