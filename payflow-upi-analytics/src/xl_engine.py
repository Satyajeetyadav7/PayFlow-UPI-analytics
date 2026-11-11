"""Builds an Excel formula string AND its pandas-computed value together.

Every formula cell in the workbook is written with a cached value, so previews (phone, GitHub,
email) show numbers; the cached values are later checked against a full LibreOffice recalculation.
"""
from __future__ import annotations

import numpy as np
import pandas as pd

KEYS = dict(month="month_key", date="txn_date", city="city", tier="tier", cat="category", bank="bank_name",
            btype="bank_type", status="status", ftype="failure_type", reason="failure_reason", net="network",
            age="age_group", hband="hour_band", aband="amount_band", tband="tenure_band", ttype="txn_type",
            hour="txn_hour")
SUMCOLS = dict(amt="amount", fraud="fraud_flag", floss="fraud_loss", cb="cashback", resp="response_ms",
               mau="is_new_user_month")


def L(v):
    """Literal criterion: Excel text/number plus the same value for pandas."""
    return (f'"{v}"' if isinstance(v, str) else str(v), v)


def R(ref: str, pv):
    """Cell-reference criterion: Excel ref plus the value it holds for pandas (None / '*' = any)."""
    return (ref, pv)


class Engine:
    def __init__(self, df: pd.DataFrame):
        self.df = df
        self.cols = {k: df[c].to_numpy() for k, c in KEYS.items()}
        self.sums = {k: df[c].to_numpy(dtype=float) for k, c in SUMCOLS.items()}

    def _mask(self, crits: dict) -> np.ndarray:
        m = np.ones(len(self.df), dtype=bool)
        for k, (_, pv) in crits.items():
            if pv is None or pv == "*":
                continue
            m &= self.cols[k] == pv
        return m

    @staticmethod
    def _args(crits: dict) -> str:
        return ",".join(f"d_{k},{expr}" for k, (expr, _) in crits.items())

    def count(self, **crits):
        return f"COUNTIFS({self._args(crits)})", float(self._mask(crits).sum())

    def sum(self, col: str, **crits):
        return f"SUMIFS(d_{col},{self._args(crits)})", float(self.sums[col][self._mask(crits)].sum())

    def avg(self, col: str, **crits):
        m = self._mask(crits)
        v = float(self.sums[col][m].mean()) if m.any() else 0.0
        return f"IFERROR(AVERAGEIFS(d_{col},{self._args(crits)}),0)", v


def inr_text(x: float) -> str:
    return f"₹{x / 1e7:.2f} Cr" if x >= 1e7 else f"₹{x / 1e5:.1f} L"


def inr_formula(ref: str) -> str:
    return (f'IF({ref}>=10000000,"₹"&TEXT({ref}/10000000,"0.00")&" Cr","₹"&TEXT({ref}/100000,"0.0")&" L")')
