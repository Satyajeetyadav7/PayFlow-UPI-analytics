"""Build the live web dashboard (web/index.html), one self-contained file.

The page needs no server and no libraries: charts are drawn as inline SVG by the page's own script.
Data is embedded as small pre-aggregated cubes (month x city x category x ...), so the three global
filters recompute every chart in the browser.

Run:  python -m src.build_web            (writes web/index.html)
Host: GitHub Pages, Netlify Drop, or any static host.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
FLAT = ROOT / "data" / "payflow_transactions_flat.csv"
WEB = ROOT / "web"


def build_data() -> dict:
    df = pd.read_csv(FLAT, keep_default_na=False)
    months = sorted(df.month_key.unique())
    cities = sorted(df.city.unique())
    cats = sorted(df.category.unique())
    bank_info = df.drop_duplicates("bank_name").set_index("bank_name").bank_type.to_dict()
    banks = sorted(bank_info)
    nets = ["5G", "WiFi", "4G", "Weak (2G/3G)"]
    reasons = sorted(df[df.status == "Failed"].failure_reason.unique())
    dates = sorted(df.txn_date.unique())
    idx = lambda lst: {v: i for i, v in enumerate(lst)}  # noqa: E731
    mi, ci, ki, bi, ni, ri, di = idx(months), idx(cities), idx(cats), idx(banks), idx(nets), idx(reasons), idx(dates)

    d = df.copy()
    d["m"], d["c"], d["k"] = d.month_key.map(mi), d.city.map(ci), d.category.map(ki)
    d["b"], d["n"], d["h"] = d.bank_name.map(bi), d.network.map(ni), d.txn_hour
    d["succ"] = (d.status == "Success").astype(int)
    d["tpv"] = d.amount * d.succ
    d["sys"] = (d.failure_type == "System").astype(int)
    d["user"] = (d.failure_type == "User").astype(int)
    d["risk"] = (d.failure_type == "Risk").astype(int)
    d["sysval"] = d.amount * d.sys
    d["one"] = 1
    rows = lambda g: g.reset_index().values.tolist()  # noqa: E731

    core = rows(d.groupby(["m", "c", "k"]).agg(t=("one", "sum"), s=("succ", "sum"), v=("tpv", "sum"), sy=("sys", "sum"), u=("user", "sum"), r=("risk", "sum"),
                                              f=("fraud_flag", "sum"), fl=("fraud_loss", "sum"), cb=("cashback", "sum"), rs=("response_ms", "sum"), sv=("sysval", "sum")))
    core = [[int(x) if i < 11 or i == 12 else round(float(x), 2) for i, x in enumerate(r)] for r in core]
    bank = rows(d.groupby(["b", "m", "c", "k"]).agg(t=("one", "sum"), s=("succ", "sum"), sy=("sys", "sum"), sv=("sysval", "sum")))
    hour = rows(d.groupby(["h", "m", "c", "k"]).agg(t=("one", "sum"), sy=("sys", "sum"), rs=("response_ms", "sum")))
    net = rows(d.groupby(["n", "m", "c", "k"]).agg(t=("one", "sum"), sy=("sys", "sum")))
    failed = d[d.status == "Failed"].copy()
    failed["r"] = failed.failure_reason.map(ri)
    reason = rows(failed.groupby(["r", "m", "c", "k"]).agg(n=("one", "sum")))

    band_defs = {"h": ("hour_band", sorted(df.hour_band.unique())), "a": ("amount_band", sorted(df.amount_band.unique())), "t": ("tenure_band", sorted(df.tenure_band.unique()))}
    bands = {}
    for key, (col, labels) in band_defs.items():
        g = d.assign(x=d[col].map(idx(labels))).groupby(["x", "m", "c", "k"]).agg(t=("one", "sum"), f=("fraud_flag", "sum"), fl=("fraud_loss", "sum"))
        bands[key] = {"labels": labels, "rows": rows(g)}

    daily = rows(d.assign(dd=d.txn_date.map(di)).groupby(["b", "dd"]).agg(t=("one", "sum"), sy=("sys", "sum")))

    def clean(x):
        return [[int(v) if float(v).is_integer() else round(float(v), 2) for v in r] for r in x]
    return {
        "dims": {"months": months, "cities": cities, "cats": cats, "banks": banks, "bankType": [bank_info[b] for b in banks], "nets": nets, "reasons": reasons, "dates": dates},
        "core": core, "bank": clean(bank), "hour": clean(hour), "net": clean(net), "reason": clean(reason),
        "bands": {k: {"labels": v["labels"], "rows": clean(v["rows"])} for k, v in bands.items()}, "daily": clean(daily),
        "meta": {"txns": int(len(df)), "customers": int(df.customer_id.nunique()), "first": months[0], "last": months[-1]},
    }


def build() -> Path:
    data = build_data()
    payload = json.dumps(data, separators=(",", ":"), ensure_ascii=False).replace("</", "<\\/")
    fragment = (WEB / "template.html").read_text().replace("__DATA__", payload)
    WEB.mkdir(exist_ok=True)
    page = ('<!doctype html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n'
            '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n</head>\n<body>\n'
            + fragment + "\n</body>\n</html>\n")
    (WEB / "index.html").write_text(page)
    if len(sys.argv) > 1:
        Path(sys.argv[1]).write_text(fragment)
    print(f"[build_web] web/index.html ({(WEB / 'index.html').stat().st_size / 1e6:.2f} MB)")
    return WEB / "index.html"


if __name__ == "__main__":
    build()
