"""Reproducible SIMULATED UPI payments dataset for the fictional fintech "PayFlow".

No real bank, company or person is represented. Bank names are invented. Realistic rules are
built in so the analysis has genuine signal to find:

  * System failures depend on the bank (small-finance and co-op banks are weakest), the hour
    (evening peak), the network (weak signal), festive load and a gradual infrastructure
    improvement over time. Two bank-outage incidents are injected so anomaly spotting works.
  * User-side failures (insufficient balance, wrong PIN, limit exceeded) depend on age group,
    customer tenure, amount and month-end.
  * Fraud flags rise at night, on large amounts, for new customers and on P2P transfers.
  * Cashback is paid on some successful merchant payments; rates differ by category.
  * Volume grows ~5.5% a month with festive, weekend and salary-day effects.

Outputs star-schema CSVs (for Power BI) and one flat table (for Excel).

Run:  python -m src.generate_data
"""
from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
PBI_DIR = ROOT / "powerbi" / "data"

SEED = 7
N_CUSTOMERS = 5_000
N_TXNS = 60_000
N_MERCHANTS = 420
START = date(2025, 1, 1)
END = date(2026, 10, 31)

# ---------------------------------------------------------------- reference data
CITIES = [  # city, tier, region, share of customers, weak-network share
    ("Mumbai", "Tier 1", "West", 0.14, 0.05), ("Delhi NCR", "Tier 1", "North", 0.15, 0.05),
    ("Bengaluru", "Tier 1", "South", 0.13, 0.04), ("Hyderabad", "Tier 1", "South", 0.09, 0.05),
    ("Chennai", "Tier 1", "South", 0.07, 0.06), ("Pune", "Tier 1", "West", 0.07, 0.05),
    ("Kolkata", "Tier 1", "East", 0.06, 0.08), ("Jaipur", "Tier 2", "North", 0.07, 0.12),
    ("Lucknow", "Tier 2", "North", 0.07, 0.14), ("Indore", "Tier 2", "Central", 0.05, 0.12),
    ("Patna", "Tier 2", "East", 0.05, 0.16), ("Coimbatore", "Tier 2", "South", 0.05, 0.11),
]

BANKS = [  # name, type, customer share, base system-failure rate
    ("Orion Private Bank", "Private", 0.17, 0.014), ("Vertex Private Bank", "Private", 0.14, 0.017),
    ("Lakeview Private Bank", "Private", 0.09, 0.020), ("Citadel National Bank", "Public", 0.16, 0.030),
    ("Dominion National Bank", "Public", 0.11, 0.036), ("Pioneer National Bank", "Public", 0.08, 0.044),
    ("Crestline National Bank", "Public", 0.06, 0.033), ("Sunrise Small Finance Bank", "Small Finance", 0.07, 0.058),
    ("Greenfield Small Finance Bank", "Small Finance", 0.05, 0.066), ("Alliance Co-operative Bank", "Co-operative", 0.07, 0.075),
]
BANK_TYPE_RESP = {"Private": 1.0, "Public": 1.15, "Small Finance": 1.3, "Co-operative": 1.45}

CATEGORIES = {  # weight, median amount, sigma, cashback rate, peak hour centre
    "P2P Transfer": dict(w=0.22, med=800, sig=1.10, cb=0.0, hr=(11, 20)),
    "Grocery & Kirana": dict(w=0.17, med=220, sig=0.80, cb=0.010, hr=(9, 18)),
    "Food Delivery": dict(w=0.11, med=340, sig=0.55, cb=0.030, hr=(13, 20)),
    "Mobile Recharge": dict(w=0.07, med=299, sig=0.45, cb=0.015, hr=(11, 19)),
    "Utilities & Bills": dict(w=0.08, med=1100, sig=0.80, cb=0.005, hr=(11, 17)),
    "Shopping & Fashion": dict(w=0.08, med=1400, sig=0.90, cb=0.025, hr=(14, 21)),
    "Travel & Transport": dict(w=0.07, med=650, sig=1.00, cb=0.015, hr=(8, 18)),
    "Fuel": dict(w=0.05, med=900, sig=0.55, cb=0.000, hr=(8, 18)),
    "Entertainment": dict(w=0.05, med=380, sig=0.70, cb=0.040, hr=(19, 22)),
    "Healthcare & Pharmacy": dict(w=0.05, med=480, sig=0.90, cb=0.010, hr=(10, 19)),
    "Education": dict(w=0.05, med=2400, sig=0.80, cb=0.000, hr=(10, 16)),
}
CAT_NAMES = list(CATEGORIES)
RECHARGE_PLANS = np.array([149, 199, 239, 299, 349, 479, 599, 719])

FESTIVE = [  # start, end, name, load multiplier
    (date(2025, 1, 20), date(2025, 1, 26), "Republic Day Sale", 1.15),
    (date(2025, 8, 10), date(2025, 8, 17), "Independence Day Sale", 1.15),
    (date(2025, 10, 13), date(2025, 10, 23), "Diwali 2025", 1.45),
    (date(2026, 1, 20), date(2026, 1, 26), "Republic Day Sale", 1.15),
    (date(2026, 8, 10), date(2026, 8, 17), "Independence Day Sale", 1.15),
    (date(2026, 10, 25), date(2026, 10, 31), "Diwali 2026 run-up", 1.40),
]
INCIDENTS = [  # bank index, start, end (inclusive), failure multiplier
    (4, date(2025, 10, 13), date(2025, 10, 15), 12.0),   # Dominion National - Diwali-week outage
    (8, date(2026, 3, 3), date(2026, 3, 5), 10.0),       # Greenfield SFB - core-banking upgrade
]
NETWORKS = ["5G", "4G", "WiFi", "Weak (2G/3G)"]
AGE_GROUPS = ["18-24", "25-34", "35-44", "45+"]
FIRST = ["Sharma", "Verma", "Iyer", "Reddy", "Patel", "Khan", "Das", "Nair", "Gupta", "Singh", "Joshi", "Mehta",
         "Rao", "Banerjee", "Kulkarni", "Chopra", "Menon", "Pillai", "Yadav", "Bhatt"]
SHOP = {"Grocery & Kirana": ["General Store", "Kirana", "Supermart", "Fresh Mart"],
        "Food Delivery": ["Kitchen", "Biryani House", "Tiffin Centre", "Cafe"],
        "Mobile Recharge": ["Mobile Hub", "Recharge Point"], "Utilities & Bills": ["Electricity Board", "Gas Agency", "Water Works", "Broadband"],
        "Shopping & Fashion": ["Fashion", "Electronics", "Footwear", "Home Decor"], "Travel & Transport": ["Cabs", "Travels", "Metro Pass", "Bus Tickets"],
        "Fuel": ["Fuel Station", "Petro Point"], "Entertainment": ["Cinemas", "Play Zone", "Streaming"],
        "Healthcare & Pharmacy": ["Pharmacy", "Diagnostics", "Clinic"], "Education": ["Academy", "Tutorials", "Coaching Centre"]}


def festive_for(d: date):
    for a, b, name, mult in FESTIVE:
        if a <= d <= b:
            return name, mult
    return "", 1.0


def make_calendar() -> pd.DataFrame:
    days = pd.date_range(START, END, freq="D")
    rows = []
    for ts in days:
        d = ts.date()
        name, _ = festive_for(d)
        rows.append({
            "date_key": int(ts.strftime("%Y%m%d")), "date": d.isoformat(), "year": ts.year, "quarter": f"Q{ts.quarter}",
            "month_num": ts.month, "month_name": ts.strftime("%b"), "month_key": ts.strftime("%Y-%m"),
            "day_of_month": ts.day, "day_name": ts.strftime("%a"), "dow": ts.weekday(),
            "is_weekend": int(ts.weekday() >= 5), "is_festive": int(bool(name)), "festive_name": name,
            "is_salary_window": int(ts.day <= 3 or ts.day >= 28),
        })
    return pd.DataFrame(rows)


def hour_weights(centres: tuple[int, int]) -> np.ndarray:
    h = np.arange(24)
    w = 0.25 + sum(np.exp(-0.5 * ((h - c) / 2.2) ** 2) for c in centres)
    w[0:5] *= 0.35
    return w / w.sum()


def make_dims(rng: np.random.Generator):
    cities = pd.DataFrame(CITIES, columns=["city", "tier", "region", "share", "weak"])
    cities.insert(0, "city_id", np.arange(1, len(cities) + 1))
    banks = pd.DataFrame(BANKS, columns=["bank_name", "bank_type", "share", "fail"])
    banks.insert(0, "bank_id", np.arange(1, len(banks) + 1))

    span = (END - START).days
    signup = (rng.beta(1.3, 1.0, N_CUSTOMERS) * (span - 5)).astype(int)
    city_idx = rng.choice(len(cities), N_CUSTOMERS, p=cities.share / cities.share.sum())
    age = rng.choice(AGE_GROUPS, N_CUSTOMERS, p=[0.30, 0.38, 0.20, 0.12])
    customers = pd.DataFrame({
        "customer_id": 200_000 + np.arange(1, N_CUSTOMERS + 1),
        "signup_date": [(START + timedelta(days=int(o))).isoformat() for o in signup],
        "city_id": cities.city_id.values[city_idx], "age_group": age,
        "device_os": rng.choice(["Android", "iOS"], N_CUSTOMERS, p=[0.78, 0.22]),
        "primary_bank_id": rng.choice(banks.bank_id, N_CUSTOMERS, p=banks.share / banks.share.sum()),
        "kyc_level": rng.choice(["Min KYC", "Full KYC"], N_CUSTOMERS, p=[0.28, 0.72]),
    })

    m_rows, mid = [], 1
    m_city = rng.choice(len(cities), N_MERCHANTS, p=cities.share / cities.share.sum())
    cat_w = np.array([CATEGORIES[c]["w"] for c in CAT_NAMES[1:]])
    m_cat = rng.choice(CAT_NAMES[1:], N_MERCHANTS, p=cat_w / cat_w.sum())
    m_size = rng.choice(["Micro", "Small", "Medium", "Large"], N_MERCHANTS, p=[0.45, 0.33, 0.16, 0.06])
    m_rows.append({"merchant_id": 0, "merchant_name": "P2P (Individual)", "category": "P2P Transfer", "city_id": 0, "merchant_size": "N/A"})
    for i in range(N_MERCHANTS):
        mid += 0
        base = SHOP[m_cat[i]][int(rng.integers(0, len(SHOP[m_cat[i]])))]
        m_rows.append({"merchant_id": 9000 + i + 1, "merchant_name": f"{FIRST[int(rng.integers(0, len(FIRST)))]} {base}",
                       "category": m_cat[i], "city_id": int(cities.city_id.values[m_city[i]]), "merchant_size": m_size[i]})
    merchants = pd.DataFrame(m_rows)
    return cities, banks, customers, merchants


def build_transactions(rng, cities, banks, customers, merchants, calendar) -> pd.DataFrame:
    n = N_TXNS
    dates = pd.to_datetime(calendar["date"])
    m_idx = ((dates.dt.year - START.year) * 12 + dates.dt.month - START.month).values
    fest_mult = np.array([festive_for(d.date())[1] for d in dates])
    day_w = (1.055 ** m_idx) * fest_mult * np.where(dates.dt.weekday >= 5, 1.12, 1.0) * np.where(dates.dt.day <= 3, 1.12, 1.0)
    day_i = rng.choice(len(calendar), n, p=day_w / day_w.sum())
    day_i.sort()
    t_date = dates.iloc[day_i].reset_index(drop=True)
    t_m = m_idx[day_i]
    t_dom = t_date.dt.day.values
    t_dow = t_date.dt.weekday.values
    t_month = t_date.dt.month.values
    t_fest = fest_mult[day_i] > 1.0

    # customers: eligible once signed up, weighted by activity propensity
    order = np.argsort(customers.signup_date.values, kind="stable")
    cust = customers.iloc[order].reset_index(drop=True)
    sign_d = pd.to_datetime(cust.signup_date)
    propensity = rng.lognormal(0, 0.9, len(cust))
    cum = np.cumsum(propensity)
    k = np.searchsorted(sign_d.values, t_date.values, side="right").clip(min=1)
    u = rng.random(n) * cum[k - 1]
    ci = np.minimum(np.searchsorted(cum, u), k - 1)
    c = cust.iloc[ci].reset_index(drop=True)
    tenure = (t_date.values - pd.to_datetime(c.signup_date).values).astype("timedelta64[D]").astype(int)

    # category (weights shift with calendar)
    W = np.tile([CATEGORIES[x]["w"] for x in CAT_NAMES], (n, 1)).astype(float)
    col = {nm: i for i, nm in enumerate(CAT_NAMES)}
    W[:, col["Utilities & Bills"]] *= np.where(t_dom <= 10, 2.0, 0.55)
    W[:, col["Mobile Recharge"]] *= np.where((t_dom <= 5) | (t_dom >= 26), 1.6, 0.8)
    W[:, col["P2P Transfer"]] *= np.where((t_dom <= 3) | (t_dom >= 28), 1.5, 1.0)
    W[:, col["Food Delivery"]] *= np.where(t_dow >= 5, 1.35, 1.0)
    W[:, col["Entertainment"]] *= np.where(t_dow >= 5, 1.6, 0.85)
    W[:, col["Shopping & Fashion"]] *= np.where(t_fest, 1.7, 1.0)
    W[:, col["Education"]] *= np.where(np.isin(t_month, [1, 6, 7]), 2.0, 0.7)
    W[:, col["Travel & Transport"]] *= np.where(np.isin(t_month, [5, 10, 12]), 1.5, 1.0)
    cdf = np.cumsum(W / W.sum(axis=1, keepdims=True), axis=1)
    cat_i = np.minimum((rng.random((n, 1)) > cdf).sum(axis=1), len(CAT_NAMES) - 1)
    category = np.array(CAT_NAMES)[cat_i]

    # hour of day (category specific) and merchant within category
    hour = np.zeros(n, dtype=int)
    merchant_id = np.zeros(n, dtype=int)
    for name in CAT_NAMES:
        sel = np.where(category == name)[0]
        hour[sel] = rng.choice(24, len(sel), p=hour_weights(CATEGORIES[name]["hr"]))
        if name != "P2P Transfer":
            pool = merchants[merchants.category == name].merchant_id.values
            mw = rng.lognormal(0, 1.0, len(pool))
            merchant_id[sel] = rng.choice(pool, len(sel), p=mw / mw.sum())
    minute, sec = rng.integers(0, 60, n), rng.integers(0, 60, n)

    # bank (85% primary), network, amount
    bank_ids = banks.bank_id.values
    other = rng.choice(bank_ids, n, p=banks.share / banks.share.sum())
    bank_id = np.where(rng.random(n) < 0.85, c.primary_bank_id.values, other)
    bank_ix = bank_id - 1
    city_ix = c.city_id.values - 1
    weak_p = cities.weak.values[city_ix]
    net_u = rng.random(n)
    network = np.where(net_u < weak_p, "Weak (2G/3G)", np.where(net_u < weak_p + 0.20, "WiFi", np.where(net_u < weak_p + 0.20 + 0.18, "5G", "4G")))

    amount = np.zeros(n)
    for name in CAT_NAMES:
        sel = np.where(category == name)[0]
        spec = CATEGORIES[name]
        amount[sel] = rng.lognormal(np.log(spec["med"]), spec["sig"], len(sel))
    rc = np.where(category == "Mobile Recharge")[0]
    amount[rc] = rng.choice(RECHARGE_PLANS, len(rc))
    amount = np.clip(np.round(amount), 1, 100_000)
    txn_type = np.where(category == "P2P Transfer", "P2P", "P2M")

    # ---------------------------------------------------------------- outcomes
    peak = np.where((hour >= 19) & (hour <= 22), 1.5, np.where((hour >= 12) & (hour <= 14), 1.15, np.where(hour <= 5, 0.85, 1.0)))
    net_m = np.select([network == "Weak (2G/3G)", network == "5G", network == "WiFi"], [2.2, 0.9, 0.95], 1.0)
    trend = 1 - 0.011 * t_m                                   # infrastructure improves ~1.1% a month
    fest_load = np.where(t_fest, 1.25, 1.0)
    incident = np.ones(n)
    in_incident = np.zeros(n, dtype=bool)
    for b, a, z, mult in INCIDENTS:
        hit = (bank_ix == b) & (t_date.dt.date.values >= a) & (t_date.dt.date.values <= z)
        incident[hit] = mult
        in_incident |= hit
    p_sys = np.clip(banks.fail.values[bank_ix] * trend * peak * net_m * fest_load * incident, 0, 0.7)
    sys_fail = rng.random(n) < p_sys

    age = c.age_group.values
    new_user = tenure <= 30
    p_bal = 0.012 * np.where(age == "18-24", 1.4, 1.0) * np.where(amount > 5000, 1.6, 1.0) * np.where(t_dom >= 26, 1.3, 1.0)
    p_pin = 0.009 * np.where(age == "45+", 1.7, 1.0) * np.where(new_user, 1.5, 1.0)
    p_lim = np.where(amount > 25_000, 0.03, np.where(amount > 10_000, 0.004, 0.0))
    bal_fail, pin_fail, lim_fail = rng.random(n) < p_bal, rng.random(n) < p_pin, rng.random(n) < p_lim

    # fraud flag
    night = hour <= 4
    p_fraud = (0.0021 * np.where(night, 4.0, 1.0) * np.where(amount > 20_000, 4.0, np.where(amount > 5_000, 2.0, 1.0))
               * np.where(new_user, 3.0, 1.0) * np.where(txn_type == "P2P", 2.5, 0.6)
               * np.where(cities.tier.values[city_ix] == "Tier 2", 1.2, 1.0))
    fraud = rng.random(n) < np.clip(p_fraud, 0, 0.5)
    blocked = fraud & (rng.random(n) < 0.55)

    # reasons
    weak = network == "Weak (2G/3G)"
    ru = rng.random(n)
    sys_reason = np.where(in_incident, "Bank server down",
                  np.where(weak, np.where(ru < 0.60, "Timeout / network error", np.where(ru < 0.95, "Technical decline", "Bank server down")),
                           np.where(ru < 0.60, "Technical decline", np.where(ru < 0.80, "Timeout / network error", "Bank server down"))))
    failure_type = np.full(n, "", dtype=object)
    reason = np.full(n, "", dtype=object)
    for mask, ftype, rs in [(blocked, "Risk", "Blocked by risk engine"), (sys_fail, "System", None), (lim_fail, "User", "Daily limit exceeded"),
                            (bal_fail, "User", "Insufficient balance"), (pin_fail, "User", "Wrong PIN")]:
        free = mask & (failure_type == "")
        failure_type[free] = ftype
        reason[free] = sys_reason[free] if rs is None else rs
    status = np.where(failure_type == "", "Success", "Failed")
    ok = status == "Success"

    fraud_loss = np.where(fraud & ok, amount, 0.0)
    rate = np.array([CATEGORIES[x]["cb"] for x in category])
    promo_p = np.where(t_fest, 0.55, 0.30)
    cashback = np.where(ok & (txn_type == "P2M") & (amount >= 100) & (rng.random(n) < promo_p), np.minimum(amount * rate, 100.0), 0.0).round(2)

    resp_mult = np.array([BANK_TYPE_RESP[t] for t in banks.bank_type.values[bank_ix]])
    resp = rng.lognormal(np.log(1800), 0.35, n) * resp_mult * np.where((hour >= 19) & (hour <= 22), 1.2, 1.0) \
        * np.where(weak, 1.9, 1.0) * np.where(failure_type == "System", 1.6, 1.0)
    resp = np.clip(resp, 300, 30_000).round(0)

    ts = [f"{d.date().isoformat()} {h:02d}:{mi:02d}:{s:02d}" for d, h, mi, s in zip(t_date, hour, minute, sec)]
    fact = pd.DataFrame({
        "txn_id": 700_000 + np.arange(1, n + 1), "txn_ts": ts, "date_key": t_date.dt.strftime("%Y%m%d").astype(int).values,
        "txn_hour": hour, "customer_id": c.customer_id.values, "merchant_id": merchant_id, "bank_id": bank_id,
        "txn_type": txn_type, "network": network, "amount": amount.astype(int), "status": status,
        "failure_type": np.where(failure_type == "", "No failure", failure_type), "failure_reason": np.where(reason == "", "No failure", reason),
        "fraud_flag": fraud.astype(int), "fraud_loss": fraud_loss.astype(int), "cashback": cashback,
        "response_ms": resp.astype(int), "tenure_days": tenure,
    })
    return fact


def to_flat(fact, cities, banks, customers, merchants) -> pd.DataFrame:
    cu = customers.merge(cities[["city_id", "city", "tier"]], on="city_id")
    flat = (fact.merge(cu[["customer_id", "city", "tier", "age_group", "device_os", "kyc_level"]], on="customer_id")
            .merge(merchants[["merchant_id", "merchant_name", "category"]], on="merchant_id")
            .merge(banks[["bank_id", "bank_name", "bank_type"]], on="bank_id"))
    flat["txn_date"] = flat.txn_ts.str[:10]
    flat["month_key"] = flat.txn_ts.str[:7]
    flat["amount_band"] = pd.cut(flat.amount, [0, 500, 2000, 10_000, 25_000, 1e9], right=True,
                                 labels=["1. Up to 500", "2. 501-2,000", "3. 2,001-10,000", "4. 10,001-25,000", "5. Above 25,000"]).astype(str)
    flat["tenure_band"] = pd.cut(flat.tenure_days, [-1, 30, 90, 180, 1e9], labels=["1. 0-30 days", "2. 31-90 days", "3. 91-180 days", "4. 180+ days"]).astype(str)
    flat["hour_band"] = pd.cut(flat.txn_hour, [-1, 4, 8, 11, 16, 21, 23],
                               labels=["1. Night (0-4)", "2. Early (5-8)", "3. Morning (9-11)", "4. Day (12-16)", "5. Evening (17-21)", "6. Late (22-23)"]).astype(str)
    flat = flat.sort_values("txn_ts").reset_index(drop=True)
    flat["is_new_user_month"] = (~flat.duplicated(["customer_id", "month_key"])).astype(int)   # 1 on a user's first txn in a month (for MAU)
    cols = ["txn_id", "txn_ts", "txn_date", "month_key", "txn_hour", "hour_band", "customer_id", "city", "tier", "age_group", "device_os",
            "kyc_level", "tenure_days", "tenure_band", "merchant_id", "merchant_name", "category", "txn_type", "bank_name", "bank_type",
            "network", "amount", "amount_band", "status", "failure_type", "failure_reason", "fraud_flag", "fraud_loss", "cashback",
            "response_ms", "is_new_user_month"]
    return flat[cols]


def main() -> dict[str, pd.DataFrame]:
    rng = np.random.default_rng(SEED)
    DATA_DIR.mkdir(exist_ok=True)
    PBI_DIR.mkdir(parents=True, exist_ok=True)
    calendar = make_calendar()
    cities, banks, customers, merchants = make_dims(rng)
    fact = build_transactions(rng, cities, banks, customers, merchants, calendar)
    flat = to_flat(fact, cities, banks, customers, merchants)

    dim_city = cities[["city_id", "city", "tier", "region"]]
    dim_bank = banks[["bank_id", "bank_name", "bank_type"]]
    dim_cust = customers.rename(columns={"primary_bank_id": "primary_bank_id"})
    star = {"fact_transactions": fact, "dim_date": calendar, "dim_customer": dim_cust, "dim_merchant": merchants,
            "dim_bank": dim_bank, "dim_city": pd.concat([pd.DataFrame([{"city_id": 0, "city": "N/A (P2P)", "tier": "N/A", "region": "N/A"}]), dim_city], ignore_index=True)}
    for name, df in star.items():
        df.to_csv(PBI_DIR / f"{name}.csv", index=False)
    flat.to_csv(DATA_DIR / "payflow_transactions_flat.csv", index=False)
    print(f"[generate_data] txns={len(fact):,} customers={len(customers):,} merchants={len(merchants) - 1} "
          f"-> {PBI_DIR} and {DATA_DIR}")
    return {"flat": flat, **star}


if __name__ == "__main__":
    main()
