"""Build PayFlow_UPI_Analytics.xlsx: formula-driven analysis workbook with an interactive dashboard.

Everything on the analysis sheets is a live Excel formula (COUNTIFS / SUMIFS / AVERAGEIFS on named
ranges over a 60,000-row Excel Table), so a recruiter can click any cell and read the logic.
Cached values are written alongside, then verified by a full recalculation (src/verify_excel.py).

Run:  python -m src.build_excel
"""
from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
import xlsxwriter
from scipy.stats import binom
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell as rc

from .xl_engine import Engine, KEYS, L, R, SUMCOLS, inr_formula, inr_text

ROOT = Path(__file__).resolve().parents[1]
FLAT = ROOT / "data" / "payflow_transactions_flat.csv"
OUT = ROOT / "excel" / "PayFlow_UPI_Analytics.xlsx"

INK, BLUE, ORANGE, MUTED, LINE, PANEL = "#1B2A3A", "#2A78D6", "#EB6834", "#6B7785", "#D5DAE1", "#F3F5F8"
INR_FMT = r"[>=10000000]##\,##\,##\,##0;[>=100000]##\,##\,##0;##,##0"


def make_formats(wb):
    base = {"font_name": "Calibri", "font_size": 10, "font_color": INK, "valign": "vcenter"}
    f = lambda **k: wb.add_format({**base, **k})  # noqa: E731
    return {
        "title": f(font_size=20, bold=True), "sub": f(font_size=10, font_color=MUTED), "h2": f(font_size=13, bold=True),
        "hdr": f(bold=True, font_color="#FFFFFF", bg_color=INK, border=1, border_color=INK, text_wrap=True, align="center"),
        "hdr_l": f(bold=True, font_color="#FFFFFF", bg_color=INK, border=1, border_color=INK, text_wrap=True),
        "txt": f(border=1, border_color=LINE), "int": f(border=1, border_color=LINE, num_format="#,##0"),
        "inr": f(border=1, border_color=LINE, num_format=INR_FMT), "pct": f(border=1, border_color=LINE, num_format="0.0%"),
        "pct2": f(border=1, border_color=LINE, num_format="0.00%"), "dec1": f(border=1, border_color=LINE, num_format="0.0"),
        "dec2": f(border=1, border_color=LINE, num_format="0.00"), "bps": f(border=1, border_color=LINE, num_format="0.0"),
        "tot_txt": f(bold=True, bg_color=PANEL, border=1, border_color=LINE), "tot_int": f(bold=True, bg_color=PANEL, border=1, border_color=LINE, num_format="#,##0"),
        "tot_inr": f(bold=True, bg_color=PANEL, border=1, border_color=LINE, num_format=INR_FMT),
        "tot_pct": f(bold=True, bg_color=PANEL, border=1, border_color=LINE, num_format="0.0%"),
        "tot_bps": f(bold=True, bg_color=PANEL, border=1, border_color=LINE, num_format="0.0"),
        "tot_dec2": f(bold=True, bg_color=PANEL, border=1, border_color=LINE, num_format="0.00"),
        "label": f(bold=True, font_color=MUTED, font_size=9), "select": f(bold=True, font_size=12, font_color=BLUE, bg_color="#FFFFFF", border=2, border_color=BLUE, align="center"),
        "kpi_lbl": f(font_size=9, bold=True, font_color=MUTED, bg_color=PANEL, align="left", indent=1),
        "kpi_val": f(font_size=22, bold=True, bg_color=PANEL, align="left", indent=1),
        "kpi_val_pct": f(font_size=22, bold=True, bg_color=PANEL, align="left", indent=1, num_format="0.0%"),
        "kpi_val_int": f(font_size=22, bold=True, bg_color=PANEL, align="left", indent=1, num_format="#,##0"),
        "kpi_val_inr": f(font_size=22, bold=True, bg_color=PANEL, align="left", indent=1, num_format="₹#,##0"),
        "kpi_val_dec": f(font_size=22, bold=True, bg_color=PANEL, align="left", indent=1, num_format="0.0"),
        "kpi_val_dec2": f(font_size=22, bold=True, bg_color=PANEL, align="left", indent=1, num_format="0.00"),
        "kpi_sub": f(font_size=8, font_color=MUTED, bg_color=PANEL, align="left", indent=1),
        "wrap": f(text_wrap=True, valign="top"), "wrap_b": f(text_wrap=True, valign="top", bold=True),
        "note": f(font_size=9, italic=True, font_color=MUTED, text_wrap=True, valign="top"),
        "link": f(font_color=BLUE, underline=1), "good": f(bg_color="#DDF1E4", font_color="#0B5B2A", bold=True, border=1, border_color=LINE),
        "bad": f(bg_color="#FBE0D4", font_color="#9A3412", bold=True, border=1, border_color=LINE),
        "big": f(font_size=11, text_wrap=True, valign="top"),
    }


def chart_style(ch, title, y_fmt=None, y_title=None, legend=False, x_rot=None, y_min=None, y_max=None, x_title=None):
    ch.set_title({"name": title, "name_font": {"size": 11, "bold": True, "color": INK}, "overlay": False})
    if not legend:
        ch.set_legend({"none": True})
    ch.set_chartarea({"border": {"color": LINE}, "fill": {"color": "#FFFFFF"}})
    ch.set_plotarea({"fill": {"none": True}})
    xa = {"num_font": {"size": 8, "color": MUTED}, "line": {"color": LINE}, "major_tick_mark": "none"}
    if x_rot is not None:
        xa["num_font"]["rotation"] = x_rot
    if x_title:
        xa["name"] = x_title
        xa["name_font"] = {"size": 8, "color": MUTED, "bold": False}
    ya = {"num_font": {"size": 8, "color": MUTED}, "line": {"none": True}, "major_gridlines": {"visible": True, "line": {"color": "#E6E9EE", "width": 0.75}}}
    if y_fmt:
        ya["num_format"] = y_fmt
    if y_title:
        ya["name"] = y_title
        ya["name_font"] = {"size": 8, "color": MUTED, "bold": False}
    if y_min is not None:
        ya["min"] = y_min
    if y_max is not None:
        ya["max"] = y_max
    ch.set_x_axis(xa)
    ch.set_y_axis(ya)


def build(path: Path = OUT) -> Path:
    df = pd.read_csv(FLAT, keep_default_na=False)
    eng = Engine(df)
    n = len(df)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = xlsxwriter.Workbook(path, {"strings_to_numbers": False})
    wb.set_properties({"title": "PayFlow UPI Payments Analytics", "subject": "Simulated data - portfolio project",
                       "comments": "All figures come from a simulated dataset."})
    F = make_formats(wb)

    order = ["Cover", "Dashboard", "Monthly", "Failures", "Segments", "Fraud_Cashback", "Incident_Finder", "Insights", "Data",
             "Dictionary", "Calc", "Lists"]
    ws = {name: wb.add_worksheet(name) for name in order}
    for name, color in [("Dashboard", BLUE), ("Monthly", INK), ("Failures", ORANGE), ("Segments", INK), ("Fraud_Cashback", ORANGE),
                        ("Incident_Finder", ORANGE), ("Insights", BLUE), ("Calc", "#9AA5B1"), ("Lists", "#9AA5B1")]:
        ws[name].set_tab_color(color)

    # ---------------------------------------------------------------- Data (Excel Table) + named ranges
    d = ws["Data"]
    cols = list(df.columns)
    numfmt = {"amount": F["inr"], "fraud_loss": F["inr"], "cashback": F["dec2"], "response_ms": F["int"], "tenure_days": F["int"]}
    d.add_table(0, 0, n, len(cols) - 1, {"name": "tblTxn", "style": "Table Style Light 9", "data": df.values.tolist(),
                                          "columns": [{"header": c, "format": numfmt.get(c)} for c in cols]})
    d.freeze_panes(1, 2)
    for i, c in enumerate(cols):
        d.set_column(i, i, 22 if c in ("merchant_name", "failure_reason", "bank_name", "txn_ts") else 14)
    colname = {c: xl_col_to_name(i) for i, c in enumerate(cols)}
    for k, c in {**KEYS, **SUMCOLS}.items():
        wb.define_name(f"d_{k}", f"=Data!${colname[c]}$2:${colname[c]}${n + 1}")

    # ---------------------------------------------------------------- Lists (dropdown sources)
    months = sorted(df.month_key.unique())
    cities = sorted(df.city.unique())
    cats = sorted(df.category.unique())
    banks_sorted = (df.assign(s=df.failure_type == "System").groupby("bank_name").s.mean().sort_values(ascending=False).index.tolist())
    li = ws["Lists"]
    for j, (title, vals) in enumerate([("Month (with All)", ["All"] + months), ("City (with All)", ["All"] + cities),
                                       ("Category (with All)", ["All"] + cats), ("Bank", banks_sorted), ("Month", months)]):
        li.write(0, j, title, F["hdr_l"])
        for i, v in enumerate(vals):
            li.write(i + 1, j, v)
        li.set_column(j, j, 24)

    # ---------------------------------------------------------------- Calc (dashboard helper tables)
    c = ws["Calc"]
    c.set_column(0, 0, 34)
    c.set_column(1, 3, 16)
    c.write(0, 0, "Helper tables that feed the Dashboard KPIs and charts (all live formulas).", F["h2"])
    c.write(2, 0, "Month criterion", F["label"]); c.write(3, 0, "City criterion", F["label"]); c.write(4, 0, "Category criterion", F["label"])
    for r_, ref in [(2, "$C$4"), (3, "$F$4"), (4, "$I$4")]:
        c.write_formula(r_, 1, f'=IF(Dashboard!{ref}="All","*",Dashboard!{ref})', F["txt"], "*")
    cm, cc, ck = R("$B$3", None), R("$B$4", None), R("$B$5", None)
    f_all = dict(month=cm, city=cc, cat=ck)

    kpi_rows = {}
    def kpi(label, row, formula, value, fmt):
        c.write(row, 0, label, F["txt"])
        c.write_formula(row, 1, "=" + formula, F[fmt], value)
        kpi_rows[label] = (row, value)

    ftx, vtx = eng.count(**f_all)
    kpi("Transactions", 7, ftx, vtx, "int")
    fsc, vsc = eng.count(**f_all, status=L("Success"))
    kpi("Successful transactions", 8, fsc, vsc, "int")
    ftp, vtp = eng.sum("amt", **f_all, status=L("Success"))
    kpi("TPV (successful value)", 9, ftp, vtp, "inr")
    kpi("Success rate", 10, "IFERROR(B9/B8,0)", vsc / vtx, "pct")
    kpi("Average ticket", 11, "IFERROR(B10/B9,0)", vtp / vsc, "inr")
    fsy, vsy = eng.count(**f_all, ftype=L("System"))
    kpi("System failure rate", 12, f"IFERROR({fsy}/B8,0)", vsy / vtx, "pct2")
    ffr, vfr = eng.sum("fraud", **f_all)
    kpi("Fraud flag rate (bps)", 13, f"IFERROR({ffr}/B8*10000,0)", vfr / vtx * 1e4, "bps")
    fcb, vcb = eng.sum("cb", **f_all)
    kpi("Cashback % of TPV", 14, f"IFERROR({fcb}/B10,0)", vcb / vtp, "pct2")
    far, var = eng.avg("resp", **f_all)
    kpi("Avg response time (sec)", 15, f"IFERROR(AVERAGEIFS(d_resp,d_month,$B$3,d_city,$B$4,d_cat,$B$5)/1000,0)", var / 1000, "dec2")
    kv = {k: v for k, (_, v) in kpi_rows.items()}

    # trend table (month filter ignored)
    r0 = 19
    c.write(r0 - 1, 0, "Monthly TPV (Rs lakh), city + category filters", F["hdr_l"]); c.write(r0 - 1, 1, "TPV (Rs lakh)", F["hdr"])
    for i, m in enumerate(months):
        r_ = r0 + i
        c.write(r_, 0, m, F["txt"])
        fm, vm = eng.sum("amt", month=R(f"$A{r_ + 1}", m), city=cc, cat=ck, status=L("Success"))
        c.write_formula(r_, 1, f"={fm}/100000", F["dec1"], vm / 1e5)
    trend_rng = (r0, r0 + len(months) - 1)

    # bank success rate (all filters)
    r1 = r0 + len(months) + 3
    c.write(r1 - 1, 0, "Success rate by bank (all filters)", F["hdr_l"]); c.write(r1 - 1, 1, "Success rate", F["hdr"])
    for i, b in enumerate(banks_sorted):
        r_ = r1 + i
        c.write(r_, 0, b, F["txt"])
        fa, va = eng.count(**f_all, bank=R(f"$A{r_ + 1}", b))
        fb, vb = eng.count(**f_all, bank=R(f"$A{r_ + 1}", b), status=L("Success"))
        c.write_formula(r_, 1, f"=IFERROR({fb}/{fa},0)", F["pct"], vb / va if va else 0)
    bank_rng = (r1, r1 + len(banks_sorted) - 1)

    # TPV by category (month + city filters)
    r2 = r1 + len(banks_sorted) + 3
    c.write(r2 - 1, 0, "TPV by category (Rs lakh), month + city filters", F["hdr_l"]); c.write(r2 - 1, 1, "TPV (Rs lakh)", F["hdr"])
    cat_order = df[df.status == "Success"].groupby("category").amount.sum().sort_values(ascending=False).index.tolist()
    for i, k in enumerate(cat_order):
        r_ = r2 + i
        c.write(r_, 0, k, F["txt"])
        fm, vm = eng.sum("amt", month=cm, city=cc, cat=R(f"$A{r_ + 1}", k), status=L("Success"))
        c.write_formula(r_, 1, f"={fm}/100000", F["dec1"], vm / 1e5)
    cat_rng = (r2, r2 + len(cat_order) - 1)

    # system failure by hour (all filters)
    r3 = r2 + len(cat_order) + 3
    c.write(r3 - 1, 0, "System failure rate by hour (all filters)", F["hdr_l"]); c.write(r3 - 1, 1, "System failure %", F["hdr"])
    for h in range(24):
        r_ = r3 + h
        c.write_number(r_, 0, h, F["txt"])
        fa, va = eng.count(**f_all, hour=R(f"$A{r_ + 1}", h))
        fs, vs = eng.count(**f_all, hour=R(f"$A{r_ + 1}", h), ftype=L("System"))
        c.write_formula(r_, 1, f"=IFERROR({fs}/{fa},0)", F["pct"], vs / va if va else 0)
    hour_rng = (r3, r3 + 23)

    # ---------------------------------------------------------------- Dashboard
    s = ws["Dashboard"]
    s.hide_gridlines(2)
    s.set_column(0, 0, 2)
    s.set_column(1, 12, 11.5)
    s.set_row(0, 34)
    s.write(0, 1, "PayFlow UPI Payments Dashboard", F["title"])
    s.write(1, 1, "Simulated data: 60,000 transactions, 5,000 customers, Jan 2025 - Oct 2026. Pick a month, city or category; every card and chart recalculates.", F["sub"])
    s.write(3, 1, "MONTH", F["label"]); s.write(3, 4, "CITY", F["label"]); s.write(3, 7, "CATEGORY", F["label"])
    s.set_row(3, 24)
    s.merge_range(3, 2, 3, 3, "All", F["select"]); s.merge_range(3, 5, 3, 6, "All", F["select"]); s.merge_range(3, 8, 3, 9, "All", F["select"])
    for col_, src in [(2, "$A$2:$A$24"), (5, "$B$2:$B$14"), (8, "$C$2:$C$13")]:
        s.data_validation(3, col_, 3, col_, {"validate": "list", "source": f"=Lists!{src}", "input_title": "Filter",
                                              "input_message": "Choose a value or All"})
    s.merge_range(4, 1, 4, 12, "Dropdowns drive the cards. The monthly trend ignores the month filter and the category chart ignores the category filter.", F["note"])

    cards = [("TOTAL PAYMENT VALUE (TPV)", "B9", "tpv", "Successful payments only"), ("TRANSACTIONS", "B8", "kpi_val_int", "All attempts"),
             ("SUCCESS RATE", "B11", "kpi_val_pct", "Successful / attempted"), ("AVERAGE TICKET", "B12", "kpi_val_inr", "Per successful payment"),
             ("SYSTEM FAILURE RATE", "B13", "kpi_val_pct", "Bank/network faults, % of attempts"), ("FRAUD FLAG RATE", "B14", "kpi_val_dec", "Flagged per 10,000 transactions"),
             ("CASHBACK COST", "B15", "kpi_val_pct", "Cashback as % of TPV"), ("AVG RESPONSE TIME", "B16", "kpi_val_dec2", "Seconds per attempt")]
    for i, (lbl, ref, fmtk, sub) in enumerate(cards):
        r_, c_ = 5 + (i // 4) * 4, 1 + (i % 4) * 3
        s.set_row(r_, 18); s.set_row(r_ + 1, 34); s.set_row(r_ + 2, 16)
        s.merge_range(r_, c_, r_, c_ + 2, lbl, F["kpi_lbl"])
        s.merge_range(r_ + 2, c_, r_ + 2, c_ + 2, sub, F["kpi_sub"])
        if fmtk == "tpv":
            s.merge_range(r_ + 1, c_, r_ + 1, c_ + 2, "", F["kpi_val"])
            s.write_formula(r_ + 1, c_, "=" + inr_formula("Calc!$B$10"), F["kpi_val"], inr_text(kv["TPV (successful value)"]))
        else:
            v = {"B8": kv["Transactions"], "B11": kv["Success rate"], "B12": kv["Average ticket"], "B13": kv["System failure rate"],
                 "B14": kv["Fraud flag rate (bps)"], "B15": kv["Cashback % of TPV"], "B16": kv["Avg response time (sec)"]}[ref]
            s.merge_range(r_ + 1, c_, r_ + 1, c_ + 2, "", F[fmtk])
            s.write_formula(r_ + 1, c_, f"=Calc!${ref[0]}${ref[1:]}", F[fmtk], v)

    def sheet_range(col_i, rng):
        return ["Calc", rng[0], col_i, rng[1], col_i]

    ch1 = wb.add_chart({"type": "column"})
    ch1.add_series({"name": "TPV (Rs lakh)", "categories": sheet_range(0, trend_rng), "values": sheet_range(1, trend_rng),
                    "fill": {"color": BLUE}, "border": {"none": True}, "gap": 40})
    chart_style(ch1, "Monthly TPV (Rs lakh)", y_fmt="#,##0", x_rot=-45); ch1.set_size({"width": 560, "height": 300})
    s.insert_chart(13, 1, ch1, {"x_offset": 2, "y_offset": 4})

    ch2 = wb.add_chart({"type": "bar"})
    ch2.add_series({"name": "Success rate", "categories": sheet_range(0, bank_rng), "values": sheet_range(1, bank_rng),
                    "fill": {"color": BLUE}, "border": {"none": True}, "gap": 35,
                    "data_labels": {"value": True, "num_format": "0.0%", "font": {"size": 8, "color": INK}}})
    chart_style(ch2, "Success rate by bank (weakest first)", y_fmt="0%")
    ch2.set_y_axis({"reverse": True, "num_font": {"size": 8, "color": MUTED}, "line": {"color": LINE}, "major_tick_mark": "none"})
    ch2.set_x_axis({"min": 0.8, "max": 1, "num_format": "0%", "num_font": {"size": 8, "color": MUTED}, "major_gridlines": {"visible": True, "line": {"color": "#E6E9EE"}}, "line": {"none": True}})
    ch2.set_size({"width": 560, "height": 300})
    s.insert_chart(13, 7, ch2, {"x_offset": 8, "y_offset": 4})

    ch3 = wb.add_chart({"type": "bar"})
    ch3.add_series({"name": "TPV (Rs lakh)", "categories": sheet_range(0, cat_rng), "values": sheet_range(1, cat_rng),
                    "fill": {"color": BLUE}, "border": {"none": True}, "gap": 35,
                    "data_labels": {"value": True, "num_format": "#,##0", "font": {"size": 8, "color": INK}}})
    chart_style(ch3, "TPV by category (Rs lakh)")
    ch3.set_y_axis({"reverse": True, "num_font": {"size": 8, "color": MUTED}, "line": {"color": LINE}, "major_tick_mark": "none"})
    ch3.set_x_axis({"num_format": "#,##0", "num_font": {"size": 8, "color": MUTED}, "major_gridlines": {"visible": True, "line": {"color": "#E6E9EE"}}, "line": {"none": True}})
    ch3.set_size({"width": 560, "height": 300})
    s.insert_chart(29, 1, ch3, {"x_offset": 2, "y_offset": 4})

    ch4 = wb.add_chart({"type": "column"})
    ch4.add_series({"name": "System failure %", "categories": sheet_range(0, hour_rng), "values": sheet_range(1, hour_rng),
                    "fill": {"color": ORANGE}, "border": {"none": True}, "gap": 30})
    chart_style(ch4, "System failure rate by hour of day", y_fmt="0.0%", x_title="Hour (0-23)"); ch4.set_size({"width": 560, "height": 300})
    s.insert_chart(29, 7, ch4, {"x_offset": 8, "y_offset": 4})

    # ---------------------------------------------------------------- Monthly
    m = ws["Monthly"]
    m.hide_gridlines(2)
    m.set_column(0, 0, 11); m.set_column(1, 11, 14)
    m.write(0, 0, "Monthly performance", F["title"]); m.set_row(0, 30)
    m.write(1, 0, "Active users = customers with at least one transaction that month (helper column is_new_user_month). CMGR = compound monthly growth rate of TPV.", F["sub"])
    heads = ["Month", "Transactions", "Successful txns", "TPV (Rs)", "Success rate", "Avg ticket (Rs)", "Active users", "Txns per user",
             "System failure %", "Fraud flag (bps)", "Cashback % of TPV", "TPV MoM growth"]
    for j, h in enumerate(heads):
        m.write(2, j, h, F["hdr"])
    m.set_row(2, 32)
    first = 3
    mon_vals = []
    for i, mo in enumerate(months):
        r_ = first + i; e = r_ + 1
        ref = R(f"$A{e}", mo)
        m.write(r_, 0, mo, F["txt"])
        fb, vb = eng.count(month=ref); fc, vc = eng.count(month=ref, status=L("Success")); fd, vd = eng.sum("amt", month=ref, status=L("Success"))
        fg, vg = eng.sum("mau", month=ref); fi, vi = eng.count(month=ref, ftype=L("System")); fj, vj = eng.sum("fraud", month=ref); fk, vk = eng.sum("cb", month=ref)
        m.write_formula(r_, 1, "=" + fb, F["int"], vb); m.write_formula(r_, 2, "=" + fc, F["int"], vc); m.write_formula(r_, 3, "=" + fd, F["inr"], vd)
        m.write_formula(r_, 4, f"=IFERROR(C{e}/B{e},0)", F["pct"], vc / vb); m.write_formula(r_, 5, f"=IFERROR(D{e}/C{e},0)", F["inr"], vd / vc)
        m.write_formula(r_, 6, "=" + fg, F["int"], vg); m.write_formula(r_, 7, f"=IFERROR(B{e}/G{e},0)", F["dec2"], vb / vg)
        m.write_formula(r_, 8, f"=IFERROR({fi}/B{e},0)", F["pct2"], vi / vb); m.write_formula(r_, 9, f"=IFERROR({fj}/B{e}*10000,0)", F["bps"], vj / vb * 1e4)
        m.write_formula(r_, 10, f"=IFERROR({fk}/D{e},0)", F["pct2"], vk / vd)
        if i == 0:
            m.write(r_, 11, "", F["txt"])
        else:
            m.write_formula(r_, 11, f"=IFERROR(D{e}/D{e - 1}-1,\"\")", F["pct"], vd / mon_vals[-1]["tpv"] - 1)
        mon_vals.append(dict(m=mo, txn=vb, succ=vc, tpv=vd, active=vg))
    last = first + len(months) - 1
    tr = last + 1; fe, le = first + 1, last + 1
    m.write(tr, 0, "Total", F["tot_txt"])
    tot = dict(txn=sum(x["txn"] for x in mon_vals), succ=sum(x["succ"] for x in mon_vals), tpv=sum(x["tpv"] for x in mon_vals))
    for j, key, fm_ in [(1, "txn", "tot_int"), (2, "succ", "tot_int"), (3, "tpv", "tot_inr")]:
        col_ = xl_col_to_name(j)
        m.write_formula(tr, j, f"=SUM({col_}{fe}:{col_}{le})", F[fm_], tot[key])
    m.write_formula(tr, 4, f"=IFERROR(C{tr + 1}/B{tr + 1},0)", F["tot_pct"], tot["succ"] / tot["txn"])
    m.write_formula(tr, 5, f"=IFERROR(D{tr + 1}/C{tr + 1},0)", F["tot_inr"], tot["tpv"] / tot["succ"])
    for j in (6, 7, 8, 9, 10, 11):
        m.write(tr, j, "", F["tot_txt"])
    cmgr = (mon_vals[-1]["tpv"] / mon_vals[0]["tpv"]) ** (1 / (len(months) - 1)) - 1
    m.write(tr + 2, 0, "TPV CMGR (Jan 2025 to Oct 2026)", F["label"])
    m.write_formula(tr + 2, 3, f"=(D{le}/D{fe})^(1/{len(months) - 1})-1", F["pct"], cmgr)
    m.conditional_format(first, 8, last, 8, {"type": "data_bar", "bar_color": "#F4B79F", "bar_solid": True})
    m.conditional_format(first, 11, last, 11, {"type": "3_color_scale", "min_color": "#F7C4B0", "mid_color": "#FFFFFF", "max_color": "#BBD5F4"})
    m.freeze_panes(3, 1)
    ca = wb.add_chart({"type": "column"})
    ca.add_series({"name": "TPV", "categories": ["Monthly", first, 0, last, 0], "values": ["Monthly", first, 3, last, 3], "fill": {"color": BLUE}, "border": {"none": True}, "gap": 40})
    chart_style(ca, "Monthly TPV (Rs)", y_fmt='#,##0,,"M"', x_rot=-45); ca.set_size({"width": 620, "height": 300}); m.insert_chart(tr + 4, 0, ca)
    cb_ = wb.add_chart({"type": "line"})
    cb_.add_series({"name": "Success rate", "categories": ["Monthly", first, 0, last, 0], "values": ["Monthly", first, 4, last, 4],
                    "line": {"color": BLUE, "width": 2.25}, "marker": {"type": "circle", "size": 6, "fill": {"color": BLUE}, "border": {"color": "#FFFFFF"}}})
    chart_style(cb_, "Success rate by month", y_fmt="0.0%", x_rot=-45, y_min=0.9, y_max=1.0); cb_.set_size({"width": 620, "height": 300}); m.insert_chart(tr + 4, 6, cb_)

    # ---------------------------------------------------------------- Failures
    fa_ = ws["Failures"]
    fa_.hide_gridlines(2)
    fa_.set_column(0, 0, 30); fa_.set_column(1, 10, 14)
    fa_.write(0, 0, "Where and why payments fail", F["title"]); fa_.set_row(0, 30)
    fa_.write(1, 0, "System failure = bank / network fault (the company can fix it). User failure = balance, PIN or limit. Risk = blocked by fraud rules. Banks sorted worst first at build time; formulas stay live.", F["sub"])
    bank_type = df.drop_duplicates("bank_name").set_index("bank_name").bank_type.to_dict()
    # A. by bank
    fa_.write(3, 0, "A. By bank", F["h2"])
    for j, h in enumerate(["Bank", "Type", "Transactions", "System fails", "System fail %", "User fails", "User fail %", "Risk blocks", "Value lost to system fails (Rs)", "Share of system-failed value"]):
        fa_.write(4, j, h, F["hdr"])
    fa_.set_row(4, 44)
    b0 = 5
    sysval_total = eng.sum("amt", ftype=L("System"))[1]
    for i, b in enumerate(banks_sorted):
        r_ = b0 + i; e = r_ + 1; ref = R(f"$A{e}", b)
        fa_.write(r_, 0, b, F["txt"]); fa_.write(r_, 1, bank_type[b], F["txt"])
        ft, vt = eng.count(bank=ref); fs, vs = eng.count(bank=ref, ftype=L("System")); fu, vu = eng.count(bank=ref, ftype=L("User"))
        fr, vr = eng.count(bank=ref, ftype=L("Risk")); fv, vv = eng.sum("amt", bank=ref, ftype=L("System"))
        fa_.write_formula(r_, 2, "=" + ft, F["int"], vt); fa_.write_formula(r_, 3, "=" + fs, F["int"], vs)
        fa_.write_formula(r_, 4, f"=IFERROR(D{e}/C{e},0)", F["pct2"], vs / vt); fa_.write_formula(r_, 5, "=" + fu, F["int"], vu)
        fa_.write_formula(r_, 6, f"=IFERROR(F{e}/C{e},0)", F["pct2"], vu / vt); fa_.write_formula(r_, 7, "=" + fr, F["int"], vr)
        fa_.write_formula(r_, 8, "=" + fv, F["inr"], vv)
        fa_.write_formula(r_, 9, f"=IFERROR(I{e}/I${b0 + len(banks_sorted) + 1},0)", F["pct"], vv / sysval_total)
    bt = b0 + len(banks_sorted); be = bt + 1
    fa_.write(bt, 0, "All banks", F["tot_txt"]); fa_.write(bt, 1, "", F["tot_txt"])
    tt, ts_, tu, tr_ = len(df), int((df.failure_type == "System").sum()), int((df.failure_type == "User").sum()), int((df.failure_type == "Risk").sum())
    for j, v_ in [(2, tt), (3, ts_), (5, tu), (7, tr_)]:
        col_ = xl_col_to_name(j); fa_.write_formula(bt, j, f"=SUM({col_}{b0 + 1}:{col_}{bt})", F["tot_int"], v_)
    fa_.write_formula(bt, 4, f"=IFERROR(D{be}/C{be},0)", F["tot_pct"], ts_ / tt); fa_.write_formula(bt, 6, f"=IFERROR(F{be}/C{be},0)", F["tot_pct"], tu / tt)
    fa_.write_formula(bt, 8, f"=SUM(I{b0 + 1}:I{bt})", F["tot_inr"], sysval_total); fa_.write_formula(bt, 9, f"=SUM(J{b0 + 1}:J{bt})", F["tot_pct"], 1.0)
    fa_.conditional_format(b0, 4, bt - 1, 4, {"type": "data_bar", "bar_color": "#F4B79F", "bar_solid": True})
    fa_.conditional_format(b0, 6, bt - 1, 6, {"type": "data_bar", "bar_color": "#BBD5F4", "bar_solid": True})

    # B. by reason
    rs0 = bt + 4
    fa_.write(rs0 - 1, 0, "B. By failure reason", F["h2"])
    for j, h in enumerate(["Reason", "Type", "Count", "Share of failures", "Value (Rs)"]):
        fa_.write(rs0, j, h, F["hdr"])
    reasons = df[df.status == "Failed"].groupby(["failure_reason", "failure_type"]).size().sort_values(ascending=False).reset_index()
    nfail = int((df.status == "Failed").sum())
    for i, row in reasons.iterrows():
        r_ = rs0 + 1 + i; e = r_ + 1; ref = R(f"$A{e}", row.failure_reason)
        fa_.write(r_, 0, row.failure_reason, F["txt"]); fa_.write(r_, 1, row.failure_type, F["txt"])
        fc, vc = eng.count(reason=ref); fv, vv = eng.sum("amt", reason=ref)
        fa_.write_formula(r_, 2, "=" + fc, F["int"], vc)
        fa_.write_formula(r_, 3, f"=IFERROR(C{e}/C${rs0 + 2 + len(reasons)},0)", F["pct"], vc / nfail); fa_.write_formula(r_, 4, "=" + fv, F["inr"], vv)
    rt = rs0 + 1 + len(reasons)
    fa_.write(rt, 0, "All failures", F["tot_txt"]); fa_.write(rt, 1, "", F["tot_txt"])
    fa_.write_formula(rt, 2, f"=SUM(C{rs0 + 2}:C{rt})", F["tot_int"], nfail); fa_.write_formula(rt, 3, f"=SUM(D{rs0 + 2}:D{rt})", F["tot_pct"], 1.0)
    fa_.write_formula(rt, 4, f"=SUM(E{rs0 + 2}:E{rt})", F["tot_inr"], float(df[df.status == "Failed"].amount.sum()))

    # C. by network
    n0 = rt + 4
    fa_.write(n0 - 1, 0, "C. By network quality", F["h2"])
    for j, h in enumerate(["Network", "Transactions", "System fails", "System fail %", "Avg response (sec)"]):
        fa_.write(n0, j, h, F["hdr"])
    nets = ["5G", "WiFi", "4G", "Weak (2G/3G)"]
    for i, nw in enumerate(nets):
        r_ = n0 + 1 + i; e = r_ + 1; ref = R(f"$A{e}", nw)
        fa_.write(r_, 0, nw, F["txt"])
        ft, vt = eng.count(net=ref); fs, vs = eng.count(net=ref, ftype=L("System")); fr_, vr_ = eng.avg("resp", net=ref)
        fa_.write_formula(r_, 1, "=" + ft, F["int"], vt); fa_.write_formula(r_, 2, "=" + fs, F["int"], vs)
        fa_.write_formula(r_, 3, f"=IFERROR(C{e}/B{e},0)", F["pct2"], vs / vt)
        fa_.write_formula(r_, 4, f"=IFERROR(AVERAGEIFS(d_resp,d_net,$A{e})/1000,0)", F["dec2"], vr_ / 1000)

    # D. by hour
    h0 = n0 + len(nets) + 4
    fa_.write(h0 - 1, 0, "D. By hour of day", F["h2"])
    for j, h in enumerate(["Hour", "Transactions", "System fail %", "User fail %", "Avg response (sec)"]):
        fa_.write(h0, j, h, F["hdr"])
    for hr in range(24):
        r_ = h0 + 1 + hr; e = r_ + 1; ref = R(f"$A{e}", hr)
        fa_.write_number(r_, 0, hr, F["txt"])
        ft, vt = eng.count(hour=ref); fs, vs = eng.count(hour=ref, ftype=L("System")); fu, vu = eng.count(hour=ref, ftype=L("User")); fr_, vr_ = eng.avg("resp", hour=ref)
        fa_.write_formula(r_, 1, "=" + ft, F["int"], vt); fa_.write_formula(r_, 2, f"=IFERROR({fs}/B{e},0)", F["pct2"], vs / vt)
        fa_.write_formula(r_, 3, f"=IFERROR({fu}/B{e},0)", F["pct2"], vu / vt)
        fa_.write_formula(r_, 4, f"=IFERROR(AVERAGEIFS(d_resp,d_hour,$A{e})/1000,0)", F["dec2"], vr_ / 1000)
    fa_.conditional_format(h0 + 1, 2, h0 + 24, 2, {"type": "2_color_scale", "min_color": "#FFFFFF", "max_color": "#F08A5D"})
    fa_.conditional_format(h0 + 1, 4, h0 + 24, 4, {"type": "2_color_scale", "min_color": "#FFFFFF", "max_color": "#8DB8EC"})

    # E. by age
    a0 = h0 + 24 + 4
    fa_.write(a0 - 1, 0, "E. User-side failures by age group and tenure", F["h2"])
    for j, h in enumerate(["Segment", "Transactions", "User fail %", "Wrong PIN %", "Insufficient balance %"]):
        fa_.write(a0, j, h, F["hdr"])
    seg_rows = [("age", a, a) for a in ["18-24", "25-34", "35-44", "45+"]] + [("tband", t, t) for t in sorted(df.tenure_band.unique())]
    for i, (key, val, label) in enumerate(seg_rows):
        r_ = a0 + 1 + i; e = r_ + 1; ref = R(f"$A{e}", val)
        fa_.write(r_, 0, label, F["txt"])
        ft, vt = eng.count(**{key: ref}); fu, vu = eng.count(**{key: ref, "ftype": L("User")})
        fp, vp = eng.count(**{key: ref, "reason": L("Wrong PIN")}); fbl, vbl = eng.count(**{key: ref, "reason": L("Insufficient balance")})
        fa_.write_formula(r_, 1, "=" + ft, F["int"], vt); fa_.write_formula(r_, 2, f"=IFERROR({fu}/B{e},0)", F["pct2"], vu / vt)
        fa_.write_formula(r_, 3, f"=IFERROR({fp}/B{e},0)", F["pct2"], vp / vt); fa_.write_formula(r_, 4, f"=IFERROR({fbl}/B{e},0)", F["pct2"], vbl / vt)

    c1 = wb.add_chart({"type": "bar"})
    c1.add_series({"name": "System fail %", "categories": ["Failures", b0, 0, bt - 1, 0], "values": ["Failures", b0, 4, bt - 1, 4], "fill": {"color": ORANGE}, "border": {"none": True}, "gap": 35,
                   "data_labels": {"value": True, "num_format": "0.0%", "font": {"size": 8, "color": INK}}})
    chart_style(c1, "System failure rate by bank")
    c1.set_y_axis({"reverse": True, "num_font": {"size": 8, "color": MUTED}, "line": {"color": LINE}, "major_tick_mark": "none"})
    c1.set_x_axis({"num_format": "0%", "num_font": {"size": 8, "color": MUTED}, "major_gridlines": {"visible": True, "line": {"color": "#E6E9EE"}}, "line": {"none": True}})
    c1.set_size({"width": 560, "height": 330}); fa_.insert_chart(3, 11, c1)
    c2 = wb.add_chart({"type": "column"})
    c2.add_series({"name": "System fail %", "categories": ["Failures", h0 + 1, 0, h0 + 24, 0], "values": ["Failures", h0 + 1, 2, h0 + 24, 2], "fill": {"color": ORANGE}, "border": {"none": True}, "gap": 30})
    chart_style(c2, "System failure rate by hour", y_fmt="0.0%", x_title="Hour (0-23)"); c2.set_size({"width": 560, "height": 300}); fa_.insert_chart(21, 11, c2)
    c3 = wb.add_chart({"type": "column"})
    c3.add_series({"name": "System fail %", "categories": ["Failures", n0 + 1, 0, n0 + 4, 0], "values": ["Failures", n0 + 1, 3, n0 + 4, 3], "fill": {"color": ORANGE}, "border": {"none": True}, "gap": 60,
                   "data_labels": {"value": True, "num_format": "0.0%", "font": {"size": 8, "color": INK}}})
    chart_style(c3, "System failure rate by network", y_fmt="0%"); c3.set_size({"width": 560, "height": 280}); fa_.insert_chart(37, 11, c3)

    # ---------------------------------------------------------------- Segments
    sg = ws["Segments"]
    sg.hide_gridlines(2)
    sg.set_column(0, 0, 26); sg.set_column(1, 9, 14)
    sg.write(0, 0, "Who pays and where", F["title"]); sg.set_row(0, 30)
    sg.write(1, 0, "Category, city and tier views. TPV counts successful payments only.", F["sub"])
    succ_total = float(df[df.status == "Success"].amount.sum())
    sg.write(3, 0, "A. By category", F["h2"])
    for j, h in enumerate(["Category", "Transactions", "TPV (Rs)", "TPV share", "Avg ticket (Rs)", "Success rate", "Cashback % of TPV", "Fraud flag (bps)"]):
        sg.write(4, j, h, F["hdr"])
    sg.set_row(4, 32)
    k0 = 5
    for i, k in enumerate(cat_order):
        r_ = k0 + i; e = r_ + 1; ref = R(f"$A{e}", k)
        sg.write(r_, 0, k, F["txt"])
        ft, vt = eng.count(cat=ref); fs, vs = eng.count(cat=ref, status=L("Success")); fp, vp = eng.sum("amt", cat=ref, status=L("Success"))
        fcb2, vcb2 = eng.sum("cb", cat=ref); ffr2, vfr2 = eng.sum("fraud", cat=ref)
        sg.write_formula(r_, 1, "=" + ft, F["int"], vt); sg.write_formula(r_, 2, "=" + fp, F["inr"], vp)
        sg.write_formula(r_, 3, f"=IFERROR(C{e}/C${k0 + len(cat_order) + 1},0)", F["pct"], vp / succ_total)
        sg.write_formula(r_, 4, f"=IFERROR(C{e}/{fs},0)", F["inr"], vp / vs); sg.write_formula(r_, 5, f"=IFERROR({fs}/B{e},0)", F["pct"], vs / vt)
        sg.write_formula(r_, 6, f"=IFERROR({fcb2}/C{e},0)", F["pct2"], vcb2 / vp if vp else 0); sg.write_formula(r_, 7, f"=IFERROR({ffr2}/B{e}*10000,0)", F["bps"], vfr2 / vt * 1e4)
    kt = k0 + len(cat_order)
    sg.write(kt, 0, "All categories", F["tot_txt"])
    sg.write_formula(kt, 1, f"=SUM(B{k0 + 1}:B{kt})", F["tot_int"], len(df)); sg.write_formula(kt, 2, f"=SUM(C{k0 + 1}:C{kt})", F["tot_inr"], succ_total)
    sg.write_formula(kt, 3, f"=SUM(D{k0 + 1}:D{kt})", F["tot_pct"], 1.0)
    for j in (4, 5, 6, 7):
        sg.write(kt, j, "", F["tot_txt"])
    sg.conditional_format(k0, 3, kt - 1, 3, {"type": "data_bar", "bar_color": "#BBD5F4", "bar_solid": True})

    cy0 = kt + 4
    sg.write(cy0 - 1, 0, "B. By city", F["h2"])
    for j, h in enumerate(["City", "Tier", "Transactions", "TPV (Rs)", "TPV share", "Success rate", "System fail %", "Weak-network share", "Avg ticket (Rs)"]):
        sg.write(cy0, j, h, F["hdr"])
    sg.set_row(cy0, 32)
    city_tier = df.drop_duplicates("city").set_index("city").tier.to_dict()
    city_order = df[df.status == "Success"].groupby("city").amount.sum().sort_values(ascending=False).index.tolist()
    for i, ct in enumerate(city_order):
        r_ = cy0 + 1 + i; e = r_ + 1; ref = R(f"$A{e}", ct)
        sg.write(r_, 0, ct, F["txt"]); sg.write(r_, 1, city_tier[ct], F["txt"])
        ft, vt = eng.count(city=ref); fs, vs = eng.count(city=ref, status=L("Success")); fp, vp = eng.sum("amt", city=ref, status=L("Success"))
        fy, vy = eng.count(city=ref, ftype=L("System")); fw, vw = eng.count(city=ref, net=L("Weak (2G/3G)"))
        sg.write_formula(r_, 2, "=" + ft, F["int"], vt); sg.write_formula(r_, 3, "=" + fp, F["inr"], vp)
        sg.write_formula(r_, 4, f"=IFERROR(D{e}/D${cy0 + 2 + len(city_order)},0)", F["pct"], vp / succ_total)
        sg.write_formula(r_, 5, f"=IFERROR({fs}/C{e},0)", F["pct"], vs / vt); sg.write_formula(r_, 6, f"=IFERROR({fy}/C{e},0)", F["pct2"], vy / vt)
        sg.write_formula(r_, 7, f"=IFERROR({fw}/C{e},0)", F["pct"], vw / vt); sg.write_formula(r_, 8, f"=IFERROR(D{e}/{fs},0)", F["inr"], vp / vs)
    cyt = cy0 + 1 + len(city_order)
    sg.write(cyt, 0, "All cities", F["tot_txt"]); sg.write(cyt, 1, "", F["tot_txt"])
    sg.write_formula(cyt, 2, f"=SUM(C{cy0 + 2}:C{cyt})", F["tot_int"], len(df)); sg.write_formula(cyt, 3, f"=SUM(D{cy0 + 2}:D{cyt})", F["tot_inr"], succ_total)
    sg.write_formula(cyt, 4, f"=SUM(E{cy0 + 2}:E{cyt})", F["tot_pct"], 1.0)
    for j in (5, 6, 7, 8):
        sg.write(cyt, j, "", F["tot_txt"])
    sg.conditional_format(cy0 + 1, 6, cyt - 1, 6, {"type": "2_color_scale", "min_color": "#FFFFFF", "max_color": "#F08A5D"})

    t0 = cyt + 4
    sg.write(t0 - 1, 0, "C. By city tier and payment type", F["h2"])
    for j, h in enumerate(["Segment", "Transactions", "TPV (Rs)", "TPV share", "Success rate", "System fail %", "Avg ticket (Rs)"]):
        sg.write(t0, j, h, F["hdr"])
    sg.set_row(t0, 32)
    for i, (key, val) in enumerate([("tier", "Tier 1"), ("tier", "Tier 2"), ("ttype", "P2M"), ("ttype", "P2P")]):
        r_ = t0 + 1 + i; e = r_ + 1; ref = R(f"$A{e}", val)
        sg.write(r_, 0, val, F["txt"])
        ft, vt = eng.count(**{key: ref}); fs, vs = eng.count(**{key: ref, "status": L("Success")}); fp, vp = eng.sum("amt", **{key: ref}, status=L("Success"))
        fy, vy = eng.count(**{key: ref, "ftype": L("System")})
        sg.write_formula(r_, 1, "=" + ft, F["int"], vt); sg.write_formula(r_, 2, "=" + fp, F["inr"], vp)
        sg.write_formula(r_, 3, f'=IFERROR(C{e}/SUMIFS(d_amt,d_status,"Success"),0)', F["pct"], vp / succ_total)
        sg.write_formula(r_, 4, f"=IFERROR({fs}/B{e},0)", F["pct"], vs / vt); sg.write_formula(r_, 5, f"=IFERROR({fy}/B{e},0)", F["pct2"], vy / vt)
        sg.write_formula(r_, 6, f"=IFERROR(C{e}/{fs},0)", F["inr"], vp / vs)
    s1 = wb.add_chart({"type": "bar"})
    s1.add_series({"name": "TPV", "categories": ["Segments", k0, 0, kt - 1, 0], "values": ["Segments", k0, 2, kt - 1, 2], "fill": {"color": BLUE}, "border": {"none": True}, "gap": 35})
    chart_style(s1, "TPV by category (Rs)")
    s1.set_y_axis({"reverse": True, "num_font": {"size": 8, "color": MUTED}, "line": {"color": LINE}, "major_tick_mark": "none"})
    s1.set_x_axis({"num_format": '#,##0.0,,"M"', "num_font": {"size": 8, "color": MUTED}, "major_gridlines": {"visible": True, "line": {"color": "#E6E9EE"}}, "line": {"none": True}})
    s1.set_size({"width": 560, "height": 330}); sg.insert_chart(3, 10, s1)
    s2 = wb.add_chart({"type": "column"})
    s2.add_series({"name": "System fail %", "categories": ["Segments", cy0 + 1, 0, cyt - 1, 0], "values": ["Segments", cy0 + 1, 6, cyt - 1, 6], "fill": {"color": ORANGE}, "border": {"none": True}, "gap": 40})
    chart_style(s2, "System failure rate by city (largest TPV first)", y_fmt="0.0%", x_rot=-45); s2.set_size({"width": 560, "height": 330}); sg.insert_chart(21, 10, s2)

    # ---------------------------------------------------------------- Fraud & Cashback
    fc_ = ws["Fraud_Cashback"]
    fc_.hide_gridlines(2)
    fc_.set_column(0, 0, 26); fc_.set_column(1, 7, 15)
    fc_.write(0, 0, "Fraud flags and cashback cost", F["title"]); fc_.set_row(0, 30)
    fc_.write(1, 0, "Fraud flag = transaction marked suspicious (bps = per 10,000 transactions). Loss = value of flagged payments that still went through. Small groups (for example amounts above Rs 25,000) are noisy.", F["sub"])
    row = 3
    fraud_blocks = {}
    for title, key, vals in [("A. By hour of day", "hband", sorted(df.hour_band.unique())), ("B. By amount band", "aband", sorted(df.amount_band.unique())),
                             ("C. By customer tenure", "tband", sorted(df.tenure_band.unique())), ("D. By payment type", "ttype", ["P2M", "P2P"])]:
        fc_.write(row, 0, title, F["h2"])
        for j, h in enumerate(["Group", "Transactions", "Flagged", "Fraud flag (bps)", "Loss (Rs)", "Loss % of TPV (bps)"]):
            fc_.write(row + 1, j, h, F["hdr"])
        fc_.set_row(row + 1, 32)
        start = row + 2
        for i, v in enumerate(vals):
            r_ = start + i; e = r_ + 1; ref = R(f"$A{e}", v)
            fc_.write(r_, 0, v, F["txt"])
            ft, vt = eng.count(**{key: ref}); ffg, vfg = eng.sum("fraud", **{key: ref}); fls, vls = eng.sum("floss", **{key: ref}); ftp2, vtp2 = eng.sum("amt", **{key: ref}, status=L("Success"))
            fc_.write_formula(r_, 1, "=" + ft, F["int"], vt); fc_.write_formula(r_, 2, "=" + ffg, F["int"], vfg)
            fc_.write_formula(r_, 3, f"=IFERROR(C{e}/B{e}*10000,0)", F["bps"], vfg / vt * 1e4); fc_.write_formula(r_, 4, "=" + fls, F["inr"], vls)
            fc_.write_formula(r_, 5, f"=IFERROR(E{e}/{ftp2}*10000,0)", F["bps"], vls / vtp2 * 1e4 if vtp2 else 0)
        fc_.conditional_format(start, 3, start + len(vals) - 1, 3, {"type": "data_bar", "bar_color": "#F4B79F", "bar_solid": True})
        fraud_blocks[key] = (start, start + len(vals) - 1)
        row = start + len(vals) + 2
    fc_.write(row, 0, "E. Cashback cost by category", F["h2"])
    for j, h in enumerate(["Category", "Cashback paid (Rs)", "TPV (Rs)", "Cashback % of TPV", "Share of cashback", "Share of TPV"]):
        fc_.write(row + 1, j, h, F["hdr"])
    fc_.set_row(row + 1, 32)
    cb_start = row + 2
    cb_total = float(df.cashback.sum())
    cb_order = df.groupby("category").cashback.sum().sort_values(ascending=False).index.tolist()
    for i, k in enumerate(cb_order):
        r_ = cb_start + i; e = r_ + 1; ref = R(f"$A{e}", k)
        fc_.write(r_, 0, k, F["txt"])
        fcb3, vcb3 = eng.sum("cb", cat=ref); ftp3, vtp3 = eng.sum("amt", cat=ref, status=L("Success"))
        fc_.write_formula(r_, 1, "=" + fcb3, F["dec2"], vcb3); fc_.write_formula(r_, 2, "=" + ftp3, F["inr"], vtp3)
        fc_.write_formula(r_, 3, f"=IFERROR(B{e}/C{e},0)", F["pct2"], vcb3 / vtp3 if vtp3 else 0)
        fc_.write_formula(r_, 4, f"=IFERROR(B{e}/SUM(d_cb),0)", F["pct"], vcb3 / cb_total); fc_.write_formula(r_, 5, f"=IFERROR(C{e}/SUMIFS(d_amt,d_status,\"Success\"),0)", F["pct"], vtp3 / succ_total)
    fc_.conditional_format(cb_start, 3, cb_start + len(cb_order) - 1, 3, {"type": "data_bar", "bar_color": "#BBD5F4", "bar_solid": True})
    f1 = wb.add_chart({"type": "column"})
    f1.add_series({"name": "Fraud flag (bps)", "categories": ["Fraud_Cashback", fraud_blocks["hband"][0], 0, fraud_blocks["hband"][1], 0], "values": ["Fraud_Cashback", fraud_blocks["hband"][0], 3, fraud_blocks["hband"][1], 3],
                   "fill": {"color": ORANGE}, "border": {"none": True}, "gap": 50, "data_labels": {"value": True, "num_format": "0", "font": {"size": 8, "color": INK}}})
    chart_style(f1, "Fraud flag rate by time of day (bps)", x_rot=-30); f1.set_size({"width": 520, "height": 290}); fc_.insert_chart(3, 8, f1)
    f2 = wb.add_chart({"type": "column"})
    f2.add_series({"name": "Fraud flag (bps)", "categories": ["Fraud_Cashback", fraud_blocks["tband"][0], 0, fraud_blocks["tband"][1], 0], "values": ["Fraud_Cashback", fraud_blocks["tband"][0], 3, fraud_blocks["tband"][1], 3],
                   "fill": {"color": ORANGE}, "border": {"none": True}, "gap": 60, "data_labels": {"value": True, "num_format": "0", "font": {"size": 8, "color": INK}}})
    chart_style(f2, "Fraud flag rate by customer tenure (bps)"); f2.set_size({"width": 520, "height": 290}); fc_.insert_chart(19, 8, f2)
    f3 = wb.add_chart({"type": "bar"})
    f3.add_series({"name": "Cashback % of TPV", "categories": ["Fraud_Cashback", cb_start, 0, cb_start + len(cb_order) - 1, 0], "values": ["Fraud_Cashback", cb_start, 3, cb_start + len(cb_order) - 1, 3],
                   "fill": {"color": BLUE}, "border": {"none": True}, "gap": 35})
    chart_style(f3, "Cashback cost as % of TPV, by category")
    f3.set_y_axis({"reverse": True, "num_font": {"size": 8, "color": MUTED}, "line": {"color": LINE}, "major_tick_mark": "none"})
    f3.set_x_axis({"num_format": "0.0%", "num_font": {"size": 8, "color": MUTED}, "major_gridlines": {"visible": True, "line": {"color": "#E6E9EE"}}, "line": {"none": True}})
    f3.set_size({"width": 520, "height": 320}); fc_.insert_chart(35, 8, f3)

    # ---------------------------------------------------------------- Incident finder
    inc = ws["Incident_Finder"]
    inc.hide_gridlines(2)
    inc.set_column(0, 0, 9); inc.set_column(1, 1, 24); inc.set_column(2, 3, 14); inc.set_column(4, 4, 24); inc.set_column(5, 6, 16)
    inc.write(0, 0, "Incident finder", F["title"]); inc.set_row(0, 30)
    inc.merge_range(1, 0, 1, 12, "Pick a bank and a month. A day is flagged when the bank has at least 5 transactions and the chance of that many system failures at its normal rate (all months) is below 0.1%. Try Dominion National Bank in 2025-10 and Greenfield Small Finance Bank in 2026-03.", F["sub"])
    inc.write(3, 0, "BANK", F["label"]); inc.write(4, 0, "MONTH", F["label"])
    sel_bank, sel_mon = "Dominion National Bank", "2025-10"
    inc.merge_range(3, 1, 3, 2, sel_bank, F["select"]); inc.merge_range(4, 1, 4, 2, sel_mon, F["select"])
    inc.data_validation(3, 1, 3, 1, {"validate": "list", "source": "=Lists!$D$2:$D$11"})
    inc.data_validation(4, 1, 4, 1, {"validate": "list", "source": "=Lists!$E$2:$E$23"})
    for j, h in enumerate(["Day", "Date", "Transactions", "System fails", "System fail %", "Chance by luck", "Flag"]):
        inc.write(6, j, h, F["hdr"])
    bank_all = df[df.bank_name == sel_bank]
    baseline = float((bank_all.failure_type == "System").mean())
    inc.write(4, 4, "Bank normal rate (all months)", F["label"])
    inc.write_formula(4, 5, '=IFERROR(COUNTIFS(d_bank,$B$4,d_ftype,"System")/COUNTIFS(d_bank,$B$4),0)', F["pct2"], baseline)
    sub = df[(df.bank_name == sel_bank) & (df.month_key == sel_mon)]
    day_first = 7
    for dd in range(1, 32):
        r_ = day_first + dd - 1; e = r_ + 1
        date_s = f"{sel_mon}-{dd:02d}"
        inc.write_number(r_, 0, dd, F["txt"])
        inc.write_formula(r_, 1, f'=$B$5&"-"&TEXT(A{e},"00")', F["txt"], date_s)
        dd_sub = sub[sub.txn_date == date_s]
        vt, vs = len(dd_sub), int((dd_sub.failure_type == "System").sum())
        pval = "" if vt < 5 else (1.0 if vs == 0 else float(1 - binom.cdf(vs - 1, vt, baseline)))
        inc.write_formula(r_, 2, f"=COUNTIFS(d_date,$B{e},d_bank,$B$4)", F["int"], vt)
        inc.write_formula(r_, 3, f'=COUNTIFS(d_date,$B{e},d_bank,$B$4,d_ftype,"System")', F["int"], vs)
        inc.write_formula(r_, 4, f"=IFERROR(D{e}/C{e},0)", F["pct"], vs / vt if vt else 0)
        inc.write_formula(r_, 5, f'=IF(C{e}<5,"",IF(D{e}=0,1,1-_xlfn.BINOM.DIST(D{e}-1,C{e},$F$5,TRUE)))', wb.add_format({"font_name": "Calibri", "font_size": 10, "font_color": INK, "border": 1, "border_color": LINE, "num_format": "0.00000"}), pval)
        inc.write_formula(r_, 6, f'=IF(C{e}<5,"",IF(F{e}<0.001,"SPIKE",""))', F["txt"], "SPIKE" if (pval != "" and pval < 0.001) else "")
    inc.conditional_format(day_first, 6, day_first + 30, 6, {"type": "cell", "criteria": "==", "value": '"SPIKE"', "format": F["bad"]})
    inc.conditional_format(day_first, 4, day_first + 30, 4, {"type": "data_bar", "bar_color": "#F4B79F", "bar_solid": True})
    i1 = wb.add_chart({"type": "column"})
    i1.add_series({"name": "System fail %", "categories": ["Incident_Finder", day_first, 0, day_first + 30, 0], "values": ["Incident_Finder", day_first, 4, day_first + 30, 4],
                   "fill": {"color": ORANGE}, "border": {"none": True}, "gap": 30})
    chart_style(i1, "Daily system failure rate for the selected bank", y_fmt="0%", x_title="Day of month"); i1.set_size({"width": 620, "height": 340}); inc.insert_chart(3, 8, i1)

    # ---------------------------------------------------------------- Insights (text written from the same data)
    ins = ws["Insights"]
    ins.hide_gridlines(2)
    ins.set_column(0, 0, 4); ins.set_column(1, 1, 26); ins.set_column(2, 2, 110)
    ins.write(0, 1, "Key findings and recommendations", F["title"]); ins.set_row(0, 30)
    ins.write(1, 1, "Computed from the simulated dataset when the workbook was built. Open the sheet named in the last column of each row to see the formulas behind it.", F["sub"])
    findings = build_findings(df, mon_vals, cmgr)
    ins.write(3, 1, "Finding", F["hdr_l"]); ins.write(3, 2, "What the data shows", F["hdr_l"])
    r_ = 4
    for title, text in findings["findings"]:
        ins.write(r_, 1, title, F["wrap_b"]); ins.write(r_, 2, text, F["wrap"]); ins.set_row(r_, 48); r_ += 1
    r_ += 1
    ins.write(r_, 1, "Recommendation", F["hdr_l"]); ins.write(r_, 2, "Why and expected effect", F["hdr_l"]); r_ += 1
    for title, text in findings["recs"]:
        ins.write(r_, 1, title, F["wrap_b"]); ins.write(r_, 2, text, F["wrap"]); ins.set_row(r_, 48); r_ += 1
    r_ += 1
    ins.merge_range(r_, 1, r_, 2, "Limits: the data is simulated, so effects are built into it by design. On real data the same formulas would apply, but conclusions would need statistical testing and domain review.", F["note"])
    ins.set_row(r_, 32)

    # ---------------------------------------------------------------- Dictionary
    dc = ws["Dictionary"]
    dc.set_column(0, 0, 22); dc.set_column(1, 1, 16); dc.set_column(2, 2, 90)
    dc.write(0, 0, "Data dictionary (sheet: Data)", F["title"]); dc.set_row(0, 28)
    for j, h in enumerate(["Column", "Type", "Meaning"]):
        dc.write(2, j, h, F["hdr_l"])
    for i, (a, b, c_) in enumerate(DICTIONARY):
        dc.write(3 + i, 0, a, F["txt"]); dc.write(3 + i, 1, b, F["txt"]); dc.write(3 + i, 2, c_, F["txt"])

    # ---------------------------------------------------------------- Cover
    cv = ws["Cover"]
    cv.hide_gridlines(2)
    cv.set_column(0, 0, 3); cv.set_column(1, 1, 22); cv.set_column(2, 2, 92)
    cv.set_row(1, 40)
    cv.write(1, 1, "PayFlow UPI Payments Analytics", F["title"])
    cv.write(2, 1, "Excel workbook: formula-driven analysis of payment success, failures, fraud flags and cashback cost.", F["sub"])
    cv.merge_range(4, 1, 4, 2, "Simulated data. PayFlow is a fictional company and the bank names are invented. The dataset (60,000 transactions, Jan 2025 to Oct 2026) was generated with realistic rules so the analysis has something to find.", F["note"])
    cv.set_row(4, 34)
    cv.write(6, 1, "Sheet", F["hdr_l"]); cv.write(6, 2, "What it answers", F["hdr_l"])
    guide = [("Dashboard", "Interactive KPIs and charts. Choose month, city and category in the blue boxes."),
             ("Monthly", "Growth, success rate, active users, fraud flags and cashback cost month by month."),
             ("Failures", "Which banks, hours, networks and customer groups fail most, and what the lost value is."),
             ("Segments", "Category, city, tier and payment-type performance."),
             ("Fraud_Cashback", "When and for whom fraud flags rise; how much cashback each category costs."),
             ("Incident_Finder", "Choose a bank and month; the daily chart exposes outages."),
             ("Insights", "Written findings and recommendations."),
             ("Data", "60,000 rows in an Excel Table (tblTxn). The source for every formula."),
             ("Dictionary", "Meaning of every column."), ("Calc", "Helper tables behind the dashboard."), ("Lists", "Dropdown sources.")]
    for i, (a, b) in enumerate(guide):
        cv.write_url(7 + i, 1, f"internal:'{a}'!A1", F["link"], a); cv.write(7 + i, 2, b, F["txt"])
    k = 7 + len(guide) + 1
    cv.write(k, 1, "Excel skills shown", F["hdr_l"]); cv.write(k, 2, "", F["hdr_l"])
    for i, t in enumerate(["COUNTIFS, SUMIFS and AVERAGEIFS over named ranges (d_month, d_amt, ...), wildcard 'All' filters",
                           "Excel Table with 60,000 rows, helper columns for bands, tenure and monthly active users",
                           "Data validation dropdowns driving live KPI cards and charts",
                           "Conditional formatting (data bars, colour scales, spike flags), charts, internal hyperlinks",
                           "Totals that reconcile: every breakdown sums back to the Data table"]):
        cv.write(k + 1 + i, 2, t, F["txt"])
    for name in ["Cover", "Dashboard", "Monthly", "Failures", "Segments", "Fraud_Cashback", "Incident_Finder", "Insights", "Dictionary"]:
        ws[name].set_landscape(); ws[name].set_paper(9); ws[name].fit_to_pages(1, 0); ws[name].set_margins(0.4, 0.4, 0.5, 0.5)
    cv.activate()
    wb.close()
    print(f"[build_excel] {path}  ({path.stat().st_size / 1e6:.1f} MB)")
    return path


DICTIONARY = [
    ("txn_id", "Number", "Unique transaction id"), ("txn_ts", "Text", "Timestamp, YYYY-MM-DD HH:MM:SS"), ("txn_date", "Text", "Date, YYYY-MM-DD (text so SUMIFS can match it)"),
    ("month_key", "Text", "YYYY-MM, used for monthly grouping"), ("txn_hour", "Number", "Hour of day, 0 to 23"), ("hour_band", "Text", "Night, Early, Morning, Day, Evening, Late"),
    ("customer_id", "Number", "Customer identifier"), ("city / tier", "Text", "Customer city and Tier 1 / Tier 2"), ("age_group", "Text", "18-24, 25-34, 35-44, 45+"),
    ("device_os", "Text", "Android or iOS"), ("kyc_level", "Text", "Min KYC or Full KYC"), ("tenure_days", "Number", "Days between signup and this transaction"),
    ("tenure_band", "Text", "0-30, 31-90, 91-180, 180+ days since signup"), ("merchant_id / merchant_name", "Text", "Merchant (0 = P2P transfer to an individual)"),
    ("category", "Text", "Spend category, including P2P Transfer"), ("txn_type", "Text", "P2M (merchant) or P2P (person to person)"), ("bank_name / bank_type", "Text", "Customer's bank (fictional) and its type"),
    ("network", "Text", "5G, 4G, WiFi or Weak (2G/3G)"), ("amount", "Number", "Transaction value in Rs"), ("amount_band", "Text", "Value band"), ("status", "Text", "Success or Failed"),
    ("failure_type", "Text", "System, User, Risk, or No failure"), ("failure_reason", "Text", "Specific reason, or No failure"), ("fraud_flag", "0 / 1", "1 if flagged suspicious"),
    ("fraud_loss", "Number", "Rs value of a flagged payment that was approved"), ("cashback", "Number", "Cashback paid to the customer in Rs"), ("response_ms", "Number", "Time to complete in milliseconds"),
    ("is_new_user_month", "0 / 1", "1 on a customer's first transaction in a month; SUM gives monthly active users"),
]


def build_findings(df: pd.DataFrame, mon_vals: list, cmgr: float) -> dict:
    ok = df.status == "Success"
    tpv = df[ok].amount.sum()
    sysf = df.failure_type == "System"
    by_bank = sysf.groupby(df.bank_name).mean().sort_values()
    worst, best = by_bank.index[-1], by_bank.index[0]
    peak = df.txn_hour.between(19, 22)
    weak = df.network == "Weak (2G/3G)"
    sfb_coop = df.bank_type.isin(["Small Finance", "Co-operative"])
    share_txn = sfb_coop.mean(); share_sys = sysf[sfb_coop].sum() / sysf.sum()
    lost = df[sysf].amount.sum()
    first, last = mon_vals[0], mon_vals[-1]
    d = df.assign(s=sysf)
    inc1 = d[(d.bank_name == "Dominion National Bank") & (d.txn_date.between("2025-10-13", "2025-10-15"))]
    inc1_base = d[(d.bank_name == "Dominion National Bank") & (d.month_key == "2025-10") & ~d.txn_date.between("2025-10-13", "2025-10-15")]
    inc2 = d[(d.bank_name == "Greenfield Small Finance Bank") & (d.txn_date.between("2026-03-03", "2026-03-05"))]
    inc2_base = d[(d.bank_name == "Greenfield Small Finance Bank") & (d.month_key == "2026-03") & ~d.txn_date.between("2026-03-03", "2026-03-05")]
    pin45 = (df[df.age_group == "45+"].failure_reason == "Wrong PIN").mean(); pin_o = (df[df.age_group != "45+"].failure_reason == "Wrong PIN").mean()
    new = df.tenure_days <= 30
    pin_new = (df[new].failure_reason == "Wrong PIN").mean(); pin_old = (df[~new].failure_reason == "Wrong PIN").mean()
    night = df.txn_hour <= 4
    cb_rate = df.cashback.sum() / tpv
    cbc = df.groupby("category").cashback.sum(); tpvc = df[ok].groupby("category").amount.sum()
    cb_share = (cbc / cbc.sum()); tp_share = (tpvc / tpvc.sum())
    top_cb = (cbc / tpvc).sort_values(ascending=False)
    t1, t2 = df[df.tier == "Tier 1"], df[df.tier == "Tier 2"]
    private_rate = sysf[df.bank_type == "Private"].mean()
    attempts = df.groupby("bank_name").amount.sum()
    bank_rate = sysf.groupby(df.bank_name).mean()
    bank_kind = df.drop_duplicates("bank_name").set_index("bank_name").bank_type
    recover = float(sum(0.5 * max(bank_rate[b] - private_rate, 0) * attempts[b] for b in bank_rate.index if bank_kind[b] != "Private"))
    gd = d.groupby(["bank_name", "txn_date"]).agg(t=("s", "size"), sy=("s", "sum")).reset_index()
    gd["p"] = gd.bank_name.map(d.groupby("bank_name").s.mean())
    gd["pt"] = np.where(gd.sy == 0, 1.0, 1 - binom.cdf(gd.sy - 1, gd.t, gd.p))
    gd["flag"] = (gd.t >= 5) & (gd.pt < 0.001)
    gd["outage"] = ((gd.bank_name == "Dominion National Bank") & gd.txn_date.between("2025-10-13", "2025-10-15")) | ((gd.bank_name == "Greenfield Small Finance Bank") & gd.txn_date.between("2026-03-03", "2026-03-05"))
    spike_elig, spike_n = int((gd.t >= 5).sum()), int(gd.flag.sum())
    spike_out, spike_hit = int((gd.outage & (gd.t >= 5)).sum()), int((gd.outage & gd.flag).sum())
    t2r, t1r = (t2.failure_type == "System").mean(), (t1.failure_type == "System").mean()
    w2, w1 = (t2.network == "Weak (2G/3G)").mean(), (t1.network == "Weak (2G/3G)").mean()
    if t2r > t1r * 1.05:
        tier_text = (f"Tier 2 cities are {len(t2) / len(df):.0%} of transactions. Their system failure rate is {t2r:.1%} against {t1r:.1%} in Tier 1, and weak-network attempts are {w2:.0%} against {w1:.0%} (sheet: Segments).")
    else:
        tier_text = (f"Tier 2 cities are {len(t2) / len(df):.0%} of transactions. Their system failure rate is {t2r:.1%} against {t1r:.1%} in Tier 1, so tier alone does not explain failures even though weak-network attempts are {w2:.0%} against {w1:.0%}. Bank and network matter more than city (sheet: Segments).")
    findings = [
        ("Scale and growth", f"{len(df):,} payments worth {inr_text(tpv)} succeeded at a {ok.mean():.1%} success rate. Monthly TPV grew from {inr_text(first['tpv'])} to {inr_text(last['tpv'])}, a compound growth rate of {cmgr:.1%} a month (sheet: Monthly)."),
        ("Bank reliability gap", f"{worst} has a {by_bank[worst]:.1%} system failure rate against {by_bank[best]:.1%} at {best}, {by_bank[worst] / by_bank[best]:.1f}x higher. Small-finance and co-operative banks carry {share_txn:.0%} of transactions but {share_sys:.0%} of system failures (sheet: Failures)."),
        ("Money left on the table", f"System failures put {inr_text(lost)} of attempted payment value at risk, about {lost / (tpv + lost):.1%} of all attempted value. Customers may retry, so this is value delayed or lost, not all lost (sheet: Failures)."),
        ("Evening peak and weak signal", f"System failure is {sysf[peak].mean():.1%} between 19:00 and 22:59 against {sysf[~peak].mean():.1%} at other times. On weak 2G/3G networks it is {sysf[weak].mean():.1%} against {sysf[~weak].mean():.1%} elsewhere (sheet: Failures)."),
        ("Outages are visible", f"Dominion National Bank's system failure rate was {inr_pct(inc1.s.mean())} on 13-15 Oct 2025 (Diwali week) against {inr_pct(inc1_base.s.mean())} for the rest of that month. Greenfield Small Finance Bank hit {inr_pct(inc2.s.mean())} on 3-5 Mar 2026 against {inr_pct(inc2_base.s.mean())} (sheet: Incident_Finder)."),
        ("User errors cluster", f"Customers aged 45+ enter a wrong PIN on {pin45:.2%} of payments against {pin_o:.2%} for others; customers in their first 30 days do so on {pin_new:.2%} against {pin_old:.2%} later. These are fixable with guidance, not infrastructure (sheet: Failures, section E)."),
        ("Fraud flags concentrate", f"Fraud flags run at {df[new].fraud_flag.mean() * 1e4:.0f} bps for customers in their first 30 days against {df[~new].fraud_flag.mean() * 1e4:.0f} bps later, and {df[night].fraud_flag.mean() * 1e4:.0f} bps between midnight and 04:59 against {df[~night].fraud_flag.mean() * 1e4:.0f} bps otherwise (sheet: Fraud_Cashback)."),
        ("Cashback cost", f"Cashback costs {cb_rate:.2%} of TPV. {top_cb.index[0]} and {top_cb.index[1]} are the most expensive at {top_cb.iloc[0]:.2%} and {top_cb.iloc[1]:.2%} of their TPV; together they take {cb_share[top_cb.index[:2]].sum():.0%} of cashback spend for {tp_share[top_cb.index[:2]].sum():.0%} of TPV (sheet: Fraud_Cashback)."),
        ("Tier 2 cities", tier_text),
    ]
    recs = [
        ("Route around weak banks", f"Add smart retry or alternate-bank prompts for {worst} and the other small-finance and co-operative banks. Closing half of each non-private bank's gap to the private-bank rate ({private_rate:.1%}) would protect roughly {inr_text(recover)} of attempted value over the 22 months (simple estimate: excess failure rate x attempted value x 50%)."),
        ("Scale for 19:00-22:59", "Add capacity and pre-warm connections before the evening peak and festive weeks, where failure and response time both rise."),
        ("Alert on daily failure spikes", f"Use the Incident_Finder rule as an automated alert: at least 5 transactions and under 0.1% chance of that many system failures at the bank's normal rate. On this data it flags {spike_n} of {spike_elig:,} bank-days and catches {spike_hit} of the {spike_out} outage days that have enough traffic to test. A simple 2.5x-average rule flagged 17% of bank-days, which is why it was rejected."),
        ("Guide new and 45+ users", "Show PIN-entry help and a retry nudge for the first 30 days and for users aged 45+."),
        ("Step-up checks for risky moments", "Apply extra verification for new customers and for night-time transfers, where flag rates are highest, and review the loss it prevents."),
        ("Re-target cashback", "Review cashback on the categories with the highest cost per rupee of TPV and test whether it changes repeat behaviour before keeping it."),
    ]
    return {"findings": findings, "recs": recs}


def inr_pct(x: float) -> str:
    return f"{x:.0%}" if x >= 0.1 else f"{x:.1%}"


if __name__ == "__main__":
    build()
