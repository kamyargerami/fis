#!/usr/bin/env python3
"""FIS (Fundamental Investing System) scorer.

Computes the 11 fundamental variables from raw financial-statement data and
scores each one with the threshold-based (non-linear, no partial credit)
scoring function described in Khosravani & Ajorlou (JAMV, 1405).

Usage:
    python3 fis_score.py input.json            # human-readable report
    python3 fis_score.py input.json --json     # machine-readable output

Input JSON: see references/input_template.json. Any variable can also be
given pre-computed under "precomputed": {"forward_pe": 4.2, ...}.

NOTE: the paper publishes the max points per variable (Table 1) but NOT the
thresholds T_i,k / points P_i,k. The bins below are calibrated assumptions
for the Iranian market (high inflation, risk-free ~35-40%) and can be edited
in THRESHOLDS.
"""
import json
import sys

# Each entry: (max_points, direction, bins)
# bins: list of (upper_bound, points); value <= upper_bound -> points.
# The last bin uses float("inf"). Points per bin are a step function
# S_i(X_i) exactly as in the paper: no interpolation inside a band.
INF = float("inf")
THRESHOLDS = {
    "forward_pe": {
        "label": "P/E آینده‌نگر",
        "max": 15,
        "bins": [(0, 0), (4, 15), (5, 12), (6, 9), (8, 6), (10, 3), (INF, 0)],
    },
    "sales_growth": {
        "label": "رشد فروش تجمعی (%)",
        "max": 15,
        "bins": [(0, 0), (20, 3), (35, 6), (50, 9), (75, 12), (INF, 15)],
    },
    "fis_margin": {
        "label": "حاشیه سود FIS (%)",
        "max": 10,
        "bins": [(0, 0), (5, 2), (10, 4), (20, 6), (30, 8), (INF, 10)],
    },
    "margin_trend": {
        "label": "رشد روند حاشیه سود فصلی (%)",
        "max": 10,
        "bins": [(-20, 0), (-10, 2), (0, 4), (10, 7), (INF, 10)],
    },
    "forward_ps": {
        "label": "P/S آینده‌نگر",
        "max": 10,
        "bins": [(0, 0), (0.5, 10), (1, 8), (1.5, 6), (2.5, 4), (4, 2), (INF, 0)],
    },
    "dividend_payout": {
        "label": "نسبت تقسیم سود (%)",
        "max": 10,
        "bins": [(0, 0), (20, 2), (40, 4), (60, 7), (INF, 10)],
    },
    "receivables_to_assets": {
        "label": "مطالبات به دارایی (%)",
        "max": 10,
        "bins": [(10, 10), (20, 8), (30, 6), (40, 3), (INF, 0)],
    },
    "earnings_expectation_gap": {
        "label": "شکاف انتظارات سود (%)",
        "max": 5,
        "bins": [(10, 5), (25, 4), (50, 2), (INF, 0)],
    },
    "collection_days": {
        "label": "دوره وصول مطالبات (روز)",
        "max": 5,
        "bins": [(60, 5), (120, 4), (180, 2), (270, 1), (INF, 0)],
    },
    "inventory_days": {
        "label": "دوره گردش کالا (روز)",
        "max": 5,
        "bins": [(60, 5), (120, 4), (180, 2), (270, 1), (INF, 0)],
    },
    "fiscal_year_coverage": {
        "label": "سال مالی (ماه‌های واقعی گزارش‌شده)",
        "max": 5,
        "bins": [(2, 1), (5, 2), (8, 4), (INF, 5)],
    },
}

ORDER = list(THRESHOLDS)


def score_value(key, x):
    if x is None:
        return None
    for upper, pts in THRESHOLDS[key]["bins"]:
        if x <= upper:
            return pts
    return 0


def _get(d, *path):
    for p in path:
        if not isinstance(d, dict) or d.get(p) is None:
            return None
        d = d[p]
    return d


def compute_variables(data):
    """Derive the 11 variables from raw inputs; precomputed values win."""
    pre = data.get("precomputed", {}) or {}
    v, notes = {}, []

    price = data.get("price")
    shares = data.get("shares_outstanding")
    market_cap = data.get("market_cap") or (price * shares if price and shares else None)

    # Forward annualisation from monthly reports: avg monthly x 12
    months = data.get("months_elapsed")
    ytd_earn = data.get("ytd_net_income_fis")
    ytd_sales = data.get("ytd_sales")
    fwd_earn = data.get("forward_net_income")
    fwd_sales = data.get("forward_sales")
    if fwd_earn is None and ytd_earn is not None and months:
        fwd_earn = ytd_earn / months * 12
    if fwd_sales is None and ytd_sales is not None and months:
        fwd_sales = ytd_sales / months * 12

    # 1. Forward P/E
    if market_cap and fwd_earn:
        v["forward_pe"] = market_cap / fwd_earn if fwd_earn > 0 else -1
    # 2. Cumulative sales growth vs same period last year
    cur = data.get("sales_period_current", ytd_sales)
    prev = data.get("sales_period_prior_year")
    if cur is not None and prev:
        v["sales_growth"] = (cur - prev) / prev * 100
    # 3. FIS margin: profit ex non-operating / non-recurring over operating revenue
    q = data.get("quarters", {}) or {}
    rev, fis_profit = _get(q, "current", "revenue"), _get(q, "current", "fis_profit")
    if rev and fis_profit is not None:
        v["fis_margin"] = fis_profit / rev * 100
    # 4. Margin trend QoQ
    prev_rev, prev_profit = _get(q, "previous", "revenue"), _get(q, "previous", "fis_profit")
    if "fis_margin" in v and prev_rev and prev_profit is not None:
        pm = prev_profit / prev_rev * 100
        if pm > 0:
            v["margin_trend"] = (v["fis_margin"] - pm) / pm * 100
        else:
            notes.append("حاشیه فصل قبل ≤ صفر؛ رشد روند حاشیه معنادار نیست.")
    # 5. Forward P/S
    if market_cap and fwd_sales:
        v["forward_ps"] = market_cap / fwd_sales
    # 6. Dividend payout: mean of (dividend / accumulated retained earnings) over 2 years
    d = data.get("dividends", {}) or {}
    ratios = []
    for yr in ("last_year", "this_year"):
        div, acc = _get(d, yr, "dividend"), _get(d, yr, "retained_earnings")
        if div is not None and acc:
            ratios.append(div / acc * 100)
    if len(ratios) == 2:
        v["dividend_payout"] = sum(ratios) / 2
    elif len(ratios) == 1:
        v["dividend_payout"] = ratios[0]
        notes.append("نسبت تقسیم سود فقط با داده یک سال محاسبه شد.")
    # 7. Receivables / assets
    bs = data.get("balance_sheet", {}) or {}
    ar, ta = bs.get("receivables"), bs.get("total_assets")
    if ar is not None and ta:
        v["receivables_to_assets"] = ar / ta * 100
    # 8. Earnings expectation gap |PE_fwd - PE_ttm| / PE_ttm
    ttm = data.get("ttm_net_income")
    if market_cap and ttm and ttm > 0 and v.get("forward_pe", -1) > 0:
        pe_ttm = market_cap / ttm
        v["pe_ttm"] = pe_ttm
        v["earnings_expectation_gap"] = abs(v["forward_pe"] - pe_ttm) / pe_ttm * 100
    # 9. Collection period
    days = data.get("period_days", 365)
    sales_total = data.get("period_sales")
    ar_avg = bs["receivables_avg"] if bs.get("receivables_avg") is not None else ar  # opening/closing average
    if ar_avg is not None and sales_total:
        v["collection_days"] = ar_avg / sales_total * days
    # 10. Inventory period
    inv = bs.get("average_inventory")
    if inv is None and bs.get("inventory_open") is not None and bs.get("inventory_close") is not None:
        inv = (bs["inventory_open"] + bs["inventory_close"]) / 2
    cogs = data.get("period_cogs")
    if inv is not None and cogs:
        v["inventory_days"] = inv / cogs * days
    # 11. Fiscal year coverage (interpretation, see SKILL.md)
    if months is not None:
        v["fiscal_year_coverage"] = months

    for k, val in pre.items():
        if val is not None:
            v[k] = val
    return v, notes


def rating(total, max_avail):
    # Paper: the 30 selected stocks (top of those above market-average score)
    # scored about 60-72 (Fig. 2). No other bands are defined in the paper.
    pct = total / max_avail * 100 if max_avail else 0
    if pct >= 60:
        return "هم‌سطح یا بالاتر از امتیاز سبد ۳۰ سهم منتخب مقاله (حدود ۶۰ تا ۷۲)"
    return "پایین‌تر از محدوده امتیاز سبد منتخب مقاله (کمتر از ۶۰)"


def evaluate(data):
    eligibility = []
    if data.get("is_loss_making"):
        eligibility.append("شرکت زیان‌ده است؛ طبق FIS از چرخه امتیازدهی حذف می‌شود.")
    if data.get("has_monthly_reports") is False:
        eligibility.append("گزارش ماهانه قابل‌اتکا ندارد؛ طبق FIS حذف می‌شود.")

    v, notes = compute_variables(data)
    rows, total, max_avail = [], 0, 0
    for k in ORDER:
        x = v.get(k)
        s = score_value(k, x)
        mx = THRESHOLDS[k]["max"]
        if s is not None:
            total += s
            max_avail += mx
        rows.append({"key": k, "label": THRESHOLDS[k]["label"], "value": x,
                     "score": s, "max": mx})

    missing = [r["label"] for r in rows if r["score"] is None]
    out = {
        "company": data.get("company"),
        "eligible": not eligibility,
        "eligibility_issues": eligibility,
        "rows": rows,
        "total": total,
        "max_available": max_avail,
        "normalized_100": round(total / max_avail * 100, 1) if max_avail else None,
        "rating": rating(total, max_avail),
        "missing": missing,
        "notes": notes,
    }

    # Entry/exit P/E band (paper): buy = 1/rf + market premium of 1-1.5 P/E units
    # (rf=40% -> 2.5 + 1..1.5 = 3.5-4); exit ~ 7-8, i.e. about twice the buy band.
    rf = data.get("risk_free_rate")
    if rf:
        base = 100 / rf
        buy = (round(base + 1, 2), round(base + 1.5, 2))
        out["pe_band"] = {"buy": buy, "exit": (round(buy[0] * 2, 2), round(buy[1] * 2, 2))}
        if v.get("forward_pe", -1) > 0:
            pe = v["forward_pe"]
            out["pe_signal"] = ("در محدوده خرید یا ارزان‌تر" if pe <= buy[1] else
                                "در محدوده خروج یا گران‌تر" if pe >= out["pe_band"]["exit"][0] else
                                "بین محدوده خرید و خروج (نگهداری، نه ورود جدید)")

    # Paper regression (n=30, scores 60-72, 1402-1405): Return% = -1357.28 + 27.51*Score
    if not missing:
        est = -1357.28 + 27.51 * total
        out["regression_3y_return_pct"] = round(est, 1)
        out["regression_in_range"] = 60 <= total <= 72
    return out


def fmt(x):
    if x is None:
        return "—"
    if isinstance(x, float):
        return f"{x:,.2f}"
    return str(x)


def print_report(o):
    print(f"# نمره FIS — {o.get('company') or 'شرکت'}\n")
    if not o["eligible"]:
        for e in o["eligibility_issues"]:
            print(f"⚠️ {e}")
        print()
    print("| متغیر | مقدار | امتیاز | حداکثر |")
    print("|---|---:|---:|---:|")
    for r in o["rows"]:
        print(f"| {r['label']} | {fmt(r['value'])} | {fmt(r['score'])} | {r['max']} |")
    print(f"| **جمع** | | **{o['total']}** | **{o['max_available']}** |\n")
    if o["missing"]:
        print(f"داده ناقص برای: {', '.join(o['missing'])} — نمره نرمال‌شده: {o['normalized_100']}/100\n")
    print(f"رتبه: {o['rating']}")
    if "pe_band" in o:
        b, e = o["pe_band"]["buy"], o["pe_band"]["exit"]
        print(f"محدوده P/E خرید: {b[0]}–{b[1]} | خروج: {e[0]}–{e[1]} | وضعیت: {o.get('pe_signal', '—')}")
    if "regression_3y_return_pct" in o:
        flag = "" if o["regression_in_range"] else " (خارج از بازه ۶۰–۷۲ نمونه مقاله؛ اتکاپذیر نیست)"
        print(f"برآورد بازده ۳ساله طبق رگرسیون مقاله: {o['regression_3y_return_pct']}%{flag}")
    for n in o["notes"]:
        print(f"- {n}")


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    with open(sys.argv[1], encoding="utf-8") as f:
        data = json.load(f)
    o = evaluate(data)
    if "--json" in sys.argv:
        print(json.dumps(o, ensure_ascii=False, indent=2))
    else:
        print_report(o)


if __name__ == "__main__":
    main()
