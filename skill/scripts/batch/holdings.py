"""Holdings / investment companies: NAV at today's prices + look-through forward earnings.

Usage: python3 -I holdings.py WORKDIR
Reads WORKDIR/results/*.json (from run_all.py), picks companies whose Codal statement format is
«سرمایه‌گذاری» (funds excluded), and writes WORKDIR/holdings.json.

Method (professor's rule + accounting standard for dividend recognition):
- Portfolio: listed holdings (بورس/فرابورس/پایه) from the latest monthly portfolio report (ن-۳۰,
  Excel export). Unlisted holdings are excluded everywhere (NAV and earnings).
- Value at TODAY's price (TSETMC market watch) × shares held. If a holding can't be matched to a
  listed instrument, the report's own market value is used and flagged.
- NAV = listed value today + (total assets − long/short-term investments) − total liabilities.
- Main ratio = market cap ÷ NAV (P/NAV): lower is better.
- Look-through forward earnings for the holding's CURRENT fiscal year:
  * ownership < 20%: dividend income is recognised on the investee's AGM date, so an investee
    fiscal year counts only if its AGM falls inside the holding's fiscal year. Actual DPS if that
    AGM was held, otherwise forecast EPS × the investee's last payout ratio.
  * ownership ≥ 20%: share of the investee's net income for the fiscal year that ENDS inside the
    holding's fiscal year (actual if published, otherwise forecast).
  minus the holding's own costs (last full year: revenue − net income).
- Forward dividend = forward earnings × the holding's last payout ratio; the reported (last AGM)
  dividend is shown too.
"""
import datetime
import json
import os
import statistics
import sys
import time
import traceback
import urllib.error

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import fetch  # noqa: E402
from collect import add_months, collect, expand, forward_net  # noqa: E402
from fetch import excel_tables, fa_num, market_watch, match_instrument, name_key, norm  # noqa: E402

AGM_LAG_MONTHS = 4  # Iranian law: AGM within 4 months of fiscal year end


def g2j(d):
    """Gregorian date -> Jalali 'YYYY/MM/DD'."""
    gy, gm, gd = d.year, d.month, d.day
    g_d_m = [0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334]
    gy2 = gy + 1 if gm > 2 else gy
    days = 355666 + (365 * gy) + ((gy2 + 3) // 4) - ((gy2 + 99) // 100) + ((gy2 + 399) // 400) + gd + g_d_m[gm - 1]
    jy = -1595 + (33 * (days // 12053))
    days %= 12053
    jy += 4 * (days // 1461)
    days %= 1461
    if days > 365:
        jy += (days - 1) // 365
        days = (days - 1) % 365
    if days < 186:
        jm, jd = 1 + days // 31, 1 + days % 31
    else:
        jm, jd = 7 + (days - 186) // 30, 1 + (days - 186) % 30
    return f"{jy:04d}/{jm:02d}/{jd:02d}"


TODAY = g2j(datetime.date.today())


def roll_fy(year_end):
    """(last completed FY end, current FY end) relative to today."""
    e = year_end
    while e[:7] <= TODAY[:7]:
        e = add_months(e, 12)
    while add_months(e, -12)[:7] > TODAY[:7]:
        e = add_months(e, -12)
    return add_months(e, -12), e


def in_window(date, start, end):
    return start[:7] < date[:7] <= end[:7]


def is_holding(r):
    f = (r.get("format") or "").replace("‌", " ")
    return not r.get("error") and "سرمایه گذاری" in f and "صندوق" not in f


# ---------------------------------------------------------------- portfolio report
def parse_portfolio(serial):
    """Listed holdings + subsidiaries' dividend table from a monthly portfolio report (Excel export)."""
    holdings, divs = [], {}
    for t in excel_tables(serial):
        if len(t) < 3:
            continue
        h1 = [norm(x) for x in expand(t[0])]
        h2 = [norm(x) for x, _ in t[1]]
        if norm("نام شرکت") in h1 and norm("سرمایه (میلیون ریال)") in h1 and norm("درصد مالکیت") in h2 \
                and norm("ارزش بازار") in h2 and not holdings:
            # sub-header columns start after name, capital, par
            pos = {}
            seen_shares = 0
            for i, lab in enumerate(h2):
                c = 3 + i
                if lab == norm("تعداد سهام"):
                    pos["s0" if seen_shares == 0 else "ds"] = c
                    seen_shares += 1
                elif lab == norm("درصد مالکیت"):
                    pos["own"] = c
                elif lab == norm("ارزش بازار") and "own" in pos and "mv1" not in pos:
                    pos["mv1"] = c
                elif lab == norm("بهای تمام شده") and "own" in pos and "cost1" not in pos:
                    pos["cost1"] = c  # end-of-period cost
                elif lab == norm("ارزش هر سهم (ریال)"):
                    pos["vps"] = c
            for r in t[2:]:
                cells = [x for x, _ in r]
                if len(cells) < 6:
                    continue
                name = cells[0]
                if not name or norm(name).startswith("جمع"):
                    continue
                cap, par = fa_num(cells[1]), fa_num(cells[2])
                if not cap or not par:
                    continue
                get = lambda k: fa_num(cells[pos[k]]) if k in pos and pos[k] < len(cells) else None  # noqa
                s_end = (get("s0") or 0) + (get("ds") or 0)
                if s_end <= 0:
                    continue
                total = cap * 1e6 / par
                holdings.append({"name": name, "capital": cap, "par": par, "shares_held": s_end,
                                 "total_shares": total, "own": s_end / total,
                                 "report_mv": get("mv1"), "report_vps": get("vps"), "report_cost": get("cost1")})
        elif norm("تاریخ برگزاری مجمع") in h2 or norm("تاریخ برگزاری مجمع") in h1:
            for r in t[2:]:
                c = [x for x, _ in r]
                if len(c) >= 9 and fa_num(c[3]) and "/" in c[2]:
                    divs[name_key(c[0])] = {"fy_end": c[1].translate(str.maketrans("۰۱۲۳۴۵۶۷۸۹", "0123456789")),
                                            "agm": c[2].translate(str.maketrans("۰۱۲۳۴۵۶۷۸۹", "0123456789")),
                                            "eps": fa_num(c[7]), "dps": fa_num(c[8])}
    return holdings, divs


# ---------------------------------------------------------------- investees
_inv_cache = {}
_results = None  # norm(symbol) -> path of run_all.py's result for that symbol


def collected(symbol, workdir):
    """run_all.py's result for symbol, if it has one that was not cut short by a failed fetch."""
    global _results
    if _results is None:
        res_dir = os.path.join(workdir, "results")
        _results = {norm(f[:-5]): os.path.join(res_dir, f) for f in os.listdir(res_dir) if f.endswith(".json")}
    path = _results.get(norm(symbol))
    if not path:
        return None
    try:
        r = json.load(open(path, encoding="utf-8"))
    except ValueError:
        return None
    return None if "fetch failed" in (r.get("error") or "") else r


def investee(symbol, workdir):
    if symbol in _inv_cache:
        return _inv_cache[symbol]
    path = os.path.join(workdir, "investees", f"{symbol}.json")
    r = collected(symbol, workdir)  # already collected by run_all.py: same collect() output, no requests
    if r is None and os.path.exists(path) and \
            datetime.date.fromtimestamp(os.path.getmtime(path)) == datetime.date.today():
        r = json.load(open(path, encoding="utf-8"))
    else:
        before = fetch.TRANSIENT
        if r is None:
            print(f"  investee {symbol}: collecting", flush=True)
            try:
                r = collect(symbol)
            except Exception as e:  # noqa: BLE001
                r = {"symbol": symbol, "error": repr(e)}
        # saved also when taken from results/, so the pipeline's stale mark (mtime 0) is cleared
        if fetch.TRANSIENT == before:  # a failed fetch is not saved, so the next run retries it
            os.makedirs(os.path.dirname(path), exist_ok=True)
            json.dump(r, open(path + ".tmp", "w", encoding="utf-8"), ensure_ascii=False, default=str)
            os.replace(path + ".tmp", path)
    _inv_cache[symbol] = r
    return r


def investee_net(r):
    """(net income last completed FY, forecast net income current FY, last FY end, current FY end)."""
    fy_last, fy_cur = roll_fy(r["year_end"])
    fwd, _ = forward_net(r)
    f = (r.get("format") or "").replace("‌", " ")
    if "سرمایه گذاری" in f and r.get("fy_net") is not None:
        fwd = r["fy_net"]  # nested holding: interim dividend income is seasonal, use last full year
    if (r.get("prev_fy_end") or "")[:7] == fy_last[:7]:
        last = r.get("fy_net")
    elif (r.get("year_end") or "")[:7] == fy_last[:7]:
        last = fwd  # that year is over but its annual report is not out yet
    else:
        last = r.get("ttm_net")
    return last, fwd, fy_last, fy_cur


def payout(r):
    fy = r.get("fy_net")
    d = r.get("last_dividend")
    if not fy or fy <= 0 or not d:
        return None
    return max(0.0, min(d / fy, 1.2))


# ---------------------------------------------------------------- one holding
def evaluate_holding(h, workdir, mw):
    notes = []
    mc = h["price"] * h["shares"] / 1e6 if h.get("price") and h.get("shares") else None
    if not h.get("m_serial"):
        return {"symbol": h["symbol"], "error": "گزارش پورتفوی ماهانه پیدا نشد"}
    rows, divs = parse_portfolio(h["m_serial"])
    if not rows:
        return {"symbol": h["symbol"], "error": "جدول پورتفوی بورسی در گزارش ماهانه پیدا نشد"}
    h_last, h_cur = roll_fy(h["year_end"])
    w_start, w_end = h_last, h_cur
    listed_today = 0.0
    income = 0.0
    cash_income = 0.0
    detail = []
    payouts = []
    unmatched_income = []
    unrealized, listed_cost, no_cost = 0.0, 0.0, 0
    for p in rows:
        ins, score = match_instrument(p["name"], p["total_shares"], mw)
        item = dict(p)
        if ins and ins.get("price"):
            item["symbol"], item["price_today"] = ins["symbol"], ins["price"]
            item["value_today"] = p["shares_held"] * ins["price"] / 1e6
        else:
            item["symbol"] = None
            item["value_today"] = p["report_mv"] or 0
            item["note"] = "قیمت روز پیدا نشد؛ ارزش بازار گزارش پورتفو"
        listed_today += item["value_today"] or 0
        if p.get("report_cost") is not None:  # unrealized gain of the listed portfolio at today's price
            listed_cost += p["report_cost"]
            item["unrealized"] = (item["value_today"] or 0) - p["report_cost"]
            unrealized += item["unrealized"]
        else:
            no_cost += 1

        # look-through income
        dv = divs.get(name_key(p["name"]))
        r = investee(item["symbol"], workdir) if item["symbol"] else {"error": "بدون نماد"}
        inc = None
        if not r.get("error") and r.get("year_end"):
            last, fwd, e_last, e_cur = investee_net(r)
            po = payout(r)
            if po is not None:
                payouts.append(po)
            item.update(inv_fy_last=e_last, inv_fy_cur=e_cur, inv_net_last=last, inv_net_fwd=fwd, inv_payout=po)
            # cash dividends arriving in the holding's fiscal year (any ownership): AGM-date basis
            cash = 0.0
            agm_last = dv["agm"] if dv and (dv.get("fy_end") or "")[:7] == e_last[:7] else add_months(e_last, AGM_LAG_MONTHS)
            agm_cur = add_months(e_cur, AGM_LAG_MONTHS)
            held_last = (r.get("agm_held_this_period") and (r.get("prev_fy_end") or "")[:7] == e_last[:7]) or \
                (dv is not None and dv.get("agm", "") <= TODAY)
            if in_window(agm_last, w_start, w_end):
                if held_last and dv and dv.get("dps") is not None:
                    cash += p["shares_held"] * dv["dps"] / 1e6
                elif held_last and r.get("last_dividend") and r.get("shares"):
                    cash += p["shares_held"] * (r["last_dividend"] * 1e6 / r["shares"]) / 1e6
                elif last is not None and po is not None:
                    cash += p["own"] * last * po
            if in_window(agm_cur, w_start, w_end) and fwd is not None and po is not None:
                cash += p["own"] * fwd * po
            # earnings recognised (accounting rule): <20% on AGM date (= cash), >=20% share of that year's net income
            if p["own"] >= 0.2:
                inc = 0.0
                for e, ni in ((e_last, last), (e_cur, fwd)):
                    if ni is not None and in_window(e, w_start, w_end):
                        inc += p["own"] * ni
            else:
                inc = cash
            item["cash_div_fwd"] = cash
            cash_income += cash
        elif dv and dv.get("dps") is not None and in_window(dv.get("agm", ""), w_start, w_end):
            inc = p["shares_held"] * dv["dps"] / 1e6
            item["cash_div_fwd"] = inc
            cash_income += inc
            item["note"] = (item.get("note", "") + " / صورت مالی پیدا نشد؛ سود تقسیمی از جدول گزارش پورتفو").strip(" /")
        else:
            unmatched_income.append(item)
        item["income_fwd"] = inc
        if inc is not None:
            income += inc
        detail.append(item)

    # investees without statements: TSETMC EPS × median payout (flagged)
    med_po = statistics.median(payouts) if payouts else 0.5
    for item in unmatched_income:
        ins = mw.get(norm(item["symbol"] or "")) if item.get("symbol") else None
        if ins and ins.get("eps"):
            inc = item["shares_held"] * ins["eps"] * med_po / 1e6
            item["income_fwd"] = item["cash_div_fwd"] = inc
            income += inc
            cash_income += inc
            item["note"] = (item.get("note", "") + " / سود از EPS تخمینی TSETMC × میانه نسبت تقسیم").strip(" /")
        else:
            item["note"] = (item.get("note", "") + " / سود این سهم محاسبه نشد").strip(" /")

    costs = 0.0
    if h.get("fy_rev") is not None and h.get("fy_net") is not None:
        costs = max(0.0, h["fy_rev"] - h["fy_net"])
    fwd_net = income - costs
    other_assets = (h.get("total_assets") or 0) - (h.get("lt_investments") or 0) - (h.get("st_investments") or 0)
    # NAV = equity + unrealized gain of the listed portfolio (today's value - cost); unlisted investments
    # stay at cost, which is already in equity
    cash_inc = h.get("cash_increase") or 0  # capital increase paid in cash after the report
    nav = (h.get("equity") or 0) + unrealized + cash_inc
    mc_val = mc - cash_inc if mc else mc  # earnings ratios: the new cash has not earned anything yet
    if cash_inc:
        notes.append(f"آورده نقدی افزایش سرمایه بعد از گزارش ({cash_inc / 1e4:,.0f} میلیارد تومان) به NAV اضافه "
                     "و در P/E و P/D از ارزش بازار کم شد")
    if no_cost:
        notes.append(f"{no_cost} سهم بدون بهای تمام شده در گزارش پورتفو؛ سود تحقق‌نیافته آن‌ها در NAV نیست")
    po_h = payout(h)
    # the holding pays dividends out of cash dividends it receives, not equity-method income
    cash_net = cash_income - costs
    fwd_div = cash_net * po_h if po_h is not None and cash_net > 0 else None
    last_div = h.get("last_dividend")
    if h.get("capital") and h.get("shares"):
        chg = h["shares"] / (h["capital"] * 1e6 / 1000) - 1
        if abs(chg) > 0.02:
            notes.append(f"افزایش سرمایه بعد از گزارش ({chg * 100:+.0f}٪ سهام): ارزش بازار با تعداد جدید است؛ "
                         "اگر آورده نقدی بوده، NAV آن را هنوز ندارد")
    n_unmatched = sum(1 for d in detail if not d.get("symbol"))
    if n_unmatched:
        notes.append(f"{n_unmatched} سهم بدون قیمت روز (ارزش گزارش استفاده شد)")
    if not po_h:
        notes.append("نسبت تقسیم سود هلدینگ مشخص نیست؛ سود تقسیمی آینده‌نگر محاسبه نشد")
    div = lambda a, b: (a / b) if (a is not None and b) else None  # noqa: E731
    return {
        "symbol": h["symbol"], "name": h.get("name"), "price": h.get("price"), "market_cap": mc,
        "fy_window": f"{w_start} تا {w_end}", "portfolio_report": h.get("m_title"),
        "n_listed": len(detail), "listed_value_today": listed_today,
        "listed_value_report": sum(d.get("report_mv") or 0 for d in detail),
        "other_assets": other_assets, "total_liabilities": h.get("total_liabilities"), "nav": nav,
        "equity": h.get("equity"), "listed_cost": listed_cost, "unrealized": unrealized,
        "roa": div(h.get("ttm_net"), h.get("total_assets")), "roe": div(h.get("ttm_net"), h.get("equity")),
        "p_nav": div(mc, nav) if nav and nav > 0 else None,
        "income_fwd": income, "cash_income_fwd": cash_income, "costs": costs, "net_fwd": fwd_net,
        "cash_net_fwd": cash_net, "payout": po_h,
        "div_fwd": fwd_div, "div_last": last_div,
        "dps_fwd": fwd_div * 1e6 / h["shares"] if fwd_div and h.get("shares") else None,
        "dps_last": last_div * 1e6 / h["shares"] if last_div and h.get("shares") else None,
        "pd_fwd": div(mc_val, fwd_div), "pd_last": div(mc_val, last_div), "cash_increase": cash_inc or None,
        "pe_fwd": div(mc_val, fwd_net) if fwd_net > 0 else None,
        "pe_ttm": div(mc_val, h.get("ttm_net")) if h.get("ttm_net") and h["ttm_net"] > 0 else None,
        "pb": div(mc, h.get("equity")), "pa": div(mc, h.get("total_assets")),
        "debt_ratio": div(h.get("total_liabilities"), h.get("total_assets")),
        "notes": notes, "detail": detail, "report_url": h.get("m_url"),
    }


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    nested = "--nested" in sys.argv  # also evaluate investees that are holdings themselves (slow: hours)
    work = os.path.abspath(args[0]) if args else os.getcwd()
    fetch.FRESH_SEARCH_SINCE = time.time()  # stale investees: an older letter list may miss the new letter
    res_dir = os.path.join(work, "results")
    hs = []
    for f in sorted(os.listdir(res_dir)):
        r = json.load(open(os.path.join(res_dir, f), encoding="utf-8"))
        if is_holding(r):
            hs.append(r)
    print(f"{len(hs)} holdings to evaluate", flush=True)
    mw = market_watch()
    out = []
    done = set()
    i = 0
    while hs:  # original holdings first, then investees that are holdings themselves (one level)
        h = hs.pop(0)
        if norm(h["symbol"]) in done:
            continue
        done.add(norm(h["symbol"]))
        i += 1
        print(f"[{i}] {h['symbol']}: evaluating", flush=True)
        before = fetch.TRANSIENT
        try:
            o = evaluate_holding(h, work, mw)
        except Exception as e:  # noqa: BLE001
            o = {"symbol": h["symbol"], "error": f"خطای اجرا: {e!r}", "trace": traceback.format_exc()[-600:]}
        if fetch.TRANSIENT > before:
            e = fetch.LAST_TRANSIENT
            if isinstance(e, urllib.error.HTTPError) and e.code >= 500:
                # one broken letter export, not an outage: keep going with what could be read
                o.setdefault("notes", []).append(f"خطای سرور کدال هنگام خواندن بخشی از داده‌ها: {e}")
                fetch.TRANSIENT = 0
            else:  # investees collected so far are kept and reused next run
                print(f"STOPPED at {h['symbol']}: Codal is throttling or unreachable; run again later", flush=True)
                sys.exit(75)
        o["nested_in"] = h.get("_nested_in")
        out.append(o)
        print(f"[{i}] {h['symbol']}: {o.get('error') or round(o['p_nav'] or 0, 3)}", flush=True)
        if nested and not h.get("_nested_in"):
            for d in o.get("detail", []):
                r = _inv_cache.get(d.get("symbol"))
                if r and is_holding(r) and norm(r["symbol"]) not in done:
                    hs.append(dict(r, _nested_in=h["symbol"]))
    json.dump(out, open(os.path.join(work, "holdings.json"), "w", encoding="utf-8"), ensure_ascii=False, default=str)
    print("DONE", flush=True)


if __name__ == "__main__":
    main()
