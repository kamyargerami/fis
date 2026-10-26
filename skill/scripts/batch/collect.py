"""Collect FIS inputs for one symbol from Codal + TSETMC.

Statement source per Codal letter: the codal.ir datasource pages if they are already in the
cache, otherwise the Excel export (excel.codal.ir) — one request for all sheets, and not
affected by codal.ir's «تأیید کاربر» bot-check. Both paths produce the same row format.
"""
import collections
import html as htmllib
import re
import urllib.parse

from fetch import (FA_DIG, codal_letters, excel_tables, fa_num, get, norm, page_datasource,
                   tsetmc, _path)
import os

DAYS = {3: 93, 6: 186, 9: 276, 12: 365}
DATE_RE = re.compile(r"(\d{4}/\d{2}/\d{2})")


def num(v):
    if v in (None, "", "--", "-"):
        return None
    try:
        return float(str(v).replace(",", "").translate(FA_DIG))
    except ValueError:
        return None


def cached(url):
    return os.path.exists(_path(url))


# ---------------------------------------------------------------- row helpers
def period_table(ds):
    """codal.ir datasource -> list of (norm_label, {periodEnd: value}) for the first table."""
    if not ds:
        return []
    t = ds["sheets"][0]["tables"][0]
    rows = collections.defaultdict(dict)
    labels = {}
    for c in t["cells"]:
        if c["columnSequence"] == 1:
            labels[c["rowSequence"]] = norm(c["value"])
        elif c["cellGroupName"] == "Body" and c.get("periodEndToDate"):
            v = num(c["value"])
            if v is not None:  # keep first value per period (% change columns have no date)
                rows[c["rowSequence"]].setdefault(c["periodEndToDate"], v)
    return [(labels.get(r, ""), rows[r]) for r in sorted(labels) if r in rows]


def expand(row):
    """Excel row [(text, colspan)] -> list of texts with colspans expanded (value at first slot)."""
    out = []
    for txt, span in row:
        out.append(txt)
        out.extend([""] * (span - 1))
    return out


def excel_period_rows(table):
    """Excel table -> list of (norm_label, {date: value}) using dates found in the header row."""
    if not table:
        return []
    hdr = expand(table[0])
    colmap = {}
    for i, h in enumerate(hdr):
        m = DATE_RE.search(h.translate(FA_DIG))
        if i and m and "درصد" not in h:
            colmap.setdefault(i, m.group(1))
    out = []
    for r in table[1:]:
        cells = expand(r)
        if not cells:
            continue
        vals = {}
        for i, d in colmap.items():
            if i < len(cells):
                v = fa_num(cells[i])
                if v is not None:
                    vals.setdefault(d, v)
        if vals:
            out.append((norm(cells[0]), vals))
    return out


def find(rows, *names, last=False):
    for n in [norm(x) for x in names]:  # name priority order
        hits = [v for lab, v in rows if lab == n and v]
        if hits:
            return hits[-1] if last else hits[0]
    return {}


def labels_of(rows):
    return {lab for lab, _ in rows}


# ---------------------------------------------------------------- letters
def parse_title(title):
    t = title.translate(FA_DIG)
    m = re.search(r"(\d+)\s*ماهه\s*منتهی\s*به\s*(\d{4}/\d{2}/\d{2})", t)
    if m:
        return int(m.group(1)), m.group(2)
    m = re.search(r"سال\s*مالی\s*منتهی\s*به\s*(\d{4}/\d{2}/\d{2})", t)
    if m:
        return 12, m.group(1)
    return None, None


def serial(letter):
    q = urllib.parse.urlparse(letter["Url"]).query
    return urllib.parse.quote(urllib.parse.parse_qs(q)["LetterSerial"][0], safe="")


def usable(l):
    t = l.get("Title", "")
    return (l.get("HasHtml") or l.get("HasExcel")) and "تلفیقی" not in t and "اصلاحیه" not in t


def add_months(date, n):
    y, m = int(date[:4]), int(date[5:7])
    t = y * 12 + (m - 1) + n
    yy, mm = t // 12, t % 12 + 1
    dd = "29" if mm == 12 else ("30" if mm > 6 else "31")
    return f"{yy:04d}/{mm:02d}/{dd}"


IS_SHEET, BS_SHEET, EQ_SHEET = 1, 0, 1060


def _dense(grid):
    if not grid:
        return []
    width = max(max(r) for r in grid.values())
    return [[grid[r].get(c) for c in range(1, width + 1)] for r in sorted(grid)]


def _sheet_url(ser, sid):
    return f"https://codal.ir/Reports/InterimStatement.aspx?LetterSerial={ser}&SheetId={sid}"


def load_statement(letter):
    """Parsed ن-۱۰ letter: {'p','end','yend','format','is','bs','eq'} (is/bs = period rows, eq = grid)."""
    ser = serial(letter)
    p, end = parse_title(letter["Title"])
    if all(cached(_sheet_url(ser, s)) for s in (IS_SHEET, BS_SHEET, EQ_SHEET)):
        st = _html_statement(ser, p)
        if st:
            return st
    try:
        st = _excel_statement(ser, p, end) if p else None
    except Exception:  # noqa: BLE001  (network error on excel.codal.ir)
        st = None
    if st:
        return st
    try:  # Excel export can be empty for some letters: fall back to live codal.ir pages
        return _html_statement(ser, p, attempts=2)
    except Exception:  # noqa: BLE001
        return None


def _html_statement(ser, p, attempts=4):
    ds = page_datasource(_sheet_url(ser, IS_SHEET), attempts=attempts)
    if not ds:
        return None
    eq = page_datasource(_sheet_url(ser, EQ_SHEET), attempts=attempts)
    grid = collections.defaultdict(dict)
    if eq:
        for c in eq["sheets"][0]["tables"][0]["cells"]:
            grid[c["rowSequence"]][c["columnSequence"]] = c["value"]
    return {"p": ds.get("period") or p, "end": ds["periodEndToDate"], "yend": ds["yearEndToDate"],
            "format": ds.get("title_Fa"), "is": period_table(ds),
            "bs": period_table(page_datasource(_sheet_url(ser, BS_SHEET), attempts=attempts)),
            "eq": _dense(grid), "source": "html"}


def _excel_statement(ser, p, end):
    tables = excel_tables(ser)
    is_rows = bs_rows = None
    eq = None
    for t in tables:
        if not t or len(t) < 3:
            continue
        rows = excel_period_rows(t)
        labs = labels_of(rows)
        if is_rows is None and norm("سود (زیان) خالص") in labs and (
                norm("سود (زیان) عملیاتی") in labs or norm("درآمدهای عملیاتی") in labs
                or norm("درآمد عملیاتی") in labs or norm("جمع درآمدهای عملیاتی") in labs):
            is_rows = rows
        elif bs_rows is None and norm("جمع دارایی‌ها") in labs:
            bs_rows = rows
        else:
            hdr = [norm(x) for x in expand(t[0])]
            if eq is None and (norm("سود (زيان) انباشته") in hdr or norm("سود انباشته") in hdr) \
                    and any(norm(expand(r)[0]).startswith(norm("مانده")) for r in t[1:] if r):
                eq = [expand(r) for r in t]
    if is_rows is None:
        return None
    labs = labels_of(is_rows)
    if norm("درآمد سود سهام") in labs or norm("جمع درآمدهای عملیاتی") in labs:
        fmt = "اطلاعات و صورت های مالی اصلی - سرمایه گذاری (excel)"
    elif norm("درآمد تسهیلات اعطایی") in labs:
        fmt = "اطلاعات و صورت های مالی - بانک (excel)"
    else:
        fmt = "اطلاعات و صورت‌های مالی (excel)"
    return {"p": p, "end": end, "yend": add_months(end, 12 - p), "format": fmt, "is": is_rows,
            "bs": bs_rows or [], "eq": eq or [], "source": "excel"}


def income(rows):
    return {
        "rev": find(rows, "درآمدهای عملیاتی", "جمع درآمدهای عملیاتی", "درآمد عملیاتی", "جمع درآمدها"),
        "cogs": find(rows, "بهای تمام شده درآمدهای عملیاتی"),
        "op": find(rows, "سود (زیان) عملیاتی"),
        "impair": find(rows, "هزینه کاهش ارزش دریافتنی‌ها (هزینه استثنایی)", "زیان کاهش ارزش دریافتنی‌ها"),
        "fin": find(rows, "هزینه‌های مالی"),
        "tax_cur": find(rows, "سال جاری", "هزینه مالیات بر درآمد"),
        "tax_prev": find(rows, "سال‌های قبل"),
        "net": find(rows, "سود (زیان) خالص", last=True),
        "bio": find(rows, "سود (زیان) فروش دارایی زیستی مولد"),
        "other_inc": find(rows, "سایر درآمدها", "سایر درآمدهای عملیاتی"),
    }


def fis_profit(inc, p):
    op = inc["op"].get(p)
    if op is None:
        return None
    imp = abs(inc["impair"].get(p) or 0)
    fin = abs(inc["fin"].get(p) or 0)
    tax = abs(inc["tax_cur"].get(p) or 0)
    return op + imp - fin - tax


NONOP_RULE = 2  # version of nonop_income()
METHOD = 4  # version of collect()'s output; recompute.py re-derives older results from the HTTP cache
CAPITAL_INCREASE_CODES = ("ن-۵۷", "ن-۷۳")  # extraordinary general meeting / board decisions on a capital increase


def _increase_rows(page):
    """Numeric rows of the «محل تأمین افزایش سرمایه» table of a ن-۵۷ / ن-۷۳ notice. Columns: shares,
    par, last registered capital, receivables & cash contribution, retained earnings, reserves,
    revaluation surplus, share premium, cash without pre-emption rights, total increase, percent, …
    (amounts in million rials; values sit in input fields or plain cells)."""
    i = page.find("محل تأمین")
    if i < 0:
        return []
    table = page[page.rfind("<table", 0, i):page.find("</table>", i)]
    rows = []
    for tr in re.findall(r"<tr.*?</tr>", table, re.S):
        cells = []
        for td in re.findall(r"<td[^>]*>(.*?)</td>", tr, re.S):
            v = re.findall(r'value="([^"]*)"', td)
            cells.append(fa_num(htmllib.unescape(v[0] if v else re.sub(r"<[^>]+>", "", td))))
        if len(cells) >= 11 and sum(c is not None for c in cells) >= 8:
            rows.append(cells)
    return rows


def capital_increase(symbol, capital_report, shares_now, since, as_of=None):
    """Sources of the capital increase registered after the latest report: the ن-۵۷ / ن-۷۳ notice whose
    last capital is the report's capital and whose new capital is today's (million rials). None if not found."""
    target = shares_now * 1000 / 1e6  # par 1000 rials
    letters = [l for l in codal_letters(symbol, pages=4, letter_type=-1, from_date=since)
               if l.get("LetterCode") in CAPITAL_INCREASE_CODES and published_by(l, as_of)]
    for l in letters:  # newest first
        for c in _increase_rows(get("https://codal.ir" + l["Url"], delay=4)):
            last, total = c[2], c[9]
            if last and total and abs(last / capital_report - 1) < 0.02 and abs((last + total) / target - 1) < 0.02:
                z = lambda x: x or 0  # noqa: E731
                return {"cash": z(c[3]) + z(c[8]), "retained": z(c[4]), "reserve": z(c[5]),
                        "revaluation": z(c[6]), "premium": z(c[7]), "total": total,
                        "letter": " ".join(l["Title"].split()), "published": str(l.get("PublishDateTime", ""))[:10].translate(FA_DIG),
                        "url": "https://codal.ir" + l["Url"]}
    return None


def gains_of(inc, p):
    """Gains that are not assumed to repeat: other operating income + non-operating income (each only
    if positive; expenses are always kept). Livestock sales are excluded by nonop_income()."""
    other = (inc["other_inc"] or {}).get(p) if inc.get("other_inc") is not None else None
    no = nonop_income(inc, p)
    return max(other or 0, 0) + max(no or 0, 0), other, no


def recurring_gains(cur, months, fy, fy2):
    """Annual amount of gains judged recurring: per item (other income, non-operating), positive in the
    current period and in both previous fiscal years -> the smallest of the three (current annualized)."""
    total = 0.0
    for k in (1, 2):  # (gains, other, nonop) -> other, nonop
        vals = [cur[k] * 12 / months if cur[k] is not None and months else None,
                fy[k] if fy else None, fy2[k] if fy2 else None]
        if all(v is not None and v > 0 for v in vals):
            total += min(vals)
    return total


def nonop_income(inc, p):
    """Pre-tax non-operating income of period p: pre-tax profit - operating profit + finance costs,
    without gains on selling productive biological assets (livestock): for farms these recur every
    year, and the analyst channel the user follows keeps them too.

    Pre-tax profit is rebuilt as net + taxes, so it works whatever the pre-tax row is called.
    None when operating profit or net income is missing (banks, insurers)."""
    op, net = inc["op"].get(p), inc["net"].get(p)
    if op is None or net is None:
        return None
    tax = abs(inc["tax_cur"].get(p) or 0) + abs(inc["tax_prev"].get(p) or 0)
    bio = (inc.get("bio") or {}).get(p) or 0
    return net + tax - op + abs(inc["fin"].get(p) or 0) - bio


def equity_dividends(grid):
    """grid: list of rows (lists, col 0 = label). Returns [(dividend, retained_earnings_at_start)]
    for prior FY and current period (Codal order)."""
    if not grid:
        return []
    header = [norm(x) for x in grid[0]]
    re_col = next((i for i, v in enumerate(header)
                   if v in (norm("سود (زيان) انباشته"), norm("سود انباشته"))), None)
    if re_col is None:
        return []
    tot_col = len(header) - 1
    opens, divs = [], []
    for r in grid[1:]:
        if not r:
            continue
        lab = norm(r[0])
        cell = lambda i: fa_num(r[i]) if i < len(r) else None  # noqa: E731
        if lab.startswith(norm("مانده تجدید ارائه شده در")):
            opens.append(cell(re_col))
        elif lab in (norm("سود سهام مصوب"), norm("سود تقسیمی")):
            v = cell(re_col)
            if v is None:
                v = cell(tot_col)
            divs.append(abs(v or 0))
    return list(zip(divs, opens))


# ---------------------------------------------------------------- monthly (ن-۳۰)
REVENUE_ALIASES = {"ProductionAndSales", "RevenuesServiceDeliveryRealized", "ProfitLossInvestments",
                   "DescriptionOfOperatingIncome", "ServicesAndSale"}
NON_REVENUE = ("انرژی", "مواد", "ارزی", "هزینه", "سپرده", "تسهیلات", "تعهدات", "سرمایه گذاری های شرکت",
               "سرمایه‌گذاری‌های شرکت", "خرید")


def months_elapsed(period_end, year_end):
    py, pm = int(period_end[:4]), int(period_end[5:7])
    yy, ym = int(year_end[:4]), int(year_end[5:7])
    return (py * 12 + pm) - (yy * 12 + ym) + 12


def _groups(header_cells, pe, prior_date):
    """header_cells: [(text, first_col, span)] -> (cur column ranges, prior-year column ranges)."""
    cur, prev = [], []
    for txt, c0, span in header_cells:
        v = norm(txt).translate(FA_DIG)
        start = "ابتدای" in v or "اولسال" in v
        rng = range(c0, c0 + span)
        if norm(pe) in v and start and "اصلاح" not in v and "ماهه" not in v:
            cur.append(rng)
        elif start and (norm(prior_date) in v or "سالمالیقبل" in v):
            prev.append(rng)
    return cur, prev


def _pick(total_row, groups):
    for g in groups:
        vals = [total_row[c] for c in g if c in total_row]
        if vals:
            return vals[-1]
    return None


def _monthly_from_ds(ds, pe):
    pe = ds.get("periodEndToDate") or pe
    prior_date = f"{int(pe[:4]) - 1}{pe[4:7]}"
    cur = prev = None
    used = []
    for sh in ds["sheets"]:
        for t in sh["tables"]:
            title, alias = t.get("title_Fa") or "", t.get("aliasName") or ""
            if alias not in REVENUE_ALIASES and not (
                    ("فروش" in title or "درآمد" in title) and not any(w in title for w in NON_REVENUE)):
                continue
            hc = [(c["value"], c["columnSequence"], c.get("colSpan") or 1) for c in t["cells"]
                  if c["cellGroupName"] == "Header" and c["value"]]
            gc, gp = _groups(hc, pe, prior_date)
            if not gc:
                continue
            rows, labels = collections.defaultdict(dict), {}
            for c in t["cells"]:
                if c["cellGroupName"] == "Header":
                    continue
                if c["columnSequence"] in (1, 2) and isinstance(c["value"], str) \
                        and norm(c["value"]).startswith("جمع"):
                    labels.setdefault(c["rowSequence"], norm(c["value"]))
                n = num(c["value"])
                if n is not None:
                    rows[c["rowSequence"]][c["columnSequence"]] = n
            tots = [r for r, lab in labels.items() if lab in ("جمع", "جمعکل")] or sorted(labels)[-1:]
            if not tots:
                continue
            cv, pv = _pick(rows[max(tots)], gc), _pick(rows[max(tots)], gp)
            if cv is None:
                continue
            cur = (cur or 0) + cv
            used.append(title)
            if pv is not None:
                prev = (prev or 0) + pv
    return cur, prev, used, pe, ds.get("yearEndToDate"), ds.get("title_Fa")


def _monthly_from_excel(ser, pe):
    prior_date = f"{int(pe[:4]) - 1}{pe[4:7]}"
    cur = prev = None
    used = []
    for t in excel_tables(ser):
        if len(t) < 3 or not t[0]:
            continue
        first = norm(t[0][0][0])
        hdr_text = " ".join(x for r in t[:3] for x, _ in r)
        if not first.startswith("شرح") or not ("فروش" in hdr_text or "درآمد" in hdr_text) \
                or any(w in hdr_text for w in ("انرژی", "مواد اولیه", "منابع ارزی", "مصارف ارزی")):
            continue
        hc, col = [], 0
        for txt, span in t[0]:
            hc.append((txt, col, span))
            col += span
        gc, gp = _groups(hc, pe, prior_date)
        if not gc:
            continue
        total = None
        for r in t[1:]:
            cells = expand(r)
            if cells and norm(cells[0]) in ("جمع", "جمعکل"):
                total = {i: fa_num(x) for i, x in enumerate(cells) if fa_num(x) is not None}
        if total is None:
            continue
        cv, pv = _pick(total, gc), _pick(total, gp)
        if cv is None:
            continue
        cur = (cur or 0) + cv
        used.append(t[0][0][0])
        if pv is not None:
            prev = (prev or 0) + pv
    return cur, prev, used, pe, None, "گزارش فعالیت ماهانه (excel)"


def published_by(letter, as_of):
    """True if the letter was published on or before as_of (Jalali 'YYYY/MM/DD'); always True without as_of."""
    return not as_of or str(letter.get("PublishDateTime", "")).translate(FA_DIG)[:10] <= as_of


def monthly_revenue(symbol, as_of=None):
    """Cumulative revenue to date (and same period last year) from the latest ن-۳۰ monthly report.

    Columns are found by header text, not position: the cumulative-to-date group contains the
    report's period-end date and «ابتدای»/«اول سال» (not «اصلاح»/«ماهه»); the prior-year group
    contains the same month a year earlier or «سال مالی قبل». The amount is the last numeric cell
    of that group in the table's «جمع»/«جمع کل» row. Source order: cached codal.ir page ->
    Excel export -> live codal.ir page.
    """
    letters = [l for l in codal_letters(symbol, letter_type=58)
               if (l.get("HasHtml") or l.get("HasExcel")) and published_by(l, as_of)]
    if not letters:
        return None
    latest = max(letters, key=lambda l: parse_title(l["Title"])[1] or "")
    url = "https://codal.ir" + latest["Url"]
    pe = parse_title(latest["Title"])[1]
    res = None
    if cached(url):
        ds = page_datasource(url)
        res = _monthly_from_ds(ds, pe) if ds else None
    if (not res or res[0] is None) and pe:
        try:
            res = _monthly_from_excel(serial(latest), pe)
        except Exception:  # noqa: BLE001
            res = None
    if not res or res[0] is None:
        try:
            ds = page_datasource(url, attempts=2)
            res = _monthly_from_ds(ds, pe) if ds else None
        except Exception:  # noqa: BLE001
            res = None
    if not res or res[0] is None:
        return None
    cur, prev, used, pe, ye, fmt = res
    return {"m_period_end": pe, "m_months": months_elapsed(pe, ye) if ye else None, "m_cum_rev": cur,
            "m_prev_rev": prev, "m_tables": used, "m_title": latest["Title"], "m_url": url, "m_format": fmt,
            "m_serial": serial(latest), "m_excel": latest.get("ExcelUrl")}


# ---------------------------------------------------------------- main
def collect(symbol, as_of=None):
    """FIS inputs for a symbol. as_of (Jalali 'YYYY/MM/DD') uses only letters published by that day,
    to rebuild what was known on a past date (e.g. to compare with an analyst's post from that day).
    Price and share count are always today's."""
    out = {"symbol": symbol, "notes": []}
    if as_of:
        out["as_of"] = as_of
    m = tsetmc(symbol)
    if not m:
        out["error"] = "نماد در TSETMC پیدا نشد"
        return out
    out.update({"name": m["name"], "price": m["price"], "shares": m["shares"], "sector": m["sector"]})
    letters, variant = [], symbol
    for variant in dict.fromkeys([symbol, symbol.replace(" ", "‌")]):
        letters = [l for l in codal_letters(variant) if usable(l) and published_by(l, as_of)]
        if not letters:
            letters = [l for l in codal_letters(variant, pages=2) if usable(l) and published_by(l, as_of)]
        if letters:
            break
    if not letters:
        out["error"] = "صورت مالی ن-۱۰ در کدال ندارد (احتمالاً صندوق)"
        return out
    by_period = {}
    for l in letters:  # newest first
        per, e = parse_title(l["Title"])
        if per and (per, e) not in by_period:
            by_period[(per, e)] = l

    def end_key(l):
        per, e = parse_title(l["Title"])
        return (e or "", per or 0)
    latest = max(letters, key=end_key)  # max() keeps first (newest-published) on ties
    st = load_statement(latest)
    if not st:
        out["error"] = "صورت سود و زیان قابل خواندن نیست"
        return out
    out["report"] = latest["Title"]
    out["report_url"] = "https://codal.ir" + latest["Url"]
    out["format"] = st["format"]
    out["source"] = st["source"]
    if "صندوق" in (out["format"] or "") or norm(m["name"] or "").startswith(norm("صندوق")):
        out["error"] = "صندوق سرمایه‌گذاری است، نه شرکت"
        return out
    p, end, yend = st["p"], st["end"], st["yend"]
    out.update(months=p, period_end=end, year_end=yend)
    inc = income(st["is"])
    periods = sorted({k for d in inc.values() for k in d})
    prev_same = next((d for d in periods if d != end and d[5:7] == end[5:7] and int(d[:4]) == int(end[:4]) - 1), None)
    prev_fy = next((d for d in periods if d != end and d[5:7] == yend[5:7] and int(d[:4]) == int(yend[:4]) - 1), None)
    if p == 12:
        prev_same = prev_fy = next((d for d in periods if d != end), None)
    if p != 12 and prev_fy is None:  # fiscal-year change: the remaining column is the last full year
        prev_fy = next((d for d in periods if d not in (end, prev_same)), None)
    # the last annual report: the two previous fiscal years (recurring gains, dividends of two AGMs);
    # some formats (banks) also leave the prior-FY column of the interim report empty
    ast = ainc = None
    if p == 12:
        ast, ainc = st, inc
    elif prev_fy:
        for (pp, ee), l in by_period.items():
            if pp == 12 and ee[:7] == prev_fy[:7]:
                ast = load_statement(l)
                ainc = income(ast["is"]) if ast else None
                break
    if ainc and p != 12 and not inc["net"].get(prev_fy):
        for k in inc:
            if ast["end"] in ainc[k]:
                inc[k][prev_fy] = ainc[k][ast["end"]]
    out["prev_fy_end"] = prev_fy

    rev, net = inc["rev"].get(end), inc["net"].get(end)
    fis = fis_profit(inc, end)
    out.update(rev=rev, net=net, fis=fis, cogs=abs(inc["cogs"].get(end) or 0) or None,
               nonop=nonop_income(inc, end), nonop_rule=NONOP_RULE, method=METHOD)
    cur_g = gains_of(inc, end)
    out["gains"], out["other_inc"] = cur_g[0], cur_g[1]
    fy_g = fy2_g = None
    if ainc and ast:
        fy_end = ast["end"]
        fy2_end = next((d for d in sorted({k for d in ainc.values() for k in d}, reverse=True) if d < fy_end), None)
        fy_g = gains_of(ainc, fy_end) if p != 12 else None
        fy2_g = gains_of(ainc, fy2_end) if fy2_end else None
        if p == 12:  # the latest report is itself annual: compare with the two years before
            fy_g, fy2_g = fy2_g, None
            for (pp, ee), l in by_period.items():
                if pp == 12 and fy2_end and ee[:7] == fy2_end[:7] and ee != end:
                    a2 = load_statement(l)
                    if a2:
                        i2 = income(a2["is"])
                        d2 = next((d for d in sorted({k for d in i2.values() for k in d}, reverse=True) if d < a2["end"]), None)
                        fy2_g = gains_of(i2, d2) if d2 else None
                    break
    out["recurring_gains"] = recurring_gains(cur_g, p, fy_g, fy2_g) if fy_g else 0.0
    if rev is None or inc["op"].get(end) is None:
        out["notes"].append("ساختار صورت سود و زیان غیراستاندارد (بانک/بیمه/…) — درآمد یا سود عملیاتی پیدا نشد")
    out["rev_prev_same"] = inc["rev"].get(prev_same) if prev_same else None
    if prev_fy:
        out["fy_rev"], out["fy_net"] = inc["rev"].get(prev_fy), inc["net"].get(prev_fy)

    def ttm(k):
        """Last 12 months: the current period replaces the same period of last year."""
        if p == 12:
            return inc[k].get(end)
        vals = (inc[k].get(prev_fy), inc[k].get(prev_same), inc[k].get(end)) if prev_fy and prev_same else (None,)
        return vals[0] - vals[1] + vals[2] if None not in vals else None
    out["ttm_net"] = ttm("net")
    out["ttm_rev"] = ttm("rev")
    out["ttm_cogs"] = abs(ttm("cogs")) if ttm("cogs") is not None else None

    def cum(per_months, year_end):
        """(rev, fis) cumulative for a period of the given fiscal year, from its own report."""
        tries = 0
        for (pp, ee), l in by_period.items():
            if pp != per_months or tries >= 4:
                continue
            tries += 1
            s2 = load_statement(l)
            if s2 and s2["yend"][:7] == year_end[:7]:
                i = income(s2["is"])
                return i["rev"].get(s2["end"]), fis_profit(i, s2["end"])
        return None, None

    try:
        if p == 3:
            r9, f9 = cum(9, prev_fy) if prev_fy else (None, None)
            fy_r = inc["rev"].get(prev_fy) if prev_fy else None
            fy_f = fis_profit(inc, prev_fy) if prev_fy else None
            q_cur = (rev, fis)
            q_prev = (fy_r - r9, fy_f - f9) if None not in (fy_r, fy_f, r9, f9) else (None, None)
        else:
            r1, f1 = cum(p - 3, yend)
            q_cur = (rev - r1, fis - f1) if None not in (rev, fis, r1, f1) else (None, None)
            if p - 6 > 0:
                r2, f2 = cum(p - 6, yend)
                q_prev = (r1 - r2, f1 - f2) if None not in (r1, f1, r2, f2) else (None, None)
            else:
                q_prev = (r1, f1)
        out["q_cur"], out["q_prev"] = q_cur, q_prev
    except Exception as e:  # noqa: BLE001
        out["notes"].append(f"حاشیه فصل قبل محاسبه نشد: {e}")

    # balance sheet
    bs = st["bs"]
    ta = find(bs, "جمع دارایی‌ها")
    rec = find(bs, "دریافتنی‌های تجاری و سایر دریافتنی‌ها")
    out["receivables"] = rec.get(end)
    out["total_assets"] = ta.get(end)
    out["total_liabilities"] = find(bs, "جمع بدهی‌ها").get(end)
    out["equity"] = find(bs, "جمع حقوق مالکانه").get(end)
    out["lt_investments"] = find(bs, "سرمایه‌گذاری‌های بلندمدت").get(end)
    out["st_investments"] = find(bs, "سرمایه‌گذاری‌های کوتاه‌مدت").get(end)
    out["capital"] = find(bs, "سرمایه").get(end)
    if out["capital"] and out.get("shares"):
        chg = out["shares"] / (out["capital"] * 1e6 / 1000) - 1  # par value 1000 rials
        if abs(chg) > 0.02:
            out["capital_change_after_report"] = chg
            inc = None
            if chg > 0:
                try:
                    inc = capital_increase(variant, out["capital"], out["shares"], add_months(end, -12), as_of)
                except Exception as e:  # noqa: BLE001
                    out["notes"].append(f"اطلاعیه افزایش سرمایه خوانده نشد: {e}"[:200])
            if inc:
                out["capital_increase"] = inc
                out["cash_increase"] = inc["cash"]
                src = "، ".join(f"{lab} {v / 1e4:,.0f}" for lab, v in (
                    ("آورده نقدی", inc["cash"]), ("سود انباشته", inc["retained"]), ("اندوخته", inc["reserve"]),
                    ("تجدید ارزیابی", inc["revaluation"]), ("صرف سهام", inc["premium"])) if v)
                out["notes"].append(f"افزایش سرمایه {chg * 100:+.0f}٪ بعد از گزارش (میلیارد تومان: {src})"
                                    + ("؛ آورده نقدی از ارزش بازار در P/E کم شد" if inc["cash"] else ""))
            else:
                out["notes"].append(f"افزایش سرمایه بعد از گزارش: تعداد سهام فعلی {chg * 100:+.0f}٪ نسبت به سرمایه گزارش؛ "
                                    "محل آن پیدا نشد (اگر آورده نقدی بوده، سود/دارایی متناظر هنوز در گزارش نیست)")
    inv = find(bs, "موجودی مواد و کالا")
    other = [d for d in ta if d != end]
    prior_bs = max(other) if other else None
    if inv.get(end) is not None:
        vals = [inv.get(end)] + ([inv.get(prior_bs)] if prior_bs and inv.get(prior_bs) is not None else [])
        out["avg_inventory"] = sum(vals) / len(vals)
    if rec.get(end) is not None:
        vals = [rec.get(end)] + ([rec.get(prior_bs)] if prior_bs and rec.get(prior_bs) is not None else [])
        out["avg_receivables"] = sum(vals) / len(vals)

    # dividends (Codal order: prior FY, then current period)
    pairs = equity_dividends(st["eq"])
    ratios = [(d, r) for d, r in pairs if r and r > 0]
    out["div_pairs"] = ratios
    out["agm_held_this_period"] = bool(len(pairs) >= 2 and pairs[1][0] > 0)
    out["last_dividend"] = next((d for d, _ in reversed(pairs) if d > 0), None)
    if len(ratios) == 2 and ratios[1][0] == 0 and p <= 3:
        # this year's AGM is not in the report yet: the AGM before last comes from the annual report
        older = [(d, r) for d, r in equity_dividends(ast["eq"]) if r and r > 0][:1] if ast and ast is not st else []
        ratios = older + ratios[:1]
        if not older:
            out["notes"].append("مجمع سال جاری هنوز در دوره گزارش نبوده؛ نسبت تقسیم سود فقط از سال قبل")
    out["div_ratios"] = ratios
    out["period_days"] = DAYS.get(p, round(p * 30.4))

    # monthly activity report (ن-۳۰): the paper's source for forward sales/earnings
    try:
        mr = monthly_revenue(variant, as_of)
        if mr and not mr.get("m_months"):
            mr["m_months"] = months_elapsed(mr["m_period_end"], yend)
    except Exception as e:  # noqa: BLE001
        mr = None
        out["notes"].append(f"گزارش ماهانه خوانده نشد: {e}")
    if mr and mr["m_cum_rev"] > 0 and mr["m_months"] and mr["m_months"] >= 1:
        out.update(mr)
    else:
        out["notes"].append("گزارش ماهانه قابل استفاده نبود؛ آینده‌نگر از صورت میان‌دوره‌ای")
    if rev and net is not None and rev > 0:
        out["net_margin"] = net / rev
    return out


def core_margin(r):
    """Net margin of realized, recurring profit: without other operating income and non-operating
    gains (asset sales, deposit interest, FX revaluation …; livestock sales stay); all expenses kept."""
    if not r.get("rev") or r["rev"] <= 0 or r.get("net") is None:
        return r.get("net_margin")
    gains = r["gains"] if r.get("gains") is not None else max(r.get("nonop") or 0, 0)
    return (r["net"] - gains) / r["rev"]


def forward_net(r):
    """Forward (current fiscal year) net income, same rule as the FIS sheet.

    Monthly report available: monthly cumulative sales / months x 12 x core_margin() + the gains that
    recurred in each of the last three periods (recurring_gains). Holdings: prior-FY margin, because
    interim dividend income is seasonal. Otherwise interim FIS profit x 12 / months.
    """
    fmt = (r.get("format") or "").replace("‌", " ")
    if r.get("m_cum_rev") and r.get("m_months"):
        fwd_sales = r["m_cum_rev"] * 12 / r["m_months"]
        if "سرمایه گذاری" in fmt and r.get("fy_rev") and r["fy_rev"] > 0 and r.get("fy_net") is not None:
            return fwd_sales * r["fy_net"] / r["fy_rev"], fwd_sales
        margin = core_margin(r)
        if margin is not None:
            return fwd_sales * margin + (r.get("recurring_gains") or 0), fwd_sales
        return None, fwd_sales
    if r.get("fis") is not None and r.get("months"):
        return r["fis"] * 12 / r["months"], (r["rev"] * 12 / r["months"] if r.get("rev") else None)
    return None, None
