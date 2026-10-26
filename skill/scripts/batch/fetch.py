"""Cached HTTP fetchers for Codal and TSETMC (direct connection, no proxy).

Persistent cache: every response is stored under FIS_CACHE (default
~/.cache/fis-stock-scoring/http) and reused across runs and days according to
TTL rules per source (see ttl_for). Codal reports addressed by LetterSerial never
change, so they are cached forever; searches and prices expire.
If the network fails, a stale cached copy is used instead (stale-if-error).
"""
import difflib
import hashlib
import html as htmllib
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

CACHE = os.environ.get("FIS_CACHE", os.path.expanduser("~/.cache/fis-stock-scoring/http"))
os.makedirs(CACHE, exist_ok=True)
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) Chrome/130 Safari/537.36"}
_opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))  # Codal blocks foreign IPs: no proxy
H, D = 3600, 86400
FOREVER = None
# fetches in this process that failed for a temporary reason (throttling, network); batch scripts
# check it after each symbol and stop, so the next run resumes instead of saving a broken result
TRANSIENT = 0
LAST_TRANSIENT = None  # the error behind the last Throttled


class Throttled(RuntimeError):
    """A fetch failed for a temporary reason (Codal throttling or the network)."""


def _transient(e):
    if isinstance(e, urllib.error.HTTPError):
        return e.code == 429 or e.code >= 500
    return True  # bot-check page, DNS, refused/dropped connections, timeouts


def ttl_for(url):
    if "search.codal.ir" in url:
        return 24 * H                      # new letters appear daily
    if "GetMarketWatch" in url or "GetClosingPriceInfo" in url:
        return 4 * H                       # prices
    if "GetInstrumentInfo" in url:
        return 7 * D                       # share count, sector
    if "GetInstrumentSearch" in url:
        return 30 * D                      # symbol -> insCode
    if "LetterSerial=" in url or "excel.codal.ir" in url:
        return FOREVER                     # a published letter never changes
    return D


def _path(url):
    return os.path.join(CACHE, hashlib.sha1(url.encode()).hexdigest())


def get(url, delay=1.5, ttl="auto", attempts=4):
    global TRANSIENT, LAST_TRANSIENT
    path = _path(url)
    ttl = ttl_for(url) if ttl == "auto" else ttl
    offline = os.environ.get("FIS_OFFLINE") == "1"  # same-day re-runs: use any cached copy
    if os.path.exists(path) and (offline or ttl is FOREVER or time.time() - os.path.getmtime(path) < ttl):
        return open(path, encoding="utf-8").read()
    if os.environ.get("FIS_CACHE_ONLY") == "1" and "tsetmc" not in url:  # recompute.py: never ask Codal
        raise RuntimeError(f"not in cache: {url}")
    if TRANSIENT:  # already throttled in this process: fail fast, the caller stops soon
        attempts = 1
    last = None
    for attempt in range(attempts):
        try:
            req = urllib.request.Request(url, headers=UA)
            body = _opener.open(req, timeout=60).read().decode("utf-8", "replace")
            time.sleep(delay)
            if "تأیید کاربر" in body:  # Codal bot-check page: never cache, back off
                raise RuntimeError("verification page (429-like)")
            with open(path, "w", encoding="utf-8") as f:
                f.write(body)
            return body
        except Exception as e:  # noqa: BLE001
            last = e
            if os.path.exists(path):  # stale-if-error
                return open(path, encoding="utf-8").read()
            if attempt + 1 == attempts:
                break
            throttled = "429" in str(e) or "verification" in str(e)
            wait = 60 * (attempt + 1) if throttled else 3 * (attempt + 1)
            print(f"fetch: {'throttled' if throttled else 'error'} ({str(e)[:120]}), "
                  f"retry {attempt + 1}/{attempts - 1} in {wait}s: {url[:100]}", file=sys.stderr, flush=True)
            time.sleep(wait)
    if _transient(last):
        TRANSIENT += 1
        LAST_TRANSIENT = last
        raise Throttled(f"fetch failed {url}: {last}")
    raise RuntimeError(f"fetch failed {url}: {last}")


def norm(s):
    if s is None:
        return ""
    s = str(s).replace("ي", "ی").replace("ك", "ک").replace("ى", "ی")
    return re.sub(r"[\s‌‏‎‍]", "", s)


FA_DIG = str.maketrans("۰۱۲۳۴۵۶۷۸۹٠١٢٣٤٥٦٧٨٩", "01234567890123456789")


# ---------------------------------------------------------------- TSETMC
def market_watch():
    """All instruments in one request: {symbol_norm: {insCode, symbol, name, price, eps}}."""
    url = ("https://cdn.tsetmc.com/api/ClosingPrice/GetMarketWatch?market=0&industrialGroup="
           + "".join(f"&paperTypes%5B{i}%5D={i + 1}" for i in range(9))
           + "&showTraded=false&withBestLimits=false")
    out = {}
    for x in json.loads(get(url))["marketwatch"]:
        out[norm(x["lva"])] = {"insCode": x["insCode"], "symbol": x["lva"], "name": x["lvc"],
                               "price": x.get("pcl") or x.get("py"), "eps": x.get("eps")}
    return out


def instrument_info(ins_code):
    return json.loads(get(f"https://cdn.tsetmc.com/api/Instrument/GetInstrumentInfo/{ins_code}"))["instrumentInfo"]


def tsetmc(symbol):
    q = urllib.parse.quote(symbol)
    d = json.loads(get(f"https://cdn.tsetmc.com/api/Instrument/GetInstrumentSearch/{q}"))
    cands = [i for i in d.get("instrumentSearch", []) if norm(i["lVal18AFC"]) == norm(symbol)]
    if not cands:
        return None
    cands.sort(key=lambda i: -(i.get("lastDate") or 0))
    code = cands[0]["insCode"]
    info = instrument_info(code)
    cp = json.loads(get(f"https://cdn.tsetmc.com/api/ClosingPrice/GetClosingPriceInfo/{code}"))["closingPriceInfo"]
    return {
        "insCode": code,
        "name": info.get("lVal30"),
        "shares": info.get("zTitad"),
        "price": cp.get("pClosing"),
        "last_trade": cp.get("finalLastDate"),
        "yVal": info.get("yVal"),
        "sector": (info.get("sector") or {}).get("lSecVal"),
    }


def name_key(s):
    """Normalise a company name for fuzzy matching Codal full names with TSETMC 30-char names."""
    s = str(s or "").replace("ي", "ی").replace("ك", "ک").replace("ى", "ی").replace("أ", "ا").replace("ة", "ه")
    s = s.replace("‌", " ")
    for a, b in (("سرمایه گذاری", "س"), ("سرمایهگذاری", "س"), ("س.", "س "), ("شرکت", ""), ("(سهامی عام)", ""),
                 ("سهامی عام", ""), ("گروه", "گ"), ("تولیدی", ""), ("صنایع", "ص")):
        s = s.replace(a, b)
    return re.sub(r"[\s\.\-()\"'«»،]", "", s)


def match_instrument(name, total_shares=None, mw=None):
    """Find the listed instrument for a Codal company name.

    Candidates: best fuzzy matches on name_key among market-watch stocks (no rights/options/funds).
    If total_shares (capital/par from the portfolio report) is known, the candidate's share count
    from TSETMC must agree within 3% — this makes the match reliable.
    Returns (instrument dict, score) or (None, 0).
    """
    mw = mw or market_watch()
    key = name_key(name)
    scored = []
    for x in mw.values():
        sym = x["symbol"]
        if re.search(r"\d", sym) or norm(sym).endswith("ح") or norm(x["name"]).startswith(norm("صندوق")):
            continue
        k2 = name_key(x["name"])
        r = difflib.SequenceMatcher(None, key, k2).ratio()
        if k2 and (key.startswith(k2) or k2.startswith(key)):
            r = max(r, 0.9)
        scored.append((r, x))
    scored.sort(key=lambda t: -t[0])
    for r, x in scored[:5]:
        if r < 0.55:
            break
        if total_shares:
            try:
                z = instrument_info(x["insCode"]).get("zTitad") or 0
            except Exception:  # noqa: BLE001
                continue
            if z and abs(z / total_shares - 1) < 0.03:
                return x, r
        elif r >= 0.85:
            return x, r
    return None, 0


# ---------------------------------------------------------------- Codal
def codal_letters(symbol, pages=1, letter_type=6, from_date=""):
    """letter_type 6 = financial statements (ن-۱۰), 58 = monthly activity (ن-۳۰), -1 = all letters.
    from_date: Jalali 'YYYY/MM/DD' (published on or after), '' = no limit."""
    q = urllib.parse.quote(symbol)
    out = []
    for p in range(1, pages + 1):
        url = ("https://search.codal.ir/api/search/v2/q?Symbol=" + q +
               f"&LetterType={letter_type}&Audited=true&NotAudited=true&Consolidatable=true&NotConsolidatable=true"
               "&Childs=false&Mains=true&Publisher=false&Category=-1&CompanyType=-1&Length=-1&search=true"
               f"&Isic=&AuditorRef=-1&CompanyState=-1&FromDate={from_date}&ToDate=&ReportingType=-1&TracingNo=-1"
               f"&isNotAudited=false&PageNumber={p}")
        d = json.loads(get(url, delay=12))
        out += d.get("Letters", [])
        if p >= d.get("Page", 1):
            break
    return [l for l in out if norm(l.get("Symbol")) == norm(symbol)]


def page_datasource(url, attempts=4):
    h = get(url, attempts=attempts)
    m = re.search(r"var datasource = (\{.*?\});\s*\n", h, re.S)
    return json.loads(m.group(1)) if m else None


def sheet(letter_serial, sheet_id):
    return page_datasource(f"https://codal.ir/Reports/InterimStatement.aspx?LetterSerial={letter_serial}"
                           f"&SheetId={sheet_id}")


def serial_of(url):
    """LetterSerial from a Codal URL, in the canonical quoted form used for cache keys."""
    raw = urllib.parse.parse_qs(urllib.parse.urlparse(url).query)["LetterSerial"][0]
    return urllib.parse.quote(raw, safe="")


def excel_tables(letter_serial):
    """All sheets of a Codal letter via the Excel export service (excel.codal.ir).

    Works even when codal.ir shows its bot-check page, and returns every sheet in one request
    (e.g. a holding's listed-portfolio detail, which the HTML page only shows after a postback).
    Returns a list of tables; each table is a list of rows; each row a list of (text, colspan).
    """
    h = get(f"https://excel.codal.ir/service/Excel/GetAll/{letter_serial}/0")
    out = []
    for t in re.findall(r"<table.*?</table>", h, re.S):
        rows = []
        for r in re.findall(r"<tr.*?</tr>", t, re.S):
            cells = []
            for attrs, body in re.findall(r"<t([dh][^>]*)>(.*?)</t[dh]>", r, re.S):
                txt = htmllib.unescape(re.sub(r"<[^>]+>", "", body)).strip()
                m = re.search(r'colspan="?(\d+)', attrs)
                cells.append((txt, int(m.group(1)) if m else 1))
            rows.append(cells)
        out.append(rows)
    return out


def fa_num(s):
    """Parse '۱,۲۳۴' / '(۲۲)' / '70.1' into float; None if not numeric."""
    s = str(s or "").translate(FA_DIG).replace(",", "").replace("٬", "").strip()
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").strip()
    try:
        v = float(s)
    except ValueError:
        return None
    return -v if neg else v
