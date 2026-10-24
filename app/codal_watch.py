"""Find new Codal letters (ن-۱۰ statements, ن-۳۰ monthly reports) for the watched symbols.

One market-wide search per letter type, filtered by publish date, instead of one search per
symbol: a few dozen requests instead of hundreds, which keeps us under Codal's rate limit.
"""
import datetime
import json
import logging
import os
import re
import sys
import time

from . import config

sys.path.insert(0, str(config.BATCH_DIR))
from fetch import FA_DIG, Throttled, get, market_watch, norm  # noqa: E402
from holdings import g2j  # noqa: E402

log = logging.getLogger(__name__)

LETTER_TYPES = {6: "صورت مالی", 58: "گزارش ماهانه"}
MAX_PAGES = 80
SEARCH_CACHE = config.DATA_DIR / "cache" / "codal_watch.json"
SEARCH = ("https://search.codal.ir/api/search/v2/q?Symbol=&LetterType={lt}&Audited=true&NotAudited=true"
          "&Consolidatable=true&NotConsolidatable=true&Childs=false&Mains=true&Publisher=false&Category=-1"
          "&CompanyType=-1&Length=-1&search=true&Isic=&AuditorRef=-1&CompanyState=-1&FromDate={fd}&ToDate="
          "&ReportingType=-1&TracingNo=-1&isNotAudited=false&PageNumber={p}")


def fa(sym):
    """TSETMC/Codal spelling -> Persian letters, like the user's symbols.json."""
    return " ".join(str(sym).replace("ي", "ی").replace("ك", "ک").split())


def listed_stocks():
    """{normalised symbol: symbol} of shares listed on TSETMC (no funds, rights, bonds or options)."""
    out = {}
    for key, x in market_watch().items():
        if re.search(r"\d", x["symbol"]) or key.endswith("ح") or norm(x["name"]).startswith(norm("صندوق")):
            continue
        out[key] = fa(x["symbol"])
    return out


def watched_symbols():
    """{normalised symbol: {"results": name or None, "investee": name or None}}."""
    w = {}
    symbols = json.load(open(config.WORK_DIR / "symbols.json", encoding="utf-8"))
    for s in symbols:
        w.setdefault(norm(s), {"results": None, "investee": None})["results"] = s
    inv = config.WORK_DIR / "investees"
    if inv.is_dir():
        for f in inv.glob("*.json"):
            w.setdefault(norm(f.stem), {"results": None, "investee": None})["investee"] = f.stem
    return w


def relevant(letter, lt):
    title = letter.get("Title", "")
    if not (letter.get("HasHtml") or letter.get("HasExcel")):
        return False
    if lt == 6:  # the skill only reads non-consolidated, non-corrected statements
        return "تلفیقی" not in title and "اصلاحیه" not in title
    return True


def search(from_jalali):
    """All letters of the watched types published since from_jalali (inclusive).

    The whole result is cached for WATCH_SEARCH_TTL, so a run that resumes after a stop does not
    page through Codal again. (Caching pages one by one could miss letters: new ones shift the pages.)
    """
    try:
        c = json.loads(SEARCH_CACHE.read_text(encoding="utf-8"))
        age = time.time() - c["at"]
        if c["from"] <= from_jalali and age < config.WATCH_SEARCH_TTL:
            log.info("codal: reusing the search from %d min ago (%d letters since %s)",
                     age // 60, len(c["letters"]), c["from"])
            return c["letters"]
    except (FileNotFoundError, ValueError, KeyError):
        pass
    out = []
    for lt in LETTER_TYPES:
        for p in range(1, MAX_PAGES + 1):
            d = json.loads(get(SEARCH.format(lt=lt, fd=from_jalali, p=p), delay=12, ttl=0))
            for x in d.get("Letters", []):
                x["_type"] = lt
                out.append(x)
            log.info("codal search %s: page %d/%s, %d letters so far",
                     LETTER_TYPES[lt], p, d.get("Page") or 1, len(out))
            if p >= (d.get("Page") or 1):
                break
        else:
            log.warning("letter type %s: stopped after %s pages", lt, MAX_PAGES)
    SEARCH_CACHE.parent.mkdir(parents=True, exist_ok=True)
    tmp = SEARCH_CACHE.with_suffix(".tmp")
    tmp.write_text(json.dumps({"from": from_jalali, "at": time.time(), "letters": out}, ensure_ascii=False),
                   encoding="utf-8")
    os.replace(tmp, SEARCH_CACHE)
    return out


def from_date(state):
    last = state.get("last_search") or state.get("last_success")
    if last:
        start = datetime.date.fromisoformat(last) - datetime.timedelta(days=1)
    else:
        start = datetime.date.today() - datetime.timedelta(days=config.LOOKBACK_DAYS)
    return g2j(start)


def find_new(state):
    """New letters for watched symbols, one entry per report (duplicates folded into tracing_nos)."""
    fd = from_date(state)
    seen = set(state.get("seen", []))
    watched = watched_symbols()
    listed = listed_stocks()
    log.info("codal: watching %d symbols, %d listed stocks; searching since %s", len(watched), len(listed), fd)
    letters = search(fd)
    state["last_search"] = datetime.date.today().isoformat()
    log.info("codal: %d letters since %s", len(letters), fd)
    new, by_title = [], {}
    for x in letters:
        key = norm(x.get("Symbol"))
        if (key not in watched and key not in listed) or x.get("TracingNo") in seen \
                or not relevant(x, x["_type"]):
            continue
        # every listed company with a new report ends up in symbols.json (always re-collected);
        # "add" is set when it is not there yet (new company, or only known as a holding's investee)
        targets = dict(watched.get(key) or {"results": None, "investee": None})
        if not targets["results"]:
            targets["add"] = fa(x.get("Symbol"))
        first = by_title.get((key, norm(x.get("Title"))))
        if first:  # Codal sometimes lists the same report twice
            first["tracing_nos"].append(x.get("TracingNo"))
            continue
        new.append({
            "key": key,
            "symbol": x.get("Symbol"),
            "type": LETTER_TYPES[x["_type"]],
            "title": " ".join(x.get("Title", "").split()),
            "published": str(x.get("PublishDateTime", "")).translate(FA_DIG),
            "tracing_nos": [x.get("TracingNo")],
            "url": "https://codal.ir" + x.get("Url", ""),
            "targets": targets,
        })
        by_title[(key, norm(x.get("Title")))] = new[-1]
    return new
