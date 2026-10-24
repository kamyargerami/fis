"""Collect FIS inputs for many symbols from Codal + TSETMC.

Usage: python3 -I run_all.py WORKDIR
WORKDIR must contain symbols.json (JSON array of symbols). Creates WORKDIR/cache
(raw HTTP responses) and WORKDIR/results/<symbol>.json. Re-running skips symbols
that already have a result and reuses the cache, so it is safe to resume. If Codal throttles
us (or the network fails) it stops with exit code 75 without saving that symbol.
Run it in the background: ~1 min per symbol because of Codal throttling.
"""
import json
import os
import sys
import traceback
import urllib.error

HERE = os.path.dirname(os.path.abspath(__file__))
WORK = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.getcwd()
os.environ["FIS_WORKDIR"] = WORK
sys.path.insert(0, HERE)
import fetch  # noqa: E402
from collect import collect  # noqa: E402

SYMBOLS = json.load(open(os.path.join(WORK, "symbols.json"), encoding="utf-8"))
OUT = os.path.join(WORK, "results")
os.makedirs(OUT, exist_ok=True)
STOPPED = 75  # Codal throttled us or the network is down: run again later to continue
# a symbol that stopped the run this many times with a server error (5xx: often one broken letter
# export, not an outage) is saved with what could be read, so it cannot block the run forever
MAX_SERVER_ERROR_STOPS = 3
STOPS_FILE = os.path.join(WORK, "run_all_stops.json")


def needs_collect(path):
    """No result yet, a file cut off mid-write, or an old result saved after a failed fetch."""
    if not os.path.exists(path):
        return True
    try:
        err = json.load(open(path, encoding="utf-8")).get("error") or ""
    except ValueError:
        return True
    return "fetch failed" in err


def server_error(e):
    return isinstance(e, urllib.error.HTTPError) and e.code >= 500


try:
    stops = json.load(open(STOPS_FILE, encoding="utf-8"))
except (FileNotFoundError, ValueError):
    stops = {}
todo = [s for s in SYMBOLS if needs_collect(os.path.join(OUT, f"{s}.json"))]
print(f"{len(todo)} of {len(SYMBOLS)} symbols to collect", flush=True)
pos = {s: i for i, s in enumerate(SYMBOLS, 1)}
for done, s in enumerate(todo, 1):
    i = pos[s]
    print(f"[{i}/{len(SYMBOLS)}] {s}: collecting ({done}/{len(todo)} this run)", flush=True)
    before = fetch.TRANSIENT
    try:
        r = collect(s)
    except Exception as e:  # noqa: BLE001
        r = {"symbol": s, "error": f"خطای اجرا: {e!r}", "trace": traceback.format_exc()[-800:]}
    if fetch.TRANSIENT > before:  # not saved: the next run collects this symbol again
        if server_error(fetch.LAST_TRANSIENT):
            stops[s] = stops.get(s, 0) + 1
            json.dump(stops, open(STOPS_FILE, "w", encoding="utf-8"), ensure_ascii=False)
        if stops.get(s, 0) < MAX_SERVER_ERROR_STOPS:
            print(f"STOPPED at {s}: Codal is throttling or unreachable ({fetch.LAST_TRANSIENT}); "
                  f"{len(todo) - done + 1} symbols left for the next run", flush=True)
            sys.exit(STOPPED)
        print(f"[{i}/{len(SYMBOLS)}] {s}: server error {stops[s]} times in a row, saving what was read",
              flush=True)
        fetch.TRANSIENT = 0  # do not fail-fast the next symbols
        r.setdefault("notes", []).append(f"خطای سرور کدال: {fetch.LAST_TRANSIENT}")
    if stops.pop(s, None) is not None:
        json.dump(stops, open(STOPS_FILE, "w", encoding="utf-8"), ensure_ascii=False)
    path = os.path.join(OUT, f"{s}.json")
    with open(path + ".tmp", "w", encoding="utf-8") as f:
        json.dump(r, f, ensure_ascii=False, default=str)
    os.replace(path + ".tmp", path)
    print(f"[{i}/{len(SYMBOLS)}] {s}: {r.get('error') or r.get('report', '')[:60]}", flush=True)
print("DONE", flush=True)
