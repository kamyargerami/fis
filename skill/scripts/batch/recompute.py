"""Re-derive collected results with the current collect() after a method change, from the HTTP cache.

Usage: python3 -I recompute.py WORKDIR [--fetch] [--all]
Files in results/ and investees/ saved by an older collect.METHOD are collected again with Codal
requests disabled (FIS_CACHE_ONLY=1): published letters are cached forever, so this normally needs
no network. A symbol whose letters are not all cached keeps its old file. Price and share count come
from the cached TSETMC data; the pipeline refreshes prices afterwards. Safe to run on every run.
--fetch allows Codal requests for letters missing from the cache (e.g. earlier annual reports a new
rule needs); --all recomputes every file, not only older ones.
"""
import json
import os
import sys

FETCH = "--fetch" in sys.argv
if not FETCH:
    os.environ["FIS_CACHE_ONLY"] = "1"
os.environ["FIS_OFFLINE"] = "1"
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from collect import METHOD, collect  # noqa: E402

args = [a for a in sys.argv[1:] if not a.startswith("--")]
WORK = os.path.abspath(args[0]) if args else os.getcwd()
done = kept = 0
for sub in ("results", "investees"):
    d = os.path.join(WORK, sub)
    if not os.path.isdir(d):
        continue
    for name in sorted(os.listdir(d)):
        if not name.endswith(".json"):
            continue
        path = os.path.join(d, name)
        old = json.load(open(path, encoding="utf-8"))
        if (old.get("method") == METHOD and "--all" not in sys.argv) or old.get("error"):
            continue
        try:
            r = collect(old["symbol"])
        except Exception as e:  # noqa: BLE001
            r = {"error": repr(e)}
        if r.get("error"):
            print(f"{sub}/{name}: kept old result ({str(r['error'])[:120]})", flush=True)
            kept += 1
            continue
        with open(path + ".tmp", "w", encoding="utf-8") as f:
            json.dump(r, f, ensure_ascii=False, default=str)
        os.replace(path + ".tmp", path)
        done += 1
        if done % 50 == 0:
            print(f"recomputed {done}", flush=True)
print(f"recomputed: {done}, kept old: {kept}", flush=True)
