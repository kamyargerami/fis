"""Add or update the non-operating income field ("nonop") of results saved before the current rule.

Usage: python3 -I backfill_nonop.py WORKDIR
Reads each statement again from the HTTP cache (published letters are cached forever), so it
needs no new Codal requests for symbols that were collected before. Files already computed with
the current rule (collect.NONOP_RULE), or with an error, are skipped. Safe to run on every pipeline run.
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from collect import NONOP_RULE, income, load_statement, nonop_income  # noqa: E402

WORK = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.getcwd()
done = failed = 0
for sub in ("results", "investees"):
    d = os.path.join(WORK, sub)
    if not os.path.isdir(d):
        continue
    for name in sorted(os.listdir(d)):
        if not name.endswith(".json"):
            continue
        path = os.path.join(d, name)
        r = json.load(open(path, encoding="utf-8"))
        if r.get("nonop_rule") == NONOP_RULE or r.get("error") or not r.get("report_url") or not r.get("report"):
            continue
        try:
            st = load_statement({"Url": r["report_url"], "Title": r["report"]})
            r["nonop"] = nonop_income(income(st["is"]), st["end"]) if st else None
            r["nonop_rule"] = NONOP_RULE
            done += 1
        except Exception as e:  # noqa: BLE001
            print(f"{sub}/{name}: {e!r}"[:200], flush=True)
            failed += 1
            continue
        with open(path + ".tmp", "w", encoding="utf-8") as f:
            json.dump(r, f, ensure_ascii=False, default=str)
        os.replace(path + ".tmp", path)
print(f"nonop added: {done}, failed: {failed}", flush=True)
