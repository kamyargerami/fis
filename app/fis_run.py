"""Run one of the skill's batch scripts with a short TTL for Codal searches.

The skill caches search.codal.ir for 24h. When we already know a symbol has a new letter,
a cached search from yesterday could hide it, so searches are refreshed after
CODAL_SEARCH_TTL seconds instead. Everything else (letters, Excel exports) keeps the
skill's own cache rules.

Usage: python -I fis_run.py <skill script> [args...]
"""
import os
import runpy
import sys

script = os.path.abspath(sys.argv[1])
sys.path.insert(0, os.path.dirname(script))
import fetch  # noqa: E402  (the skill's module, shared with the script via sys.modules)

_ttl = int(os.environ.get("CODAL_SEARCH_TTL", "3600"))
_orig_ttl_for = fetch.ttl_for


def ttl_for(url):
    return _ttl if "search.codal.ir" in url else _orig_ttl_for(url)


fetch.ttl_for = ttl_for
sys.argv = [script] + sys.argv[2:]
runpy.run_path(script, run_name="__main__")
