"""Settings read from the environment (.env via docker compose)."""
import os
from pathlib import Path


def _bool(name, default=False):
    return os.environ.get(name, str(default)).strip().lower() in ("1", "true", "yes", "on")


def _required(name):
    v = os.environ.get(name, "").strip()
    if not v:
        raise SystemExit(f"environment variable {name} is not set")
    return v


DATA_DIR = Path(os.environ.get("DATA_DIR", "/data"))
WORK_DIR = DATA_DIR / "work"            # symbols.json, results/, investees/, holdings.json
OUTPUT_DIR = DATA_DIR / "output"        # dated Excel files + FIS_ranking_latest.xlsx
STATE_FILE = DATA_DIR / "state.json"
LOCK_FILE = DATA_DIR / ".run.lock"

SKILL_DIR = Path(os.environ.get("SKILL_DIR", "/app/skill"))
BATCH_DIR = SKILL_DIR / "scripts" / "batch"

TZ_NAME = os.environ.get("TZ", "Asia/Tehran")
SCHEDULER_ENABLED = _bool("SCHEDULER_ENABLED", True)
# minute hour day month weekday, in TZ
CRON_SCHEDULE = os.environ.get("CRON_SCHEDULE", "").strip() or "0 15 * * *"
RUN_ON_START = _bool("RUN_ON_START", False)

# Proxy only for Telegram and the Claude API. Codal/TSETMC are always fetched directly
# (the skill's fetch.py ignores proxies because Codal blocks foreign IPs).
OUTBOUND_PROXY = os.environ.get("OUTBOUND_PROXY", "").strip() or None

LOOKBACK_DAYS = int(os.environ.get("LOOKBACK_DAYS", "2"))      # first run only
RUN_HOLDINGS = _bool("RUN_HOLDINGS", True)
# listed companies not in the workbook yet that are added per run, until the whole market is covered
BACKFILL_PER_DAY = int(os.environ.get("BACKFILL_PER_DAY", "30"))
SEND_WHEN_NO_CHANGE = _bool("SEND_WHEN_NO_CHANGE", False)
KEEP_EXCEL_FILES = int(os.environ.get("KEEP_EXCEL_FILES", "30"))
STEP_TIMEOUT = int(os.environ.get("STEP_TIMEOUT_SECONDS", str(5 * 3600)))
# market-wide Codal search result is reused for this long (a resumed run does not search again)
WATCH_SEARCH_TTL = int(os.environ.get("WATCH_SEARCH_TTL_SECONDS", str(20 * 3600)))
# a run stopped by Codal throttling is retried after this many minutes (0 = wait for the next cron run)
RESUME_AFTER_MINUTES = int(os.environ.get("RESUME_AFTER_MINUTES", "1"))

CLAUDE_MODEL = os.environ.get("CLAUDE_MODEL", "claude-opus-5-5")
CLAUDE_EFFORT = os.environ.get("CLAUDE_EFFORT", "medium")


def telegram_token():
    return _required("TELEGRAM_BOT_TOKEN")


def telegram_chat_id():
    return _required("TELEGRAM_CHAT_ID")


def anthropic_key():
    return os.environ.get("ANTHROPIC_API_KEY", "").strip() or None
