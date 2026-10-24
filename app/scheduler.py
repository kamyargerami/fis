"""In-container cron: runs the pipeline on CRON_SCHEDULE in TZ (default every day at 15:00 Asia/Tehran).

With SCHEDULER_ENABLED=false nothing runs on a schedule; the container stays up for manual runs
(`docker compose exec fis-bot python -m app.pipeline`), and RUN_ON_START still works.

Each run is a separate process, so a crash or memory leak in one run cannot affect the next.
"""
import datetime
import logging
import subprocess
import sys
import time
from pathlib import Path
from zoneinfo import ZoneInfo

from croniter import croniter

from . import config

log = logging.getLogger("scheduler")
RESUMABLE = 75  # pipeline exit code: stopped early, progress saved
HEARTBEAT = Path("/tmp/scheduler.heartbeat")


def run_pipeline():
    log.info("starting pipeline")
    p = subprocess.Popen([sys.executable, "-m", "app.pipeline"], cwd="/app")
    deadline = time.monotonic() + config.STEP_TIMEOUT * 3
    while p.poll() is None:
        HEARTBEAT.touch()   # a long run is not a dead scheduler
        if time.monotonic() > deadline:
            p.kill()
            log.error("pipeline timed out and was killed")
        time.sleep(30)
    log.info("pipeline exited with %s", p.returncode)
    return p.returncode


def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(name)s: %(message)s")
    tz = ZoneInfo(config.TZ_NAME)
    if not croniter.is_valid(config.CRON_SCHEDULE):
        raise SystemExit(f"invalid CRON_SCHEDULE: {config.CRON_SCHEDULE!r}")
    if config.SCHEDULER_ENABLED:
        log.info("schedule: %r (%s)", config.CRON_SCHEDULE, config.TZ_NAME)
    else:
        log.info("scheduler disabled (SCHEDULER_ENABLED=false); runs only on start or by hand")
    rc = run_pipeline() if config.RUN_ON_START else 0
    while True:
        now = datetime.datetime.now(tz)
        nxt = None
        if config.SCHEDULER_ENABLED:
            nxt = croniter(config.CRON_SCHEDULE, now).get_next(datetime.datetime)
        if rc == RESUMABLE and config.RESUME_AFTER_MINUTES > 0:  # finish a stopped run even when disabled
            retry = now + datetime.timedelta(minutes=config.RESUME_AFTER_MINUTES)
            nxt = min(nxt, retry) if nxt else retry
            log.info("run stopped early; continuing it at %s", nxt.isoformat())
        if nxt is None:
            log.info("no run scheduled; idling")
            while True:
                HEARTBEAT.touch()
                time.sleep(60)
        log.info("next run: %s", nxt.isoformat())
        while (left := (nxt - datetime.datetime.now(tz)).total_seconds()) > 0:
            HEARTBEAT.touch()
            time.sleep(min(left, 60))
        rc = run_pipeline()


if __name__ == "__main__":
    main()
