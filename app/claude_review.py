"""Ask Claude to review the day's changes against the skill's rules and write the Telegram summary.

Claude only receives numbers already computed by the skill's scripts; it does not run code
on the server. If no API key is set or the call fails, a plain summary is used instead.
"""
import json
import logging

import anthropic
from anthropic import DefaultHttpxClient

from . import config

log = logging.getLogger(__name__)

INSTRUCTIONS = """تو بازبین خروجی روزانه skill «fis-stock-scoring» هستی. متن کامل skill پایین آمده است.
هر روز گزارش‌های جدید کدال برای نمادهای فهرست کاربر پیدا می‌شود، اسکریپت‌های skill داده آن نمادها را دوباره
جمع می‌کنند و اکسل رتبه‌بندی از نو ساخته می‌شود. داده امروز را به شکل JSON می‌گیری.

یک پیام تلگرام به فارسی بنویس:
- متن ساده، بدون Markdown و بدون جدول؛ حداکثر حدود ۲۵۰۰ کاراکتر.
- اول: کدام نمادها گزارش جدید داشتند (نوع گزارش و دوره).
- بعد: تغییر نمره و رتبه همان نمادها؛ تغییرهای مهم (ورود یا خروج از بالای میانگین، تغییر وضعیت P/E) را برجسته کن.
- مقادیر پرت یا مشکوک را طبق مرحله ۶ «اجرای گروهی» skill بگو (رشد بیش از ۵۰۰٪، حاشیه منفی شدید، خطا یا یادداشت داده ناقص).
- ۱۰ نماد اول و میانگین گروه را کوتاه بیاور.
- عددی را که در داده نیست نساز.
- در یک خط آخر بگو این تحلیل پژوهشی است و توصیه سرمایه‌گذاری نیست.

=== SKILL.md ===
"""


def _system():
    skill = (config.SKILL_DIR / "SKILL.md").read_text(encoding="utf-8")
    return [{"type": "text", "text": INSTRUCTIONS + skill, "cache_control": {"type": "ephemeral"}}]


def summarize(facts):
    """Persian summary text, or None if Claude is not configured or the call fails."""
    key = config.anthropic_key()
    if not key:
        return None
    http = DefaultHttpxClient(proxy=config.OUTBOUND_PROXY) if config.OUTBOUND_PROXY else None
    client = anthropic.Anthropic(api_key=key, http_client=http, timeout=600, max_retries=3)
    log.info("claude: asking %s (effort=%s) for the summary", config.CLAUDE_MODEL, config.CLAUDE_EFFORT)
    try:
        resp = client.beta.messages.create(
            model=config.CLAUDE_MODEL,
            max_tokens=16000,
            betas=["server-side-fallback-2026-07-01"],
            fallbacks="default",
            output_config={"effort": config.CLAUDE_EFFORT},
            system=_system(),
            messages=[{"role": "user", "content": json.dumps(facts, ensure_ascii=False, default=str)}],
        )
    except anthropic.APIConnectionError as e:
        log.error("claude: connection failed: %s", e)
        return None
    except anthropic.RateLimitError:
        log.error("claude: rate limited")
        return None
    except anthropic.APIStatusError as e:
        log.error("claude: HTTP %s: %s", e.status_code, e.message)
        return None
    if resp.stop_reason == "refusal":
        log.error("claude: refused (%s)", resp.stop_details.category if resp.stop_details else "?")
        return None
    text = "".join(b.text for b in resp.content if b.type == "text").strip()
    log.info("claude: %s, in=%s out=%s", resp.model, resp.usage.input_tokens, resp.usage.output_tokens)
    return text or None
