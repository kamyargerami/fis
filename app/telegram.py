"""Minimal Telegram Bot API client. The token never appears in logs or raised errors."""
import logging
import time

import requests

from . import config

log = logging.getLogger(__name__)
MAX_TEXT = 4000      # Telegram limit is 4096
MAX_CAPTION = 1000   # Telegram limit is 1024


class TelegramError(RuntimeError):
    pass


def _call(method, data, files=None, attempts=4):
    token = config.telegram_token()
    proxies = {"https": config.OUTBOUND_PROXY} if config.OUTBOUND_PROXY else None
    for attempt in range(attempts):
        if files:  # a file handle is consumed by each request
            for f in files.values():
                f[1].seek(0)
        try:
            r = requests.post(f"https://api.telegram.org/bot{token}/{method}", data=data, files=files,
                              proxies=proxies, timeout=120)
            body = r.json()
        except (requests.RequestException, ValueError) as e:
            raise TelegramError(f"{method} failed: {str(e).replace(token, '***')}") from None
        if body.get("ok"):
            return body["result"]
        wait = (body.get("parameters") or {}).get("retry_after")
        if body.get("error_code") == 429 and wait and attempt + 1 < attempts:
            log.info("telegram: rate limited, retrying %s in %ss", method, wait)
            time.sleep(wait + 1)
            continue
        raise TelegramError(f"{method} failed: {body.get('error_code')} {body.get('description')}")


def send_text(text):
    chat = config.telegram_chat_id()
    for i in range(0, len(text), MAX_TEXT):
        if i:
            time.sleep(1)  # stay under Telegram's per-chat rate limit
        _call("sendMessage", {"chat_id": chat, "text": text[i:i + MAX_TEXT],
                              "disable_web_page_preview": "true"})


def send_document(path, caption=""):
    with open(path, "rb") as f:
        _call("sendDocument", {"chat_id": config.telegram_chat_id(), "caption": caption[:MAX_CAPTION]},
              files={"document": (path.name, f,
                                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")})
