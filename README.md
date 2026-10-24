# FIS Codal Bot

طبق زمان‌بندی `.env` (پیش‌فرض: هر روز ساعت ۱۵ به وقت تهران):

1. کدال را برای گزارش‌های جدید (صورت مالی ن-۱۰ و گزارش ماهانه ن-۳۰) نمادهای `data/work/symbols.json` و زیرمجموعه‌های هلدینگ‌ها می‌گردد.
2. هر شرکت بورسی/فرابورسی که گزارش جدید داده و در اکسل نیست به `symbols.json` اضافه می‌شود؛ هلدینگ‌ها و شرکت‌های سرمایه‌گذاری خودکار به تب خودشان می‌روند. هر اجرا هم `BACKFILL_PER_DAY` شرکتِ هنوز‌نیامده را اضافه می‌کند تا کم‌کم کل بازار پوشش داده شود.
3. اگر گزارش جدیدی بود، داده همان نمادها را با اسکریپت‌های skill «fis-stock-scoring» دوباره جمع می‌کند، قیمت‌ها را به‌روز می‌کند، هلدینگ‌ها را دوباره ارزیابی می‌کند و اکسل را از نو می‌سازد.
4. Claude (با API key) تغییرات را طبق قواعد skill بررسی می‌کند و خلاصه فارسی می‌نویسد.
5. اکسل و خلاصه به تلگرام فرستاده می‌شود.

## راه‌اندازی

```bash
cp .env.example .env        # اگر .env ندارید
chmod 600 .env
# ANTHROPIC_API_KEY و APP_UID/APP_GID (خروجی id -u و id -g) را پر کنید
docker compose up -d --build
```

سرور باید IP ایران داشته باشد (کدال IP خارجی را قطع می‌کند). اگر از سرور به `api.telegram.org` و `api.anthropic.com` دسترسی نیست، `OUTBOUND_PROXY` را تنظیم کنید؛ این پروکسی فقط برای تلگرام و Claude استفاده می‌شود و کدال همیشه مستقیم است. اگر پروکسی روی خود سرور و روی `127.0.0.1` گوش می‌دهد، از داخل کانتینر در دسترس نیست؛ آن را روی IP پل داکر (مثلاً `172.17.0.1`) باز کنید.

## دستورها

```bash
docker compose logs -f                                         # لاگ
docker compose exec fis-bot python -m app.pipeline             # اجرای دستی
docker compose exec fis-bot python -m app.pipeline --force     # ساخت و ارسال حتی بدون گزارش جدید
docker compose exec fis-bot python -m app.pipeline --dry-run   # بدون Claude، تلگرام و ذخیره وضعیت
```

زمان‌بندی در `.env` است (به وقت تهران): `CRON_SCHEDULE` یک عبارت cron است (دقیقه ساعت روز ماه روزِهفته)، مثلاً `0 15 * * *` هر روز ساعت ۱۵ یا `0 */6 * * *` هر ۶ ساعت. برای غیرفعال کردن، `SCHEDULER_ENABLED=false` بگذارید؛ آن وقت فقط با `RUN_ON_START` یا دستی (`docker compose exec fis-bot python -m app.pipeline`) اجرا می‌شود.

## داده‌ها (`data/`)

| مسیر | محتوا |
|---|---|
| `work/symbols.json` | فهرست نمادها؛ خودکار بزرگ می‌شود. نماد دلخواه هم اضافه کنید، در اجرای بعدی جمع‌آوری می‌شود. |
| `work/symbols.user.json` | نسخه اصلی فهرست دستی شما (قبل از اضافه‌شدن خودکار) |
| `work/results`, `work/investees`, `work/holdings.json` | خروجی اسکریپت‌های skill |
| `cache/http` | کش کدال و TSETMC |
| `output/` | اکسل هر روز و `FIS_ranking_latest.xlsx` |
| `state.json` | آخرین اجرای موفق و گزارش‌های دیده‌شده |

## به‌روزرسانی skill

پوشه `skill/` کپی skill است. بعد از تغییر skill:

```bash
rm -rf skill && cp -r ~/.claude/skills/fis-stock-scoring skill && docker compose up -d --build
```

## امنیت

- کانتینر با کاربر غیر root، فایل‌سیستم فقط‌خواندنی، بدون capability، با `no-new-privileges` و محدودیت حافظه و پردازه اجرا می‌شود و هیچ پورتی باز نمی‌کند.
- اسرار فقط در `.env` (دسترسی 600) هستند و وارد image نمی‌شوند (`.dockerignore`). اسکریپت‌های skill بدون این متغیرها اجرا می‌شوند و توکن تلگرام در لاگ و پیام خطا نمی‌آید.
- Claude فقط اعداد محاسبه‌شده را می‌گیرد و روی سرور کدی اجرا نمی‌کند.
- نسخه همه پکیج‌ها ثابت است.
