FROM python:3.12-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1 \
    HOME=/tmp

WORKDIR /app
COPY requirements.txt .
# optional mirror / build proxy (HTTP_PROXY, HTTPS_PROXY build args are not kept in the image)
ARG PIP_INDEX_URL=https://pypi.org/simple
RUN pip install --timeout 60 --index-url "$PIP_INDEX_URL" -r requirements.txt

# code is owned by root and read-only for the runtime user
COPY skill/ /app/skill/
COPY app/ /app/app/
RUN chmod -R a-w /app

USER 10001:10001
HEALTHCHECK --interval=5m --timeout=10s --start-period=2m \
    CMD python -c "import os,sys,time; sys.exit(time.time()-os.path.getmtime('/tmp/scheduler.heartbeat')>600)"
CMD ["python", "-m", "app.scheduler"]
