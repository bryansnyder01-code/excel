# SheetCraft AI backend.  docker build -t sheetcraft-backend backend
# Add --build-arg WITH_PDF=1 for "Download as PDF" (installs LibreOffice Calc, ~400 MB).
FROM python:3.12-slim

ARG WITH_PDF=0
RUN if [ "$WITH_PDF" = "1" ]; then \
      apt-get update && apt-get install -y --no-install-recommends libreoffice-calc fonts-dejavu-core \
      && rm -rf /var/lib/apt/lists/*; \
    fi

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    SHEETCRAFT_DATA_DIR=/data

WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt

COPY app ./app

RUN useradd --create-home --uid 1000 sheetcraft && mkdir -p /data && chown sheetcraft /data
USER sheetcraft
VOLUME /data

EXPOSE 8000
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s \
  CMD python -c "import urllib.request; urllib.request.urlopen('http://127.0.0.1:8000/health', timeout=4)"
# --proxy-headers: behind a TLS-terminating proxy. ws-max-size leaves room for 5 MB live snapshots.
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8000", "--proxy-headers", "--forwarded-allow-ips", "*", "--ws-max-size", "6291456"]
