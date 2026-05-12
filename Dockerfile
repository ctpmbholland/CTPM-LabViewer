FROM python:3.12-slim

WORKDIR /app

# System deps needed by openpyxl / lxml / pgeocode
RUN apt-get update && apt-get install -y --no-install-recommends \
        build-essential \
        curl \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# App source — copy only what the app needs
COPY CTPM-LabViewer_latest.py .
COPY wo_efficiency.py .
COPY ["CTPM-Logo_thumbnail reduced.png", "."]
COPY CTPM-Weekly-Report/utils/ ./CTPM-Weekly-Report/utils/

# Pre-create cache directories so the app can write to them
RUN mkdir -p .ctpm_cache/uploads .ctpm_cache/parquet

EXPOSE 8501

ENV CTPM_CACHE_DIR=/app/.ctpm_cache

HEALTHCHECK --interval=30s --timeout=10s --start-period=60s \
    CMD curl --fail http://localhost:8501/_stcore/health || exit 1

CMD ["streamlit", "run", "CTPM-LabViewer_latest.py", \
     "--server.port=8501", \
     "--server.address=0.0.0.0", \
     "--server.headless=true", \
     "--browser.gatherUsageStats=false"]
