FROM mcr.microsoft.com/playwright/python:v1.44.0-jammy

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    fonts-dejavu-core \
    fonts-freefont-ttf \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

RUN python -m playwright install chromium

COPY . .

CMD ["/bin/sh", "-c", "gunicorn app:app --bind 0.0.0.0:${PORT:-8080} --workers 1 --timeout 180"]
