FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    wget curl gnupg ca-certificates \
    libnss3 libnspr4 libatk1.0-0 \
    libatk-bridge2.0-0 libcups2 libdbus-1-3 \
    libxcomposite1 libxdamage1 libxrandr2 \
    libgbm1 libxshmfence1 libasound2 \
    libx11-xcb1 libxfixes3 libxext6 \
    fonts-liberation xdg-utils \
    --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

ENV PLAYWRIGHT_BROWSERS_PATH=/ms-playwright

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
RUN python -m playwright install --with-deps chromium

COPY . .

CMD gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --timeout 180
