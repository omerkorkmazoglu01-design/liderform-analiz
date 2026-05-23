FROM python:3.11-slim

# Sistem bağımlılıkları (Playwright Chromium için gerekli)
RUN apt-get update && apt-get install -y \
    wget curl gnupg ca-certificates \
    chromium \
    chromium-driver \
    fonts-liberation \
    libappindicator3-1 \
    libasound2 \
    libatk-bridge2.0-0 \
    libatk1.0-0 \
    libcups2 \
    libdbus-1-3 \
    libgdk-pixbuf2.0-0 \
    libnspr4 \
    libnss3 \
    libx11-xcb1 \
    libxcomposite1 \
    libxdamage1 \
    libxrandr2 \
    xdg-utils \
    libgbm1 \
    libxshmfence1 \
    --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

# Playwright'a sistem Chromium'unu kullandır
ENV PLAYWRIGHT_BROWSERS_PATH=/usr/bin
ENV PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD=1
ENV CHROMIUM_PATH=/usr/bin/chromium

WORKDIR /app

# Python paketleri
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Playwright Chromium'u ayrıca kur (sistem chromium fallback olarak kalır)
RUN python -m playwright install chromium || true
RUN python -m playwright install-deps chromium || true

# Uygulama dosyaları
COPY . .

EXPOSE 8080

CMD gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --timeout 180
