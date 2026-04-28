FROM python:3.10-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    curl \
    wget \
    gnupg \
    ca-certificates \
    fonts-noto-cjk \
    fonts-dejavu \
    fonts-wqy-zenhei \
    fonts-wqy-microhei \
    fontconfig \
    poppler-utils \
    nginx \
    supervisor \
    && fc-cache -f -v

RUN curl -fsSL https://deb.nodesource.com/setup_20.x | bash - \
    && apt-get install -y nodejs \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package*.json ./
RUN npm install --production && npm cache clean --force

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p /app/data /app/data/backups /app/output /app/logs /app/uploads /app/fonts \
    && chmod 755 /app/data /app/data/backups /app/output /app/logs /app/uploads /app/fonts

RUN mkdir -p /var/log/supervisor \
    && cp /app/nginx.conf /etc/nginx/nginx.conf

# 初始化数据和配置
RUN if [ ! -f /app/data/data.json ]; then \
    echo '{"venues":[],"attendees":[]}' > /app/data/data.json; \
    fi && \
    if [ ! -f /app/data/config.json ]; then \
    cp /app/config.json /app/data/config.json; \
    fi

EXPOSE 80 8506

HEALTHCHECK --interval=30s --timeout=3s --start-period=10s --retries=3 \
    CMD wget --no-verbose --tries=1 --spider http://localhost:3000/api/health || exit 1

CMD ["/usr/bin/supervisord", "-c", "/app/supervisord.conf"]
