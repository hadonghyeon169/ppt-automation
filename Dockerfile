FROM python:3.11-slim

# LibreOffice(렌더링 미리보기용) + PDF->PNG 변환 도구 + 다국어 폰트
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    poppler-utils \
    fonts-noto-core \
    fonts-noto-cjk \
    fonts-noto-extra \
    fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

ENV DATA_DIR=/data
VOLUME ["/data"]

EXPOSE 8000
CMD ["gunicorn", "-w", "2", "--threads", "4", "--timeout", "300", "-b", "0.0.0.0:8000", "wsgi:app"]
