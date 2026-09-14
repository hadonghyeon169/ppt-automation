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
# 영구 저장공간은 호스팅 플랫폼의 볼륨(Railway Volume 등)을 /data에 마운트해서 제공합니다.
# (Docker VOLUME 명령은 일부 플랫폼 빌더에서 지원하지 않아 제거했습니다.)

EXPOSE 8000
# 쉘 형식 CMD: Railway 등 호스팅 플랫폼이 주입하는 $PORT 환경변수를 그대로 바인딩한다.
# (없으면 로컬 실행 등을 위해 8000으로 기본값 처리)
CMD gunicorn -w 2 --threads 4 --timeout 300 -b 0.0.0.0:${PORT:-8000} wsgi:app
