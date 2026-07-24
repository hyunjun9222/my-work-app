# 포털(app.py) 컨테이너 이미지 — Fly.io 배포용
FROM python:3.11-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

# Fly 는 아래 포트로 트래픽을 보낸다(fly.toml 의 internal_port 와 같게).
ENV PORT=8080
EXPOSE 8080

# app.py 는 PORT 가 있으면 0.0.0.0:PORT 로 바인딩한다(클라우드 모드).
CMD ["python", "app.py"]
