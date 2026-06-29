# 코스닥 액티브 ETF PDF 변화 — 클라우드(서울 리전) 일일 작업
FROM python:3.12-slim

# 한글 폰트(NanumGothic) + rclone(드라이브 업로드) + 타임존(KST)
RUN apt-get update && apt-get install -y --no-install-recommends \
        fonts-nanum rclone tzdata ca-certificates curl \
    && fc-cache -f \
    && ln -snf /usr/share/zoneinfo/Asia/Seoul /etc/localtime \
    && echo "Asia/Seoul" > /etc/timezone \
    && rm -rf /var/lib/apt/lists/*
ENV TZ=Asia/Seoul

WORKDIR /app
COPY cloud/requirements.txt /app/cloud/requirements.txt
RUN pip install --no-cache-dir -r /app/cloud/requirements.txt

# 앱 코드 (pdf_change.py, portfolio_chart.py, cloud/)
COPY pdf_change.py portfolio_chart.py /app/
COPY cloud/ /app/cloud/

# rclone 설정·스냅샷은 볼륨으로 마운트 (비밀/상태 영속)
#   docker run -v $PWD/state:/app/state -v $PWD/rclone:/root/.config/rclone ...
ENV RCLONE_REMOTE=gdrive:etf_reports

CMD ["python", "/app/cloud/run_daily.py"]
