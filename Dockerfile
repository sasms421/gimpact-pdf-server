# G-IMPACT PDF Generator Server
# Docker 이미지

FROM python:3.11-slim

# 시스템 패키지 설치 (폰트 포함)
RUN apt-get update && apt-get install -y \
    fonts-nanum \
    fonts-nanum-coding \
    fontconfig \
    && rm -rf /var/lib/apt/lists/* \
    && fc-cache -fv

# 작업 디렉토리
WORKDIR /app

# 의존성 설치
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 소스 코드 복사
COPY pdf_api_server.py .
COPY analysis_report_generator.py .
COPY real_sample_data.py .

# 포트 설정
ENV PORT=8080
EXPOSE 8080

# 실행
CMD ["python", "pdf_api_server.py"]
