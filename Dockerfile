FROM python:3.11-slim

# Only ghostscript needed now (for PDF, no ImageMagick)
RUN apt-get update && apt-get install -y --no-install-recommends \
    ghostscript \
    libgl1 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .
RUN mkdir -p uploads

EXPOSE 10000
CMD ["gunicorn","app:app","--workers","1","--threads","4","--timeout","300","--bind","0.0.0.0:10000","--max-requests","100","--max-requests-jitter","20"]
