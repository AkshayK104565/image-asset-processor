FROM python:3.11-slim

# Install ImageMagick + ghostscript (for PDF conversion)
RUN apt-get update && apt-get install -y --no-install-recommends \
    imagemagick \
    ghostscript \
    && rm -rf /var/lib/apt/lists/*

# Allow ImageMagick to read/write PDFs (disabled by default in Ubuntu)
RUN sed -i 's/rights="none" pattern="PDF"/rights="read|write" pattern="PDF"/' \
    /etc/ImageMagick-6/policy.xml 2>/dev/null || true

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p uploads

EXPOSE 10000
CMD ["gunicorn", "app:app", "--workers", "2", "--threads", "4", "--timeout", "300", "--bind", "0.0.0.0:10000"]
