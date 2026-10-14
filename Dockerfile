# Dockerfile - untuk Fly.io deployment
# Single-stage: dependensi sistem + pip dalam satu image (lebih andal).
FROM python:3.11-slim

# System deps: PyMuPDF (fitz) + libxml2 untuk lxml
RUN apt-get update && apt-get install -y --no-install-recommends \
    libmupdf-dev \
    libxml2 \
    libxslt1.1 \
    curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Dependensi Python dulu (cache layer)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Kode aplikasi
COPY . .

# Pastikan folder pustaka & output ada
RUN mkdir -p backend/variable_library output

ENV PYTHONUNBUFFERED=1
# Fly.io menyetel FLY_APP_NAME -> server.py masuk mode CLOUD (HTTP saja).
EXPOSE 5050

# Server berbasis http.server (BUKAN Flask/gunicorn).
CMD ["python", "-u", "backend/server.py"]
