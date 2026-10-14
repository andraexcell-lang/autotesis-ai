# Dockerfile - untuk Fly.io deployment
# Multi-stage build untuk image kecil

FROM python:3.11-slim as builder

# Install system dependencies untuk PyMuPDF
RUN apt-get update && apt-get install -y --no-install-recommends \
    libmupdf-dev \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# Install python dependencies
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir --user -r requirements.txt

# Final stage
FROM python:3.11-slim

# Install runtime dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    libmupdf-dev \
    && rm -rf /var/lib/apt/lists/*

# Copy installed packages from builder
COPY --from=builder /root/.local /root/.local

# Set PATH
ENV PATH=/root/.local/bin:$PATH

WORKDIR /app
COPY . .

# Create variable_library directory
RUN mkdir -p backend/variable_library

EXPOSE 5050

CMD ["gunicorn", "backend.server:app", "--bind", "0.0.0.0:5050", "--workers", "2", "--timeout", "120"]