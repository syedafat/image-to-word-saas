# Base Python image
FROM python:3.12-slim

# System Dependencies & Poppler/Tesseract OCR
RUN apt-get update && apt-get install -y \
    poppler-utils \
    tesseract-ocr \
    libtesseract-dev \
    libsm6 \
    libxext6 \
    && rm -rf /var/lib/apt-get/lists/*

# Working directory
WORKDIR /app

# Requirements installation
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application source code
COPY . .

# Environment Variables
ENV TESSERACT_CMD=/usr/bin/tesseract
ENV PYTHONUNBUFFERED=1

# Expose port
EXPOSE 5000

# Run Gunicorn WSGI Server with generous worker timeout for large documents
CMD ["gunicorn", "--bind", "0.0.0.0:5000", "--workers", "2", "--timeout", "120", "app:app"]