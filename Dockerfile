FROM python:3.11-slim

# System deps: LibreOffice (renders source slides), poppler (pdftoppm), fonts
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-impress \
        poppler-utils \
        fonts-liberation fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

ENV PORT=8000
EXPOSE 8000
CMD ["sh", "-c", "uvicorn server:app --host 0.0.0.0 --port ${PORT}"]
