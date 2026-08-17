FROM python:3.11-slim

# LibreOffice (para exportar a PDF) + Carlito, fuente metric-compatible con
# Calibri, para que el PDF se vea igual a como se ve en Word.
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-writer \
    fonts-crosextra-carlito \
    fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Render inyecta la variable PORT; el fallback 8000 es solo para correr local.
ENV PORT=8000

# Render termina el TLS en su proxy y le habla al contenedor por HTTP plano.
# Por defecto uvicorn solo confía en 127.0.0.1, así que descarta el header
# X-Forwarded-Proto y arma las URLs con http:// -> el navegador bloquea la
# descarga por mixed content. Con esto uvicorn confía en el proxy y el
# scheme queda en https.
ENV FORWARDED_ALLOW_IPS=*

CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT} --proxy-headers --forwarded-allow-ips '*'"]
