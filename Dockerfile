FROM python:3.11-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# 1. Instalar dependencias del sistema necesarias para Oracle Instant Client
RUN apt-get update && apt-get install -y --no-install-recommends \
    libaio1 \
    wget \
    unzip \
    && rm -rf /var/lib/apt/lists/*

# 2. Descargar y descomprimir Oracle Instant Client (Modo Thick)
RUN mkdir -p /opt/oracle && \
    cd /opt/oracle && \
    wget -q https://download.oracle.com/otn_software/linux/instantclient/1921000/instantclient-basic-linux.x64-19.21.0.0.0dbru.zip && \
    unzip instantclient-basic-linux.x64-19.21.0.0.0dbru.zip && \
    rm instantclient-basic-linux.x64-19.21.0.0.0dbru.zip && \
    mv instantclient_* instantclient

# Configurar la ruta de las librerías en el sistema
ENV LD_LIBRARY_PATH=/opt/oracle/instantclient:$LD_LIBRARY_PATH

# 3. Instalar dependencias de Python
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 4. Copiar el resto del código de la aplicación
COPY . .

EXPOSE 11020

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "11020", "--workers", "4"]