FROM python:3.11-slim

# Avoid interactive prompts during package installation
ENV DEBIAN_FRONTEND=noninteractive

# Install LibreOffice and minimal runtime dependencies for PDF conversion
RUN apt-get update \
     && apt-get install -y --no-install-recommends \
         libreoffice-writer \
         libreoffice-common \
         default-jre-headless \
         fonts-dejavu-core \
         ca-certificates \
         wget \
     && apt-get clean \
     && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy and install Python dependencies
COPY requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

# Copy app code
COPY . /app

# Ensure temp folder exists
RUN mkdir -p /app/temp

ENV PORT=8000
EXPOSE 8000

CMD ["gunicorn", "--bind", "0.0.0.0:8000", "main:app", "--timeout", "120"]
