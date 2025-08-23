# Use a slim Python image
FROM python:3.11-slim

# Avoid interactive prompts during apt installs
ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1

# Install LibreOffice (for DOCX -> PDF)
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice \
        libreoffice-writer \
        fonts-dejavu \
        fonts-liberation \
        fonts-crosextra-carlito \
        fonts-crosextra-caladea \
        fonts-noto-core \
        fonts-noto-cjk \
        fonts-noto-color-emoji \
        locales \
    && fc-cache -f -v \
    && rm -rf /var/lib/apt/lists/*

# Configure a UTF-8 locale (some PDF conversions rely on it)
RUN sed -i 's/# en_US.UTF-8 UTF-8/en_US.UTF-8 UTF-8/' /etc/locale.gen \
    && locale-gen
ENV LANG=en_US.UTF-8 \
    LANGUAGE=en_US:en \
    LC_ALL=en_US.UTF-8

# App directory
WORKDIR /app

# Install Python deps first (better layer caching)
COPY requirements.txt ./
RUN pip install --upgrade pip \
    && pip install -r requirements.txt

# Copy the rest of the app
COPY . .

# Render provides $PORT
ENV PORT=10000
EXPOSE 10000

# Start Streamlit
CMD ["bash", "-lc", "streamlit run app.py --server.port $PORT --server.address 0.0.0.0 --server.headless true"]
