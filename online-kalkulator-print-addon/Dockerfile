FROM debian:stable

# 1. Added 'locales' to the package list
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice \
        python3 \
        python3-venv \
        python3-pip \
        ca-certificates \
        curl \
        locales && \
    # 2. Generate the de_AT locale
    sed -i -e 's/# de_AT.UTF-8 UTF-8/de_AT.UTF-8 UTF-8/' /etc/locale.gen && \
    locale-gen de_AT.UTF-8 && \
    # Clean up
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# 3. Set Environment Variables for the whole container
# This ensures LibreOffice uses the comma as decimal separator
ENV LANG=de_AT.UTF-8
ENV LC_ALL=de_AT.UTF-8
ENV LANGUAGE=de_AT:de

# Create app directory
WORKDIR /app

# Create virtual environment
RUN python3 -m venv /app/venv

# Copy requirements first (better caching)
COPY requirements.txt /app/

# Install all Python dependencies inside venv
RUN /app/venv/bin/pip install --no-cache-dir -r requirements.txt

# Copy the rest of the app
COPY . /app/

# Make run.sh executable
RUN chmod +x /app/run.sh

CMD ["/app/run.sh"]