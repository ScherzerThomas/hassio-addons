FROM debian:stable

RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice \
        python3 \
        python3-venv \
        python3-pip \
        ca-certificates \
        curl && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

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
