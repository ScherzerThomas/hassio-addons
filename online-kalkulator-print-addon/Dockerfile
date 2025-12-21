FROM debian:stable

# Install system dependencies
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice \
        python3 \
        python3-pip \
        python3-venv \
        ca-certificates \
        curl && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Create app directory
WORKDIR /app

# Install Python dependencies
COPY requirements.txt /app/
RUN pip3 install --no-cache-dir -r requirements.txt

# Copy application code
COPY . /app/
RUN chmod +x /app/run.sh

CMD ["/app/run.sh"]
