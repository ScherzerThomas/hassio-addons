FROM debian:stable

RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        libreoffice \
        python3 \
        python3-venv \
        python3-fastapi \
        python3-pydantic \
        python3-uvicorn \
        ca-certificates \
        curl && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Create app directory
WORKDIR /app

# Create virtual environment
RUN python3 -m venv /app/venv

# Install only missing packages inside venv (allowed)
RUN /app/venv/bin/pip install openpyxl

COPY . /app/
RUN chmod +x /app/run.sh

CMD ["/app/run.sh"]
