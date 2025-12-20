# Home Assistant will automatically inject the correct base image
ARG BUILD_FROM
FROM $BUILD_FROM

# Install Python and required system packages
RUN apk add --no-cache \
    python3 \
    py3-pip \
    py3-setuptools \
    py3-wheel

# Create working directory
WORKDIR /app

# Install Python dependencies first (better caching)
RUN python3 -m venv /app/venv
ENV PATH="/app/venv/bin:$PATH"

COPY requirements.txt /app/
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files
COPY . /app/

# Ensure run script is executable
RUN chmod a+x /app/run.sh

# Default command
CMD ["/app/run.sh"]