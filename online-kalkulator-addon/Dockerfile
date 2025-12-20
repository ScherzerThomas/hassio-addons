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
COPY requirements.txt /app/
RUN pip3 install --no-cache-dir -r requirements.txt

# Copy application files
COPY run.sh /app/run.sh

# Ensure run script is executable
RUN chmod a+x /app/run.sh

# Default command
CMD ["/app/run.sh"]