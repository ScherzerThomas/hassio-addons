ARG BUILD_FROM=ghcr.io/home-assistant/aarch64-addon-php:latest
FROM $BUILD_FROM

RUN apk add --no-cache \
    nginx \
    mariadb \
    mariadb-client \
    certbot \
    curl \
    tar

# Install bashio manually
RUN curl -sSL https://github.com/hassio-addons/bashio/archive/refs/tags/v0.17.5.tar.gz \
    | tar -xz -C /usr/lib/ && \
    mv /usr/lib/bashio-* /usr/lib/bashio


COPY rootfs/ /

CMD ["/init"]
