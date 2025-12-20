ARG BUILD_FROM=ghcr.io/home-assistant/aarch64-addon-php:latest
FROM $BUILD_FROM

RUN apk add --no-cache \
    nginx \
    mariadb \
    mariadb-client \
    certbot \
    curl \
    tar

COPY rootfs/ /

CMD ["/init"]
