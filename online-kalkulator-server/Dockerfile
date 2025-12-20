ARG BUILD_FROM
FROM $BUILD_FROM

RUN apk add --no-cache \
    nginx \
    php81 \
    php81-fpm \
    php81-mysqli \
    php81-json \
    php81-session \
    php81-curl \
    php81-dom \
    php81-xml \
    php81-zip \
    php81-gd \
    php81-mbstring \
    php81-opcache \
    php81-phar \
    php81-iconv \
    php81-simplexml \
    php81-tokenizer \
    php81-xmlwriter \
    php81-xmlreader \
    php81-fileinfo \
    php81-zip \
    mariadb mariadb-client \
    certbot \
    curl \
    tar


# Copy rootfs (configs + s6 services)
COPY rootfs/ /

CMD ["/init"]
