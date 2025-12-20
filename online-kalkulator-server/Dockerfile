ARG BUILD_FROM
FROM $BUILD_FROM

RUN apk add --no-cache \
    nginx \
    php82 \
    php82-fpm \
    php82-mysqli \
    php82-json \
    php82-session \
    php82-curl \
    php82-dom \
    php82-xml \
    php82-zip \
    php82-gd \
    php82-mbstring \
    php82-opcache \
    php82-phar \
    php82-iconv \
    php82-simplexml \
    php82-tokenizer \
    php82-xmlwriter \
    php82-xmlreader \
    php82-fileinfo \
    mariadb mariadb-client \
    certbot \
    curl \
    tar


COPY rootfs/ /

CMD ["/init"]
