FROM php:8.2-cli

WORKDIR /var/www

# Install system dependencies and php extensions via script for speed
RUN apt-get update && apt-get install -y \
    libpng-dev \
    libjpeg62-turbo-dev \
    libfreetype6-dev \
    zip \
    unzip \
    git \
    curl \
    libonig-dev \
    libxml2-dev \
    libzip-dev \
    libpq-dev \
    && curl -sSL https://github.com/mlocati/docker-php-extension-installer/releases/latest/download/install-php-extensions -o /usr/local/bin/install-php-extensions \
    && chmod +x /usr/local/bin/install-php-extensions \
    && install-php-extensions pdo_pgsql mbstring zip exif pcntl bcmath gd \
    && apt-get clean && rm -rf /var/lib/apt/lists/*

COPY --from=composer:latest /usr/bin/composer /usr/bin/composer

COPY . /var/www

RUN composer install --optimize-autoloader --no-dev

RUN cp .env.example .env || true
RUN php artisan key:generate --force || true

# Force drivers to 'file' to avoid database dependency
RUN sed -i 's/SESSION_DRIVER=database/SESSION_DRIVER=file/g' .env || echo "SESSION_DRIVER=file" >> .env
RUN sed -i 's/CACHE_STORE=database/CACHE_STORE=file/g' .env || echo "CACHE_STORE=file" >> .env

RUN chown -R www-data:www-data /var/www/storage /var/www/bootstrap/cache
RUN chmod -R 775 /var/www/storage /var/www/bootstrap/cache

EXPOSE 10000

CMD php artisan serve --host=0.0.0.0 --port=10000
