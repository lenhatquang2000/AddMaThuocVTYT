# Sử dụng PHP 8.2 FPM làm base image
FROM php:8.2-fpm

# Thiết lập thư mục làm việc
WORKDIR /var/www

# Cài đặt các system dependencies cần thiết
RUN apt-get update && apt-get install -y \
    build-essential \
    libpng-dev \
    libjpeg62-turbo-dev \
    libfreetype6-dev \
    locales \
    zip \
    jpegoptim optipng pngquant gifsicle \
    vim \
    unzip \
    git \
    curl \
    libonig-dev \
    libxml2-dev \
    libzip-dev \
    libpq-dev

# Xóa cache
RUN apt-get clean && rm -rf /var/lib/apt/lists/*

# Cài đặt các PHP extensions cần thiết cho Laravel và PhpSpreadsheet
RUN docker-php-ext-install pdo_pgsql mbstring zip exif pcntl bcmath gd

# Cài đặt Composer
COPY --from=composer:latest /usr/bin/composer /usr/bin/composer

# Copy mã nguồn vào container
COPY . /var/www

# Cài đặt các PHP dependencies
RUN composer install --no-scripts --no-autoloader

# Phân quyền cho Laravel
RUN chown -R www-data:www-data /var/www/storage /var/www/bootstrap/cache

# Expose port 9000
EXPOSE 9000

CMD ["php-fpm"]
