#!/bin/bash

# Установка SSL сертификатов

echo "Настройка HTTPS для портала..."

# Создаем конфигурацию nginx с SSL
cat > nginx/ssl.conf << EOF
server {
    listen 80;
    server_name your-domain.com;
    location /.well-known/acme-challenge/ {
        root /var/www/certbot;
    }
    location / {
        return 301 https://\$host\$request_uri;
    }
}

server {
    listen 443 ssl http2;
    server_name your-domain.com;
    
    ssl_certificate /etc/letsencrypt/live/your-domain.com/fullchain.pem;
    ssl_certificate_key /etc/letsencrypt/live/your-domain.com/privkey.pem;
    
    # Настройки безопасности SSL
    ssl_protocols TLSv1.2 TLSv1.3;
    ssl_ciphers ECDHE-RSA-AES256-GCM-SHA512:DHE-RSA-AES256-GCM-SHA512;
    ssl_prefer_server_ciphers off;
    
    location / {
        proxy_pass http://portal:80;
        # ... остальные настройки
    }
}
EOF

echo "Для получения реальных сертификатов выполните:"
echo "docker-compose run --rm certbot certonly --webroot -w /var/www/certbot -d your-domain.com"