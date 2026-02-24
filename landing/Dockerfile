FROM nginx:alpine
COPY index.html /usr/share/nginx/html/
# Копируем картинки (имена должны совпадать с тем, что в HTML)
COPY 1.png /usr/share/nginx/html/1.png
COPY 2.png /usr/share/nginx/html/2.png
COPY 3.png /usr/share/nginx/html/3.png
# Копируем остальные файлы
COPY nginx.conf /etc/nginx/conf.d/default.conf
COPY login.html /usr/share/nginx/html/
COPY admin.html /usr/share/nginx/html/