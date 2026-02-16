FROM nginx:alpine
COPY index.html /usr/share/nginx/html/
# Копируем логотип (который мы добавили ранее)
COPY 1.png /usr/share/nginx/html/logo.png
# Копируем картинки модулей
COPY 2.png /usr/share/nginx/html/flask.png
COPY 3.png /usr/share/nginx/html/react.png
# Копируем остальные файлы
COPY nginx.conf /etc/nginx/conf.d/default.conf
COPY login.html /usr/share/nginx/html/
COPY admin.html /usr/share/nginx/html/