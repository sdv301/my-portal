# Этап 1: Сборка фронтенда
FROM node:20-alpine as build
WORKDIR /app
COPY package*.json ./
# Устанавливаем все зависимости (включая devDependencies для сборки vite)
RUN npm install 
COPY . .
RUN npm run build

# Этап 2: Запуск сервера (Node.js)
FROM node:20-alpine
WORKDIR /app

# Копируем зависимости и серверный код
COPY package*.json ./
# Устанавливаем только production-зависимости (без vite и т.д.)
RUN npm install --omit=dev

# Копируем скомпилированный фронтенд из первого этапа в папку, 
# которую сервер сможет раздавать (обычно public или dist)
COPY --from=build /app/dist ./dist
COPY src/server.cjs ./src/
COPY data.sqlite ./data.sqlite

# Если серверу нужны дополнительные файлы (например, geojson)
COPY public/districts.geojson ./public/

# Открываем порт, который указан в server.cjs (5000)
EXPOSE 5000

# Запускаем сервер
CMD ["node", "src/server.cjs"]