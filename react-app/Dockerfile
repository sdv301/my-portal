# Этап 1 — сборка фронтенда
FROM node:20-slim AS builder
WORKDIR /app
COPY package*.json ./
RUN npm ci
COPY . .
RUN npm run build


# Этап 2 — запуск (отдаём через твой express сервер)
FROM node:20-slim
WORKDIR /app
COPY --from=builder /app/dist ./dist
COPY --from=builder /app/package*.json ./
COPY --from=builder /app/src/server.cjs ./src/
COPY --from=builder /app/data.sqlite ./
COPY --from=builder /app/schema.sql ./
COPY --from=builder /app/init_db.cjs ./
# скопируй другие нужные файлы бэкенда, если они есть
RUN npm ci --omit=dev
RUN mkdir -p uploads
EXPOSE 4000
ENV NODE_ENV=production
CMD ["node", "src/server.cjs"]