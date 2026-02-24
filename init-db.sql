-- Инициализация базы данных portal_db
CREATE TABLE IF NOT EXISTS test_table (id SERIAL PRIMARY KEY);

-- Разрешаем пользователю читать все текущие и будущие таблицы
-- (Grafana использует этого же пользователя для SELECT-запросов)
ALTER DEFAULT PRIVILEGES IN SCHEMA public GRANT SELECT ON TABLES TO portal_user;