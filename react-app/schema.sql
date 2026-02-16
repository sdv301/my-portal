-- 1. Таблица районов и получателей (включая ведомства вроде ГБУ "Служба спасения")
CREATE TABLE IF NOT EXISTS districts (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,           -- Официальное название для отображения
    geojson_name TEXT,            -- Имя для связи с картой (если это район)
    type TEXT DEFAULT 'district'  -- 'district' (район) или 'agency' (ведомство)
);

-- 2. Таблица категорий имущества
CREATE TABLE IF NOT EXISTS item_categories (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE -- "Продовольствие...", "Техника" и т.д.
);

-- 3. Справочник номенклатуры (сами товары)
CREATE TABLE IF NOT EXISTS items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    category_id INTEGER,
    name TEXT NOT NULL,       -- "Бензопилы", "Индивидуальный рацион питания"
    unit TEXT,                -- "шт", "кг", "компл"
    unit_price REAL,          -- Балансовая цена за единицу
    FOREIGN KEY (category_id) REFERENCES item_categories(id)
);

-- 4. Таблица распределения (Главная таблица, куда льются данные из Excel)
CREATE TABLE IF NOT EXISTS distributions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    district_id TEXT NOT NULL,    -- Кому выдали (связь с districts)
    item_id INTEGER NOT NULL,     -- Что выдали (связь с items)
    issue_year INTEGER,           -- Год выдачи (например, 2025)
    issue_date DATE,              -- Точная дата (если есть)
    quantity REAL DEFAULT 0,      -- Количество
    total_cost REAL DEFAULT 0,    -- Общая стоимость
    FOREIGN KEY (district_id) REFERENCES districts(id),
    FOREIGN KEY (item_id) REFERENCES items(id)
);