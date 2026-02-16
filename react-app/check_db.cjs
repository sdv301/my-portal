// check_db.cjs
const sqlite3 = require('sqlite3').verbose();
const db = new sqlite3.Database('data.sqlite');

console.log('🔍 Проверка базы данных материальных резервов...\n');

db.serialize(() => {
  // 1. Проверяем районы
  db.all("SELECT id, name, type, population FROM districts LIMIT 5", [], (err, rows) => {
    if (err) return console.error(err.message);
    console.log('--- 📍 РАЙОНЫ (Первые 5 записей) ---');
    console.table(rows);
  });

  // 2. Проверяем категории
  db.all("SELECT * FROM item_categories", [], (err, rows) => {
    if (err) return console.error(err.message);
    console.log('\n--- 📂 КАТЕГОРИИ ИМУЩЕСТВА ---');
    if (rows.length === 0) console.log('Категорий пока нет.');
    else console.table(rows);
  });

  // 3. Проверяем товары (номенклатуру)
  const itemsSql = `
    SELECT i.id, c.name as category, i.name, i.unit, i.unit_price 
    FROM items i 
    JOIN item_categories c ON i.category_id = c.id 
    LIMIT 5
  `;
  db.all(itemsSql, [], (err, rows) => {
    if (err) return console.error(err.message);
    console.log('\n--- 📦 НОМЕНКЛАТУРА (Первые 5 товаров) ---');
    if (rows.length === 0) console.log('Товаров пока нет.');
    else console.table(rows);
  });

  // 4. Проверяем распределение (кто, что и сколько получил)
  const distSql = `
    SELECT 
      d.name as district_name,
      c.name as category,
      i.name as item_name,
      dist.quantity,
      i.unit,
      dist.total_cost,
      dist.issue_year
    FROM distributions dist
    JOIN districts d ON dist.district_id = d.id
    JOIN items i ON dist.item_id = i.id
    JOIN item_categories c ON i.category_id = c.id
    LIMIT 10
  `;
  db.all(distSql, [], (err, rows) => {
    if (err) return console.error(err.message);
    console.log('\n--- 🚚 ВЫДАЧА РЕЗЕРВОВ (Первые 10 записей) ---');
    if (rows.length === 0) {
      console.log('Данных о выдаче пока нет. Загрузите Excel-файл через админку!');
    } else {
      console.table(rows);
    }
  });
});

db.close();