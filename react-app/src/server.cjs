// src/server.cjs
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const xlsx = require('xlsx');
const sqlite3 = require('sqlite3').verbose();
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ dest: 'uploads/' });

if (!fs.existsSync('uploads')) {
  fs.mkdirSync('uploads');
}

const PORT = 5000;

app.use(cors());
app.use(express.json());

const dbPath = path.resolve(__dirname, '../data.sqlite');
const db = new sqlite3.Database(dbPath);

// --- API эндпоинты ---

// 1. Получить все районы
app.get('/api/districts', (req, res) => {
  db.all('SELECT * FROM districts ORDER BY name', [], (err, rows) => {
    if (err) return res.status(500).json({ error: err.message });
    res.json(rows);
  });
});

// 2. Получить справочник индикаторов
app.get('/api/indicators', (req, res) => {
  const sql = `
    SELECT i.id, i.name, i.unit, c.name as category 
    FROM indicators i 
    JOIN indicator_categories c ON i.category_id = c.id
    ORDER BY c.name, i.name`;
  db.all(sql, [], (err, rows) => {
    if (err) return res.status(500).json({ error: err.message });
    res.json(rows);
  });
});

// 3. Получить информацию о районе (ищем и по ID, и по имени)
app.get('/api/district/:id', (req, res) => {
  const { id } = req.params;
  db.get('SELECT * FROM districts WHERE id = ? OR name = ?', [id, id], (err, row) => {
    if (err) return res.status(500).json({ error: err.message });
    if (!row) return res.status(404).json({ error: 'Район не найден' });
    res.json(row);
  });
});

// API: Получить информацию о районе (ИСПРАВЛЕНО ДЛЯ SQLITE)
app.get('/api/district/:id', (req, res) => {
  const { id } = req.params;
  
  db.get('SELECT * FROM districts WHERE id = ? OR name = ?', [id, id], (err, row) => {
    if (err) return res.status(500).json({ error: err.message });
    if (!row) return res.status(404).json({ error: 'Район не найден' });
    
    // Возвращаем данные района из базы
    res.json(row);
  });
});

// --- ИСПРАВЛЕНО: Получение списка материальных резервов для InfoPanel ---
// --- ИСПРАВЛЕНО: Получение списка материальных резервов с учетом фильтра дат ---
app.get('/api/district/:id/data', (req, res) => {
  const { id } = req.params;
  const { startDate, endDate } = req.query;

  db.get('SELECT id FROM districts WHERE id = ? OR name = ?', [id, id], (err, district) => {
    if (err || !district) return res.status(404).json({ error: 'Район не найден' });

    let sql = `
      SELECT 
        c.name as category_name, 
        i.name as item_name, 
        i.unit, 
        d.quantity, 
        d.total_cost, 
        d.issue_year
      FROM distributions d
      JOIN items i ON d.item_id = i.id
      JOIN item_categories c ON i.category_id = c.id
      WHERE d.district_id = ? AND d.quantity > 0
    `;
    const params = [district.id];

    // Применяем фильтр по датам с фронтенда
    if (startDate && endDate) {
      const startYear = parseInt(startDate.split('-')[0]);
      const endYear = parseInt(endDate.split('-')[0]);
      sql += " AND d.issue_year BETWEEN ? AND ?";
      params.push(startYear, endYear);
    }

    sql += " ORDER BY c.name, i.name";

    db.all(sql, params, (err, rows) => {
      if (err) return res.status(500).json({ error: err.message });

      const inventory = {};
      let grandTotal = 0;
      let minYear = null;
      let maxYear = null;

      rows.forEach(row => {
        const cat = row.category_name || 'Разное';
        if (!inventory[cat]) inventory[cat] = [];
        
        inventory[cat].push({
          name: row.item_name,
          unit: row.unit,
          quantity: row.quantity,
          cost: row.total_cost,
          year: row.issue_year
        });
        grandTotal += row.total_cost;

        // Вычисляем реальный диапазон дат для этих данных
        if (minYear === null || row.issue_year < minYear) minYear = row.issue_year;
        if (maxYear === null || row.issue_year > maxYear) maxYear = row.issue_year;
      });

      res.json({
        inventory,
        statistics: {
          total_cost: grandTotal,
          total_items: rows.length,
          earliest_date: minYear,
          latest_date: maxYear
        }
      });
    });
  });
});

// 5. Получить доп. инфо (статьи)
app.get('/api/district/:id/info', (req, res) => {
  const { id } = req.params;
  
  db.get('SELECT id FROM districts WHERE id = ? OR name = ?', [id, id], (err, district) => {
    if (err || !district) return res.status(404).json({});
    
    // Берем данные из таблицы district_info (если она у тебя так называется)
    db.all('SELECT category, title, content, updated_at FROM district_info WHERE district_id = ?', [district.id], (err, rows) => {
      if (err) return res.status(500).json({});
      
      const info = {};
      rows.forEach(r => {
        if (!info[r.category]) info[r.category] = [];
        info[r.category].push({ 
          title: r.title, 
          content: r.content, 
          updatedAt: r.updated_at 
        });
      });
      res.json(info);
    });
  });
});

// 6. Добавить данные вручную
app.post('/api/data', (req, res) => {
  const { district_id, indicator_id, date, value, source } = req.body;
  const sql = `INSERT OR REPLACE INTO data_values (district_id, indicator_id, date, value, source) VALUES (?, ?, ?, ?, ?)`;
  db.run(sql, [district_id, indicator_id, date, value, source], function(err) {
    if (err) return res.status(500).json({ success: false, error: err.message });
    res.json({ success: true, id: this.lastID });
  });
});

// 7. Импорт из Excel
/// --- ИСПРАВЛЕННЫЙ ПАРСЕР EXCEL (Синхронное сохранение) ---
app.post('/api/upload-excel', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'Файл не загружен' });

  try {
    const workbook = xlsx.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const data = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });

    // Индексы колонок по твоему файлу (6 - Жиганский, 8 - Оймяконский, 10 - Абыйский)
    const colToDistrict = {
      6: 'zhigansky',
      8: 'oymyakonsky',
      10: 'abysky'
    };

    // Вспомогательные функции, чтобы заставить БД ждать (Promises)
    const runDb = (sql, params) => new Promise((resolve, reject) => {
      db.run(sql, params, function(err) {
        if (err) reject(err); else resolve(this.lastID);
      });
    });

    const getDb = (sql, params) => new Promise((resolve, reject) => {
      db.get(sql, params, (err, row) => {
        if (err) reject(err); else resolve(row);
      });
    });

    let currentCategoryId = null;
    let recordsAdded = 0;

    // Идем по строкам Excel
    for (let i = 5; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length === 0) continue;

      const colA = row[0]; // Номер
      const colB = row[1]; // Наименование

      // Если это Категория (нет номера, но есть текст)
      if (!colA && colB) {
        await runDb(`INSERT OR IGNORE INTO item_categories (name) VALUES (?)`, [colB]);
        const cat = await getDb(`SELECT id FROM item_categories WHERE name = ?`, [colB]);
        if (cat) currentCategoryId = cat.id;
        continue;
      }

      // Если это Товар (есть и номер, и наименование)
      if (colA && colB) {
        const itemName = colB;
        const unit = row[2];
        const price = parseFloat(row[4]) || 0;

        // 1. Сохраняем товар с ПРАВИЛЬНОЙ категорией
        const itemId = await runDb(
          `INSERT INTO items (category_id, name, unit, unit_price) VALUES (?, ?, ?, ?)`, 
          [currentCategoryId, itemName, unit, price]
        );

        // 2. Проверяем колонки районов и записываем выдачу
        for (const colIndex of Object.keys(colToDistrict)) {
          const qtyCol = parseInt(colIndex);
          const costCol = qtyCol + 1;
          
          const quantity = parseFloat(row[qtyCol]) || 0;
          const cost = parseFloat(row[costCol]) || 0;

          if (quantity > 0) {
            await runDb(
              `INSERT INTO distributions (district_id, item_id, issue_year, quantity, total_cost) VALUES (?, ?, 2025, ?, ?)`, 
              [colToDistrict[colIndex], itemId, quantity, cost]
            );
            recordsAdded++;
          }
        }
      }
    }

    fs.unlinkSync(req.file.path); // Удаляем файл
    res.json({ success: true, message: `Загружено записей о выдаче: ${recordsAdded}`, count: recordsAdded });

  } catch (error) {
    console.error('Ошибка обработки Excel:', error);
    res.status(500).json({ error: error.message });
  }
});

// 8. Экспорт в Excel за период
// --- ДОБАВЛЕНО: Проверка наличия данных для Excel ---
app.get('/api/check-export', (req, res) => {
  const { startDate, endDate, district_id } = req.query;
  
  let baseSql = "SELECT COUNT(*) as count, MAX(issue_year) as latest_year FROM distributions WHERE quantity > 0";
  const baseParams = [];
  
  if (district_id && district_id !== 'all') {
    baseSql += " AND district_id = ?";
    baseParams.push(district_id);
  }

  let checkSql = baseSql;
  const checkParams = [...baseParams];

  // Проверяем выбранный период
  if (startDate && endDate) {
    const startYear = parseInt(startDate.split('-')[0]);
    const endYear = parseInt(endDate.split('-')[0]);
    checkSql += " AND issue_year BETWEEN ? AND ?";
    checkParams.push(startYear, endYear);
  }

  db.get(checkSql, checkParams, (err, row) => {
    if (err) return res.status(500).json({ error: err.message });
    
    // Если данные за этот период есть
    if (row && row.count > 0) {
      res.json({ hasData: true });
    } else {
      // Если данных нет, ищем, в каком году были самые свежие записи
      db.get(baseSql, baseParams, (err2, row2) => {
         res.json({ 
           hasData: false, 
           latest_year: row2 && row2.latest_year ? row2.latest_year : null 
         });
      });
    }
  });
});

app.get('/api/date-range', (req, res) => {
  db.get('SELECT MIN(issue_year) as min_year, MAX(issue_year) as max_year FROM distributions WHERE quantity > 0', [], (err, row) => {
    if (err) return res.status(500).json({ error: err.message });
    res.json({
      min_year: row?.min_year || new Date().getFullYear(),
      max_year: row?.max_year || new Date().getFullYear()
    });
  });
});

// --- ИСПРАВЛЕННЫЙ ЭКСПОРТ В EXCEL (Материальные резервы) ---
app.get('/api/export-excel', (req, res) => {
  const { startDate, endDate, district_id } = req.query;
  
  let sql = `
    SELECT 
      d.name as "Район/Ведомство", 
      c.name as "Категория имущества", 
      i.name as "Наименование (Номенклатура)", 
      dist.issue_year as "Год выдачи",
      dist.quantity as "Количество", 
      i.unit as "Ед. изм.", 
      dist.total_cost as "Общая стоимость (руб)"
    FROM distributions dist
    JOIN districts d ON dist.district_id = d.id
    JOIN items i ON dist.item_id = i.id
    JOIN item_categories c ON i.category_id = c.id
    WHERE dist.quantity > 0
  `;
  
  const params = [];
  
  // Фильтр по датам (пока фильтруем по году, так как в базе лежит issue_year)
  if (startDate && endDate) {
    const startYear = parseInt(startDate.split('-')[0]);
    const endYear = parseInt(endDate.split('-')[0]);
    sql += " AND dist.issue_year BETWEEN ? AND ?";
    params.push(startYear, endYear);
  }

  // Фильтр по конкретному району (если выбран не "Все")
  if (district_id && district_id !== 'all') {
    sql += " AND dist.district_id = ?";
    params.push(district_id);
  }

  sql += " ORDER BY d.name ASC, c.name ASC, i.name ASC";

  db.all(sql, params, (err, rows) => {
    if (err) return res.status(500).send(err.message);
    if (rows.length === 0) return res.status(404).send('Нет данных для выгрузки за этот период/регион');

    try {
      // Формируем Excel
      const ws = xlsx.utils.json_to_sheet(rows);
      
      // Настраиваем ширину колонок для красоты
      ws['!cols'] = [
        { wch: 25 }, // Район
        { wch: 35 }, // Категория
        { wch: 30 }, // Наименование
        { wch: 12 }, // Год
        { wch: 12 }, // Кол-во
        { wch: 10 }, // Ед. изм.
        { wch: 20 }  // Стоимость
      ];

      const wb = xlsx.utils.book_new();
      xlsx.utils.book_append_sheet(wb, ws, "Материальные резервы");
      const buf = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
      
      // Кодируем имя файла, чтобы русские буквы не ломались в браузере
      const fileName = encodeURIComponent('Выдача_МЦ_Якутия.xlsx');
      
      res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${fileName}`);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(buf);
    } catch (excelErr) {
      res.status(500).send('Ошибка при создании Excel файла');
    }
  });
});

// 9. Скачать шаблон
app.get('/api/download-template', (req, res) => {
  const template = [{ district_id: 'yakutsk', indicator_id: 1, date: '2023-01-01', value: 330000, source: 'Росстат' }];
  const ws = xlsx.utils.json_to_sheet(template);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, "Template");
  const buf = xlsx.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Disposition', 'attachment; filename="template.xlsx"');
  res.send(buf);
});

app.get('/', (req, res) => res.json({ message: '🚀 API Yakutia Map Online' }));

app.listen(PORT, () => console.log(`✅ Server: http://localhost:${PORT}`));
