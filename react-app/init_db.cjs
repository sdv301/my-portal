// init_db.cjs
const sqlite3 = require('sqlite3').verbose();
const fs = require('fs');

const dbFile = 'data.sqlite';

if (fs.existsSync(dbFile)) {
  fs.unlinkSync(dbFile);
  console.log('Старая база данных удалена.');
}

const db = new sqlite3.Database(dbFile);

// Твой полный список статичных данных
const allDistricts = {
  'город Якутск': { id: 'yakutsk', name: 'Якутск', population: 330000, area_km2: 122, capital: 'Якутск', code: 'YAK' },
  'Жатай': { id: 'zhatay', name: 'Жатай', population: 11000, area_km2: 8, capital: 'Жатай', code: 'ZHT' },
  'Абыйский район': { id: 'abysky', name: 'Абыйский район', population: 4500, area_km2: 69300, capital: 'Белая Гора', code: 'ABY' },
  'Алданский район': { id: 'aldansky', name: 'Алданский район', population: 41000, area_km2: 156800, capital: 'Алдан', code: 'ALD' },
  'Аллаиховский район': { id: 'allaikhovsky', name: 'Аллаиховский район', population: 3000, area_km2: 107300, capital: 'Чокурдах', code: 'ALL' },
  'Амгинский район': { id: 'amginsky', name: 'Амгинский район', population: 17000, area_km2: 29300, capital: 'Амга', code: 'AMG' },
  'Анабарский район': { id: 'anabarsky', name: 'Анабарский район', population: 3400, area_km2: 55900, capital: 'Саскылах', code: 'ANA' },
  'Булунский район': { id: 'bulunsky', name: 'Булунский район', population: 8500, area_km2: 223600, capital: 'Тикси', code: 'BUL' },
  'Верхневилюйский район': { id: 'verkhnevilyuysky', name: 'Верхневилюйский район', population: 21000, area_km2: 42000, capital: 'Верхневилюйск', code: 'VVL' },
  'Верхнеколымский район': { id: 'verkhnekolymsky', name: 'Верхнеколымский район', population: 4300, area_km2: 67800, capital: 'Зырянка', code: 'VKL' },
  'Верхоянский район': { id: 'verkhoyansky', name: 'Верхоянский район', population: 11500, area_km2: 137400, capital: 'Батагай', code: 'VRY' },
  'Вилюйский район': { id: 'vilyuysky', name: 'Вилюйский район', population: 25000, area_km2: 55200, capital: 'Вилюйск', code: 'VIL' },
  'Горный район': { id: 'gorny', name: 'Горный район', population: 11000, area_km2: 45600, capital: 'Бердигестях', code: 'GRN' },
  'Жиганский район': { id: 'zhigansky', name: 'Жиганский район', population: 4200, area_km2: 140200, capital: 'Жиганск', code: 'ZHI' },
  'Кобяйский район': { id: 'kobyaysky', name: 'Кобяйский район', population: 13000, area_km2: 108000, capital: 'Сангар', code: 'KOB' },
  'Ленский район': { id: 'lensky', name: 'Ленский район', population: 39000, area_km2: 77000, capital: 'Ленск', code: 'LEN' },
  'Мегино-Кангаласский район': { id: 'megino-kangalassky', name: 'Мегино-Кангаласский район', population: 31000, area_km2: 11700, capital: 'Нижний Бестях', code: 'MEG' },
  'Мирнинский район': { id: 'mirninsky', name: 'Мирнинский район', population: 72000, area_km2: 165800, capital: 'Мирный', code: 'MIR' },
  'Момский район': { id: 'omsky', name: 'Момский район', population: 4500, area_km2: 104600, capital: 'Хонуу', code: 'MOM' },
  'Намский район': { id: 'namsky', name: 'Намский район', population: 25000, area_km2: 11900, capital: 'Намцы', code: 'NAM' },
  'Нерюнгринский район': { id: 'neryungrinsky', name: 'Нерюнгринский район', population: 75000, area_km2: 93000, capital: 'Нерюнгри', code: 'NER' },
  'Нижнеколымский район': { id: 'nizhnekolymsky', name: 'Нижнеколымский район', population: 4500, area_km2: 87600, capital: 'Черский', code: 'NKL' },
  'Нюрбинский район': { id: 'nyurbinsky', name: 'Нюрбинский район', population: 25000, area_km2: 52400, capital: 'Нюрба', code: 'NYU' },
  'Оймяконский район': { id: 'oymyakonsky', name: 'Оймяконский район', population: 10500, area_km2: 92000, capital: 'Усть-Нера', code: 'OYM' },
  'Олёкминский район': { id: 'olekminsky', name: 'Олёкминский район', population: 26000, area_km2: 161300, capital: 'Олёкминск', code: 'OLE' },
  'Оленёкский район': { id: 'olenek', name: 'Оленёкский район', population: 4300, area_km2: 318000, capital: 'Оленёк', code: 'OLK' },
  'Среднеколымский район': { id: 'srednekolymsky', name: 'Среднеколымский район', population: 7800, area_km2: 125200, capital: 'Среднеколымск', code: 'SKL' },
  'Сунтарский район': { id: 'suntarsky', name: 'Сунтарский район', population: 25000, area_km2: 57800, capital: 'Сунтар', code: 'SUN' },
  'Таттинский район': { id: 'tattinsky', name: 'Таттинский район', population: 17000, area_km2: 18900, capital: 'Ытык-Кюёль', code: 'TAT' },
  'Томпонский район': { id: 'tomponsky', name: 'Томпонский район', population: 13500, area_km2: 135800, capital: 'Хандыга', code: 'TOM' },
  'Усть-Алданский район': { id: 'ust-aldansky', name: 'Усть-Алданский район', population: 22000, area_km2: 18300, capital: 'Борогонцы', code: 'UAL' },
  'Усть-Майский район': { id: 'ust-maysky', name: 'Усть-Майский район', population: 8500, area_km2: 95300, capital: 'Усть-Мая', code: 'UMA' },
  'Усть-Янский район': { id: 'ust-yansky', name: 'Усть-Янский район', population: 7300, area_km2: 120300, capital: 'Депутатский', code: 'UYA' },
  'Хангаласский район': { id: 'khangalassky', name: 'Хангаласский район', population: 34000, area_km2: 24700, capital: 'Покровск', code: 'KHA' },
  'Чурапчинский район': { id: 'churapchinsky', name: 'Чурапчинский район', population: 20000, area_km2: 12600, capital: 'Чурапча', code: 'CHU' },
  'Эвено-Бытантайский район': { id: 'eveno-bytantaysky', name: 'Эвено-Бытантайский район', population: 2800, area_km2: 55300, capital: 'Батагай-Алыта', code: 'EVB' }
};

db.serialize(() => {
  // 1. Таблица районов и ведомств (Добавлены новые колонки для статики)
  db.run(`CREATE TABLE IF NOT EXISTS districts (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    population INTEGER,
    area_km2 REAL,
    capital TEXT,
    code TEXT,
    type TEXT DEFAULT 'district'
  )`);

  // 2. Таблица категорий
  db.run(`CREATE TABLE IF NOT EXISTS item_categories (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE
  )`);

  // 3. Таблица номенклатуры
  db.run(`CREATE TABLE IF NOT EXISTS items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    category_id INTEGER,
    name TEXT NOT NULL,
    unit TEXT,
    unit_price REAL,
    FOREIGN KEY (category_id) REFERENCES item_categories(id)
  )`);

  // 4. Таблица распределения
  db.run(`CREATE TABLE IF NOT EXISTS distributions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    district_id TEXT NOT NULL,
    item_id INTEGER NOT NULL,
    issue_year INTEGER,
    issue_date DATE,
    quantity REAL DEFAULT 0,
    total_cost REAL DEFAULT 0,
    FOREIGN KEY (district_id) REFERENCES districts(id),
    FOREIGN KEY (item_id) REFERENCES items(id)
  )`);

  // Заполняем районы
  const stmt = db.prepare("INSERT INTO districts (id, name, population, area_km2, capital, code, type) VALUES (?, ?, ?, ?, ?, ?, ?)");
  
  Object.values(allDistricts).forEach(d => {
    stmt.run(d.id, d.name, d.population, d.area_km2, d.capital, d.code, 'district');
  });

  // Ведомства, которые не отображаются на карте
  stmt.run('spas_rsy', 'Служба спасения РС(Я)', null, null, null, null, 'agency');
  stmt.run('gps_rsy', 'Государственная противопожарная служба РС(Я)', null, null, null, null, 'agency');
  stmt.finalize();

  console.log('✅ Новая база данных data.sqlite успешно создана со всеми статичными данными!');
});

db.close();