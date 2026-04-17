const express = require('express');
const session = require('express-session');
const SQLiteStore = require('connect-sqlite3')(session);
const Database = require('better-sqlite3');
const bcrypt = require('bcryptjs');
const ExcelJS = require('exceljs');
const path = require('path');
const crypto = require('crypto');

const multer = require('multer');

const app = express();
const PORT = process.env.PORT || 3000;
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

// --- Database Setup ---
const db = new Database(path.join(__dirname, 'lab_data.db'));
db.pragma('journal_mode = WAL');
db.pragma('foreign_keys = ON');

db.exec(`
  CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    display_name TEXT NOT NULL,
    password_hash TEXT NOT NULL,
    created_at TEXT DEFAULT (datetime('now'))
  );

  CREATE TABLE IF NOT EXISTS data_dumps (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    entry_count INTEGER NOT NULL DEFAULT 0,
    created_by INTEGER NOT NULL,
    created_at TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (created_by) REFERENCES users(id)
  );

  CREATE TABLE IF NOT EXISTS dump_entries (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    dump_id INTEGER NOT NULL,
    rat_id TEXT NOT NULL,
    age INTEGER NOT NULL,
    sex TEXT NOT NULL,
    weight REAL NOT NULL,
    strain TEXT NOT NULL,
    diet_group TEXT NOT NULL,
    drug_name TEXT NOT NULL,
    dose TEXT NOT NULL,
    route TEXT NOT NULL,
    brain_region TEXT NOT NULL,
    notes TEXT,
    original_created_by TEXT,
    original_created_at TEXT,
    FOREIGN KEY (dump_id) REFERENCES data_dumps(id) ON DELETE CASCADE
  );

  CREATE TABLE IF NOT EXISTS entries (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    rat_id TEXT NOT NULL,
    age INTEGER NOT NULL,
    sex TEXT NOT NULL,
    weight REAL NOT NULL,
    strain TEXT NOT NULL,
    diet_group TEXT NOT NULL,
    drug_name TEXT NOT NULL,
    dose TEXT NOT NULL,
    route TEXT NOT NULL,
    brain_region TEXT NOT NULL,
    notes TEXT,
    created_by INTEGER NOT NULL,
    created_at TEXT DEFAULT (datetime('now')),
    updated_at TEXT DEFAULT (datetime('now')),
    FOREIGN KEY (created_by) REFERENCES users(id)
  );
`);

// Create default admin account if no users exist
const userCount = db.prepare('SELECT COUNT(*) as count FROM users').get().count;
if (userCount === 0) {
  const hash = bcrypt.hashSync('admin123', 10);
  db.prepare('INSERT INTO users (username, display_name, password_hash) VALUES (?, ?, ?)')
    .run('admin', 'Administrator', hash);
  console.log('Default admin account created (username: admin, password: admin123)');
}

// --- Middleware ---
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(session({
  store: new SQLiteStore({ dir: __dirname, db: 'sessions.db' }),
  secret: crypto.randomBytes(32).toString('hex'),
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 7 * 24 * 60 * 60 * 1000 } // 7 days
}));

app.use(express.static(path.join(__dirname, 'public')));

function requireAuth(req, res, next) {
  if (!req.session.userId) return res.status(401).json({ error: 'Not authenticated' });
  next();
}

// --- Auth Routes ---
app.post('/api/login', (req, res) => {
  const { username, password } = req.body;
  const user = db.prepare('SELECT * FROM users WHERE username = ?').get(username);
  if (!user || !bcrypt.compareSync(password, user.password_hash)) {
    return res.status(401).json({ error: 'Invalid credentials' });
  }
  req.session.userId = user.id;
  req.session.displayName = user.display_name;
  res.json({ id: user.id, username: user.username, displayName: user.display_name });
});

app.post('/api/logout', (req, res) => {
  req.session.destroy();
  res.json({ ok: true });
});

app.get('/api/me', (req, res) => {
  if (!req.session.userId) return res.status(401).json({ error: 'Not authenticated' });
  const user = db.prepare('SELECT id, username, display_name FROM users WHERE id = ?').get(req.session.userId);
  if (!user) return res.status(401).json({ error: 'Not authenticated' });
  res.json({ id: user.id, username: user.username, displayName: user.display_name });
});

app.post('/api/register', (req, res) => {
  const { username, displayName, password } = req.body;
  if (!username || !displayName || !password) {
    return res.status(400).json({ error: 'All fields are required' });
  }
  if (password.length < 6) {
    return res.status(400).json({ error: 'Password must be at least 6 characters' });
  }
  const existing = db.prepare('SELECT id FROM users WHERE username = ?').get(username);
  if (existing) return res.status(409).json({ error: 'Username already taken' });

  const hash = bcrypt.hashSync(password, 10);
  const result = db.prepare('INSERT INTO users (username, display_name, password_hash) VALUES (?, ?, ?)')
    .run(username, displayName, hash);
  req.session.userId = result.lastInsertRowid;
  req.session.displayName = displayName;
  res.json({ id: result.lastInsertRowid, username, displayName });
});

// --- Entry CRUD ---
app.get('/api/entries', requireAuth, (req, res) => {
  const entries = db.prepare(`
    SELECT e.*, u.display_name as created_by_name
    FROM entries e JOIN users u ON e.created_by = u.id
    ORDER BY e.created_at DESC
  `).all();
  res.json(entries);
});

app.post('/api/entries', requireAuth, (req, res) => {
  const { ratId, age, sex, weight, strain, dietGroup, drugName, dose, route, brainRegion, notes } = req.body;
  if (!ratId || !age || !sex || !weight || !strain || !dietGroup || !drugName || !dose || !route || !brainRegion) {
    return res.status(400).json({ error: 'Missing required fields' });
  }
  const result = db.prepare(`
    INSERT INTO entries (rat_id, age, sex, weight, strain, diet_group, drug_name, dose, route, brain_region, notes, created_by)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).run(ratId, age, sex, weight, strain, dietGroup, drugName, dose, route, brainRegion, notes || '', req.session.userId);

  const entry = db.prepare(`
    SELECT e.*, u.display_name as created_by_name
    FROM entries e JOIN users u ON e.created_by = u.id WHERE e.id = ?
  `).get(result.lastInsertRowid);
  res.json(entry);
});

app.put('/api/entries/:id', requireAuth, (req, res) => {
  const { ratId, age, sex, weight, strain, dietGroup, drugName, dose, route, brainRegion, notes } = req.body;
  db.prepare(`
    UPDATE entries SET rat_id=?, age=?, sex=?, weight=?, strain=?, diet_group=?, drug_name=?, dose=?, route=?, brain_region=?, notes=?, updated_at=datetime('now')
    WHERE id=?
  `).run(ratId, age, sex, weight, strain, dietGroup, drugName, dose, route, brainRegion, notes || '', req.params.id);

  const entry = db.prepare(`
    SELECT e.*, u.display_name as created_by_name
    FROM entries e JOIN users u ON e.created_by = u.id WHERE e.id = ?
  `).get(req.params.id);
  res.json(entry);
});

app.delete('/api/entries/:id', requireAuth, (req, res) => {
  db.prepare('DELETE FROM entries WHERE id = ?').run(req.params.id);
  res.json({ ok: true });
});

// --- Excel Export ---
app.get('/api/export/excel', requireAuth, async (req, res) => {
  const entries = db.prepare(`
    SELECT e.*, u.display_name as created_by_name
    FROM entries e JOIN users u ON e.created_by = u.id
    ORDER BY e.created_at DESC
  `).all();

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'Obesogenic Diet Brain Study';
  workbook.created = new Date();

  const sheet = workbook.addWorksheet('Lab Results');

  sheet.columns = [
    { header: 'Rat ID', key: 'rat_id', width: 12 },
    { header: 'Age (weeks)', key: 'age', width: 12 },
    { header: 'Sex', key: 'sex', width: 8 },
    { header: 'Weight (g)', key: 'weight', width: 12 },
    { header: 'Strain', key: 'strain', width: 16 },
    { header: 'Diet Group', key: 'diet_group', width: 14 },
    { header: 'Drug/Compound', key: 'drug_name', width: 18 },
    { header: 'Dose', key: 'dose', width: 14 },
    { header: 'Route', key: 'route', width: 10 },
    { header: 'Brain Region', key: 'brain_region', width: 18 },
    { header: 'Notes', key: 'notes', width: 24 },
    { header: 'Entered By', key: 'created_by_name', width: 16 },
    { header: 'Date', key: 'created_at', width: 18 },
  ];

  // Style header row
  sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF16213E' } };
  sheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

  entries.forEach(e => sheet.addRow(e));

  // Alternate row colors
  for (let i = 2; i <= entries.length + 1; i++) {
    if (i % 2 === 0) {
      sheet.getRow(i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF7F8FC' } };
    }
  }

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename=lab_results_${new Date().toISOString().slice(0,10)}.xlsx`);
  await workbook.xlsx.write(res);
  res.end();
});

// --- Excel Import ---
app.post('/api/import/excel', requireAuth, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(req.file.buffer);
    const sheet = workbook.worksheets[0];
    if (!sheet) return res.status(400).json({ error: 'No worksheet found in file' });

    // Build a header map from the first row (flexible matching)
    const headerMap = {};
    const columnAliases = {
      rat_id: ['rat id', 'rat_id', 'ratid', 'subject id', 'subject_id', 'id'],
      age: ['age', 'age (weeks)', 'age_weeks', 'age weeks'],
      sex: ['sex', 'gender'],
      weight: ['weight', 'weight (g)', 'weight_g', 'weight grams', 'body weight'],
      strain: ['strain', 'rat strain'],
      diet_group: ['diet group', 'diet_group', 'dietgroup', 'diet', 'group'],
      drug_name: ['drug', 'drug name', 'drug_name', 'drugname', 'compound', 'drug / compound', 'drug/compound'],
      dose: ['dose', 'dosage'],
      route: ['route', 'route of administration', 'administration route', 'roa'],
      brain_region: ['brain region', 'brain_region', 'brainregion', 'region'],
      notes: ['notes', 'note', 'comments', 'comment'],
    };

    const firstRow = sheet.getRow(1);
    firstRow.eachCell((cell, colNumber) => {
      const val = String(cell.value || '').toLowerCase().trim();
      for (const [field, aliases] of Object.entries(columnAliases)) {
        if (aliases.includes(val)) {
          headerMap[field] = colNumber;
          break;
        }
      }
    });

    // Check required columns
    const required = ['rat_id', 'age', 'sex', 'weight', 'strain', 'diet_group', 'drug_name', 'dose', 'route', 'brain_region'];
    const missing = required.filter(f => !(f in headerMap));
    if (missing.length > 0) {
      const friendlyNames = { rat_id: 'Rat ID', age: 'Age', sex: 'Sex', weight: 'Weight', strain: 'Strain', diet_group: 'Diet Group', drug_name: 'Drug', dose: 'Dose', route: 'Route', brain_region: 'Brain Region' };
      return res.status(400).json({
        error: `Missing required columns: ${missing.map(f => friendlyNames[f]).join(', ')}. Check that your header row matches the expected column names.`
      });
    }

    const insertStmt = db.prepare(`
      INSERT INTO entries (rat_id, age, sex, weight, strain, diet_group, drug_name, dose, route, brain_region, notes, created_by)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);

    let imported = 0;
    const errors = [];

    for (let i = 2; i <= sheet.rowCount; i++) {
      const row = sheet.getRow(i);
      const getCellVal = (field) => {
        const col = headerMap[field];
        if (!col) return '';
        const val = row.getCell(col).value;
        if (val === null || val === undefined) return '';
        return String(val).trim();
      };

      const ratId = getCellVal('rat_id');
      if (!ratId) continue; // skip empty rows

      const age = parseInt(getCellVal('age'));
      const weight = parseFloat(getCellVal('weight'));
      const sex = getCellVal('sex');
      const strain = getCellVal('strain');
      const dietGroup = getCellVal('diet_group');
      const drugName = getCellVal('drug_name');
      const dose = getCellVal('dose');
      const route = getCellVal('route');
      const brainRegion = getCellVal('brain_region');
      const notes = getCellVal('notes');

      if (isNaN(age) || isNaN(weight) || !sex || !strain || !dietGroup || !drugName || !dose || !route || !brainRegion) {
        errors.push(`Row ${i}: missing or invalid data (skipped)`);
        continue;
      }

      try {
        insertStmt.run(ratId, age, sex, weight, strain, dietGroup, drugName, dose, route, brainRegion, notes, req.session.userId);
        imported++;
      } catch (e) {
        errors.push(`Row ${i}: ${e.message}`);
      }
    }

    res.json({ imported, errors, total: sheet.rowCount - 1 });
  } catch (e) {
    res.status(400).json({ error: 'Failed to parse Excel file: ' + e.message });
  }
});

// --- Import Template ---
app.get('/api/import/template', requireAuth, async (req, res) => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Lab Results Template');

  sheet.columns = [
    { header: 'Rat ID', key: 'rat_id', width: 12 },
    { header: 'Age (weeks)', key: 'age', width: 12 },
    { header: 'Sex', key: 'sex', width: 8 },
    { header: 'Weight (g)', key: 'weight', width: 12 },
    { header: 'Strain', key: 'strain', width: 16 },
    { header: 'Diet Group', key: 'diet_group', width: 14 },
    { header: 'Drug/Compound', key: 'drug_name', width: 18 },
    { header: 'Dose', key: 'dose', width: 14 },
    { header: 'Route', key: 'route', width: 10 },
    { header: 'Brain Region', key: 'brain_region', width: 18 },
    { header: 'Notes', key: 'notes', width: 24 },
  ];

  sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF16213E' } };

  // Add one example row
  sheet.addRow({ rat_id: 'R-001', age: 12, sex: 'Male', weight: 320.5, strain: 'Sprague Dawley', diet_group: 'Obesogenic', drug_name: 'Liraglutide', dose: '0.2 mg/kg', route: 'SC', brain_region: 'Hypothalamus', notes: 'Example entry' });

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=import_template.xlsx');
  await workbook.xlsx.write(res);
  res.end();
});

// --- Data Dumps ---
app.get('/api/dumps', requireAuth, (req, res) => {
  const dumps = db.prepare(`
    SELECT d.*, u.display_name as created_by_name
    FROM data_dumps d JOIN users u ON d.created_by = u.id
    ORDER BY d.created_at DESC
  `).all();
  res.json(dumps);
});

app.post('/api/dumps', requireAuth, (req, res) => {
  const { name } = req.body;
  if (!name) return res.status(400).json({ error: 'Dump name is required' });

  const entries = db.prepare(`
    SELECT e.*, u.display_name as created_by_name
    FROM entries e JOIN users u ON e.created_by = u.id
  `).all();

  if (entries.length === 0) return res.status(400).json({ error: 'No entries to store' });

  const result = db.prepare('INSERT INTO data_dumps (name, entry_count, created_by) VALUES (?, ?, ?)')
    .run(name, entries.length, req.session.userId);

  const insertDumpEntry = db.prepare(`
    INSERT INTO dump_entries (dump_id, rat_id, age, sex, weight, strain, diet_group, drug_name, dose, route, brain_region, notes, original_created_by, original_created_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  for (const e of entries) {
    insertDumpEntry.run(result.lastInsertRowid, e.rat_id, e.age, e.sex, e.weight, e.strain, e.diet_group, e.drug_name, e.dose, e.route, e.brain_region, e.notes, e.created_by_name, e.created_at);
  }

  const dump = db.prepare(`
    SELECT d.*, u.display_name as created_by_name
    FROM data_dumps d JOIN users u ON d.created_by = u.id WHERE d.id = ?
  `).get(result.lastInsertRowid);
  res.json(dump);
});

app.get('/api/dumps/:id/entries', requireAuth, (req, res) => {
  const entries = db.prepare('SELECT * FROM dump_entries WHERE dump_id = ? ORDER BY id').all(req.params.id);
  res.json(entries);
});

app.delete('/api/dumps/:id', requireAuth, (req, res) => {
  db.prepare('DELETE FROM dump_entries WHERE dump_id = ?').run(req.params.id);
  db.prepare('DELETE FROM data_dumps WHERE id = ?').run(req.params.id);
  res.json({ ok: true });
});

app.get('/api/dumps/:id/export', requireAuth, async (req, res) => {
  const dump = db.prepare('SELECT * FROM data_dumps WHERE id = ?').get(req.params.id);
  if (!dump) return res.status(404).json({ error: 'Dump not found' });

  const entries = db.prepare('SELECT * FROM dump_entries WHERE dump_id = ? ORDER BY id').all(req.params.id);

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(dump.name);

  sheet.columns = [
    { header: 'Rat ID', key: 'rat_id', width: 12 },
    { header: 'Age (weeks)', key: 'age', width: 12 },
    { header: 'Sex', key: 'sex', width: 8 },
    { header: 'Weight (g)', key: 'weight', width: 12 },
    { header: 'Strain', key: 'strain', width: 16 },
    { header: 'Diet Group', key: 'diet_group', width: 14 },
    { header: 'Drug/Compound', key: 'drug_name', width: 18 },
    { header: 'Dose', key: 'dose', width: 14 },
    { header: 'Route', key: 'route', width: 10 },
    { header: 'Brain Region', key: 'brain_region', width: 18 },
    { header: 'Notes', key: 'notes', width: 24 },
    { header: 'Entered By', key: 'original_created_by', width: 16 },
    { header: 'Date', key: 'original_created_at', width: 18 },
  ];

  sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF16213E' } };
  sheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
  entries.forEach(e => sheet.addRow(e));

  for (let i = 2; i <= entries.length + 1; i++) {
    if (i % 2 === 0) {
      sheet.getRow(i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF7F8FC' } };
    }
  }

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename=${dump.name.replace(/[^a-zA-Z0-9]/g, '_')}_${new Date().toISOString().slice(0,10)}.xlsx`);
  await workbook.xlsx.write(res);
  res.end();
});

// --- Chart Data Aggregation ---
app.get('/api/charts/weight-by-diet', requireAuth, (req, res) => {
  const data = db.prepare(`
    SELECT diet_group, AVG(weight) as avg_weight, MIN(weight) as min_weight, MAX(weight) as max_weight, COUNT(*) as count
    FROM entries GROUP BY diet_group
  `).all();
  res.json(data);
});

app.get('/api/charts/weight-distribution', requireAuth, (req, res) => {
  const data = db.prepare(`
    SELECT diet_group, weight FROM entries ORDER BY diet_group, weight
  `).all();
  res.json(data);
});

app.get('/api/charts/brain-regions', requireAuth, (req, res) => {
  const data = db.prepare(`
    SELECT brain_region, diet_group, COUNT(*) as count
    FROM entries GROUP BY brain_region, diet_group ORDER BY brain_region
  `).all();
  res.json(data);
});

app.get('/api/charts/drugs', requireAuth, (req, res) => {
  const data = db.prepare(`
    SELECT drug_name, COUNT(*) as count FROM entries GROUP BY drug_name ORDER BY count DESC
  `).all();
  res.json(data);
});

app.get('/api/charts/sex-distribution', requireAuth, (req, res) => {
  const data = db.prepare(`
    SELECT sex, diet_group, COUNT(*) as count FROM entries GROUP BY sex, diet_group
  `).all();
  res.json(data);
});

app.get('/api/charts/weight-by-age', requireAuth, (req, res) => {
  const data = db.prepare(`
    SELECT age, weight, diet_group, rat_id FROM entries ORDER BY age
  `).all();
  res.json(data);
});

// --- Start ---
app.listen(PORT, '0.0.0.0', () => {
  console.log(`\n  Lab Results App running at http://localhost:${PORT}\n`);
});
