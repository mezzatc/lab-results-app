const express = require('express');
const session = require('express-session');
const pgSession = require('connect-pg-simple')(session);
const { Pool } = require('pg');
const bcrypt = require('bcryptjs');
const ExcelJS = require('exceljs');
const multer = require('multer');
const path = require('path');
const crypto = require('crypto');

const app = express();
const PORT = process.env.PORT || 3000;
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

// --- PostgreSQL Connection ---
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false
});

// --- Database Setup ---
async function initDB() {
  await pool.query(`
    CREATE TABLE IF NOT EXISTS users (
      id SERIAL PRIMARY KEY,
      username TEXT UNIQUE NOT NULL,
      display_name TEXT NOT NULL,
      password_hash TEXT NOT NULL,
      created_at TIMESTAMPTZ DEFAULT NOW()
    );

    CREATE TABLE IF NOT EXISTS data_dumps (
      id SERIAL PRIMARY KEY,
      name TEXT NOT NULL,
      entry_count INTEGER NOT NULL DEFAULT 0,
      created_by INTEGER NOT NULL REFERENCES users(id),
      created_at TIMESTAMPTZ DEFAULT NOW()
    );

    CREATE TABLE IF NOT EXISTS dump_entries (
      id SERIAL PRIMARY KEY,
      dump_id INTEGER NOT NULL REFERENCES data_dumps(id) ON DELETE CASCADE,
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
      original_created_at TEXT
    );

    CREATE TABLE IF NOT EXISTS entries (
      id SERIAL PRIMARY KEY,
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
      created_by INTEGER NOT NULL REFERENCES users(id),
      created_at TIMESTAMPTZ DEFAULT NOW(),
      updated_at TIMESTAMPTZ DEFAULT NOW()
    );
  `);

  // Create default admin if no users exist
  const { rows } = await pool.query('SELECT COUNT(*) as count FROM users');
  if (parseInt(rows[0].count) === 0) {
    const hash = bcrypt.hashSync('admin123', 10);
    await pool.query(
      'INSERT INTO users (username, display_name, password_hash) VALUES ($1, $2, $3)',
      ['admin', 'Administrator', hash]
    );
    console.log('Default admin account created (username: admin, password: admin123)');
  }
}

// --- Middleware ---
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(session({
  store: new pgSession({ pool, createTableIfMissing: true }),
  secret: process.env.SESSION_SECRET || crypto.randomBytes(32).toString('hex'),
  resave: false,
  saveUninitialized: false,
  cookie: { maxAge: 7 * 24 * 60 * 60 * 1000 }
}));

app.use(express.static(path.join(__dirname, 'public')));

function requireAuth(req, res, next) {
  if (!req.session.userId) return res.status(401).json({ error: 'Not authenticated' });
  next();
}

// --- Auth Routes ---
app.post('/api/login', async (req, res) => {
  const { username, password } = req.body;
  const { rows } = await pool.query('SELECT * FROM users WHERE username = $1', [username]);
  const user = rows[0];
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

app.get('/api/me', async (req, res) => {
  if (!req.session.userId) return res.status(401).json({ error: 'Not authenticated' });
  const { rows } = await pool.query('SELECT id, username, display_name FROM users WHERE id = $1', [req.session.userId]);
  if (!rows[0]) return res.status(401).json({ error: 'Not authenticated' });
  res.json({ id: rows[0].id, username: rows[0].username, displayName: rows[0].display_name });
});

app.post('/api/register', async (req, res) => {
  const { username, displayName, password } = req.body;
  if (!username || !displayName || !password)
    return res.status(400).json({ error: 'All fields are required' });
  if (password.length < 6)
    return res.status(400).json({ error: 'Password must be at least 6 characters' });

  const existing = await pool.query('SELECT id FROM users WHERE username = $1', [username]);
  if (existing.rows.length > 0) return res.status(409).json({ error: 'Username already taken' });

  const hash = bcrypt.hashSync(password, 10);
  const { rows } = await pool.query(
    'INSERT INTO users (username, display_name, password_hash) VALUES ($1, $2, $3) RETURNING id',
    [username, displayName, hash]
  );
  req.session.userId = rows[0].id;
  req.session.displayName = displayName;
  res.json({ id: rows[0].id, username, displayName });
});

// --- Entry CRUD ---
app.get('/api/entries', requireAuth, async (req, res) => {
  const { rows } = await pool.query(`
    SELECT e.*, u.display_name as created_by_name
    FROM entries e JOIN users u ON e.created_by = u.id
    ORDER BY e.created_at DESC
  `);
  res.json(rows);
});

app.post('/api/entries', requireAuth, async (req, res) => {
  const { ratId, age, sex, weight, strain, dietGroup, drugName, dose, route, brainRegion, notes } = req.body;
  if (!ratId || !age || !sex || !weight || !strain || !dietGroup || !drugName || !dose || !route || !brainRegion)
    return res.status(400).json({ error: 'Missing required fields' });

  const { rows } = await pool.query(`
    INSERT INTO entries (rat_id, age, sex, weight, strain, diet_group, drug_name, dose, route, brain_region, notes, created_by)
    VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12) RETURNING id
  `, [ratId, age, sex, weight, strain, dietGroup, drugName, dose, route, brainRegion, notes || '', req.session.userId]);

  const { rows: entry } = await pool.query(`
    SELECT e.*, u.display_name as created_by_name
    FROM entries e JOIN users u ON e.created_by = u.id WHERE e.id = $1
  `, [rows[0].id]);
  res.json(entry[0]);
});

app.put('/api/entries/:id', requireAuth, async (req, res) => {
  const { ratId, age, sex, weight, strain, dietGroup, drugName, dose, route, brainRegion, notes } = req.body;
  await pool.query(`
    UPDATE entries SET rat_id=$1, age=$2, sex=$3, weight=$4, strain=$5, diet_group=$6,
    drug_name=$7, dose=$8, route=$9, brain_region=$10, notes=$11, updated_at=NOW()
    WHERE id=$12
  `, [ratId, age, sex, weight, strain, dietGroup, drugName, dose, route, brainRegion, notes || '', req.params.id]);

  const { rows } = await pool.query(`
    SELECT e.*, u.display_name as created_by_name
    FROM entries e JOIN users u ON e.created_by = u.id WHERE e.id = $1
  `, [req.params.id]);
  res.json(rows[0]);
});

app.delete('/api/entries/:id', requireAuth, async (req, res) => {
  await pool.query('DELETE FROM entries WHERE id = $1', [req.params.id]);
  res.json({ ok: true });
});

// --- Excel Export ---
app.get('/api/export/excel', requireAuth, async (req, res) => {
  const { rows } = await pool.query(`
    SELECT e.*, u.display_name as created_by_name
    FROM entries e JOIN users u ON e.created_by = u.id ORDER BY e.created_at DESC
  `);

  const workbook = new ExcelJS.Workbook();
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
  sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF16213E' } };
  sheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
  rows.forEach(e => sheet.addRow(e));
  for (let i = 2; i <= rows.length + 1; i++) {
    if (i % 2 === 0) sheet.getRow(i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF7F8FC' } };
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
    sheet.getRow(1).eachCell((cell, col) => {
      const val = String(cell.value || '').toLowerCase().trim();
      for (const [field, aliases] of Object.entries(columnAliases)) {
        if (aliases.includes(val)) { headerMap[field] = col; break; }
      }
    });
    const required = ['rat_id', 'age', 'sex', 'weight', 'strain', 'diet_group', 'drug_name', 'dose', 'route', 'brain_region'];
    const missing = required.filter(f => !(f in headerMap));
    if (missing.length > 0) {
      const names = { rat_id: 'Rat ID', age: 'Age', sex: 'Sex', weight: 'Weight', strain: 'Strain', diet_group: 'Diet Group', drug_name: 'Drug', dose: 'Dose', route: 'Route', brain_region: 'Brain Region' };
      return res.status(400).json({ error: `Missing required columns: ${missing.map(f => names[f]).join(', ')}` });
    }

    let imported = 0;
    const errors = [];
    for (let i = 2; i <= sheet.rowCount; i++) {
      const row = sheet.getRow(i);
      const get = (f) => { const v = row.getCell(headerMap[f] || 0).value; return v == null ? '' : String(v).trim(); };
      const ratId = get('rat_id');
      if (!ratId) continue;
      const age = parseInt(get('age')), weight = parseFloat(get('weight'));
      if (isNaN(age) || isNaN(weight)) { errors.push(`Row ${i}: invalid age or weight (skipped)`); continue; }
      try {
        await pool.query(
          `INSERT INTO entries (rat_id,age,sex,weight,strain,diet_group,drug_name,dose,route,brain_region,notes,created_by) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12)`,
          [ratId, age, get('sex'), weight, get('strain'), get('diet_group'), get('drug_name'), get('dose'), get('route'), get('brain_region'), get('notes'), req.session.userId]
        );
        imported++;
      } catch (e) { errors.push(`Row ${i}: ${e.message}`); }
    }
    res.json({ imported, errors, total: sheet.rowCount - 1 });
  } catch (e) { res.status(400).json({ error: 'Failed to parse file: ' + e.message }); }
});

// --- Import Template ---
app.get('/api/import/template', requireAuth, async (req, res) => {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Lab Results Template');
  sheet.columns = [
    { header: 'Rat ID', key: 'rat_id', width: 12 }, { header: 'Age (weeks)', key: 'age', width: 12 },
    { header: 'Sex', key: 'sex', width: 8 }, { header: 'Weight (g)', key: 'weight', width: 12 },
    { header: 'Strain', key: 'strain', width: 16 }, { header: 'Diet Group', key: 'diet_group', width: 14 },
    { header: 'Drug/Compound', key: 'drug_name', width: 18 }, { header: 'Dose', key: 'dose', width: 14 },
    { header: 'Route', key: 'route', width: 10 }, { header: 'Brain Region', key: 'brain_region', width: 18 },
    { header: 'Notes', key: 'notes', width: 24 },
  ];
  sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF16213E' } };
  sheet.addRow({ rat_id: 'R-001', age: 12, sex: 'Male', weight: 320.5, strain: 'Sprague Dawley', diet_group: 'Obesogenic', drug_name: 'Liraglutide', dose: '0.2 mg/kg', route: 'SC', brain_region: 'Hypothalamus', notes: 'Example entry' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename=import_template.xlsx');
  await workbook.xlsx.write(res);
  res.end();
});

// --- Data Dumps ---
app.get('/api/dumps', requireAuth, async (req, res) => {
  const { rows } = await pool.query(`
    SELECT d.*, u.display_name as created_by_name
    FROM data_dumps d JOIN users u ON d.created_by = u.id ORDER BY d.created_at DESC
  `);
  res.json(rows);
});

app.post('/api/dumps', requireAuth, async (req, res) => {
  const { name } = req.body;
  if (!name) return res.status(400).json({ error: 'Dump name is required' });

  const { rows: entries } = await pool.query(`
    SELECT e.*, u.display_name as created_by_name FROM entries e JOIN users u ON e.created_by = u.id
  `);
  if (entries.length === 0) return res.status(400).json({ error: 'No entries to store' });

  const { rows: dump } = await pool.query(
    'INSERT INTO data_dumps (name, entry_count, created_by) VALUES ($1,$2,$3) RETURNING id',
    [name, entries.length, req.session.userId]
  );
  const dumpId = dump[0].id;

  for (const e of entries) {
    await pool.query(
      `INSERT INTO dump_entries (dump_id,rat_id,age,sex,weight,strain,diet_group,drug_name,dose,route,brain_region,notes,original_created_by,original_created_at)
       VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14)`,
      [dumpId, e.rat_id, e.age, e.sex, e.weight, e.strain, e.diet_group, e.drug_name, e.dose, e.route, e.brain_region, e.notes, e.created_by_name, e.created_at]
    );
  }

  const { rows: result } = await pool.query(`
    SELECT d.*, u.display_name as created_by_name FROM data_dumps d JOIN users u ON d.created_by = u.id WHERE d.id = $1
  `, [dumpId]);
  res.json(result[0]);
});

app.get('/api/dumps/:id/entries', requireAuth, async (req, res) => {
  const { rows } = await pool.query('SELECT * FROM dump_entries WHERE dump_id = $1 ORDER BY id', [req.params.id]);
  res.json(rows);
});

app.delete('/api/dumps/:id', requireAuth, async (req, res) => {
  await pool.query('DELETE FROM data_dumps WHERE id = $1', [req.params.id]);
  res.json({ ok: true });
});

app.get('/api/dumps/:id/export', requireAuth, async (req, res) => {
  const { rows: dumps } = await pool.query('SELECT * FROM data_dumps WHERE id = $1', [req.params.id]);
  if (!dumps[0]) return res.status(404).json({ error: 'Dump not found' });
  const dump = dumps[0];
  const { rows: entries } = await pool.query('SELECT * FROM dump_entries WHERE dump_id = $1 ORDER BY id', [req.params.id]);

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet(dump.name);
  sheet.columns = [
    { header: 'Rat ID', key: 'rat_id', width: 12 }, { header: 'Age (weeks)', key: 'age', width: 12 },
    { header: 'Sex', key: 'sex', width: 8 }, { header: 'Weight (g)', key: 'weight', width: 12 },
    { header: 'Strain', key: 'strain', width: 16 }, { header: 'Diet Group', key: 'diet_group', width: 14 },
    { header: 'Drug/Compound', key: 'drug_name', width: 18 }, { header: 'Dose', key: 'dose', width: 14 },
    { header: 'Route', key: 'route', width: 10 }, { header: 'Brain Region', key: 'brain_region', width: 18 },
    { header: 'Notes', key: 'notes', width: 24 }, { header: 'Entered By', key: 'original_created_by', width: 16 },
    { header: 'Date', key: 'original_created_at', width: 18 },
  ];
  sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF16213E' } };
  sheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };
  entries.forEach(e => sheet.addRow(e));
  for (let i = 2; i <= entries.length + 1; i++) {
    if (i % 2 === 0) sheet.getRow(i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF7F8FC' } };
  }
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename=${dump.name.replace(/[^a-zA-Z0-9]/g, '_')}_${new Date().toISOString().slice(0,10)}.xlsx`);
  await workbook.xlsx.write(res);
  res.end();
});

// --- Chart Data ---
app.get('/api/charts/weight-by-diet', requireAuth, async (req, res) => {
  const { rows } = await pool.query(`SELECT diet_group, AVG(weight) as avg_weight, MIN(weight) as min_weight, MAX(weight) as max_weight, COUNT(*) as count FROM entries GROUP BY diet_group`);
  res.json(rows);
});
app.get('/api/charts/weight-distribution', requireAuth, async (req, res) => {
  const { rows } = await pool.query(`SELECT diet_group, weight FROM entries ORDER BY diet_group, weight`);
  res.json(rows);
});
app.get('/api/charts/brain-regions', requireAuth, async (req, res) => {
  const { rows } = await pool.query(`SELECT brain_region, diet_group, COUNT(*) as count FROM entries GROUP BY brain_region, diet_group ORDER BY brain_region`);
  res.json(rows);
});
app.get('/api/charts/drugs', requireAuth, async (req, res) => {
  const { rows } = await pool.query(`SELECT drug_name, COUNT(*) as count FROM entries GROUP BY drug_name ORDER BY count DESC`);
  res.json(rows);
});
app.get('/api/charts/sex-distribution', requireAuth, async (req, res) => {
  const { rows } = await pool.query(`SELECT sex, diet_group, COUNT(*) as count FROM entries GROUP BY sex, diet_group`);
  res.json(rows);
});
app.get('/api/charts/weight-by-age', requireAuth, async (req, res) => {
  const { rows } = await pool.query(`SELECT age, weight, diet_group, rat_id FROM entries ORDER BY age`);
  res.json(rows);
});

// --- Start ---
initDB()
  .then(() => {
    app.listen(PORT, '0.0.0.0', () => {
      console.log(`\n  Lab Results App running at http://localhost:${PORT}\n`);
    });
  })
  .catch(err => {
    console.error('Failed to initialize database:', err);
    process.exit(1);
  });
