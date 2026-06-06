import Fastify from 'fastify';
import fastifyStatic from '@fastify/static';
import fastifyMultipart from '@fastify/multipart';
import pg from 'pg';
import XLSX from 'xlsx';
import { fileURLToPath } from 'url';
import path from 'path';

const { Pool } = pg;

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = Fastify({ logger: true });
const port = Number(process.env.PORT || 3000);
const host = '0.0.0.0';

const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false
});

const FIELD_ALIASES = {
  jobNumber: ['Job Nr', 'Job No', 'Job Number', 'Job', 'Job Nr.', 'Job #'],
  invoiceNumber: ['Invoice No', 'Invoice Number', 'Invoice', 'Invoice No.'],
  jobDate: ['Date', 'Job Date'],
  division: ['Division'],
  opsManager: ['OPS Manager', 'OPS manager', 'Operations Manager', 'Ops Manager'],
  hours: ['Total Hours', 'Hours', 'Hour'],
  revenue: ['Revenue Excl VAT', 'Revenue Excl', 'Revenue', 'Revenue Ex VAT', 'Revenue excluding VAT'],
  labourCost: ['Labour Costs', 'Labour Cost', 'Cost Labour', 'Labor Costs', 'Labor Cost'],
  equipmentCost: ['Equipment Costs', 'Equipment Cost', 'Cost Equipment'],
  workshopCost: ['Workshop Cost', 'Workshop Costs'],
  totalCost: ['Total Costs', 'Total Cost']
};

function normalizeHeader(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '');
}

function headerScore(headers) {
  const normalized = headers.map(normalizeHeader);
  const aliasGroups = Object.values(FIELD_ALIASES);
  let score = 0;
  for (const aliases of aliasGroups) {
    if (aliases.some((alias) => normalized.includes(normalizeHeader(alias)))) score += 1;
  }
  return score;
}

function findHeaderRow(sheet) {
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
  let best = { rowIndex: 0, score: 0, headers: [] };
  matrix.slice(0, 50).forEach((row, rowIndex) => {
    const score = headerScore(row);
    if (score > best.score) best = { rowIndex, score, headers: row };
  });
  return best;
}

function chooseBestSheet(workbook) {
  let best = null;
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const header = findHeaderRow(sheet);
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
    const nonEmptyRows = rows.filter((row) => row.some((cell) => String(cell || '').trim() !== '')).length;
    const result = { sheetName, sheet, header, nonEmptyRows };
    if (!best || header.score > best.header.score || (header.score === best.header.score && nonEmptyRows > best.nonEmptyRows)) best = result;
  }
  return best;
}

function toNumber(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const cleaned = String(value).replace(/[^0-9.-]/g, '');
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function toDate(value) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      const month = String(parsed.m).padStart(2, '0');
      const day = String(parsed.d).padStart(2, '0');
      return `${parsed.y}-${month}-${day}`;
    }
  }
  const date = new Date(value);
  if (!Number.isNaN(date.getTime())) return date.toISOString().slice(0, 10);
  return null;
}

function pick(row, names) {
  const keys = Object.keys(row);
  const normalizedLookup = new Map(keys.map((key) => [normalizeHeader(key), key]));
  for (const wanted of names) {
    const key = normalizedLookup.get(normalizeHeader(wanted));
    if (key) return row[key];
  }
  return undefined;
}

async function initDatabase() {
  if (!process.env.DATABASE_URL) {
    app.log.warn('DATABASE_URL is not set. Database features will fail until Railway variable is connected.');
    return;
  }

  await pool.query(`
    CREATE TABLE IF NOT EXISTS jobs (
      id SERIAL PRIMARY KEY,
      job_number VARCHAR(50),
      invoice_number VARCHAR(50),
      job_date DATE,
      division VARCHAR(100),
      ops_manager VARCHAR(100),
      hours NUMERIC,
      revenue NUMERIC,
      labour_cost NUMERIC,
      equipment_cost NUMERIC,
      workshop_cost NUMERIC,
      total_cost NUMERIC,
      gross_profit NUMERIC,
      created_at TIMESTAMP DEFAULT NOW()
    );
  `);

  await pool.query(`
    CREATE TABLE IF NOT EXISTS imports (
      id SERIAL PRIMARY KEY,
      filename VARCHAR(255),
      imported_rows INTEGER DEFAULT 0,
      status VARCHAR(50) DEFAULT 'completed',
      imported_at TIMESTAMP DEFAULT NOW()
    );
  `);
}

await app.register(fastifyMultipart, { limits: { fileSize: 25 * 1024 * 1024 } });
await app.register(fastifyStatic, { root: path.join(__dirname, 'public'), prefix: '/' });

app.get('/health', async () => ({ status: 'ok', app: 'B4 Nautilus Operations' }));

app.get('/api/status', async () => {
  try {
    const db = await pool.query('SELECT NOW() AS now');
    return { app: 'B4 Nautilus Operations', status: 'running', database: 'connected', time: db.rows[0].now };
  } catch (error) {
    return { app: 'B4 Nautilus Operations', status: 'running', database: 'not connected', error: error.message };
  }
});

app.get('/api/jobs', async (request) => {
  const limit = Math.min(Number(request.query?.limit || 100), 500);
  const result = await pool.query(`SELECT * FROM jobs ORDER BY job_date DESC NULLS LAST, id DESC LIMIT $1`, [limit]);
  return result.rows;
});

app.get('/api/dashboard', async () => {
  const totals = await pool.query(`
    SELECT COALESCE(SUM(revenue),0) AS revenue, COALESCE(SUM(labour_cost),0) AS labour_cost,
           COALESCE(SUM(equipment_cost),0) AS equipment_cost, COALESCE(SUM(workshop_cost),0) AS workshop_cost,
           COALESCE(SUM(total_cost),0) AS total_cost, COALESCE(SUM(gross_profit),0) AS gross_profit, COUNT(*) AS jobs
    FROM jobs
  `);
  const divisions = await pool.query(`SELECT COALESCE(division,'Unassigned') AS division, COALESCE(SUM(gross_profit),0) AS gross_profit FROM jobs GROUP BY division ORDER BY gross_profit DESC`);
  const opsManagers = await pool.query(`SELECT COALESCE(ops_manager,'Unassigned') AS ops_manager, COALESCE(SUM(gross_profit),0) AS gross_profit FROM jobs GROUP BY ops_manager ORDER BY gross_profit DESC`);
  return { totals: totals.rows[0], divisions: divisions.rows, opsManagers: opsManagers.rows };
});

app.get('/api/imports', async () => {
  const result = await pool.query(`SELECT * FROM imports ORDER BY imported_at DESC LIMIT 50`);
  return result.rows;
});

app.post('/api/imports/test-seed', async () => {
  await pool.query('DELETE FROM jobs');
  const sampleJobs = [
    ['#1','IN001','2024-01-01','Pollution','Jolanda',15,170500,6110,13731,2763],
    ['#2','IN002','2024-01-01','Pollution','Jolanda',1.9,17500,1200,1900,339],
    ['#3','IN003','2024-01-02','Pollution','Jolanda',4,17500,2000,3800,1154],
    ['#4','IN004','2024-01-03','Diving','Eugen',5.4,38280,4500,9000,6255],
    ['#5','IN005','2024-01-03','Civil','Carmon',21,178200,6200,9700,9958]
  ];
  for (const job of sampleJobs) {
    const totalCost = Number(job[7]) + Number(job[8]) + Number(job[9]);
    const grossProfit = Number(job[6]) - totalCost;
    await pool.query(`INSERT INTO jobs (job_number,invoice_number,job_date,division,ops_manager,hours,revenue,labour_cost,equipment_cost,workshop_cost,total_cost,gross_profit) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12)`, [...job,totalCost,grossProfit]);
  }
  await pool.query(`INSERT INTO imports (filename, imported_rows, status) VALUES ($1,$2,$3)`, ['test-seed', sampleJobs.length, 'completed']);
  return { status: 'seeded', rows: sampleJobs.length };
});

app.post('/api/imports/excel', async (request, reply) => {
  const file = await request.file();
  if (!file) return reply.code(400).send({ error: 'No Excel file uploaded' });

  const buffer = await file.toBuffer();
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const best = chooseBestSheet(workbook);
  if (!best || best.header.score < 3) {
    return reply.code(400).send({
      error: 'Could not find a valid job data sheet. Please upload a workbook with columns like Date, Job Nr, Invoice No, Revenue, Costs, Division and OPS Manager.',
      detectedSheets: workbook.SheetNames
    });
  }

  const rows = XLSX.utils.sheet_to_json(best.sheet, { defval: '', range: best.header.rowIndex });
  const client = await pool.connect();
  let importedRows = 0;
  let skippedRows = 0;

  try {
    await client.query('BEGIN');
    await client.query('DELETE FROM jobs');

    for (const row of rows) {
      const jobNumber = pick(row, FIELD_ALIASES.jobNumber);
      const invoiceNumber = pick(row, FIELD_ALIASES.invoiceNumber);
      const jobDate = toDate(pick(row, FIELD_ALIASES.jobDate));
      const division = pick(row, FIELD_ALIASES.division);
      const opsManager = pick(row, FIELD_ALIASES.opsManager);
      const hours = toNumber(pick(row, FIELD_ALIASES.hours));
      const revenue = toNumber(pick(row, FIELD_ALIASES.revenue));
      const labourCost = toNumber(pick(row, FIELD_ALIASES.labourCost));
      const equipmentCost = toNumber(pick(row, FIELD_ALIASES.equipmentCost));
      const workshopCost = toNumber(pick(row, FIELD_ALIASES.workshopCost));
      const totalCostFromSheet = toNumber(pick(row, FIELD_ALIASES.totalCost));
      const totalCost = totalCostFromSheet || labourCost + equipmentCost + workshopCost;
      const grossProfit = revenue - totalCost;

      const hasFinancialData = revenue !== 0 || totalCost !== 0 || labourCost !== 0 || equipmentCost !== 0 || workshopCost !== 0;
      const hasRealIdentifier = Boolean(jobNumber || invoiceNumber || jobDate);
      if (!hasRealIdentifier && !hasFinancialData) { skippedRows += 1; continue; }
      if (!jobDate && !hasFinancialData) { skippedRows += 1; continue; }

      await client.query(
        `INSERT INTO jobs (job_number,invoice_number,job_date,division,ops_manager,hours,revenue,labour_cost,equipment_cost,workshop_cost,total_cost,gross_profit) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12)`,
        [jobNumber ? String(jobNumber) : null, invoiceNumber ? String(invoiceNumber) : null, jobDate, division ? String(division) : null, opsManager ? String(opsManager) : null, hours, revenue, labourCost, equipmentCost, workshopCost, totalCost, grossProfit]
      );
      importedRows += 1;
    }

    await client.query(`INSERT INTO imports (filename, imported_rows, status) VALUES ($1,$2,$3)`, [file.filename || 'uploaded workbook', importedRows, 'completed']);
    await client.query('COMMIT');

    return { status: 'imported', filename: file.filename, sheet: best.sheetName, headerRow: best.header.rowIndex + 1, headerScore: best.header.score, rowsFound: rows.length, importedRows, skippedRows };
  } catch (error) {
    await client.query('ROLLBACK');
    app.log.error(error);
    return reply.code(500).send({ error: error.message });
  } finally {
    client.release();
  }
});

app.setNotFoundHandler((request, reply) => { reply.sendFile('index.html'); });

try {
  await initDatabase();
  await app.listen({ port, host });
} catch (err) {
  app.log.error(err);
  process.exit(1);
}
