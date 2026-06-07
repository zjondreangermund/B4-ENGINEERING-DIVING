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
  ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false,
});

const FIELD_ALIASES = {
  jobNumber: ['Job Nr','Job No','Job Number','Job','Job Nr.','Job #','Job_No','Job_Nr','JobNo','JobNr'],
  invoiceNumber: ['Invoice No','Invoice Number','Invoice','Invoice No.','Inv No','Invoice_No'],
  jobDate: ['Date','Job Date','Start Date'],
  division: ['Division','Division ','Divisio'],
  opsManager: ['OPS Manager','OPS manager','Operations Manager','Ops Manager','Divison Head','Division Head','Divison H'],
  description: ['Description','Job/Project_Description','Job Project Description','Job Description','Column1'],
  client: ['Client','Customer','Client Name'],
  quoteNumber: ['Quote_Nr','Quote Nr','Quote Number','Quote No'],
  poNumber: ['PO_No','PO No','PO Number','PO'],
  hours: ['Total Hours','Hours','Hour'],
  revenue: ['Revenue Excl VAT','Revenue Excl','Revenue','Revenue Ex VAT','Revenue excluding VAT',' Revenue Excl VAT','Value_excl_Va','Value_excl_Vat','Value Excl Vat','Value Excl VAT'],
  revenueInclVat: ['Revenue Incl VAT','Revenue Incl','Revenue Inc VAT','Revenue Including VAT','Value_incl_Vat','Value Incl VAT'],
  labourCost: ['Labour Costs','Labour Cost','Cost Labour','Labor Costs','Labor Cost',' Labour_Costs'],
  equipmentCost: ['Equipment Costs','Equipment Cost','Cost Equipment',' Equipment_Costs'],
  workshopCost: ['Workshop Cost','Workshop Costs',' Workshop Cost'],
  totalCost: ['Total Costs','Total Cost',' Total_Costs'],
  paymentsMade: ['Payments Made','Payment Made','Amount Paid','Paid','Payments','Receipts','payments_made','payment_made','amount_paid'],
};

function normalizeHeader(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '');
}
function toNumber(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const parsed = Number(String(value).replace(/[^0-9.-]/g, ''));
  return Number.isFinite(parsed) ? parsed : 0;
}
function toDate(value) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);
  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return `${parsed.y}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`;
  }
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date.toISOString().slice(0, 10);
}
function toTime(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number') {
    const totalMinutes = Math.round((value % 1) * 24 * 60);
    return `${String(Math.floor(totalMinutes / 60)).padStart(2, '0')}:${String(totalMinutes % 60).padStart(2, '0')}`;
  }
  if (value instanceof Date && !Number.isNaN(value.getTime())) return `${String(value.getUTCHours()).padStart(2, '0')}:${String(value.getUTCMinutes()).padStart(2, '0')}`;
  const match = String(value).match(/(\d{1,2}:\d{2})/);
  return match ? match[1] : String(value);
}
function pick(row, names) {
  const keys = Object.keys(row);
  const lookup = new Map(keys.map((key) => [normalizeHeader(key), key]));
  for (const wanted of names) {
    const key = lookup.get(normalizeHeader(wanted));
    if (key) return row[key];
  }
  return undefined;
}
function headerScore(headers) {
  const normalized = headers.map(normalizeHeader);
  let score = 0;
  for (const aliases of Object.values(FIELD_ALIASES)) if (aliases.some((alias) => normalized.includes(normalizeHeader(alias)))) score += 1;
  return score;
}
function findHeaderRow(sheet) {
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false });
  let best = { rowIndex: 0, score: 0, headers: [] };
  matrix.slice(0, 80).forEach((row, rowIndex) => {
    const score = headerScore(row);
    if (score > best.score) best = { rowIndex, score, headers: row };
  });
  return best;
}
function rowsFromSheet(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];
  const header = findHeaderRow(sheet);
  return XLSX.utils.sheet_to_json(sheet, { defval: '', range: header.rowIndex });
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
function findSheet(workbook, candidates) {
  const sheetNames = workbook.SheetNames.map((name) => ({ name, key: normalizeHeader(name) }));
  for (const candidate of candidates) {
    const target = normalizeHeader(candidate);
    const exact = sheetNames.find((sheet) => sheet.key === target);
    if (exact) return exact.name;
  }
  for (const candidate of candidates) {
    const target = normalizeHeader(candidate);
    const partial = sheetNames.find((sheet) => sheet.key.includes(target) || target.includes(sheet.key));
    if (partial) return partial.name;
  }
  return null;
}

async function initDatabase() {
  if (!process.env.DATABASE_URL) {
    app.log.warn('DATABASE_URL is not set.');
    return;
  }
  await pool.query(`CREATE TABLE IF NOT EXISTS jobs (id SERIAL PRIMARY KEY, job_number VARCHAR(50), invoice_number VARCHAR(50), job_date DATE, division VARCHAR(100), ops_manager VARCHAR(100), hours NUMERIC, revenue NUMERIC, labour_cost NUMERIC, equipment_cost NUMERIC, workshop_cost NUMERIC, total_cost NUMERIC, gross_profit NUMERIC, created_at TIMESTAMP DEFAULT NOW());`);
  await pool.query(`ALTER TABLE jobs ADD COLUMN IF NOT EXISTS description TEXT;`);
  await pool.query(`ALTER TABLE jobs ADD COLUMN IF NOT EXISTS client_name TEXT;`);
  await pool.query(`ALTER TABLE jobs ADD COLUMN IF NOT EXISTS quote_number TEXT;`);
  await pool.query(`ALTER TABLE jobs ADD COLUMN IF NOT EXISTS po_number TEXT;`);
  await pool.query(`CREATE TABLE IF NOT EXISTS imports (id SERIAL PRIMARY KEY, filename VARCHAR(255), imported_rows INTEGER DEFAULT 0, status VARCHAR(50) DEFAULT 'completed', imported_at TIMESTAMP DEFAULT NOW());`);
  await pool.query(`CREATE TABLE IF NOT EXISTS job_register_entries (id SERIAL PRIMARY KEY, job_number VARCHAR(50), job_date DATE, ops_manager TEXT, division TEXT, client_name TEXT, description TEXT, completion_date DATE, quote_number TEXT, po_number TEXT, report_reference TEXT, invoice_number TEXT, client_feedback TEXT, value_incl_vat NUMERIC DEFAULT 0, value_excl_vat NUMERIC DEFAULT 0, payments_made NUMERIC DEFAULT 0, imported_at TIMESTAMP DEFAULT NOW());`);
  await pool.query(`ALTER TABLE job_register_entries ADD COLUMN IF NOT EXISTS value_incl_vat NUMERIC DEFAULT 0;`);
  await pool.query(`ALTER TABLE job_register_entries ADD COLUMN IF NOT EXISTS value_excl_vat NUMERIC DEFAULT 0;`);
  await pool.query(`ALTER TABLE job_register_entries ADD COLUMN IF NOT EXISTS payments_made NUMERIC DEFAULT 0;`);
  await pool.query(`CREATE TABLE IF NOT EXISTS job_card_entries (id SERIAL PRIMARY KEY, job_date DATE, month TEXT, year INTEGER, fy TEXT, job_number VARCHAR(50), ops_manager TEXT, description TEXT, start_time TEXT, end_time TEXT, hours NUMERIC, labour_cost NUMERIC, equipment_cost NUMERIC, workshop_cost NUMERIC, division TEXT, imported_at TIMESTAMP DEFAULT NOW());`);
  await pool.query(`CREATE TABLE IF NOT EXISTS salaries (id SERIAL PRIMARY KEY, salary_year INTEGER, salary_month TEXT, total_salaries NUMERIC, imported_at TIMESTAMP DEFAULT NOW());`);
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
  const limit = Math.min(Number(request.query?.limit || 5000), 5000);
  const result = await pool.query(`SELECT * FROM jobs ORDER BY job_date DESC NULLS LAST, id DESC LIMIT $1`, [limit]);
  return result.rows;
});
app.get('/api/dashboard', async () => {
  const totals = await pool.query(`SELECT COALESCE(SUM(revenue),0) AS revenue, COALESCE(SUM(labour_cost),0) AS labour_cost, COALESCE(SUM(equipment_cost),0) AS equipment_cost, COALESCE(SUM(workshop_cost),0) AS workshop_cost, COALESCE(SUM(total_cost),0) AS total_cost, COALESCE(SUM(gross_profit),0) AS gross_profit, COUNT(*) AS jobs FROM jobs`);
  const divisions = await pool.query(`SELECT COALESCE(division,'Unassigned') AS division, COALESCE(SUM(gross_profit),0) AS gross_profit FROM jobs GROUP BY division ORDER BY gross_profit DESC`);
  const opsManagers = await pool.query(`SELECT COALESCE(ops_manager,'Unassigned') AS ops_manager, COALESCE(SUM(gross_profit),0) AS gross_profit FROM jobs GROUP BY ops_manager ORDER BY gross_profit DESC`);
  return { totals: totals.rows[0], divisions: divisions.rows, opsManagers: opsManagers.rows };
});
app.get('/api/imports', async () => {
  const result = await pool.query(`SELECT * FROM imports ORDER BY imported_at DESC LIMIT 50`);
  return result.rows;
});
app.get('/api/job-register', async (request) => {
  const limit = Math.min(Number(request.query?.limit || 5000), 5000);
  const result = await pool.query(`SELECT * FROM job_register_entries ORDER BY job_date DESC NULLS LAST, id DESC LIMIT $1`, [limit]);
  return result.rows;
});
app.get('/api/job-cards', async (request) => {
  const limit = Math.min(Number(request.query?.limit || 5000), 5000);
  const result = await pool.query(`SELECT * FROM job_card_entries ORDER BY job_date DESC NULLS LAST, id DESC LIMIT $1`, [limit]);
  return result.rows;
});
app.get('/api/salaries', async () => {
  const result = await pool.query(`SELECT * FROM salaries ORDER BY salary_year DESC, imported_at DESC`);
  return result.rows;
});
app.post('/api/imports/test-seed', async () => ({ status: 'disabled', rows: 0, message: 'Seed test data is disabled on this build.' }));

async function importWorkbook(buffer, filename) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const best = chooseBestSheet(workbook);
  if (!best || best.header.score < 3) throw new Error('Could not find a valid job data sheet.');
  const rows = XLSX.utils.sheet_to_json(best.sheet, { defval: '', range: best.header.rowIndex });
  const client = await pool.connect();
  let importedRows = 0, skippedRows = 0, registerRows = 0, cardRows = 0, salaryRows = 0;
  try {
    await client.query('BEGIN');
    await client.query('DELETE FROM jobs');
    await client.query('DELETE FROM job_register_entries');
    await client.query('DELETE FROM job_card_entries');
    await client.query('DELETE FROM salaries');
    for (const row of rows) {
      const jobNumber = pick(row, FIELD_ALIASES.jobNumber);
      const invoiceNumber = pick(row, FIELD_ALIASES.invoiceNumber);
      const jobDate = toDate(pick(row, FIELD_ALIASES.jobDate));
      const division = pick(row, FIELD_ALIASES.division);
      const opsManager = pick(row, FIELD_ALIASES.opsManager);
      const description = pick(row, FIELD_ALIASES.description);
      const clientName = pick(row, FIELD_ALIASES.client);
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
      if ((!hasRealIdentifier && !hasFinancialData) || (!jobDate && !hasFinancialData)) { skippedRows += 1; continue; }
      await client.query(`INSERT INTO jobs (job_number, invoice_number, job_date, division, ops_manager, hours, revenue, labour_cost, equipment_cost, workshop_cost, total_cost, gross_profit, description, client_name, quote_number, po_number) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16)`, [jobNumber ? String(jobNumber) : null, invoiceNumber ? String(invoiceNumber) : null, jobDate, division ? String(division) : null, opsManager ? String(opsManager) : null, hours, revenue, labourCost, equipmentCost, workshopCost, totalCost, grossProfit, description ? String(description) : null, clientName ? String(clientName) : null, pick(row, FIELD_ALIASES.quoteNumber) || null, pick(row, FIELD_ALIASES.poNumber) || null]);
      importedRows += 1;
    }
    const jrSheet = workbook.SheetNames.find((name) => name === 'Job Register ') || findSheet(workbook, ['Job Register', 'JobRegister', 'Register']);
    if (jrSheet) {
      app.log.info(`Found Job Register sheet: ${jrSheet}`);
      const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[jrSheet], { header: 1, defval: '', blankrows: false });
      const headerRow = matrix[1] || [];
      const normalizedHeader = new Map(headerRow.map((header, index) => [normalizeHeader(header), index]));
      const paymentHeaderIndex = FIELD_ALIASES.paymentsMade.map((alias) => normalizedHeader.get(normalizeHeader(alias))).find((index) => Number.isInteger(index));
      const paymentsColumnIndex = Number.isInteger(paymentHeaderIndex) ? paymentHeaderIndex : 14;
      const dataRows = matrix.slice(2);
      for (const row of dataRows) {
        const jobNumber = row[0];
        const jobDate = toDate(row[1]);
        const opsManager = row[2];
        const division = row[3];
        const clientName = row[4];
        const description = row[5];
        const completionDate = toDate(row[6]);
        const quoteNumber = row[7];
        const poNumber = row[8];
        const issued = row[9];
        const invoiceNumber = row[10];
        const form = row[11];
        const valueInclVat = toNumber(row[12]);
        const valueExclVat = toNumber(row[13]);
        const paymentsMade = toNumber(row[paymentsColumnIndex]);
        if (!jobNumber && !description && !invoiceNumber) continue;
        await client.query(`INSERT INTO job_register_entries (job_number, job_date, ops_manager, division, client_name, description, completion_date, quote_number, po_number, report_reference, invoice_number, client_feedback, value_incl_vat, value_excl_vat, payments_made) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15)`, [jobNumber ? String(jobNumber) : null, jobDate, opsManager ? String(opsManager) : null, division ? String(division) : null, clientName ? String(clientName) : null, description ? String(description) : null, completionDate, quoteNumber ? String(quoteNumber) : null, poNumber ? String(poNumber) : null, issued ? String(issued) : null, invoiceNumber ? String(invoiceNumber) : null, form ? String(form) : null, valueInclVat, valueExclVat, paymentsMade]);
        registerRows += 1;
      }
    }
    const jcSheet = findSheet(workbook, ['Job Card Conversion', 'JobCardConversion']);
    if (jcSheet) {
      for (const row of rowsFromSheet(workbook, jcSheet)) {
        const jobNumber = pick(row, FIELD_ALIASES.jobNumber);
        if (!jobNumber) continue;
        await client.query(`INSERT INTO job_card_entries (job_date, month, year, fy, job_number, ops_manager, description, start_time, end_time, hours, labour_cost, equipment_cost, workshop_cost, division) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14)`, [toDate(pick(row, FIELD_ALIASES.jobDate)), row.Month || row['Month'] || null, toNumber(row.Year || row['Year']) || null, row.FY ? String(row.FY) : null, String(jobNumber), pick(row, FIELD_ALIASES.opsManager) || null, pick(row, FIELD_ALIASES.description) || null, toTime(row['Start Time '] ?? row['Start Time']), toTime(row['End Time']), toNumber(pick(row, FIELD_ALIASES.hours)), toNumber(pick(row, FIELD_ALIASES.labourCost)), toNumber(pick(row, FIELD_ALIASES.equipmentCost)), toNumber(pick(row, FIELD_ALIASES.workshopCost)), pick(row, FIELD_ALIASES.division) || null]);
        cardRows += 1;
      }
    }
    const salSheet = findSheet(workbook, ['Salaries', 'Salary']);
    if (salSheet) {
      const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[salSheet], { header: 1, defval: '' });
      const years = [];
      matrix.forEach((row) => row.forEach((cell, idx) => { if (Number(cell) >= 2020 && Number(cell) <= 2035) years.push({ year: Number(cell), col: idx }); }));
      for (const item of years) {
        for (const row of matrix) {
          const month = row[item.col];
          const value = row[item.col + 1];
          if (typeof month === 'string' && month.length > 2 && toNumber(value) > 0) {
            await client.query(`INSERT INTO salaries (salary_year, salary_month, total_salaries) VALUES ($1,$2,$3)`, [item.year, month, toNumber(value)]);
            salaryRows += 1;
          }
        }
      }
    }
    await client.query(`INSERT INTO imports (filename, imported_rows, status) VALUES ($1,$2,$3)`, [filename || 'uploaded workbook', importedRows, 'completed']);
    await client.query('COMMIT');
    return { status: 'imported', filename, sheet: best.sheetName, headerRow: best.header.rowIndex + 1, headerScore: best.header.score, rowsFound: rows.length, importedRows, skippedRows, registerRows, cardRows, salaryRows, sheets: workbook.SheetNames };
  } catch (error) {
    await client.query('ROLLBACK');
    throw error;
  } finally {
    client.release();
  }
}

app.post('/api/imports/excel', async (request, reply) => {
  try {
    const file = await request.file();
    if (!file) return reply.code(400).send({ error: 'No Excel file uploaded' });
    return await importWorkbook(await file.toBuffer(), file.filename);
  } catch (error) {
    app.log.error(error);
    return reply.code(500).send({ error: error.message });
  }
});

app.setNotFoundHandler((request, reply) => {
  reply.sendFile('index.html');
});

try {
  await initDatabase();
  await app.listen({ port, host });
} catch (err) {
  app.log.error(err);
  process.exit(1);
}
