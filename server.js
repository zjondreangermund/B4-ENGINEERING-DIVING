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
  invoiceNumber: ['Invoice No','Invoice Number','Invoice','Invoice No.','Inv No','Invoice_No','Invoice No.'],
  jobDate: ['Date','Job Date','Start Date'],
  division: ['Division','Division ','Divisio','Main Division'],
  opsManager: ['OPS Manager','OPS manager','Operations Manager','Ops Manager','Divison Head','Division Head','Divison H'],
  description: ['Description','Job/Project_Description','Job Project Description','Job Description','Column1'],
  client: ['Client','Customer','Client Name'],
  quoteNumber: ['Quote_Nr','Quote Nr','Quote Number','Quote No'],
  poNumber: ['PO_No','PO No','PO Number','PO'],
  hours: ['Total Hours','Hours','Hour',' Total Hours'],
  revenue: ['Revenue Excl VAT','Revenue Excl','Revenue','Revenue Ex VAT','Revenue excluding VAT',' Revenue Excl VAT','Value_excl_Va','Value_excl_Vat','Value Excl Vat','Value Excl VAT'],
  revenueInclVat: ['Revenue Incl VAT','Revenue Incl','Revenue Inc VAT','Revenue Including VAT','Value_incl_Vat','Value Incl VAT'],
  labourCost: ['Labour Costs','Labour Cost','Cost Labour','Labor Costs','Labor Cost',' Labour_Costs'],
  equipmentCost: ['Equipment Costs','Equipment Cost','Cost Equipment',' Equipment_Costs'],
  workshopCost: ['Workshop Cost','Workshop Costs',' Workshop Cost'],
  totalCost: ['Total Costs','Total Cost',' Total_Costs'],
  paymentsMade: ['Payments Made','Payment Made','Amount Paid','Paid','Payments','Receipts','payments_made','payment_made','amount_paid'],
  expenseDate: ['Date','Expense Date','Transaction Date','Month'],
  expenseSupplier: ['Supplier','Vendor','Paid To','Creditor','Company'],
  expenseCategory: ['Category','Expense Category','Type','Expense Type','Account'],
  expenseDescription: ['Description','Details','Narrative','Item','Expense Description'],
  expenseAmount: ['Amount','Value','Cost','Expense','Total','Total Cost','Incl VAT','Excl VAT'],
  vessel: ['Vessel','Boat','Vessel Name'],
};

const VESSEL_COLUMNS = ['Dingy', 'Ingwegwe', 'DSV Saturn', 'Flatcat', 'Ingwena', 'Pumba'];

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
function toHours(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return value > 0 && value < 1 ? value * 24 : value;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.getUTCHours() + value.getUTCMinutes() / 60 + value.getUTCSeconds() / 3600;
  const text = String(value).trim();
  const time = text.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (time) return Number(time[1]) + Number(time[2]) / 60 + Number(time[3] || 0) / 3600;
  return toNumber(value);
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
function isUsed(value) {
  if (value === true) return true;
  if (typeof value === 'number') return value === 1;
  return ['1', 'yes', 'y', 'true', 'x'].includes(String(value || '').trim().toLowerCase());
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
  await pool.query(`CREATE TABLE IF NOT EXISTS vessel_entries (id SERIAL PRIMARY KEY, job_date DATE, month TEXT, client_name TEXT, job_number VARCHAR(50), ops_manager TEXT, description TEXT, start_time TEXT, end_time TEXT, hours NUMERIC, vessel_name TEXT, labour_cost NUMERIC DEFAULT 0, equipment_cost NUMERIC DEFAULT 0, workshop_cost NUMERIC DEFAULT 0, division TEXT, imported_at TIMESTAMP DEFAULT NOW());`);
  await pool.query(`CREATE INDEX IF NOT EXISTS vessel_entries_job_date_idx ON vessel_entries (job_date);`);
  await pool.query(`CREATE INDEX IF NOT EXISTS vessel_entries_vessel_name_idx ON vessel_entries (vessel_name);`);
  await pool.query(`CREATE TABLE IF NOT EXISTS expense_entries (id SERIAL PRIMARY KEY, expense_date DATE, month TEXT, supplier TEXT, category TEXT, description TEXT, division TEXT, vessel_name TEXT, amount NUMERIC DEFAULT 0, imported_at TIMESTAMP DEFAULT NOW());`);
  await pool.query(`CREATE INDEX IF NOT EXISTS expense_entries_expense_date_idx ON expense_entries (expense_date);`);
  await pool.query(`CREATE INDEX IF NOT EXISTS expense_entries_category_idx ON expense_entries (category);`);
}

await app.register(fastifyMultipart, { limits: { fileSize: 30 * 1024 * 1024 } });
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
app.get('/api/expenses', async (request) => {
  const limit = Math.min(Number(request.query?.limit || 5000), 10000);
  const result = await pool.query(`SELECT * FROM expense_entries ORDER BY expense_date DESC NULLS LAST, id DESC LIMIT $1`, [limit]);
  return result.rows;
});
app.get('/api/expense-dashboard', async () => {
  const totals = await pool.query(`SELECT COUNT(*) AS entries, COALESCE(SUM(amount),0) AS amount FROM expense_entries`);
  const byCategory = await pool.query(`SELECT COALESCE(category,'Unassigned') AS category, COUNT(*) AS entries, COALESCE(SUM(amount),0) AS amount FROM expense_entries GROUP BY category ORDER BY amount DESC`);
  const byVessel = await pool.query(`SELECT COALESCE(vessel_name,'Unassigned') AS vessel_name, COUNT(*) AS entries, COALESCE(SUM(amount),0) AS amount FROM expense_entries GROUP BY vessel_name ORDER BY amount DESC`);
  const byMonth = await pool.query(`SELECT TO_CHAR(expense_date,'YYYY-MM') AS month, COUNT(*) AS entries, COALESCE(SUM(amount),0) AS amount FROM expense_entries WHERE expense_date IS NOT NULL GROUP BY TO_CHAR(expense_date,'YYYY-MM') ORDER BY month`);
  return { totals: totals.rows[0], byCategory: byCategory.rows, byVessel: byVessel.rows, byMonth: byMonth.rows };
});
app.get('/api/vessels', async (request) => {
  const limit = Math.min(Number(request.query?.limit || 5000), 10000);
  const result = await pool.query(`SELECT * FROM vessel_entries ORDER BY job_date DESC NULLS LAST, id DESC LIMIT $1`, [limit]);
  return result.rows;
});
app.get('/api/vessel-dashboard', async () => {
  const totals = await pool.query(`SELECT COUNT(*) AS entries, COUNT(DISTINCT job_number) AS jobs, COALESCE(SUM(hours),0) AS hours, COALESCE(SUM(labour_cost),0) AS labour_cost, COALESCE(SUM(equipment_cost),0) AS equipment_cost, COALESCE(SUM(workshop_cost),0) AS workshop_cost FROM vessel_entries`);
  const byVessel = await pool.query(`SELECT COALESCE(vessel_name,'Unassigned') AS vessel_name, COUNT(*) AS entries, COUNT(DISTINCT job_number) AS jobs, COALESCE(SUM(hours),0) AS hours, COALESCE(SUM(labour_cost),0) AS labour_cost, COALESCE(SUM(equipment_cost),0) AS equipment_cost, COALESCE(SUM(workshop_cost),0) AS workshop_cost FROM vessel_entries GROUP BY vessel_name ORDER BY hours DESC`);
  const byClient = await pool.query(`SELECT COALESCE(client_name,'Unassigned') AS client_name, COUNT(*) AS entries, COUNT(DISTINCT job_number) AS jobs, COALESCE(SUM(hours),0) AS hours FROM vessel_entries GROUP BY client_name ORDER BY hours DESC LIMIT 20`);
  const byOps = await pool.query(`SELECT COALESCE(ops_manager,'Unassigned') AS ops_manager, COUNT(*) AS entries, COUNT(DISTINCT job_number) AS jobs, COALESCE(SUM(hours),0) AS hours FROM vessel_entries GROUP BY ops_manager ORDER BY hours DESC`);
  const byMonth = await pool.query(`SELECT TO_CHAR(job_date,'YYYY-MM') AS month, COUNT(*) AS entries, COUNT(DISTINCT job_number) AS jobs, COALESCE(SUM(hours),0) AS hours FROM vessel_entries WHERE job_date IS NOT NULL GROUP BY TO_CHAR(job_date,'YYYY-MM') ORDER BY month`);
  return { totals: totals.rows[0], byVessel: byVessel.rows, byClient: byClient.rows, byOps: byOps.rows, byMonth: byMonth.rows };
});
app.post('/api/imports/test-seed', async () => ({ status: 'disabled', rows: 0, message: 'Seed test data is disabled on this build.' }));

async function importWorkbook(buffer, filename) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const best = chooseBestSheet(workbook);
  if (!best || best.header.score < 3) throw new Error('Could not find a valid job data sheet.');
  const rows = XLSX.utils.sheet_to_json(best.sheet, { defval: '', range: best.header.rowIndex });
  const client = await pool.connect();
  let importedRows = 0, skippedRows = 0, registerRows = 0, cardRows = 0, salaryRows = 0, vesselRows = 0, expenseRows = 0;
  try {
    await client.query('BEGIN');
    await client.query('DELETE FROM jobs');
    await client.query('DELETE FROM job_register_entries');
    await client.query('DELETE FROM job_card_entries');
    await client.query('DELETE FROM salaries');
    await client.query('DELETE FROM vessel_entries');
    await client.query('DELETE FROM expense_entries');

    for (const row of rows) {
      const jobNumber = pick(row, FIELD_ALIASES.jobNumber);
      const invoiceNumber = pick(row, FIELD_ALIASES.invoiceNumber);
      const jobDate = toDate(pick(row, FIELD_ALIASES.jobDate));
      const division = pick(row, FIELD_ALIASES.division);
      const opsManager = pick(row, FIELD_ALIASES.opsManager);
      const description = pick(row, FIELD_ALIASES.description);
      const clientName = pick(row, FIELD_ALIASES.client);
      const hours = toHours(pick(row, FIELD_ALIASES.hours));
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
      const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[jrSheet], { header: 1, defval: '', blankrows: false });
      const headerRow = matrix[1] || [];
      const normalizedHeader = new Map(headerRow.map((header, index) => [normalizeHeader(header), index]));
      const paymentHeaderIndex = FIELD_ALIASES.paymentsMade.map((alias) => normalizedHeader.get(normalizeHeader(alias))).find((index) => Number.isInteger(index));
      const paymentsColumnIndex = Number.isInteger(paymentHeaderIndex) ? paymentHeaderIndex : 14;
      for (const row of matrix.slice(2)) {
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
        await client.query(`INSERT INTO job_card_entries (job_date, month, year, fy, job_number, ops_manager, description, start_time, end_time, hours, labour_cost, equipment_cost, workshop_cost, division) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14)`, [toDate(pick(row, FIELD_ALIASES.jobDate)), row.Month || row['Month'] || null, toNumber(row.Year || row['Year']) || null, row.FY ? String(row.FY) : null, String(jobNumber), pick(row, FIELD_ALIASES.opsManager) || null, pick(row, FIELD_ALIASES.description) || null, toTime(row['Start Time '] ?? row['Start Time']), toTime(row['End Time']), toHours(pick(row, FIELD_ALIASES.hours)), toNumber(pick(row, FIELD_ALIASES.labourCost)), toNumber(pick(row, FIELD_ALIASES.equipmentCost)), toNumber(pick(row, FIELD_ALIASES.workshopCost)), pick(row, FIELD_ALIASES.division) || null]);
        cardRows += 1;
      }
    }

    const vesselSheet = findSheet(workbook, ['Vessel Info', 'VesselInfo', 'Vessels']);
    if (vesselSheet) {
      const vesselData = rowsFromSheet(workbook, vesselSheet);
      for (const row of vesselData) {
        const jobNumber = pick(row, FIELD_ALIASES.jobNumber);
        const usedVessels = VESSEL_COLUMNS.filter((vessel) => isUsed(row[vessel]));
        if (!jobNumber || usedVessels.length === 0) continue;
        const baseHours = toHours(pick(row, FIELD_ALIASES.hours));
        const perVesselHours = usedVessels.length > 1 ? baseHours / usedVessels.length : baseHours;
        const labourCost = toNumber(pick(row, FIELD_ALIASES.labourCost));
        const equipmentCost = toNumber(pick(row, FIELD_ALIASES.equipmentCost));
        const workshopCost = toNumber(pick(row, FIELD_ALIASES.workshopCost));
        for (const vessel of usedVessels) {
          await client.query(`INSERT INTO vessel_entries (job_date, month, client_name, job_number, ops_manager, description, start_time, end_time, hours, vessel_name, labour_cost, equipment_cost, workshop_cost, division) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14)`, [toDate(pick(row, FIELD_ALIASES.jobDate)), row.Month || row['Month'] || null, pick(row, FIELD_ALIASES.client) ? String(pick(row, FIELD_ALIASES.client)) : null, String(jobNumber), pick(row, FIELD_ALIASES.opsManager) || null, pick(row, FIELD_ALIASES.description) || null, toTime(row['Start Time '] ?? row['Start Time']), toTime(row['End Time']), perVesselHours, vessel, usedVessels.length > 1 ? labourCost / usedVessels.length : labourCost, usedVessels.length > 1 ? equipmentCost / usedVessels.length : equipmentCost, usedVessels.length > 1 ? workshopCost / usedVessels.length : workshopCost, pick(row, FIELD_ALIASES.division) || null]);
          vesselRows += 1;
        }
      }
    }


    const expSheet = findSheet(workbook, ['Expenses', 'Expense']);
    if (expSheet) {
      for (const row of rowsFromSheet(workbook, expSheet)) {
        const amount = toNumber(pick(row, FIELD_ALIASES.expenseAmount));
        const description = pick(row, FIELD_ALIASES.expenseDescription) || pick(row, FIELD_ALIASES.description);
        if (!amount && !description) continue;
        await client.query(
          `INSERT INTO expense_entries (expense_date, month, supplier, category, description, division, vessel_name, amount) VALUES ($1,$2,$3,$4,$5,$6,$7,$8)`,
          [
            toDate(pick(row, FIELD_ALIASES.expenseDate)),
            row.Month || row['Month'] || null,
            pick(row, FIELD_ALIASES.expenseSupplier) ? String(pick(row, FIELD_ALIASES.expenseSupplier)) : null,
            pick(row, FIELD_ALIASES.expenseCategory) ? String(pick(row, FIELD_ALIASES.expenseCategory)) : 'Unassigned',
            description ? String(description) : null,
            pick(row, FIELD_ALIASES.division) ? String(pick(row, FIELD_ALIASES.division)) : null,
            pick(row, FIELD_ALIASES.vessel) ? String(pick(row, FIELD_ALIASES.vessel)) : null,
            amount,
          ]
        );
        expenseRows += 1;
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
    return { status: 'imported', filename, sheet: best.sheetName, headerRow: best.header.rowIndex + 1, headerScore: best.header.score, rowsFound: rows.length, importedRows, skippedRows, registerRows, cardRows, salaryRows, vesselRows, expenseRows, sheets: workbook.SheetNames };
  } catch (error) {
    await client.query('ROLLBACK');
    throw error;
  } finally {
    client.release();
  }
}

const EXPORT_DATE_FORMAT = 'yyyy-mm-dd';
const EXPORT_CURRENCY_FORMAT = 'N$ #,##0.00';
function cleanExportRows(rows) { return rows.map((row) => Object.fromEntries(Object.entries(row).map(([key, value]) => [key, value === null || value === undefined ? '' : value]))); }
function applySheetFormats(worksheet, currencyColumns = [], dateColumns = []) {
  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1');
  for (let row = range.s.r + 1; row <= range.e.r; row += 1) {
    for (const col of currencyColumns) { const cell = worksheet[XLSX.utils.encode_cell({ r: row, c: col })]; if (cell && typeof cell.v === 'number') cell.z = EXPORT_CURRENCY_FORMAT; }
    for (const col of dateColumns) { const cell = worksheet[XLSX.utils.encode_cell({ r: row, c: col })]; if (cell && cell.v) cell.z = EXPORT_DATE_FORMAT; }
  }
  worksheet['!cols'] = Array.from({ length: Math.max(1, range.e.c + 1) }, () => ({ wch: 18 }));
}
function appendJsonSheet(workbook, sheetName, rows, currencyColumns = [], dateColumns = []) { const worksheet = XLSX.utils.json_to_sheet(cleanExportRows(rows)); applySheetFormats(worksheet, currencyColumns, dateColumns); XLSX.utils.book_append_sheet(workbook, worksheet, sheetName); }
function sendWorkbook(reply, workbook, filename) { const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }); reply.header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet').header('Content-Disposition', `attachment; filename="${filename}"`).send(buffer); }

async function getJobRegisterExportRows() {
  const result = await pool.query(`SELECT job_number AS "Job Nr", job_date AS "Date", ops_manager AS "Division Head", division AS "Division", client_name AS "Client", description AS "Description", completion_date AS "Completion Date", quote_number AS "Quote Nr", po_number AS "PO No", report_reference AS "Report Reference", invoice_number AS "Invoice No", client_feedback AS "Client Feedback", value_incl_vat AS "Value Incl VAT", value_excl_vat AS "Value Excl VAT", payments_made AS "Payments Made" FROM job_register_entries ORDER BY job_date DESC NULLS LAST, id DESC`);
  return result.rows;
}
async function getJobCardsExportRows() {
  const result = await pool.query(`SELECT job_date AS "Date", month AS "Month", year AS "Year", fy AS "FY", job_number AS "Job Nr", ops_manager AS "OPS Manager", description AS "Description", start_time AS "Start Time", end_time AS "End Time", hours AS "Total Hours", labour_cost AS "Labour Costs", equipment_cost AS "Equipment Costs", workshop_cost AS "Workshop Cost", division AS "Division" FROM job_card_entries ORDER BY job_date DESC NULLS LAST, id DESC`);
  return result.rows;
}
async function getSalariesExportRows() {
  const result = await pool.query(`SELECT salary_year AS "Year", salary_month AS "Month", total_salaries AS "Total Salaries" FROM salaries ORDER BY salary_year DESC, imported_at DESC, id DESC`);
  return result.rows;
}
async function getJobsExportRows() {
  const result = await pool.query(`SELECT job_number AS "Job Nr", invoice_number AS "Invoice No", job_date AS "Date", division AS "Division", ops_manager AS "OPS Manager", client_name AS "Client", description AS "Description", quote_number AS "Quote Nr", po_number AS "PO No", hours AS "Total Hours", revenue AS "Revenue Excl VAT", labour_cost AS "Labour Costs", equipment_cost AS "Equipment Costs", workshop_cost AS "Workshop Cost", total_cost AS "Total Costs", gross_profit AS "Gross Profit" FROM jobs ORDER BY job_date DESC NULLS LAST, id DESC`);
  return result.rows;
}
async function getVesselExportRows() {
  const result = await pool.query(`SELECT job_date AS "Date", month AS "Month", client_name AS "Client", job_number AS "Job Nr", ops_manager AS "OPS Manager", description AS "Description", start_time AS "Start Time", end_time AS "End Time", hours AS "Hours", vessel_name AS "Vessel", labour_cost AS "Cost Labour", equipment_cost AS "Cost Equipment", workshop_cost AS "Workshop Cost", division AS "Division" FROM vessel_entries ORDER BY job_date DESC NULLS LAST, id DESC`);
  return result.rows;
}
async function getExpensesExportRows() {
  const result = await pool.query(`SELECT expense_date AS "Date", month AS "Month", supplier AS "Supplier", category AS "Category", description AS "Description", vessel_name AS "Vessel", division AS "Division", amount AS "Amount" FROM expense_entries ORDER BY expense_date DESC NULLS LAST, id DESC`);
  return result.rows;
}
function buildAnalysisRows(jobs) {
  const totals = jobs.reduce((acc, job) => { acc.jobs += 1; acc.hours += Number(job['Total Hours'] || 0); acc.revenue += Number(job['Revenue Excl VAT'] || 0); acc.labour += Number(job['Labour Costs'] || 0); acc.equipment += Number(job['Equipment Costs'] || 0); acc.workshop += Number(job['Workshop Cost'] || 0); acc.cost += Number(job['Total Costs'] || 0); acc.profit += Number(job['Gross Profit'] || 0); return acc; }, { jobs: 0, hours: 0, revenue: 0, labour: 0, equipment: 0, workshop: 0, cost: 0, profit: 0 });
  return [{ Metric: 'Jobs', Value: totals.jobs }, { Metric: 'Total Hours', Value: totals.hours }, { Metric: 'Revenue Excl VAT', Value: totals.revenue }, { Metric: 'Labour Costs', Value: totals.labour }, { Metric: 'Equipment Costs', Value: totals.equipment }, { Metric: 'Workshop Cost', Value: totals.workshop }, { Metric: 'Total Costs', Value: totals.cost }, { Metric: 'Gross Profit', Value: totals.profit }, { Metric: 'Gross Profit %', Value: totals.revenue ? totals.profit / totals.revenue : 0 }];
}
function buildProfitPerJobRows(jobs) { return jobs.map((job) => { const revenue = Number(job['Revenue Excl VAT'] || 0); const profit = Number(job['Gross Profit'] || 0); return { 'Job Nr': job['Job Nr'], Date: job.Date, Client: job.Client, Division: job.Division, 'OPS Manager': job['OPS Manager'], 'Revenue Excl VAT': revenue, 'Total Costs': Number(job['Total Costs'] || 0), 'Gross Profit': profit, 'GP %': revenue ? profit / revenue : 0 }; }); }
function buildSettingsRows() { return [{ Setting: 'Company', Value: 'B4 Engineering & Diving / Nautilus Operations' }, { Setting: 'Currency', Value: 'N$' }, { Setting: 'Workbook Export Version', Value: 'Phase 1 + Vessel Info + Expenses' }, { Setting: 'Debtors Logic', Value: 'Monitoring only - current invoices treated as paid by default unless payment data exists' }]; }

app.get('/api/export/job-register.xlsx', async (request, reply) => { try { const workbook = XLSX.utils.book_new(); appendJsonSheet(workbook, 'Job Register', await getJobRegisterExportRows(), [12, 13, 14], [1, 6]); return sendWorkbook(reply, workbook, 'job-register.xlsx'); } catch (error) { app.log.error(error); return reply.code(500).send({ error: error.message }); } });
app.get('/api/export/job-cards.xlsx', async (request, reply) => { try { const workbook = XLSX.utils.book_new(); appendJsonSheet(workbook, 'Job Card Conversion', await getJobCardsExportRows(), [10, 11, 12], [0]); return sendWorkbook(reply, workbook, 'job-cards.xlsx'); } catch (error) { app.log.error(error); return reply.code(500).send({ error: error.message }); } });
app.get('/api/export/salaries.xlsx', async (request, reply) => { try { const workbook = XLSX.utils.book_new(); appendJsonSheet(workbook, 'Salaries', await getSalariesExportRows(), [2], []); return sendWorkbook(reply, workbook, 'salaries.xlsx'); } catch (error) { app.log.error(error); return reply.code(500).send({ error: error.message }); } });
app.get('/api/export/vessels.xlsx', async (request, reply) => { try { const workbook = XLSX.utils.book_new(); appendJsonSheet(workbook, 'Vessel Info', await getVesselExportRows(), [10, 11, 12], [0]); return sendWorkbook(reply, workbook, 'vessel-info.xlsx'); } catch (error) { app.log.error(error); return reply.code(500).send({ error: error.message }); } });
app.get('/api/export/expenses.xlsx', async (request, reply) => { try { const workbook = XLSX.utils.book_new(); appendJsonSheet(workbook, 'Expenses', await getExpensesExportRows(), [7], [0]); return sendWorkbook(reply, workbook, 'expenses.xlsx'); } catch (error) { app.log.error(error); return reply.code(500).send({ error: error.message }); } });
app.get('/api/export/full-workbook.xlsx', async (request, reply) => {
  try {
    const workbook = XLSX.utils.book_new();
    const jobRegisterRows = await getJobRegisterExportRows();
    const jobRows = await getJobsExportRows();
    const jobCardRows = await getJobCardsExportRows();
    const vesselRows = await getVesselExportRows();
    const expenseRows = await getExpensesExportRows();
    const salaryRows = await getSalariesExportRows();
    appendJsonSheet(workbook, 'Job Register', jobRegisterRows, [12, 13, 14], [1, 6]);
    appendJsonSheet(workbook, 'Analysis', buildAnalysisRows(jobRows), [1], []);
    appendJsonSheet(workbook, 'Profit per Job', buildProfitPerJobRows(jobRows), [5, 6, 7], [1]);
    appendJsonSheet(workbook, 'Data', jobRows, [10, 11, 12, 13, 14, 15], [2]);
    appendJsonSheet(workbook, 'Jobs', jobRows, [10, 11, 12, 13, 14, 15], [2]);
    appendJsonSheet(workbook, 'Job Card Conversion', jobCardRows, [10, 11, 12], [0]);
    appendJsonSheet(workbook, 'Vessel Info', vesselRows, [10, 11, 12], [0]);
    appendJsonSheet(workbook, 'Expenses', expenseRows, [7], [0]);
    appendJsonSheet(workbook, 'Settings', buildSettingsRows(), [], []);
    appendJsonSheet(workbook, 'Salaries', salaryRows, [2], []);
    return sendWorkbook(reply, workbook, 'b4-nautilus-full-workbook.xlsx');
  } catch (error) { app.log.error(error); return reply.code(500).send({ error: error.message }); }
});

app.post('/api/imports/excel', async (request, reply) => {
  try {
    const data = await request.file();
    if (!data) return reply.code(400).send({ error: 'No file uploaded.' });
    const chunks = [];
    for await (const chunk of data.file) chunks.push(chunk);
    const buffer = Buffer.concat(chunks);
    const result = await importWorkbook(buffer, data.filename);
    return result;
  } catch (error) {
    app.log.error(error);
    return reply.code(500).send({ error: error.message });
  }
});

app.setNotFoundHandler((request, reply) => {
  if (request.raw.url?.startsWith('/api/')) return reply.code(404).send({ error: 'Not found' });
  return reply.sendFile('index.html');
});

try {
  await initDatabase();
  await app.listen({ port, host });
} catch (error) {
  app.log.error(error);
  process.exit(1);
}