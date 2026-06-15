import Fastify from 'fastify';
import fastifyStatic from '@fastify/static';
import fastifyMultipart from '@fastify/multipart';
import pg from 'pg';
import XLSX from 'xlsx';
import { fileURLToPath } from 'url';
import path from 'path';
import crypto from 'crypto';

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


const AUTH_SESSIONS = new Map();
const DEFAULT_USERS = [
  { name: 'Jo-Ann', username: 'jo-ann', pin: '9999', role: 'ADMIN' },
  { name: 'Joann', username: 'joann', pin: '2468', role: 'ADMIN' },
  { name: 'Bookkeeper', username: 'bookkeeper', pin: '1357', role: 'BOOKKEEPER' },
  { name: 'Jolanda', username: 'jolanda', pin: '1111', role: 'OPS_MANAGER' },
  { name: 'Eugen', username: 'eugen', pin: '2222', role: 'OPS_MANAGER' },
  { name: 'Carmon', username: 'carmon', pin: '3333', role: 'OPS_MANAGER' },
  { name: 'Bresler', username: 'bresler', pin: '4444', role: 'OPS_MANAGER' },
  { name: 'Supervisor', username: 'supervisor', pin: '5555', role: 'SUPERVISOR' },
  { name: 'Crew', username: 'crew', pin: '0000', role: 'CREW' },
];
const FINANCE_ROLES = new Set(['ADMIN', 'BOOKKEEPER', 'OPS_MANAGER']);
const ADMIN_ROLES = new Set(['ADMIN', 'BOOKKEEPER']);
const JOB_CARD_ROLES = new Set(['ADMIN', 'BOOKKEEPER', 'OPS_MANAGER', 'SUPERVISOR', 'CREW']);

const STAFF_CREW_USERS = [
  'A B de Beer',
  'W Van Wyk',
  'J Titus',
  'B Moses',
  'E Matamu',
  'K Luhl',
  'H Hamupunda',
  'L Rhode',
  'F De Barros',
  'E Green',
  'D De Waal',
  'D Lackey',
  'L Naweseb',
  'D Nickel',
  'J Kennedy',
  'L Visagie',
  'M Benz',
  'M Job',
  'J Shaanyenge',
  'I Armandu',
  'S Titus',
  'M S Siluka',
  'J M Muyenga',
  'M Matheus',
  'R Coetzee',
  'B E Masuku',
  'E Uarije',
  'A Thiele',
  'E Croza',
  'J Loock',
  'C September',
  'F Kavera',
  'L Dias',
  'P Swartz',
  'J Rootman',
  'A Morkel',
  'M C Nel',
  'T Joel',
  'H Matias',
  'A Ndanyengwa',
  'W Samuel',
  'R Hendricks',
  'I Shifela',
  'A Fredericks',
  'I Amulungu',
  'B Bester',
  'S Ntinda',
  'E Wildschudt',
  'A Hammond',
  'R Lensley',
  'N Botha',
  'P Sikongo',
  'T Mouton',
  'P Shekudja',
  'W Shilunga',
  'V Sakkeus',
  'T Haulyowike',
  'J Olivier',
  'R Simushi',
  'E Fourie',
  'T De Beer',
  'S Ugulu',
  'L Haikali',
  'C Shimbilinga',
  'F Shindongo',
  'E Haikela',
  'A Katambo',
  'H Goseb',
  'E C Mouton',
  'Z Van Wyk',
  'T van Wyk',
  'Y Hekulanu',
  'J Steyn',
  'G Hoeses',
  'D Wallis',
  'I Swartz',
  'S Jerobeam'
];
function usernameFromName(name) {
  return String(name || '').toLowerCase().replace(/[^a-z0-9]+/g, '.').replace(/^\.+|\.+$/g, '');
}


function normalizeUsername(value) { return String(value || '').trim().toLowerCase(); }
function makeToken() { return crypto.randomBytes(32).toString('hex'); }
function cookieValue(header, name) {
  return String(header || '').split(';').map(v => v.trim()).find(v => v.startsWith(name + '='))?.slice(name.length + 1) || '';
}
function tokenFromRequest(request) {
  const auth = String(request.headers.authorization || '');
  if (auth.toLowerCase().startsWith('bearer ')) return auth.slice(7).trim();
  return cookieValue(request.headers.cookie, 'b4_token');
}
function currentUser(request) {
  const token = tokenFromRequest(request);
  return token ? AUTH_SESSIONS.get(token) || null : null;
}
function safeUser(user) { return user ? { username: user.username, name: user.name, role: user.role } : null; }
function hasRole(user, roles) { return Boolean(user && roles.has(user.role)); }
function forbidden(reply, message = 'You do not have permission for this action.') { return reply.code(403).send({ error: message }); }
function unauthorized(reply) { return reply.code(401).send({ error: 'Please log in.' }); }
function isPublicApi(url) {
  return url.startsWith('/api/status') || url.startsWith('/api/auth/login') || url.startsWith('/api/auth/me') || url.startsWith('/api/auth/logout');
}
function safeJobRow(row, user) {
  if (hasRole(user, FINANCE_ROLES)) return row;
  return {
    id: row.id,
    job_number: row.job_number,
    job_date: row.job_date,
    division: row.division,
    ops_manager: row.ops_manager,
    description: row.description,
    client_name: row.client_name,
    quote_number: row.quote_number,
    po_number: row.po_number,
    hours: row.hours,
  };
}
function safeCardRow(row, user) {
  if (hasRole(user, FINANCE_ROLES)) return row;
  return {
    id: row.id,
    job_date: row.job_date,
    month: row.month,
    year: row.year,
    fy: row.fy,
    job_number: row.job_number,
    ops_manager: row.ops_manager,
    description: row.description,
    start_time: row.start_time,
    end_time: row.end_time,
    hours: row.hours,
    division: row.division,
    client_name: row.client_name,
    location: row.location,
    source: row.source,
    vessel_name: row.vessel_name,
    contact_person: row.contact_person,
    activity_type: row.activity_type,
    staff_names: row.staff_names,
    equipment_used: row.equipment_used,
    completed_by_name: row.completed_by_name,
  };
}

const FIELD_ALIASES = {
  jobNumber: ['Job Nr','Job No','Job Number','Job','Job Nr.','Job #','Job_No','Job_Nr','JobNo','JobNr'],
  invoiceNumber: ['Invoice No','Invoice Number','Invoice','Invoice No.','Inv No','Invoice_No','Invoice No.'],
  jobDate: ['Date','Job Date','Start Date'],
  division: ['Main Division','Division','Division ','Divisio'],
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
  jobExpenses: ['Job Expenses','Job Expense','Job_Expenses',' Job Expenses'],
  officeExpenses: ['Office Expenses','Office Expense','Office_Expenses','Expenses NO JOB #','Expenses No Job #'],
  salaryCost: ['Salaries','Salary','Salary Cost','Salary_Cost'],
  totalCost: ['Total Costs','Total Cost',' Total_Costs'],
  paymentsMade: ['Payments Made','Payment Made','Amount Paid','Paid','Payments','Receipts','payments_made','payment_made','amount_paid'],
  expenseDate: ['Date','Expense Date','Transaction Date','Month'],
  expenseSupplier: ['Supplier','Vendor','Paid To','Creditor','Company'],
  expenseCategory: ['Category','Expense Category','Type','Expense Type','Account'],
  expenseDescription: ['Description','Details','Narrative','Item','Expense Description'],
  expenseAmount: ['Total (without Tax)','Total Without Tax','Excl VAT','Total (with Tax)','Total With Tax','Incl VAT','Amount','Value','Cost','Expense','Total','Total Cost'],
  expenseNoJob: ['Expenses NO JOB #','Expenses No Job #','Expenses No Job','Expenses without Job','No Job Expense'],
  vessel: ['Vessel','Boat','Vessel Name'],
};

const VESSEL_COLUMNS = ['Dingy', 'Ingwegwe', 'DSV Saturn', 'Flatcat', 'Ingwena', 'Pumba'];
const KNOWN_OPS_MANAGERS = ['Jolanda', 'Eugen', 'Carmon', 'Bresler', 'Elliotte'];

function normalizeHeader(value) { return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, ''); }
function toNumber(value) { if (value === null || value === undefined || value === '') return 0; if (typeof value === 'number') return Number.isFinite(value) ? value : 0; const parsed = Number(String(value).replace(/[^0-9.-]/g, '')); return Number.isFinite(parsed) ? parsed : 0; }
function toDate(value) { if (!value) return null; if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10); if (typeof value === 'number') { const parsed = XLSX.SSF.parse_date_code(value); if (parsed) return `${parsed.y}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`; } const date = new Date(value); return Number.isNaN(date.getTime()) ? null : date.toISOString().slice(0, 10); }
function toTime(value) { if (value === null || value === undefined || value === '') return null; if (typeof value === 'number') { const totalMinutes = Math.round((value % 1) * 24 * 60); return `${String(Math.floor(totalMinutes / 60)).padStart(2, '0')}:${String(totalMinutes % 60).padStart(2, '0')}`; } if (value instanceof Date && !Number.isNaN(value.getTime())) return `${String(value.getUTCHours()).padStart(2, '0')}:${String(value.getUTCMinutes()).padStart(2, '0')}`; const match = String(value).match(/(\d{1,2}:\d{2})/); return match ? match[1] : String(value); }
function toHours(value) { if (value === null || value === undefined || value === '') return 0; if (typeof value === 'number') return value > 0 && value < 1 ? value * 24 : value; if (value instanceof Date && !Number.isNaN(value.getTime())) return value.getUTCHours() + value.getUTCMinutes() / 60 + value.getUTCSeconds() / 3600; const text = String(value).trim(); const time = text.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/); if (time) return Number(time[1]) + Number(time[2]) / 60 + Number(time[3] || 0) / 3600; return toNumber(value); }
function pick(row, names) { const keys = Object.keys(row); const lookup = new Map(keys.map((key) => [normalizeHeader(key), key])); for (const wanted of names) { const key = lookup.get(normalizeHeader(wanted)); if (key) return row[key]; } return undefined; }
function headerScore(headers) { const normalized = headers.map(normalizeHeader); let score = 0; for (const aliases of Object.values(FIELD_ALIASES)) if (aliases.some((alias) => normalized.includes(normalizeHeader(alias)))) score += 1; return score; }
function findHeaderRow(sheet) { const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false }); let best = { rowIndex: 0, score: 0, headers: [] }; matrix.slice(0, 80).forEach((row, rowIndex) => { const score = headerScore(row); if (score > best.score) best = { rowIndex, score, headers: row }; }); return best; }
function rowsFromSheet(workbook, sheetName) { const sheet = workbook.Sheets[sheetName]; if (!sheet) return []; const header = findHeaderRow(sheet); return XLSX.utils.sheet_to_json(sheet, { defval: '', range: header.rowIndex }); }
function chooseBestSheet(workbook) { let best = null; for (const sheetName of workbook.SheetNames) { const sheet = workbook.Sheets[sheetName]; const header = findHeaderRow(sheet); const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: false }); const nonEmptyRows = rows.filter((row) => row.some((cell) => String(cell || '').trim() !== '')).length; const result = { sheetName, sheet, header, nonEmptyRows }; if (!best || header.score > best.header.score || (header.score === best.header.score && nonEmptyRows > best.nonEmptyRows)) best = result; } return best; }
function findSheet(workbook, candidates) { const sheetNames = workbook.SheetNames.map((name) => ({ name, key: normalizeHeader(name) })); for (const candidate of candidates) { const target = normalizeHeader(candidate); const exact = sheetNames.find((sheet) => sheet.key === target); if (exact) return exact.name; } for (const candidate of candidates) { const target = normalizeHeader(candidate); const partial = sheetNames.find((sheet) => sheet.key.includes(target) || target.includes(sheet.key)); if (partial) return partial.name; } return null; }
function isUsed(value) { if (value === true) return true; if (typeof value === 'number') return value === 1; return ['1', 'yes', 'y', 'true', 'x'].includes(String(value || '').trim().toLowerCase()); }

function normalizeVesselName(value) { const text = String(value || '').trim(); const key = normalizeHeader(text); if (!key) return null; if (key === 'saturn' || key === 'dsvsaturn') return 'DSV Saturn'; if (key === 'ingwegwe') return 'Ingwegwe'; if (key === 'ingwena') return 'Ingwena'; if (key === 'dingy' || key === 'dinghy') return 'Dingy'; if (key === 'flatcat') return 'Flatcat'; if (key === 'pumba' || key === 'pumbaa') return 'Pumba'; return text; }
function isVesselName(value) { return Boolean(normalizeVesselName(value) && ['Dingy','Ingwegwe','DSV Saturn','Flatcat','Ingwena','Pumba'].includes(normalizeVesselName(value))); }
function cleanDivision(value) { const text = String(value || '').trim(); return isVesselName(text) ? null : (text || null); }
function isCoreDivision(value) { return ['pollution','diving','civils','wsmaintenance'].includes(normalizeHeader(value)); }
function chooseFinancialSheet(workbook) { const preferred = findSheet(workbook, ['Data', 'Jobs', 'Analysis', 'Profit per Job']); if (preferred && !['Vessel Info', 'Vessels', 'Expenses', 'Expense'].map(normalizeHeader).includes(normalizeHeader(preferred))) return { sheetName: preferred, sheet: workbook.Sheets[preferred], header: findHeaderRow(workbook.Sheets[preferred]) }; let best = null; for (const sheetName of workbook.SheetNames) { if (['Vessel Info', 'VesselInfo', 'Vessels', 'Expenses', 'Expense', 'Salaries', 'Salary'].map(normalizeHeader).includes(normalizeHeader(sheetName))) continue; const sheet = workbook.Sheets[sheetName]; const header = findHeaderRow(sheet); const result = { sheetName, sheet, header }; if (!best || header.score > best.header.score) best = result; } return best || chooseBestSheet(workbook); }
function importKey(parts) { return parts.map((part) => String(part ?? '').trim().toLowerCase()).join('|'); }
function coalesceText(...values) { for (const value of values) { const text = String(value ?? '').trim(); if (text) return text; } return null; }
function detectExpenseVessel(row, description) { const direct = normalizeVesselName(pick(row, FIELD_ALIASES.vessel)); if (direct && isVesselName(direct)) return direct; const text = `${pick(row, FIELD_ALIASES.expenseCategory) || ''} ${row['Placement'] || ''} ${description || ''}`; const key = normalizeHeader(text); if (key.includes('dsvsaturn') || key.includes('saturn')) return 'DSV Saturn'; if (key.includes('ingwegwe')) return 'Ingwegwe'; if (key.includes('ingwena')) return 'Ingwena'; if (key.includes('dinghy') || key.includes('dingy')) return 'Dingy'; if (key.includes('flatcat')) return 'Flatcat'; if (key.includes('pumba')) return 'Pumba'; return null; }

async function initDatabase() {
  if (!process.env.DATABASE_URL) { app.log.warn('DATABASE_URL is not set.'); return; }
  await pool.query(`CREATE TABLE IF NOT EXISTS app_users (id SERIAL PRIMARY KEY, username TEXT UNIQUE NOT NULL, name TEXT NOT NULL, pin TEXT NOT NULL, role TEXT NOT NULL, is_active BOOLEAN DEFAULT TRUE, created_at TIMESTAMP DEFAULT NOW());`);
  for (const u of DEFAULT_USERS) {
    await pool.query(`INSERT INTO app_users (username, name, pin, role, is_active) VALUES ($1,$2,$3,$4,TRUE) ON CONFLICT (username) DO UPDATE SET name=EXCLUDED.name, role=EXCLUDED.role, is_active=TRUE`, [u.username, u.name, u.pin, u.role]);
  }
  for (const name of STAFF_CREW_USERS) {
    await pool.query(`INSERT INTO app_users (username, name, pin, role, is_active) VALUES ($1,$2,$3,'CREW',TRUE) ON CONFLICT (username) DO NOTHING`, [usernameFromName(name), name, '0000']);
  }
  await pool.query(`CREATE TABLE IF NOT EXISTS jobs (id SERIAL PRIMARY KEY, job_number TEXT, invoice_number TEXT, job_date DATE, division VARCHAR(100), ops_manager VARCHAR(100), hours NUMERIC, revenue NUMERIC, labour_cost NUMERIC, equipment_cost NUMERIC, workshop_cost NUMERIC, total_cost NUMERIC, gross_profit NUMERIC, created_at TIMESTAMP DEFAULT NOW());`);
  await pool.query(`ALTER TABLE jobs ADD COLUMN IF NOT EXISTS description TEXT;`);
  await pool.query(`ALTER TABLE jobs ADD COLUMN IF NOT EXISTS client_name TEXT;`);
  await pool.query(`ALTER TABLE jobs ADD COLUMN IF NOT EXISTS quote_number TEXT;`);
  await pool.query(`ALTER TABLE jobs ADD COLUMN IF NOT EXISTS po_number TEXT;`);
  await pool.query(`ALTER TABLE jobs ADD COLUMN IF NOT EXISTS job_expenses NUMERIC DEFAULT 0;`);
  await pool.query(`ALTER TABLE jobs ADD COLUMN IF NOT EXISTS office_expenses NUMERIC DEFAULT 0;`);
  await pool.query(`ALTER TABLE jobs ADD COLUMN IF NOT EXISTS salary_cost NUMERIC DEFAULT 0;`);
  await pool.query(`CREATE TABLE IF NOT EXISTS imports (id SERIAL PRIMARY KEY, filename TEXT, imported_rows INTEGER DEFAULT 0, status VARCHAR(50) DEFAULT 'completed', imported_at TIMESTAMP DEFAULT NOW());`);
  await pool.query(`ALTER TABLE imports ADD COLUMN IF NOT EXISTS register_rows INTEGER DEFAULT 0;`);
  await pool.query(`ALTER TABLE imports ADD COLUMN IF NOT EXISTS card_rows INTEGER DEFAULT 0;`);
  await pool.query(`ALTER TABLE imports ADD COLUMN IF NOT EXISTS vessel_rows INTEGER DEFAULT 0;`);
  await pool.query(`ALTER TABLE imports ADD COLUMN IF NOT EXISTS expense_rows INTEGER DEFAULT 0;`);
  await pool.query(`ALTER TABLE imports ADD COLUMN IF NOT EXISTS salary_rows INTEGER DEFAULT 0;`);
  await pool.query(`ALTER TABLE imports ADD COLUMN IF NOT EXISTS skipped_rows INTEGER DEFAULT 0;`);
  await pool.query(`ALTER TABLE imports ADD COLUMN IF NOT EXISTS ws_rows INTEGER DEFAULT 0;`);
  // Existing Railway databases may still have VARCHAR(50) from older builds.
  // Long invoice lists / job references from the workbook can exceed 50 chars, so widen safely before import.
  for (const stmt of [
    `ALTER TABLE jobs ALTER COLUMN job_number TYPE TEXT`,
    `ALTER TABLE jobs ALTER COLUMN invoice_number TYPE TEXT`,
    `ALTER TABLE job_register_entries ALTER COLUMN job_number TYPE TEXT`,
    `ALTER TABLE job_register_entries ALTER COLUMN invoice_number TYPE TEXT`,
    `ALTER TABLE job_card_entries ALTER COLUMN job_number TYPE TEXT`,
    `ALTER TABLE vessel_entries ALTER COLUMN job_number TYPE TEXT`,
    `ALTER TABLE imports ALTER COLUMN filename TYPE TEXT`
  ]) {
    try { await pool.query(stmt); } catch (e) { app.log.warn({ err: e.message, stmt }, 'column widen skipped'); }
  }
  await pool.query(`CREATE TABLE IF NOT EXISTS job_register_entries (id SERIAL PRIMARY KEY, job_number TEXT, job_date DATE, ops_manager TEXT, division TEXT, client_name TEXT, description TEXT, completion_date DATE, quote_number TEXT, po_number TEXT, report_reference TEXT, invoice_number TEXT, client_feedback TEXT, value_incl_vat NUMERIC DEFAULT 0, value_excl_vat NUMERIC DEFAULT 0, payments_made NUMERIC DEFAULT 0, imported_at TIMESTAMP DEFAULT NOW());`);
  await pool.query(`ALTER TABLE job_register_entries ADD COLUMN IF NOT EXISTS value_incl_vat NUMERIC DEFAULT 0;`);
  await pool.query(`ALTER TABLE job_register_entries ADD COLUMN IF NOT EXISTS value_excl_vat NUMERIC DEFAULT 0;`);
  await pool.query(`ALTER TABLE job_register_entries ADD COLUMN IF NOT EXISTS payments_made NUMERIC DEFAULT 0;`);
  await pool.query(`CREATE TABLE IF NOT EXISTS job_card_entries (id SERIAL PRIMARY KEY, job_date DATE, month TEXT, year INTEGER, fy TEXT, job_number TEXT, ops_manager TEXT, description TEXT, start_time TEXT, end_time TEXT, hours NUMERIC, labour_cost NUMERIC, equipment_cost NUMERIC, workshop_cost NUMERIC, division TEXT, imported_at TIMESTAMP DEFAULT NOW());`);
  await pool.query(`CREATE TABLE IF NOT EXISTS salaries (id SERIAL PRIMARY KEY, salary_year INTEGER, salary_month TEXT, total_salaries NUMERIC, imported_at TIMESTAMP DEFAULT NOW());`);
  await pool.query(`CREATE TABLE IF NOT EXISTS vessel_entries (id SERIAL PRIMARY KEY, job_date DATE, month TEXT, client_name TEXT, job_number TEXT, ops_manager TEXT, description TEXT, start_time TEXT, end_time TEXT, hours NUMERIC, vessel_name TEXT, labour_cost NUMERIC DEFAULT 0, equipment_cost NUMERIC DEFAULT 0, workshop_cost NUMERIC DEFAULT 0, division TEXT, imported_at TIMESTAMP DEFAULT NOW());`);
  await pool.query(`CREATE INDEX IF NOT EXISTS vessel_entries_job_date_idx ON vessel_entries (job_date);`);
  await pool.query(`CREATE INDEX IF NOT EXISTS vessel_entries_vessel_name_idx ON vessel_entries (vessel_name);`);
  await pool.query(`CREATE TABLE IF NOT EXISTS expense_entries (id SERIAL PRIMARY KEY, expense_date DATE, month TEXT, supplier TEXT, category TEXT, description TEXT, division TEXT, vessel_name TEXT, amount NUMERIC DEFAULT 0, imported_at TIMESTAMP DEFAULT NOW());`);
  await pool.query(`ALTER TABLE expense_entries ADD COLUMN IF NOT EXISTS amount_excl_vat NUMERIC DEFAULT 0;`);
  await pool.query(`ALTER TABLE expense_entries ADD COLUMN IF NOT EXISTS amount_incl_vat NUMERIC DEFAULT 0;`);
  await pool.query(`UPDATE expense_entries SET amount_excl_vat = amount WHERE COALESCE(amount_excl_vat,0) = 0 AND COALESCE(amount,0) <> 0;`);
  await pool.query(`WITH ranked AS (
    SELECT id, ROW_NUMBER() OVER (
      PARTITION BY
        COALESCE(expense_date::text,''),
        LOWER(TRIM(COALESCE(supplier,''))),
        LOWER(TRIM(COALESCE(category,''))),
        LOWER(TRIM(COALESCE(description,''))),
        LOWER(TRIM(COALESCE(division,''))),
        LOWER(TRIM(COALESCE(vessel_name,''))),
        ROUND(COALESCE(NULLIF(amount_excl_vat,0), amount,0)::numeric, 2),
        ROUND(COALESCE(NULLIF(amount_incl_vat,0),0)::numeric, 2)
      ORDER BY id DESC
    ) AS rn
    FROM expense_entries
  )
  DELETE FROM expense_entries e USING ranked r WHERE e.id = r.id AND r.rn > 1;`);
  await pool.query(`CREATE INDEX IF NOT EXISTS expense_entries_expense_date_idx ON expense_entries (expense_date);`);
  await pool.query(`CREATE INDEX IF NOT EXISTS expense_entries_category_idx ON expense_entries (category);`);
  await pool.query(`ALTER TABLE job_card_entries ADD COLUMN IF NOT EXISTS import_key TEXT;`);
  await pool.query(`ALTER TABLE job_card_entries ADD COLUMN IF NOT EXISTS client_name TEXT;`);
  await pool.query(`ALTER TABLE job_card_entries ADD COLUMN IF NOT EXISTS location TEXT;`);
  await pool.query(`ALTER TABLE job_card_entries ADD COLUMN IF NOT EXISTS material_cost NUMERIC DEFAULT 0;`);
  await pool.query(`ALTER TABLE job_card_entries ADD COLUMN IF NOT EXISTS other_cost NUMERIC DEFAULT 0;`);
  await pool.query(`ALTER TABLE job_card_entries ADD COLUMN IF NOT EXISTS total_cost NUMERIC DEFAULT 0;`);
  await pool.query(`ALTER TABLE job_card_entries ADD COLUMN IF NOT EXISTS source TEXT DEFAULT 'import';`);
  await pool.query(`ALTER TABLE job_card_entries ADD COLUMN IF NOT EXISTS job_link_status TEXT DEFAULT 'linked';`);
  await pool.query(`ALTER TABLE job_card_entries ADD COLUMN IF NOT EXISTS linked_job_found BOOLEAN DEFAULT TRUE;`);
  await pool.query(`ALTER TABLE job_card_entries ADD COLUMN IF NOT EXISTS vessel_name TEXT;`);
  await pool.query(`ALTER TABLE job_card_entries ADD COLUMN IF NOT EXISTS contact_person TEXT;`);
  await pool.query(`ALTER TABLE job_card_entries ADD COLUMN IF NOT EXISTS activity_type TEXT;`);
  await pool.query(`ALTER TABLE job_card_entries ADD COLUMN IF NOT EXISTS staff_names TEXT;`);
  await pool.query(`ALTER TABLE job_card_entries ADD COLUMN IF NOT EXISTS equipment_used TEXT;`);
  await pool.query(`ALTER TABLE job_card_entries ADD COLUMN IF NOT EXISTS completed_by_user TEXT;`);
  await pool.query(`ALTER TABLE job_card_entries ADD COLUMN IF NOT EXISTS completed_by_name TEXT;`);
  await pool.query(`ALTER TABLE job_card_entries ADD COLUMN IF NOT EXISTS review_status TEXT DEFAULT 'pending';`);
  await pool.query(`ALTER TABLE job_card_entries ADD COLUMN IF NOT EXISTS reviewed_by TEXT;`);
  await pool.query(`ALTER TABLE job_card_entries ADD COLUMN IF NOT EXISTS reviewed_at TIMESTAMP;`);
  await pool.query(`UPDATE job_card_entries SET review_status='pending' WHERE COALESCE(source,'import')='manual' AND (review_status IS NULL OR review_status='');`);
  await pool.query(`UPDATE job_card_entries SET review_status='accepted' WHERE COALESCE(source,'import')<>'manual' AND (review_status IS NULL OR review_status='');`);
  await pool.query(`ALTER TABLE vessel_entries ADD COLUMN IF NOT EXISTS import_key TEXT;`);
  await pool.query(`ALTER TABLE expense_entries ADD COLUMN IF NOT EXISTS import_key TEXT;`);
  await pool.query(`DELETE FROM jobs a USING jobs b WHERE a.id < b.id AND a.job_number IS NOT NULL AND b.job_number IS NOT NULL AND a.job_number = b.job_number;`);
  await pool.query(`DELETE FROM job_register_entries a USING job_register_entries b WHERE a.id < b.id AND a.job_number IS NOT NULL AND b.job_number IS NOT NULL AND a.job_number = b.job_number;`);
  await pool.query(`DELETE FROM job_card_entries a USING job_card_entries b WHERE a.id < b.id AND a.import_key IS NOT NULL AND b.import_key IS NOT NULL AND a.import_key = b.import_key;`);
  await pool.query(`DELETE FROM vessel_entries a USING vessel_entries b WHERE a.id < b.id AND a.import_key IS NOT NULL AND b.import_key IS NOT NULL AND a.import_key = b.import_key;`);
  await pool.query(`DELETE FROM expense_entries a USING expense_entries b WHERE a.id < b.id AND a.import_key IS NOT NULL AND b.import_key IS NOT NULL AND a.import_key = b.import_key;`);
  await pool.query(`DELETE FROM salaries a USING salaries b WHERE a.id < b.id AND COALESCE(a.salary_year,0) = COALESCE(b.salary_year,0) AND COALESCE(a.salary_month,'') = COALESCE(b.salary_month,'');`);
  await pool.query(`CREATE UNIQUE INDEX IF NOT EXISTS jobs_job_number_unique ON jobs (job_number) WHERE job_number IS NOT NULL;`);
  await pool.query(`CREATE UNIQUE INDEX IF NOT EXISTS job_register_job_number_unique ON job_register_entries (job_number) WHERE job_number IS NOT NULL;`);
  await pool.query(`CREATE UNIQUE INDEX IF NOT EXISTS job_cards_import_key_unique ON job_card_entries (import_key) WHERE import_key IS NOT NULL;`);
  await pool.query(`CREATE UNIQUE INDEX IF NOT EXISTS vessel_entries_import_key_unique ON vessel_entries (import_key) WHERE import_key IS NOT NULL;`);
  await pool.query(`CREATE UNIQUE INDEX IF NOT EXISTS expense_entries_import_key_unique ON expense_entries (import_key) WHERE import_key IS NOT NULL;`);
  await pool.query(`CREATE UNIQUE INDEX IF NOT EXISTS salaries_year_month_unique ON salaries (salary_year, salary_month);`);
}

await app.register(fastifyMultipart, { limits: { fileSize: 30 * 1024 * 1024 } });
await app.register(fastifyStatic, { root: path.join(__dirname, 'public'), prefix: '/' });


app.addHook('preHandler', async (request, reply) => {
  const url = request.raw.url || '';
  if (!url.startsWith('/api/') || isPublicApi(url)) return;
  const user = currentUser(request);
  if (!user) return unauthorized(reply);
  request.user = user;
  if (url.startsWith('/api/dashboard') || url.startsWith('/api/job-profit') || url.startsWith('/api/expenses') || url.startsWith('/api/expense-dashboard') || url.startsWith('/api/salaries') || url.startsWith('/api/vessels') || url.startsWith('/api/vessel-dashboard')) {
    if (!hasRole(user, FINANCE_ROLES)) return forbidden(reply, 'Financial dashboards are restricted to Admin, Bookkeeper and OPS Managers.');
  }
  if (url.startsWith('/api/imports') || url.startsWith('/api/export')) {
    if (!hasRole(user, ADMIN_ROLES)) return forbidden(reply, 'Imports and exports are restricted to Admin and Bookkeeper.');
  }
  if (url.startsWith('/api/job-cards/manual')) {
    if (!hasRole(user, JOB_CARD_ROLES)) return forbidden(reply);
  }
});

app.post('/api/auth/login', async (request, reply) => {
  try {
    const username = normalizeUsername(request.body?.username);
    const pin = String(request.body?.pin || '').trim();
    if (!username || !pin) return reply.code(400).send({ error: 'Username and PIN are required.' });
    const result = await pool.query(`SELECT username, name, role FROM app_users WHERE LOWER(username)=LOWER($1) AND pin=$2 AND is_active=TRUE LIMIT 1`, [username, pin]);
    const user = result.rows[0];
    if (!user) return reply.code(401).send({ error: 'Invalid username or PIN.' });
    const token = makeToken();
    AUTH_SESSIONS.set(token, user);
    reply.header('Set-Cookie', `b4_token=${token}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${60 * 60 * 24 * 14}`);
    return { token, user: safeUser(user) };
  } catch (error) {
    app.log.error(error);
    return reply.code(500).send({ error: error.message });
  }
});
app.get('/api/auth/me', async (request) => ({ user: safeUser(currentUser(request)) }));
app.post('/api/auth/logout', async (request, reply) => {
  const token = tokenFromRequest(request);
  if (token) AUTH_SESSIONS.delete(token);
  reply.header('Set-Cookie', 'b4_token=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0');
  return { status: 'logged_out' };
});

app.get('/health', async () => ({ status: 'ok', app: 'B4 Nautilus Operations' }));
app.get('/api/status', async () => { try { const db = await pool.query('SELECT NOW() AS now'); return { app: 'B4 Nautilus Operations', status: 'running', database: 'connected', time: db.rows[0].now }; } catch (error) { return { app: 'B4 Nautilus Operations', status: 'running', database: 'not connected', error: error.message }; } });
app.get('/api/jobs', async (request) => { const limit = Math.min(Number(request.query?.limit || 5000), 5000); const result = await pool.query(`SELECT * FROM jobs ORDER BY job_date DESC NULLS LAST, id DESC LIMIT $1`, [limit]); return result.rows.map(row => safeJobRow(row, request.user)); });
app.get('/api/dashboard', async () => {
  const totals = await pool.query(`SELECT COALESCE(SUM(revenue),0) AS revenue, COALESCE(SUM(labour_cost),0) AS labour_cost, COALESCE(SUM(equipment_cost),0) AS equipment_cost, COALESCE(SUM(workshop_cost),0) AS workshop_cost, COALESCE(SUM(job_expenses),0) AS job_expenses, COALESCE(SUM(office_expenses),0) AS office_expenses, COALESCE(SUM(salary_cost),0) AS salary_cost, COALESCE(SUM(total_cost),0) AS total_cost, COALESCE(SUM(gross_profit),0) AS gross_profit, COUNT(*) AS jobs FROM jobs`);
  const expenses = await pool.query(`SELECT COALESCE(SUM(COALESCE(NULLIF(amount_excl_vat,0), amount)),0) AS expenses, COALESCE(SUM(amount_excl_vat),0) AS expenses_excl_vat, COALESCE(SUM(amount_incl_vat),0) AS expenses_incl_vat FROM expense_entries`);
  const salaryTotals = await pool.query(`SELECT COALESCE(SUM(total_salaries),0) AS salaries FROM salaries`);
  const manualCards = await pool.query(`SELECT COALESCE(SUM(total_cost),0) AS manual_card_cost FROM job_card_entries WHERE source='manual' AND COALESCE(review_status,'pending')='accepted'`);
  const base = totals.rows[0];
  const importedExpenseTotal = Number(expenses.rows[0]?.expenses || 0);
  const sheetOfficeExpenses = Number(base.office_expenses || 0);
  const expenseTotal = sheetOfficeExpenses || importedExpenseTotal;
  const expenseExclVat = expenseTotal;
  const expenseInclVat = Number(expenses.rows[0]?.expenses_incl_vat || 0);
  const salaryTotal = Number(base.salary_cost || 0) || Number(salaryTotals.rows[0]?.salaries || 0);
  const manualCardCost = Number(manualCards.rows[0]?.manual_card_cost || 0);
  const revenue = Number(base.revenue || 0);
  // Match the workbook dashboard: Data!Total_Costs already includes direct costs, job expenses, office expenses and salaries.
  const totalCostInclOverheads = Number(base.total_cost || 0) + manualCardCost;
  const directJobCost = Number(base.labour_cost || 0) + Number(base.equipment_cost || 0) + Number(base.workshop_cost || 0) + Number(base.job_expenses || 0) + manualCardCost;
  const netProfit = revenue - totalCostInclOverheads;
  base.expenses = expenseTotal;
  base.expenses_excl_vat = expenseExclVat;
  base.expenses_incl_vat = expenseInclVat;
  base.salaries = salaryTotal;
  base.job_cost = directJobCost;
  base.manual_card_cost = manualCardCost;
  base.total_cost_incl_overheads = totalCostInclOverheads;
  base.net_profit = netProfit;
  base.net_profit_pct = revenue ? netProfit / revenue : 0;
  const divisions = await pool.query(`SELECT division, COALESCE(SUM(gross_profit),0) AS gross_profit FROM jobs WHERE division IN ('Pollution','Diving','Civils','WS Maintenance') GROUP BY division ORDER BY gross_profit DESC`);
  const opsManagers = await pool.query(`SELECT ops_manager, COALESCE(SUM(gross_profit),0) AS gross_profit FROM jobs WHERE ops_manager = ANY($1) GROUP BY ops_manager ORDER BY gross_profit DESC`, [KNOWN_OPS_MANAGERS]);
  return { totals: base, divisions: divisions.rows, opsManagers: opsManagers.rows };
});
app.get('/api/users', async (request, reply) => { if (!hasRole(request.user, ADMIN_ROLES)) return forbidden(reply); const result = await pool.query(`SELECT username, name, role, is_active FROM app_users ORDER BY role, name`); return result.rows; });
app.get('/api/imports', async () => { const result = await pool.query(`SELECT * FROM imports ORDER BY imported_at DESC LIMIT 50`); return result.rows; });
app.get('/api/job-register', async (request) => { const limit = Math.min(Number(request.query?.limit || 5000), 5000); const result = await pool.query(`SELECT * FROM job_register_entries ORDER BY job_date DESC NULLS LAST, id DESC LIMIT $1`, [limit]); return result.rows; });
app.get('/api/job-cards', async (request) => {
  const limit = Math.min(Number(request.query?.limit || 5000), 5000);
  const result = await pool.query(`SELECT * FROM job_card_entries ORDER BY job_date DESC NULLS LAST, id DESC LIMIT $1`, [limit]);
  return result.rows.map(row => safeCardRow(row, request.user));
});


app.get('/api/job-cards/pending', async (request, reply) => {
  try {
    if (!hasRole(request.user, FINANCE_ROLES)) return forbidden(reply, 'Only Admin, Bookkeeper and OPS Managers may review job cards.');
    const result = await pool.query(`SELECT * FROM job_card_entries
      WHERE COALESCE(review_status,'pending')='pending'
         OR COALESCE(job_link_status,'')='pending_review'
      ORDER BY imported_at DESC NULLS LAST, id DESC
      LIMIT 1000`);
    return result.rows;
  } catch (error) {
    app.log.error(error);
    return reply.code(500).send({ error: error.message });
  }
});

app.post('/api/job-cards/manual', async (request, reply) => {
  try {
    const body = request.body || {};
    const jobNumber = String(body.job_number || '').trim();
    if (!jobNumber) return reply.code(400).send({ error: 'Job Number is required.' });
    const linked = await pool.query(`SELECT job_number FROM jobs WHERE job_number=$1 UNION SELECT job_number FROM job_register_entries WHERE job_number=$1 LIMIT 1`, [jobNumber]);
    const linkedJobFound = linked.rows.length > 0;
    const jobLinkStatus = linkedJobFound ? 'linked' : 'pending_review';
    const jobDate = toDate(body.job_date) || new Date().toISOString().slice(0, 10);
    const clientName = coalesceText(body.client_name, body.client);
    const location = coalesceText(body.location);
    const contactPerson = coalesceText(body.contact_person);
    const vesselName = normalizeVesselName(body.vessel_name);
    const activityType = coalesceText(body.activity_type);
    const staffNames = Array.isArray(body.staff_names) ? body.staff_names.join(', ') : coalesceText(body.staff_names);
    const equipmentUsed = Array.isArray(body.equipment_used) ? body.equipment_used.join(', ') : coalesceText(body.equipment_used);
    const description = coalesceText(body.description, activityType, 'Manual job card');
    const opsManager = coalesceText(body.ops_manager);
    const division = cleanDivision(body.division);
    const hours = toHours(body.hours);
    // Option A workflow: supervisors/crew capture operational job cards only.
    // Costs are hidden and always zero until an OPS/Admin/Bookkeeper reviews and accepts the card.
    const labourCost = 0;
    const equipmentCost = 0;
    const workshopCost = 0;
    const materialCost = 0;
    const otherCost = 0;
    const completedByUser = request.user?.username || null;
    const completedByName = request.user?.name || null;
    const totalCost = labourCost + equipmentCost + workshopCost + materialCost + otherCost;
    const manualKey = importKey(['manual-job-card', jobNumber, jobDate, vesselName, clientName, location, activityType, description, hours, totalCost]);
    const result = await pool.query(`INSERT INTO job_card_entries
      (job_date, month, year, fy, job_number, ops_manager, description, start_time, end_time, hours, labour_cost, equipment_cost, workshop_cost, material_cost, other_cost, total_cost, division, client_name, location, vessel_name, contact_person, activity_type, staff_names, equipment_used, completed_by_user, completed_by_name, import_key, source, job_link_status, linked_job_found, review_status)
      VALUES ($1,TO_CHAR($1::date,'Month'),EXTRACT(YEAR FROM $1::date),NULL,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19,$20,$21,$22,$23,$24,'manual',$25,$26,'pending')
      ON CONFLICT (import_key) WHERE import_key IS NOT NULL DO UPDATE SET
        ops_manager=EXCLUDED.ops_manager,
        description=EXCLUDED.description,
        start_time=EXCLUDED.start_time,
        end_time=EXCLUDED.end_time,
        hours=EXCLUDED.hours,
        labour_cost=EXCLUDED.labour_cost,
        equipment_cost=EXCLUDED.equipment_cost,
        workshop_cost=EXCLUDED.workshop_cost,
        material_cost=EXCLUDED.material_cost,
        other_cost=EXCLUDED.other_cost,
        total_cost=EXCLUDED.total_cost,
        division=EXCLUDED.division,
        client_name=EXCLUDED.client_name,
        location=EXCLUDED.location,
        vessel_name=EXCLUDED.vessel_name,
        contact_person=EXCLUDED.contact_person,
        activity_type=EXCLUDED.activity_type,
        staff_names=EXCLUDED.staff_names,
        equipment_used=EXCLUDED.equipment_used,
        completed_by_user=EXCLUDED.completed_by_user,
        completed_by_name=EXCLUDED.completed_by_name,
        job_link_status=EXCLUDED.job_link_status,
        linked_job_found=EXCLUDED.linked_job_found,
        review_status='pending',
        reviewed_by=NULL,
        reviewed_at=NULL
      RETURNING *`, [jobDate, jobNumber, opsManager, description, body.start_time || null, body.end_time || null, hours, labourCost, equipmentCost, workshopCost, materialCost, otherCost, totalCost, division, clientName, location, vesselName, contactPerson, activityType, staffNames, equipmentUsed, completedByUser, completedByName, manualKey, jobLinkStatus, linkedJobFound]);
    return result.rows[0];
  } catch (error) {
    app.log.error(error);
    return reply.code(500).send({ error: error.message });
  }
});


app.patch('/api/job-cards/:id/costs', async (request, reply) => {
  try {
    if (!hasRole(request.user, FINANCE_ROLES)) return forbidden(reply, 'Only Admin, Bookkeeper and OPS Managers may add job-card costs.');
    const id = Number(request.params.id);
    if (!Number.isFinite(id)) return reply.code(400).send({ error: 'Valid job card id is required.' });
    const labourCost = toNumber(request.body?.labour_cost);
    const equipmentCost = toNumber(request.body?.equipment_cost);
    const workshopCost = toNumber(request.body?.workshop_cost);
    const materialCost = toNumber(request.body?.material_cost);
    const otherCost = toNumber(request.body?.other_cost);
    const totalCost = labourCost + equipmentCost + workshopCost + materialCost + otherCost;
    const result = await pool.query(`UPDATE job_card_entries SET labour_cost=$1, equipment_cost=$2, workshop_cost=$3, material_cost=$4, other_cost=$5, total_cost=$6, review_status='accepted', reviewed_by=$8, reviewed_at=NOW() WHERE id=$7 RETURNING *`, [labourCost, equipmentCost, workshopCost, materialCost, otherCost, totalCost, id, request.user?.username || null]);
    if (!result.rows[0]) return reply.code(404).send({ error: 'Job card not found.' });
    return result.rows[0];
  } catch (error) {
    app.log.error(error);
    return reply.code(500).send({ error: error.message });
  }
});


app.patch('/api/job-cards/:id/reject', async (request, reply) => {
  try {
    if (!hasRole(request.user, FINANCE_ROLES)) return forbidden(reply, 'Only Admin, Bookkeeper and OPS Managers may review job cards.');
    const id = Number(request.params.id);
    if (!Number.isFinite(id)) return reply.code(400).send({ error: 'Valid job card id is required.' });
    const result = await pool.query(`UPDATE job_card_entries SET review_status='rejected', reviewed_by=$2, reviewed_at=NOW() WHERE id=$1 RETURNING *`, [id, request.user?.username || null]);
    if (!result.rows[0]) return reply.code(404).send({ error: 'Job card not found.' });
    return result.rows[0];
  } catch (error) {
    app.log.error(error);
    return reply.code(500).send({ error: error.message });
  }
});

app.get('/api/job-profit', async (request) => {
  const limit = Math.min(Number(request.query?.limit || 5000), 5000);
  const result = await pool.query(`SELECT
      j.job_number,
      j.invoice_number,
      j.job_date,
      j.client_name,
      j.division,
      j.ops_manager,
      j.description,
      COALESCE(j.revenue,0) AS revenue,
      COALESCE(j.total_cost,0) AS imported_job_cost,
      COALESCE(j.gross_profit,0) AS imported_gross_profit,
      COALESCE(c.imported_card_cost,0) AS imported_card_cost,
      COALESCE(c.manual_card_cost,0) AS manual_card_cost,
      COALESCE(c.total_card_cost,0) AS linked_card_cost,
      COALESCE(c.card_hours,0) AS linked_card_hours,
      COALESCE(c.labour_cost,0) AS linked_labour_cost,
      COALESCE(c.equipment_cost,0) AS linked_equipment_cost,
      COALESCE(c.workshop_cost,0) AS linked_workshop_cost,
      COALESCE(c.material_cost,0) AS linked_material_cost,
      COALESCE(c.other_cost,0) AS linked_other_cost,
      COALESCE(j.revenue,0) - COALESCE(c.total_card_cost,0) AS profit_with_cards,
      CASE WHEN COALESCE(j.revenue,0) = 0 THEN 0 ELSE (COALESCE(j.revenue,0) - COALESCE(c.total_card_cost,0)) / COALESCE(j.revenue,0) END AS margin_with_cards
    FROM jobs j
    LEFT JOIN (
      SELECT job_number,
        SUM(CASE WHEN COALESCE(source,'import')='manual' AND COALESCE(review_status,'pending')='accepted' THEN COALESCE(NULLIF(total_cost,0),COALESCE(labour_cost,0)+COALESCE(equipment_cost,0)+COALESCE(workshop_cost,0)+COALESCE(material_cost,0)+COALESCE(other_cost,0)) ELSE 0 END) AS manual_card_cost,
        SUM(CASE WHEN COALESCE(source,'import')<>'manual' THEN COALESCE(NULLIF(total_cost,0),COALESCE(labour_cost,0)+COALESCE(equipment_cost,0)+COALESCE(workshop_cost,0)+COALESCE(material_cost,0)+COALESCE(other_cost,0)) ELSE 0 END) AS imported_card_cost,
        SUM(CASE WHEN COALESCE(source,'import')<>'manual' OR COALESCE(review_status,'pending')='accepted' THEN COALESCE(NULLIF(total_cost,0),COALESCE(labour_cost,0)+COALESCE(equipment_cost,0)+COALESCE(workshop_cost,0)+COALESCE(material_cost,0)+COALESCE(other_cost,0)) ELSE 0 END) AS total_card_cost,
        SUM(COALESCE(hours,0)) AS card_hours,
        SUM(COALESCE(labour_cost,0)) AS labour_cost,
        SUM(COALESCE(equipment_cost,0)) AS equipment_cost,
        SUM(COALESCE(workshop_cost,0)) AS workshop_cost,
        SUM(COALESCE(material_cost,0)) AS material_cost,
        SUM(COALESCE(other_cost,0)) AS other_cost
      FROM job_card_entries
      GROUP BY job_number
    ) c ON c.job_number = j.job_number
    ORDER BY j.job_date DESC NULLS LAST, j.id DESC
    LIMIT $1`, [limit]);
  return result.rows;
});
app.get('/api/salaries', async () => { const result = await pool.query(`SELECT * FROM salaries ORDER BY salary_year DESC, imported_at DESC`); return result.rows; });
app.get('/api/expenses', async (request) => { const limit = Math.min(Number(request.query?.limit || 5000), 10000); const result = await pool.query(`SELECT * FROM expense_entries ORDER BY expense_date DESC NULLS LAST, id DESC LIMIT $1`, [limit]); return result.rows; });
app.get('/api/expense-dashboard', async () => { const totals = await pool.query(`SELECT COUNT(*) AS entries, COALESCE(SUM(COALESCE(NULLIF(amount_excl_vat,0), amount)),0) AS amount, COALESCE(SUM(amount_excl_vat),0) AS amount_excl_vat, COALESCE(SUM(amount_incl_vat),0) AS amount_incl_vat FROM expense_entries`); const byCategory = await pool.query(`SELECT COALESCE(category,'Unassigned') AS category, COUNT(*) AS entries, COALESCE(SUM(COALESCE(NULLIF(amount_excl_vat,0), amount)),0) AS amount FROM expense_entries GROUP BY category ORDER BY amount DESC`); const byVessel = await pool.query(`SELECT COALESCE(vessel_name,'Unassigned') AS vessel_name, COUNT(*) AS entries, COALESCE(SUM(COALESCE(NULLIF(amount_excl_vat,0), amount)),0) AS amount FROM expense_entries GROUP BY vessel_name ORDER BY amount DESC`); const byMonth = await pool.query(`SELECT TO_CHAR(expense_date,'YYYY-MM') AS month, COUNT(*) AS entries, COALESCE(SUM(COALESCE(NULLIF(amount_excl_vat,0), amount)),0) AS amount FROM expense_entries WHERE expense_date IS NOT NULL GROUP BY TO_CHAR(expense_date,'YYYY-MM') ORDER BY month`); return { totals: totals.rows[0], byCategory: byCategory.rows, byVessel: byVessel.rows, byMonth: byMonth.rows }; });
app.get('/api/vessels', async (request) => { const limit = Math.min(Number(request.query?.limit || 5000), 10000); const result = await pool.query(`SELECT * FROM vessel_entries ORDER BY job_date DESC NULLS LAST, id DESC LIMIT $1`, [limit]); return result.rows; });
app.get('/api/vessel-dashboard', async () => { const totals = await pool.query(`SELECT COUNT(*) AS entries, COUNT(DISTINCT job_number) AS jobs, COALESCE(SUM(hours),0) AS hours, COALESCE(SUM(labour_cost),0) AS labour_cost, COALESCE(SUM(equipment_cost),0) AS equipment_cost, COALESCE(SUM(workshop_cost),0) AS workshop_cost FROM vessel_entries`); const byVessel = await pool.query(`SELECT COALESCE(vessel_name,'Unassigned') AS vessel_name, COUNT(*) AS entries, COUNT(DISTINCT job_number) AS jobs, COALESCE(SUM(hours),0) AS hours, COALESCE(SUM(labour_cost),0) AS labour_cost, COALESCE(SUM(equipment_cost),0) AS equipment_cost, COALESCE(SUM(workshop_cost),0) AS workshop_cost FROM vessel_entries GROUP BY vessel_name ORDER BY hours DESC`); const byClient = await pool.query(`SELECT COALESCE(client_name,'Unassigned') AS client_name, COUNT(*) AS entries, COUNT(DISTINCT job_number) AS jobs, COALESCE(SUM(hours),0) AS hours FROM vessel_entries GROUP BY client_name ORDER BY hours DESC LIMIT 20`); const byOps = await pool.query(`SELECT COALESCE(ops_manager,'Unassigned') AS ops_manager, COUNT(*) AS entries, COUNT(DISTINCT job_number) AS jobs, COALESCE(SUM(hours),0) AS hours FROM vessel_entries GROUP BY ops_manager ORDER BY hours DESC`); const byMonth = await pool.query(`SELECT TO_CHAR(job_date,'YYYY-MM') AS month, COUNT(*) AS entries, COUNT(DISTINCT job_number) AS jobs, COALESCE(SUM(hours),0) AS hours FROM vessel_entries WHERE job_date IS NOT NULL GROUP BY TO_CHAR(job_date,'YYYY-MM') ORDER BY month`); return { totals: totals.rows[0], byVessel: byVessel.rows, byClient: byClient.rows, byOps: byOps.rows, byMonth: byMonth.rows }; });
app.post('/api/imports/test-seed', async () => ({ status: 'disabled', rows: 0, message: 'Seed test data is disabled on this build.' }));

async function importWorkbook(buffer, filename) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const best = chooseFinancialSheet(workbook);
  if (!best || best.header.score < 3) throw new Error('Could not find a valid job data sheet.');
  const rows = XLSX.utils.sheet_to_json(best.sheet, { defval: '', range: best.header.rowIndex });
  const client = await pool.connect();
  let importedRows = 0, skippedRows = 0, registerRows = 0, cardRows = 0, salaryRows = 0, vesselRows = 0, expenseRows = 0, wsRows = 0;
  try {
    await client.query('BEGIN');
    for (const row of rows) {
      const jobNumber = pick(row, FIELD_ALIASES.jobNumber); const invoiceNumber = pick(row, FIELD_ALIASES.invoiceNumber); const jobDate = toDate(pick(row, FIELD_ALIASES.jobDate)); const division = cleanDivision(pick(row, FIELD_ALIASES.division)); const opsManagerRaw = pick(row, FIELD_ALIASES.opsManager); const opsManager = KNOWN_OPS_MANAGERS.includes(String(opsManagerRaw || '').trim()) ? String(opsManagerRaw).trim() : null; const description = pick(row, FIELD_ALIASES.description); const clientName = pick(row, FIELD_ALIASES.client); const hours = toHours(pick(row, FIELD_ALIASES.hours)); const revenue = toNumber(pick(row, FIELD_ALIASES.revenue)); const labourCost = toNumber(pick(row, FIELD_ALIASES.labourCost)); const equipmentCost = toNumber(pick(row, FIELD_ALIASES.equipmentCost)); const workshopCost = toNumber(pick(row, FIELD_ALIASES.workshopCost)); const jobExpenses = toNumber(pick(row, FIELD_ALIASES.jobExpenses)); const officeExpenses = toNumber(pick(row, FIELD_ALIASES.officeExpenses)); const salaryCost = toNumber(pick(row, FIELD_ALIASES.salaryCost)); const totalCostFromSheet = toNumber(pick(row, FIELD_ALIASES.totalCost));
      const totalCost = totalCostFromSheet || labourCost + equipmentCost + workshopCost + jobExpenses + officeExpenses + salaryCost; const grossProfit = revenue - totalCost; const effectiveJobNumber = jobNumber || ((!invoiceNumber && !jobDate && (totalCost || officeExpenses || salaryCost)) ? importKey(['overhead', row.Month || row['Month'] || '', row.Year || row['Year'] || '', totalCost, officeExpenses, salaryCost]) : null); const hasFinancialData = revenue !== 0 || totalCost !== 0 || labourCost !== 0 || equipmentCost !== 0 || workshopCost !== 0; const hasRealIdentifier = Boolean(effectiveJobNumber || invoiceNumber || jobDate); if ((!hasRealIdentifier && !hasFinancialData) || (!jobDate && !hasFinancialData)) { skippedRows += 1; continue; }
      await client.query(`INSERT INTO jobs (job_number, invoice_number, job_date, division, ops_manager, hours, revenue, labour_cost, equipment_cost, workshop_cost, total_cost, gross_profit, description, client_name, quote_number, po_number, job_expenses, office_expenses, salary_cost) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19) ON CONFLICT (job_number) WHERE job_number IS NOT NULL DO UPDATE SET invoice_number=EXCLUDED.invoice_number, job_date=EXCLUDED.job_date, division=EXCLUDED.division, ops_manager=EXCLUDED.ops_manager, hours=EXCLUDED.hours, revenue=EXCLUDED.revenue, labour_cost=EXCLUDED.labour_cost, equipment_cost=EXCLUDED.equipment_cost, workshop_cost=EXCLUDED.workshop_cost, total_cost=EXCLUDED.total_cost, gross_profit=EXCLUDED.gross_profit, description=EXCLUDED.description, client_name=EXCLUDED.client_name, quote_number=EXCLUDED.quote_number, po_number=EXCLUDED.po_number, job_expenses=EXCLUDED.job_expenses, office_expenses=EXCLUDED.office_expenses, salary_cost=EXCLUDED.salary_cost`, [effectiveJobNumber ? String(effectiveJobNumber) : null, invoiceNumber ? String(invoiceNumber) : null, jobDate, division ? String(division) : null, opsManager ? String(opsManager) : null, hours, revenue, labourCost, equipmentCost, workshopCost, totalCost, grossProfit, description ? String(description) : null, clientName ? String(clientName) : null, pick(row, FIELD_ALIASES.quoteNumber) || null, pick(row, FIELD_ALIASES.poNumber) || null, jobExpenses, officeExpenses, salaryCost]);
      importedRows += 1;
    }

    
    const wsSheet = findSheet(workbook, ['WS Register', 'WSRegister', 'Workshop Register']);
    if (wsSheet) {
      const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[wsSheet], { header: 1, defval: '', blankrows: false });
      for (const row of matrix.slice(1)) {
        const jobNumber = coalesceText(row[3], row[16]);
        const jobDate = toDate(row[2] || row[15]);
        const description = coalesceText(row[4], row[19]);
        const supervisor = coalesceText(row[6], row[20]);
        const leftHours = toHours(row[8]);
        const rightHours = toHours(row[17]);
        const hours = leftHours || rightHours;
        const labourCost = toNumber(row[9] || row[18]);
        if (!jobNumber && !description) continue;
        if (!jobDate && !labourCost && !hours) continue;
        const wsJobNumber = String(jobNumber);
        const wsDivision = 'WS Maintenance';
        const wsKey = importKey(['ws-card', wsJobNumber, jobDate, description, labourCost]);
        await client.query(`INSERT INTO jobs (job_number, invoice_number, job_date, division, ops_manager, hours, revenue, labour_cost, equipment_cost, workshop_cost, total_cost, gross_profit, description, client_name, quote_number, po_number, job_expenses, office_expenses, salary_cost) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,$17,$18,$19) ON CONFLICT (job_number) WHERE job_number IS NOT NULL DO NOTHING`, [wsJobNumber, null, jobDate, wsDivision, null, hours, 0, labourCost, 0, 0, labourCost, -labourCost, description, 'Workshop', null, null, 0, 0, 0]);
        await client.query(`INSERT INTO job_card_entries (job_date, month, year, fy, job_number, ops_manager, description, start_time, end_time, hours, labour_cost, equipment_cost, workshop_cost, division, import_key) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15) ON CONFLICT (import_key) WHERE import_key IS NOT NULL DO UPDATE SET job_date=EXCLUDED.job_date, month=EXCLUDED.month, year=EXCLUDED.year, fy=EXCLUDED.fy, job_number=EXCLUDED.job_number, ops_manager=EXCLUDED.ops_manager, description=EXCLUDED.description, start_time=EXCLUDED.start_time, end_time=EXCLUDED.end_time, hours=EXCLUDED.hours, labour_cost=EXCLUDED.labour_cost, equipment_cost=EXCLUDED.equipment_cost, workshop_cost=EXCLUDED.workshop_cost, division=EXCLUDED.division`, [jobDate, row[0] || row[13] || null, toNumber(row[1] || row[14]) || null, null, wsJobNumber, supervisor, description, null, null, hours, labourCost, 0, 0, wsDivision, wsKey]);
        wsRows += 1;
        cardRows += 1;
        importedRows += 1;
      }
    }

    const jrSheet = workbook.SheetNames.find((name) => name === 'Job Register ') || findSheet(workbook, ['Job Register', 'JobRegister', 'Register']);
    if (jrSheet) { const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[jrSheet], { header: 1, defval: '', blankrows: false }); const headerRow = matrix[1] || []; const normalizedHeader = new Map(headerRow.map((header, index) => [normalizeHeader(header), index])); const paymentHeaderIndex = FIELD_ALIASES.paymentsMade.map((alias) => normalizedHeader.get(normalizeHeader(alias))).find((index) => Number.isInteger(index)); const paymentsColumnIndex = Number.isInteger(paymentHeaderIndex) ? paymentHeaderIndex : 14; for (const row of matrix.slice(2)) { const jobNumber = row[0]; const jobDate = toDate(row[1]); const opsManager = row[2]; const division = cleanDivision(row[3]); const clientName = row[4]; const description = row[5]; const completionDate = toDate(row[6]); const quoteNumber = row[7]; const poNumber = row[8]; const issued = row[9]; const invoiceNumber = row[10]; const form = row[11]; const valueInclVat = toNumber(row[12]); const valueExclVat = toNumber(row[13]); const paymentsMade = toNumber(row[paymentsColumnIndex]); if (!jobNumber && !description && !invoiceNumber) continue; await client.query(`INSERT INTO job_register_entries (job_number, job_date, ops_manager, division, client_name, description, completion_date, quote_number, po_number, report_reference, invoice_number, client_feedback, value_incl_vat, value_excl_vat, payments_made) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15) ON CONFLICT (job_number) WHERE job_number IS NOT NULL DO UPDATE SET job_date=EXCLUDED.job_date, ops_manager=EXCLUDED.ops_manager, division=EXCLUDED.division, client_name=EXCLUDED.client_name, description=EXCLUDED.description, completion_date=EXCLUDED.completion_date, quote_number=EXCLUDED.quote_number, po_number=EXCLUDED.po_number, report_reference=EXCLUDED.report_reference, invoice_number=EXCLUDED.invoice_number, client_feedback=EXCLUDED.client_feedback, value_incl_vat=EXCLUDED.value_incl_vat, value_excl_vat=EXCLUDED.value_excl_vat, payments_made=EXCLUDED.payments_made`, [jobNumber ? String(jobNumber) : null, jobDate, opsManager ? String(opsManager) : null, division ? String(division) : null, clientName ? String(clientName) : null, description ? String(description) : null, completionDate, quoteNumber ? String(quoteNumber) : null, poNumber ? String(poNumber) : null, issued ? String(issued) : null, invoiceNumber ? String(invoiceNumber) : null, form ? String(form) : null, valueInclVat, valueExclVat, paymentsMade]); registerRows += 1; } }

    const jcSheet = findSheet(workbook, ['Job Card Conversion', 'JobCardConversion']);
    if (jcSheet) { for (const row of rowsFromSheet(workbook, jcSheet)) { const jobNumber = pick(row, FIELD_ALIASES.jobNumber); if (!jobNumber) continue; const cardDate = toDate(pick(row, FIELD_ALIASES.jobDate)); const cardStart = toTime(row['Start Time '] ?? row['Start Time']); const cardEnd = toTime(row['End Time']); const cardDescription = pick(row, FIELD_ALIASES.description) || null; const labourCost = toNumber(pick(row, FIELD_ALIASES.labourCost)); const equipmentCost = toNumber(pick(row, FIELD_ALIASES.equipmentCost)); const workshopCost = toNumber(pick(row, FIELD_ALIASES.workshopCost)); const cardTotalCost = labourCost + equipmentCost + workshopCost; const cardKey = importKey(['job-card', jobNumber, cardDate, cardStart, cardEnd, cardDescription]); await client.query(`INSERT INTO job_card_entries (job_date, month, year, fy, job_number, ops_manager, description, start_time, end_time, hours, labour_cost, equipment_cost, workshop_cost, total_cost, division, import_key, source) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15,$16,'import') ON CONFLICT (import_key) WHERE import_key IS NOT NULL DO UPDATE SET job_date=EXCLUDED.job_date, month=EXCLUDED.month, year=EXCLUDED.year, fy=EXCLUDED.fy, job_number=EXCLUDED.job_number, ops_manager=EXCLUDED.ops_manager, description=EXCLUDED.description, start_time=EXCLUDED.start_time, end_time=EXCLUDED.end_time, hours=EXCLUDED.hours, labour_cost=EXCLUDED.labour_cost, equipment_cost=EXCLUDED.equipment_cost, workshop_cost=EXCLUDED.workshop_cost, total_cost=EXCLUDED.total_cost, division=EXCLUDED.division, source=EXCLUDED.source`, [cardDate, row.Month || row['Month'] || null, toNumber(row.Year || row['Year']) || null, row.FY ? String(row.FY) : null, String(jobNumber), pick(row, FIELD_ALIASES.opsManager) || null, cardDescription, cardStart, cardEnd, toHours(pick(row, FIELD_ALIASES.hours)), labourCost, equipmentCost, workshopCost, cardTotalCost, cleanDivision(pick(row, FIELD_ALIASES.division)), cardKey]); cardRows += 1; } }

    const vesselSheet = findSheet(workbook, ['Vessel Info', 'VesselInfo', 'Vessels']);
    if (vesselSheet) { const vesselData = rowsFromSheet(workbook, vesselSheet); for (const row of vesselData) { const jobNumber = pick(row, FIELD_ALIASES.jobNumber); const usedVessels = VESSEL_COLUMNS.filter((vessel) => isUsed(row[vessel])); if (!jobNumber || usedVessels.length === 0) continue; const baseHours = toHours(pick(row, FIELD_ALIASES.hours)); const perVesselHours = usedVessels.length > 1 ? baseHours / usedVessels.length : baseHours; const labourCost = toNumber(pick(row, FIELD_ALIASES.labourCost)); const equipmentCost = toNumber(pick(row, FIELD_ALIASES.equipmentCost)); const workshopCost = toNumber(pick(row, FIELD_ALIASES.workshopCost)); for (const vessel of usedVessels) { const vesselDate = toDate(pick(row, FIELD_ALIASES.jobDate)); const vesselStart = toTime(row['Start Time '] ?? row['Start Time']); const vesselEnd = toTime(row['End Time']); const vesselDescription = pick(row, FIELD_ALIASES.description) || null; const vesselName = normalizeVesselName(vessel); const vesselKey = importKey(['vessel', jobNumber, vesselDate, vesselStart, vesselEnd, vesselName, vesselDescription]); await client.query(`INSERT INTO vessel_entries (job_date, month, client_name, job_number, ops_manager, description, start_time, end_time, hours, vessel_name, labour_cost, equipment_cost, workshop_cost, division, import_key) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,$14,$15) ON CONFLICT (import_key) WHERE import_key IS NOT NULL DO UPDATE SET job_date=EXCLUDED.job_date, month=EXCLUDED.month, client_name=EXCLUDED.client_name, job_number=EXCLUDED.job_number, ops_manager=EXCLUDED.ops_manager, description=EXCLUDED.description, start_time=EXCLUDED.start_time, end_time=EXCLUDED.end_time, hours=EXCLUDED.hours, vessel_name=EXCLUDED.vessel_name, labour_cost=EXCLUDED.labour_cost, equipment_cost=EXCLUDED.equipment_cost, workshop_cost=EXCLUDED.workshop_cost, division=EXCLUDED.division`, [vesselDate, row.Month || row['Month'] || null, pick(row, FIELD_ALIASES.client) ? String(pick(row, FIELD_ALIASES.client)) : null, String(jobNumber), pick(row, FIELD_ALIASES.opsManager) || null, vesselDescription, vesselStart, vesselEnd, perVesselHours, vesselName, usedVessels.length > 1 ? labourCost / usedVessels.length : labourCost, usedVessels.length > 1 ? equipmentCost / usedVessels.length : equipmentCost, usedVessels.length > 1 ? workshopCost / usedVessels.length : workshopCost, cleanDivision(pick(row, FIELD_ALIASES.division)), vesselKey]); vesselRows += 1; } } }

    const expSheet = findSheet(workbook, ['Expenses', 'Expense']);
    if (expSheet) {
      for (const row of rowsFromSheet(workbook, expSheet)) {
        // The bookkeeper workbook has both job-linked expenses and overhead expenses.
        // Executive Summary must use only the bookkeeper's "Expenses NO JOB #" amount.
        // Do not filter on Job Number here: the workbook itself decides which rows belong
        // in the non-job expense bucket by placing a value in this dedicated column.
        const noJobAmount = toNumber(pick(row, FIELD_ALIASES.expenseNoJob));
        if (noJobAmount <= 0) continue;

        const amountExclVat = noJobAmount;
        const amountInclVat = toNumber(row['Total (with Tax)'] ?? row['Total With Tax'] ?? row['Incl VAT']);
        const amount = amountExclVat;
        const supplier = coalesceText(pick(row, FIELD_ALIASES.expenseSupplier), row.Company);
        const category = coalesceText(pick(row, FIELD_ALIASES.expenseCategory), row.Category, 'Unassigned');
        const placement = coalesceText(row.Placement, cleanDivision(pick(row, FIELD_ALIASES.division)), 'Expenses');
        const description = coalesceText(pick(row, FIELD_ALIASES.expenseDescription), row['Invoice Number'], row['Cash Source'], placement, category);
        const expenseDate = toDate(pick(row, FIELD_ALIASES.expenseDate));
        const vesselName = detectExpenseVessel(row, description);
        const expenseKey = importKey(['expense-no-job', expenseDate, supplier, category, description, amount, row['Cash Source'] || '']);
        await client.query(
          `INSERT INTO expense_entries (expense_date, month, supplier, category, description, division, vessel_name, amount, amount_excl_vat, amount_incl_vat, import_key)
           VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11)
           ON CONFLICT (import_key) WHERE import_key IS NOT NULL DO UPDATE SET
             expense_date=EXCLUDED.expense_date,
             month=EXCLUDED.month,
             supplier=EXCLUDED.supplier,
             category=EXCLUDED.category,
             description=EXCLUDED.description,
             division=EXCLUDED.division,
             vessel_name=EXCLUDED.vessel_name,
             amount=EXCLUDED.amount,
             amount_excl_vat=EXCLUDED.amount_excl_vat,
             amount_incl_vat=EXCLUDED.amount_incl_vat`,
          [expenseDate, row.Month || row['Month'] || null, supplier, category || 'Unassigned', description, cleanDivision(placement), vesselName, amount, amountExclVat, amountInclVat, expenseKey]
        );
        expenseRows += 1;
      }
    }

    const salSheet = findSheet(workbook, ['Salaries', 'Salary']);
    if (salSheet) {
      const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[salSheet], { header: 1, defval: '' });

      // Salaries sheet is a monthly summary table. Only import the top block headed "Total Salaries".
      for (let r = 0; r < matrix.length; r += 1) {
        const row = matrix[r] || [];
        for (let c = 0; c < row.length - 1; c += 1) {
          const isMonthHeader = normalizeHeader(row[c]) === 'month';
          const isTotalSalariesHeader = normalizeHeader(row[c + 1]) === 'totalsalaries';
          if (!isMonthHeader || !isTotalSalariesHeader) continue;

          const year = toNumber(matrix[r - 1]?.[c]);
          if (!year) continue;

          for (let rr = r + 1; rr < matrix.length; rr += 1) {
            const month = matrix[rr]?.[c];
            const value = toNumber(matrix[rr]?.[c + 1]);
            if (!month || normalizeHeader(month) === 'month') break;
            if (typeof month === 'string' && month.length > 2 && value > 0) {
              await client.query(
                `INSERT INTO salaries (salary_year, salary_month, total_salaries)
                 VALUES ($1,$2,$3)
                 ON CONFLICT (salary_year, salary_month)
                 DO UPDATE SET total_salaries=EXCLUDED.total_salaries`,
                [year, month, value]
              );
              salaryRows += 1;
            }
          }
        }
      }
    }

    await client.query(`INSERT INTO imports (filename, imported_rows, status, register_rows, card_rows, vessel_rows, expense_rows, salary_rows, skipped_rows, ws_rows) VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10)`, [filename || 'uploaded workbook', importedRows, 'completed', registerRows, cardRows, vesselRows, expenseRows, salaryRows, skippedRows, wsRows]);
    await client.query('COMMIT');
    return { status: 'imported', filename, sheet: best.sheetName, headerRow: best.header.rowIndex + 1, headerScore: best.header.score, rowsFound: rows.length, importedRows, skippedRows, registerRows, cardRows, salaryRows, vesselRows, expenseRows, wsRows, sheets: workbook.SheetNames };
  } catch (error) { await client.query('ROLLBACK'); throw error; } finally { client.release(); }
}

const EXPORT_DATE_FORMAT = 'yyyy-mm-dd'; const EXPORT_CURRENCY_FORMAT = 'N$ #,##0.00';
function cleanExportRows(rows) { return rows.map((row) => Object.fromEntries(Object.entries(row).map(([key, value]) => [key, value === null || value === undefined ? '' : value]))); }
function applySheetFormats(worksheet, currencyColumns = [], dateColumns = []) { const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:A1'); for (let row = range.s.r + 1; row <= range.e.r; row += 1) { for (const col of currencyColumns) { const cell = worksheet[XLSX.utils.encode_cell({ r: row, c: col })]; if (cell && typeof cell.v === 'number') cell.z = EXPORT_CURRENCY_FORMAT; } for (const col of dateColumns) { const cell = worksheet[XLSX.utils.encode_cell({ r: row, c: col })]; if (cell && cell.v) cell.z = EXPORT_DATE_FORMAT; } } worksheet['!cols'] = Array.from({ length: Math.max(1, range.e.c + 1) }, () => ({ wch: 18 })); }
function appendJsonSheet(workbook, sheetName, rows, currencyColumns = [], dateColumns = []) { const worksheet = XLSX.utils.json_to_sheet(cleanExportRows(rows)); applySheetFormats(worksheet, currencyColumns, dateColumns); XLSX.utils.book_append_sheet(workbook, worksheet, sheetName); }
function sendWorkbook(reply, workbook, filename) { const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' }); reply.header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet').header('Content-Disposition', `attachment; filename="${filename}"`).send(buffer); }

async function getJobRegisterExportRows() { const result = await pool.query(`SELECT job_number AS "Job Nr", job_date AS "Date", ops_manager AS "Division Head", division AS "Division", client_name AS "Client", description AS "Description", completion_date AS "Completion Date", quote_number AS "Quote Nr", po_number AS "PO No", report_reference AS "Report Reference", invoice_number AS "Invoice No", client_feedback AS "Client Feedback", value_incl_vat AS "Value Incl VAT", value_excl_vat AS "Value Excl VAT", payments_made AS "Payments Made" FROM job_register_entries ORDER BY job_date DESC NULLS LAST, id DESC`); return result.rows; }
async function getJobCardsExportRows() { const result = await pool.query(`SELECT job_date AS "Date", TRIM(month) AS "Month", year AS "Year", fy AS "FY", job_number AS "Job Nr", client_name AS "Client", contact_person AS "Contact Person / Rep", location AS "Location", vessel_name AS "Vessel", ops_manager AS "OPS Manager", division AS "Vertical / Division", activity_type AS "Activity / Work Type", description AS "Work Done / Description", start_time AS "Start Time", end_time AS "End Time", hours AS "Total Hours", staff_names AS "Staff / Crew", equipment_used AS "Equipment Used", labour_cost AS "Cost Labour", equipment_cost AS "Cost Equipment", workshop_cost AS "Workshop Cost", material_cost AS "Material Cost", other_cost AS "Other Cost", total_cost AS "Total Card Cost", completed_by_name AS "Completed By", source AS "Source", job_link_status AS "Job Link Status" FROM job_card_entries ORDER BY job_date DESC NULLS LAST, id DESC`); return result.rows; }
async function getJobCardConversionExportRows() {
  const rows = await getJobCardsExportRows();
  return rows.map(row => ({
    'Date': row['Date'],
    'Month': row['Month'],
    'Year': row['Year'],
    'FY': row['FY'],
    'Job Nr': row['Job Nr'],
    'Client': row['Client'],
    'OPS Manager': row['OPS Manager'],
    'Vertical / Division': row['Vertical / Division'],
    'Vessel': row['Vessel'],
    'Location': row['Location'],
    'Activity / Work Type': row['Activity / Work Type'],
    'Work Done / Description': row['Work Done / Description'],
    'Start Time': row['Start Time'],
    'End Time': row['End Time'],
    'Total Hours': row['Total Hours'],
    'Staff / Crew': row['Staff / Crew'],
    'Equipment Used': row['Equipment Used'],
    'Cost Labour': row['Cost Labour'],
    'Cost Equipment': row['Cost Equipment'],
    'Workshop Cost': row['Workshop Cost'],
    'Material Cost': row['Material Cost'],
    'Other Cost': row['Other Cost'],
    'Total Card Cost': row['Total Card Cost'],
    'Completed By': row['Completed By'],
    'Source': row['Source']
  }));
}
async function getSalariesExportRows() { const result = await pool.query(`SELECT salary_year AS "Year", salary_month AS "Month", total_salaries AS "Total Salaries" FROM salaries ORDER BY salary_year DESC, imported_at DESC, id DESC`); return result.rows; }
async function getJobsExportRows() { const result = await pool.query(`SELECT job_number AS "Job Nr", invoice_number AS "Invoice No", job_date AS "Date", division AS "Division", ops_manager AS "OPS Manager", client_name AS "Client", description AS "Description", quote_number AS "Quote Nr", po_number AS "PO No", hours AS "Total Hours", revenue AS "Revenue Excl VAT", labour_cost AS "Labour Costs", equipment_cost AS "Equipment Costs", workshop_cost AS "Workshop Cost", job_expenses AS "Job Expenses", office_expenses AS "Office Expenses", salary_cost AS "Salaries", total_cost AS "Total Costs", gross_profit AS "Gross Profit" FROM jobs ORDER BY job_date DESC NULLS LAST, id DESC`); return result.rows; }
async function getVesselExportRows() { const result = await pool.query(`SELECT job_date AS "Date", month AS "Month", client_name AS "Client", job_number AS "Job Nr", ops_manager AS "OPS Manager", description AS "Description", start_time AS "Start Time", end_time AS "End Time", hours AS "Hours", vessel_name AS "Vessel", labour_cost AS "Cost Labour", equipment_cost AS "Cost Equipment", workshop_cost AS "Workshop Cost", division AS "Division" FROM vessel_entries ORDER BY job_date DESC NULLS LAST, id DESC`); return result.rows; }
async function getExpensesExportRows() { const result = await pool.query(`SELECT expense_date AS "Date", month AS "Month", supplier AS "Supplier", category AS "Category", description AS "Description", vessel_name AS "Vessel", division AS "Division", amount_excl_vat AS "Amount Excl VAT", amount_incl_vat AS "Amount Incl VAT", amount AS "Amount" FROM expense_entries ORDER BY expense_date DESC NULLS LAST, id DESC`); return result.rows; }
function buildAnalysisRows(jobs) { const totals = jobs.reduce((acc, job) => { acc.jobs += 1; acc.hours += Number(job['Total Hours'] || 0); acc.revenue += Number(job['Revenue Excl VAT'] || 0); acc.labour += Number(job['Labour Costs'] || 0); acc.equipment += Number(job['Equipment Costs'] || 0); acc.workshop += Number(job['Workshop Cost'] || 0); acc.cost += Number(job['Total Costs'] || 0); acc.profit += Number(job['Gross Profit'] || 0); return acc; }, { jobs: 0, hours: 0, revenue: 0, labour: 0, equipment: 0, workshop: 0, cost: 0, profit: 0 }); return [{ Metric: 'Jobs', Value: totals.jobs }, { Metric: 'Total Hours', Value: totals.hours }, { Metric: 'Revenue Excl VAT', Value: totals.revenue }, { Metric: 'Labour Costs', Value: totals.labour }, { Metric: 'Equipment Costs', Value: totals.equipment }, { Metric: 'Workshop Cost', Value: totals.workshop }, { Metric: 'Total Costs', Value: totals.cost }, { Metric: 'Gross Profit', Value: totals.profit }, { Metric: 'Gross Profit %', Value: totals.revenue ? totals.profit / totals.revenue : 0 }]; }
function buildProfitPerJobRows(jobs) { return jobs.map((job) => { const revenue = Number(job['Revenue Excl VAT'] || 0); const profit = Number(job['Gross Profit'] || 0); return { 'Job Nr': job['Job Nr'], Date: job.Date, Client: job.Client, Division: job.Division, 'OPS Manager': job['OPS Manager'], 'Revenue Excl VAT': revenue, 'Total Costs': Number(job['Total Costs'] || 0), 'Gross Profit': profit, 'GP %': revenue ? profit / revenue : 0 }; }); }
function buildSettingsRows() { return [{ Setting: 'Company', Value: 'B4 Engineering & Diving / Nautilus Operations' }, { Setting: 'Currency', Value: 'N$' }, { Setting: 'Workbook Export Version', Value: 'Incremental Import + Expenses VAT' }, { Setting: 'Debtors Logic', Value: 'Monitoring only - current invoices treated as paid by default unless payment data exists' }]; }

app.get('/api/export/job-register.xlsx', async (request, reply) => { try { const workbook = XLSX.utils.book_new(); appendJsonSheet(workbook, 'Job Register', await getJobRegisterExportRows(), [12, 13, 14], [1, 6]); return sendWorkbook(reply, workbook, 'job-register.xlsx'); } catch (error) { app.log.error(error); return reply.code(500).send({ error: error.message }); } });
app.get('/api/export/job-cards.xlsx', async (request, reply) => { try { const workbook = XLSX.utils.book_new(); appendJsonSheet(workbook, 'Job Card Conversion', await getJobCardConversionExportRows(), [17, 18, 19, 20, 21, 22], [0]); return sendWorkbook(reply, workbook, 'job-card-conversion.xlsx'); } catch (error) { app.log.error(error); return reply.code(500).send({ error: error.message }); } });
app.get('/api/export/job-card-conversion.xlsx', async (request, reply) => { try { const workbook = XLSX.utils.book_new(); appendJsonSheet(workbook, 'Job Card Conversion', await getJobCardConversionExportRows(), [17, 18, 19, 20, 21, 22], [0]); return sendWorkbook(reply, workbook, 'job-card-conversion.xlsx'); } catch (error) { app.log.error(error); return reply.code(500).send({ error: error.message }); } });
app.get('/api/export/salaries.xlsx', async (request, reply) => { try { const workbook = XLSX.utils.book_new(); appendJsonSheet(workbook, 'Salaries', await getSalariesExportRows(), [2], []); return sendWorkbook(reply, workbook, 'salaries.xlsx'); } catch (error) { app.log.error(error); return reply.code(500).send({ error: error.message }); } });
app.get('/api/export/vessels.xlsx', async (request, reply) => { try { const workbook = XLSX.utils.book_new(); appendJsonSheet(workbook, 'Vessel Info', await getVesselExportRows(), [10, 11, 12], [0]); return sendWorkbook(reply, workbook, 'vessel-info.xlsx'); } catch (error) { app.log.error(error); return reply.code(500).send({ error: error.message }); } });
app.get('/api/export/expenses.xlsx', async (request, reply) => { try { const workbook = XLSX.utils.book_new(); appendJsonSheet(workbook, 'Expenses', await getExpensesExportRows(), [7,8,9], [0]); return sendWorkbook(reply, workbook, 'expenses.xlsx'); } catch (error) { app.log.error(error); return reply.code(500).send({ error: error.message }); } });
app.get('/api/export/full-workbook.xlsx', async (request, reply) => { try { const workbook = XLSX.utils.book_new(); const jobRegisterRows = await getJobRegisterExportRows(); const jobRows = await getJobsExportRows(); const jobCardRows = await getJobCardsExportRows(); const vesselRows = await getVesselExportRows(); const expenseRows = await getExpensesExportRows(); const salaryRows = await getSalariesExportRows(); appendJsonSheet(workbook, 'Job Register', jobRegisterRows, [12, 13, 14], [1, 6]); appendJsonSheet(workbook, 'Analysis', buildAnalysisRows(jobRows), [1], []); appendJsonSheet(workbook, 'Profit per Job', buildProfitPerJobRows(jobRows), [5, 6, 7], [1]); appendJsonSheet(workbook, 'Data', jobRows, [10, 11, 12, 13, 14, 15], [2]); appendJsonSheet(workbook, 'Jobs', jobRows, [10, 11, 12, 13, 14, 15], [2]); appendJsonSheet(workbook, 'Job Card Conversion', jobCardRows, [17, 18, 19, 20, 21, 22], [0]); appendJsonSheet(workbook, 'Vessel Info', vesselRows, [10, 11, 12], [0]); appendJsonSheet(workbook, 'Expenses', expenseRows, [7,8,9], [0]); appendJsonSheet(workbook, 'Settings', buildSettingsRows(), [], []); appendJsonSheet(workbook, 'Salaries', salaryRows, [2], []); return sendWorkbook(reply, workbook, 'b4-nautilus-full-workbook.xlsx'); } catch (error) { app.log.error(error); return reply.code(500).send({ error: error.message }); } });

app.post('/api/imports/excel', async (request, reply) => { try { const data = await request.file(); if (!data) return reply.code(400).send({ error: 'No file uploaded.' }); const chunks = []; for await (const chunk of data.file) chunks.push(chunk); const buffer = Buffer.concat(chunks); const result = await importWorkbook(buffer, data.filename); return result; } catch (error) { app.log.error(error); return reply.code(500).send({ error: error.message }); } });

app.setNotFoundHandler((request, reply) => { if (request.raw.url?.startsWith('/api/')) return reply.code(404).send({ error: 'Not found' }); return reply.sendFile('index.html'); });

try { await initDatabase(); await app.listen({ port, host }); } catch (error) { app.log.error(error); process.exit(1); }