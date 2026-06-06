import Fastify from 'fastify';
import fastifyStatic from '@fastify/static';
import fastifyMultipart from '@fastify/multipart';
import pg from 'pg';
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

await app.register(fastifyMultipart, {
  limits: {
    fileSize: 25 * 1024 * 1024
  }
});

await app.register(fastifyStatic, {
  root: path.join(__dirname, 'public'),
  prefix: '/',
});

app.get('/health', async () => ({
  status: 'ok',
  app: 'B4 Nautilus Operations'
}));

app.get('/api/status', async () => {
  try {
    const db = await pool.query('SELECT NOW() AS now');
    return {
      app: 'B4 Nautilus Operations',
      status: 'running',
      database: 'connected',
      time: db.rows[0].now
    };
  } catch (error) {
    return {
      app: 'B4 Nautilus Operations',
      status: 'running',
      database: 'not connected',
      error: error.message
    };
  }
});

app.get('/api/jobs', async () => {
  const result = await pool.query(`
    SELECT *
    FROM jobs
    ORDER BY job_date DESC NULLS LAST, id DESC
    LIMIT 100
  `);

  return result.rows;
});

app.get('/api/dashboard', async () => {
  const totals = await pool.query(`
    SELECT
      COALESCE(SUM(revenue), 0) AS revenue,
      COALESCE(SUM(labour_cost), 0) AS labour_cost,
      COALESCE(SUM(equipment_cost), 0) AS equipment_cost,
      COALESCE(SUM(workshop_cost), 0) AS workshop_cost,
      COALESCE(SUM(total_cost), 0) AS total_cost,
      COALESCE(SUM(gross_profit), 0) AS gross_profit,
      COUNT(*) AS jobs
    FROM jobs
  `);

  const divisions = await pool.query(`
    SELECT
      COALESCE(division, 'Unassigned') AS division,
      COALESCE(SUM(gross_profit), 0) AS gross_profit
    FROM jobs
    GROUP BY division
    ORDER BY gross_profit DESC
  `);

  const opsManagers = await pool.query(`
    SELECT
      COALESCE(ops_manager, 'Unassigned') AS ops_manager,
      COALESCE(SUM(gross_profit), 0) AS gross_profit
    FROM jobs
    GROUP BY ops_manager
    ORDER BY gross_profit DESC
  `);

  return {
    totals: totals.rows[0],
    divisions: divisions.rows,
    opsManagers: opsManagers.rows
  };
});

app.get('/api/imports', async () => {
  const result = await pool.query(`
    SELECT *
    FROM imports
    ORDER BY imported_at DESC
    LIMIT 50
  `);

  return result.rows;
});

app.post('/api/imports/test-seed', async () => {
  await pool.query('DELETE FROM jobs');

  const sampleJobs = [
    ['#1', 'IN001', '2024-01-01', 'Pollution', 'Jolanda', 15, 170500, 6110, 13731, 2763],
    ['#2', 'IN002', '2024-01-01', 'Pollution', 'Jolanda', 1.9, 17500, 1200, 1900, 339],
    ['#3', 'IN003', '2024-01-02', 'Pollution', 'Jolanda', 4, 17500, 2000, 3800, 1154],
    ['#4', 'IN004', '2024-01-03', 'Diving', 'Eugen', 5.4, 38280, 4500, 9000, 6255],
    ['#5', 'IN005', '2024-01-03', 'Civil', 'Carmon', 21, 178200, 6200, 9700, 9958]
  ];

  for (const job of sampleJobs) {
    const totalCost = Number(job[7]) + Number(job[8]) + Number(job[9]);
    const grossProfit = Number(job[6]) - totalCost;

    await pool.query(
      `
      INSERT INTO jobs (
        job_number,
        invoice_number,
        job_date,
        division,
        ops_manager,
        hours,
        revenue,
        labour_cost,
        equipment_cost,
        workshop_cost,
        total_cost,
        gross_profit
      )
      VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12)
      `,
      [...job, totalCost, grossProfit]
    );
  }

  await pool.query(
    `INSERT INTO imports (filename, imported_rows, status) VALUES ($1, $2, $3)`,
    ['test-seed', sampleJobs.length, 'completed']
  );

  return {
    status: 'seeded',
    rows: sampleJobs.length
  };
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
