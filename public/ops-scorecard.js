let jobs = [];
let registers = [];
let scoreRows = [];

const money = v => 'N$ ' + Math.round(Number(v || 0)).toLocaleString();
const num = v => Number(v || 0).toLocaleString();
const pct = v => (Number(v || 0) * 100).toFixed(1) + '%';
const safe = v => String(v ?? '').replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
const key = v => String(v ?? '').trim();

async function api(path) {
  const res = await fetch(path);
  if (!res.ok) throw new Error(path + ' failed: ' + res.status);
  return res.json();
}

async function checkStatus() {
  const el = document.getElementById('apiStatus');
  try {
    const status = await api('/api/status');
    el.textContent = status.database === 'connected' ? 'Database connected' : 'Database not connected';
    el.className = status.database === 'connected' ? 'status ok' : 'status error';
  } catch (err) {
    el.textContent = 'API offline';
    el.className = 'status error';
  }
}

function buildScorecard() {
  const regByJob = {};
  registers.forEach(r => { regByJob[key(r.job_number)] = r; });
  const byOps = {};

  jobs.forEach(j => {
    const r = regByJob[key(j.job_number)] || {};
    const ops = r.ops_manager || j.ops_manager || 'Unassigned';
    const revenue = Number(j.revenue || 0);
    const cost = Number(j.total_cost || 0);
    const profit = Number(j.gross_profit || 0);
    const hours = Number(j.hours || 0);
    const invoice = key(r.invoice_number || j.invoice_number || '');

    if (!byOps[ops]) {
      byOps[ops] = { ops, jobs: 0, revenue: 0, cost: 0, profit: 0, hours: 0, missingInvoices: 0, missingValue: 0, lossJobs: 0 };
    }

    byOps[ops].jobs += 1;
    byOps[ops].revenue += revenue;
    byOps[ops].cost += cost;
    byOps[ops].profit += profit;
    byOps[ops].hours += hours;
    if (profit < 0) byOps[ops].lossJobs += 1;
    if (!invoice && revenue > 0) {
      byOps[ops].missingInvoices += 1;
      byOps[ops].missingValue += revenue;
    }
  });

  scoreRows = Object.values(byOps).map(r => ({
    ...r,
    gp: r.revenue > 0 ? r.profit / r.revenue : 0,
    revenuePerHour: r.hours > 0 ? r.revenue / r.hours : 0
  }));
}

function filteredRows() {
  const search = (document.getElementById('search')?.value || '').toLowerCase();
  const sort = document.getElementById('sort')?.value || 'profit';
  return scoreRows.filter(r => r.ops.toLowerCase().includes(search)).sort((a, b) => Number(b[sort] || 0) - Number(a[sort] || 0));
}

function bars(id, rows, field) {
  const max = Math.max(...rows.map(r => Math.abs(Number(r[field] || 0))), 1);
  document.getElementById(id).innerHTML = rows.slice(0, 10).map(r => {
    const val = Number(r[field] || 0);
    const width = Math.max(3, Math.abs(val) / max * 100);
    return '<div class="bar"><span>' + safe(r.ops) + '</span><div class="track"><div class="fill ' + (val < 0 ? 'loss' : '') + '" style="width:' + width + '%"></div></div><strong class="' + (val < 0 ? 'danger' : 'good') + '">' + money(val) + '</strong></div>';
  }).join('') || '<div class="muted">No data.</div>';
}

function render() {
  const rows = filteredRows();
  const t = rows.reduce((a, r) => {
    a.revenue += r.revenue; a.cost += r.cost; a.profit += r.profit; a.hours += r.hours; a.jobs += r.jobs; a.missingValue += r.missingValue; a.lossJobs += r.lossJobs; return a;
  }, { revenue: 0, cost: 0, profit: 0, hours: 0, jobs: 0, missingValue: 0, lossJobs: 0 });

  document.getElementById('kpis').innerHTML = [
    ['OPS Managers', num(rows.length)], ['Revenue', money(t.revenue)], ['Profit', money(t.profit), t.profit >= 0 ? 'good' : 'danger'], ['GP %', pct(t.revenue > 0 ? t.profit / t.revenue : 0)],
    ['Jobs', num(t.jobs)], ['Hours', Number(t.hours).toFixed(1)], ['Missing Invoice Value', money(t.missingValue), 'danger'], ['Loss Jobs', num(t.lossJobs), 'danger']
  ].map(c => '<div class="card"><div class="label">' + c[0] + '</div><div class="value ' + (c[2] || '') + '">' + c[1] + '</div></div>').join('');

  bars('profitBars', rows, 'profit');
  bars('invoiceBars', rows, 'missingValue');

  document.getElementById('rows').innerHTML = rows.map(r => '<tr><td>' + safe(r.ops) + '</td><td>' + num(r.jobs) + '</td><td>' + money(r.revenue) + '</td><td>' + money(r.cost) + '</td><td class="' + (r.profit < 0 ? 'danger' : 'good') + '">' + money(r.profit) + '</td><td>' + pct(r.gp) + '</td><td>' + Number(r.hours).toFixed(1) + '</td><td>' + money(r.revenuePerHour) + '</td><td class="danger">' + num(r.missingInvoices) + '</td><td class="danger">' + money(r.missingValue) + '</td><td class="danger">' + num(r.lossJobs) + '</td></tr>').join('') || '<tr><td colspan="11" class="muted">No OPS managers found.</td></tr>';
}

function exportOpsCsv() {
  const rows = [['OPS Manager','Jobs','Revenue','Cost','Profit','GP %','Hours','Revenue per Hour','Missing Invoices','Missing Invoice Value','Loss Jobs']];
  filteredRows().forEach(r => rows.push([r.ops,r.jobs,r.revenue,r.cost,r.profit,(r.gp*100).toFixed(1)+'%',r.hours,r.revenuePerHour,r.missingInvoices,r.missingValue,r.lossJobs]));
  const csv = rows.map(r => r.map(v => '"' + String(v ?? '').replaceAll('"','""') + '"').join(',')).join('\n');
  const blob = new Blob([csv], { type: 'text/csv' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'b4-ops-scorecard.csv';
  a.click();
  URL.revokeObjectURL(url);
}

async function load() {
  await checkStatus();
  jobs = await api('/api/jobs?limit=5000');
  registers = await api('/api/job-register?limit=5000');
  buildScorecard();
  render();
}

load().catch(err => {
  const el = document.getElementById('apiStatus');
  el.textContent = err.message;
  el.className = 'status error';
});
