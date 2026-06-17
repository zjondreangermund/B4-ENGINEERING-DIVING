const CACHE_NAME='b4-ops-cache-v18-live-filter-hardfix';

self.addEventListener('install', event => {
  self.skipWaiting();
});

self.addEventListener('activate', event => {
  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(keys.map(k => caches.delete(k))))
      .then(() => self.clients.claim())
  );
});

const HOTFIX_SCRIPT = `
<script>
(function(){
  if(window.__B4_FILTER_HARDFIX_V18__) return;
  window.__B4_FILTER_HARDFIX_V18__ = true;
  const VALID_OPS = ['Jolanda','Eugen','Carmon','Bresler','Elliotte'];
  function nf(v){
    return String(v ?? '')
      .normalize('NFD').replace(/[\\u0300-\\u036f]/g,'')
      .trim().toLowerCase().replace(/\\s+/g,' ');
  }
  window.normFilter = nf;
  function safeEsc(v){
    if(typeof window.esc === 'function') return window.esc(v);
    return String(v ?? '').replace(/[&<>]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
  }
  function safeMoney(v){return typeof window.money==='function'?window.money(v):('N$ '+Number(v||0).toFixed(2));}
  function safeDate(v){return typeof window.date==='function'?window.date(v):(v?String(v).slice(0,10):'');}
  function safeNum(v){return typeof window.num==='function'?window.num(v):Number(v||0).toLocaleString();}
  function strictSetOptions(id, first, values){
    const el=document.getElementById(id); if(!el) return;
    const old=el.value;
    const opts=[...new Set((values||[]).map(v=>String(v||'').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
    el.innerHTML='<option value="">'+safeEsc(first)+'</option>'+opts.map(v=>'<option value="'+safeEsc(v)+'">'+safeEsc(v)+'</option>').join('');
    if(opts.includes(old)) el.value=old; else el.value='';
  }
  window.setJobCardFilterOptions = function(){
    strictSetOptions('cardsOpsFilter','All OPS managers',VALID_OPS);
    strictSetOptions('cardsVesselFilter','All vessels',(window.cards||[]).map(c=>c.vessel_name));
    strictSetOptions('cardsActivityFilter','All activities',(window.cards||[]).map(c=>c.activity_type));
  };
  function rowTotal(r){
    return Number(r.total_cost||0) || (Number(r.labour_cost||0)+Number(r.equipment_cost||0)+Number(r.workshop_cost||0)+Number(r.material_cost||0)+Number(r.other_cost||0));
  }
  function cardByLabel(r){
    if(typeof window.cardCompletedByLabel==='function') return window.cardCompletedByLabel(r);
    const source=nf(r?.source||'import');
    if(r?.completed_by_name) return r.completed_by_name;
    return source.includes('manual') ? 'Crew / App' : 'Imported Excel';
  }
  function cardDay(r){return typeof window.cardDaysLabel==='function'?window.cardDaysLabel(r):1;}
  function sortRows(rows,key,dir){
    if(typeof window.sortByDate==='function') return window.sortByDate(rows,key,dir);
    const m=dir==='asc'?1:-1; return [...rows].sort((a,b)=>((new Date(a[key]||0))-(new Date(b[key]||0)))*m);
  }
  function sliceRows(name, rows){return typeof window.pageSlice==='function'?window.pageSlice(name,rows):rows.slice(0,25);}
  function pager(name, fn, len){return typeof window.makePager==='function'?window.makePager(name,fn,len):'';}

  window.renderCards = function(){
    const q=nf(document.getElementById('cardsSearch')?.value||'');
    const src=nf(document.getElementById('cardsSource')?.value||'');
    const ops=nf(document.getElementById('cardsOpsFilter')?.value||'');
    const vessel=nf(document.getElementById('cardsVesselFilter')?.value||'');
    const activity=nf(document.getElementById('cardsActivityFilter')?.value||'');
    const sort=document.getElementById('cardsDateSort')?.value||'desc';
    let rows=(window.cards||[]).filter(r=>{
      const source=nf(r.source||'imported');
      const isManual=source.includes('manual')||source.includes('offshore');
      const hay=nf([r.job_number,r.client_name,r.location,r.vessel_name,r.ops_manager,r.activity_type,r.staff_names,r.equipment_used,r.description,r.division,cardByLabel(r)].join(' '));
      return (!q || hay.includes(q))
        && (!src || (src==='manual' ? isManual : !isManual))
        && (!ops || nf(r.ops_manager)===ops)
        && (!vessel || nf(r.vessel_name)===vessel)
        && (!activity || nf(r.activity_type)===activity);
    });
    rows=sortRows(rows,'job_date',sort);
    if(typeof window.renderPendingReviews==='function') window.renderPendingReviews();
    const pagerEl=document.getElementById('cardsPager'); if(pagerEl) pagerEl.innerHTML=pager('cards','renderCards',rows.length);
    const body=document.getElementById('cardRows'); if(!body) return;
    body.innerHTML=sliceRows('cards',rows).map(r=>'<tr>'+
      '<td>'+safeEsc(r.source||'import')+'</td>'+
      '<td>'+safeEsc(cardByLabel(r))+'</td>'+
      '<td>'+safeEsc(r.job_number)+'</td>'+
      '<td>'+safeDate(r.job_date)+'</td>'+
      '<td>'+safeNum(cardDay(r))+'</td>'+
      '<td>'+safeEsc(r.client_name)+'</td>'+
      '<td>'+safeEsc(r.location)+'</td>'+
      '<td>'+safeEsc(r.vessel_name)+'</td>'+
      '<td>'+safeEsc(r.ops_manager)+'</td>'+
      '<td>'+safeEsc(r.activity_type)+'</td>'+
      '<td>'+safeEsc(r.staff_names)+'</td>'+
      '<td>'+safeEsc(r.equipment_used)+'</td>'+
      '<td>'+safeEsc(r.description)+'</td>'+
      '<td>'+Number(r.hours||0).toFixed(2)+'</td>'+
      '<td>'+safeMoney(r.labour_cost)+'</td>'+
      '<td>'+safeMoney(r.equipment_cost)+'</td>'+
      '<td>'+safeMoney(r.workshop_cost)+'</td>'+
      '<td>'+safeMoney(r.material_cost)+'</td>'+
      '<td>'+safeMoney(r.other_cost)+'</td>'+
      '<td class="warn">'+safeMoney(rowTotal(r))+'</td>'+
      '<td>'+safeEsc(r.division)+'</td>'+
      '<td class="profit-only"><div class="row-actions"><button class="action-btn edit" onclick="editJobCardRow('+Number(r.id||0)+')">✏️ Edit</button><button class="action-btn delete" onclick="deleteJobCardRow('+Number(r.id||0)+')">🗑 Delete</button></div></td>'+
    '</tr>').join('') || '<tr><td colspan="22" class="muted">No job cards match the current search/filter selection.</td></tr>';
  };

  window.renderJobs = function(){
    const q=nf(document.getElementById('jobSearch')?.value||'');
    const d=nf(document.getElementById('jobDivision')?.value||'');
    const o=nf(document.getElementById('jobOps')?.value||'');
    const sort=document.getElementById('jobsDateSort')?.value||'desc';
    let rows=(window.jobs||[]).filter(j=>{
      const hay=nf([j.job_number,j.invoice_number,j.division,j.ops_manager,j.client_name,j.description].join(' '));
      return (!q||hay.includes(q)) && (!d||nf(j.division)===d) && (!o||nf(j.ops_manager)===o);
    });
    rows=sortRows(rows,'job_date',sort);
    const pagerEl=document.getElementById('jobsPager'); if(pagerEl) pagerEl.innerHTML=pager('jobs','renderJobs',rows.length);
    const body=document.getElementById('jobsRows'); if(!body) return;
    body.innerHTML=sliceRows('jobs',rows).map(j=> typeof window.jobRow==='function' ? window.jobRow(j) : '<tr><td>'+safeEsc(j.job_number)+'</td><td>'+safeDate(j.job_date)+'</td><td>'+safeEsc(j.invoice_number)+'</td><td>'+safeEsc(j.division)+'</td><td>'+safeEsc(j.ops_manager)+'</td><td>'+Number(j.hours||0).toFixed(1)+'</td><td>'+safeMoney(j.revenue)+'</td><td>'+safeMoney(j.total_cost)+'</td><td>'+safeMoney(j.gross_profit)+'</td></tr>').join('') || '<tr><td colspan="9" class="muted">No financial jobs match the current filters.</td></tr>';
  };

  function bind(){
    window.setJobCardFilterOptions();
    ['cardsSearch','cardsSource','cardsOpsFilter','cardsVesselFilter','cardsActivityFilter','cardsDateSort'].forEach(id=>{
      const el=document.getElementById(id); if(!el || el.dataset.v18Bound==='yes') return; el.dataset.v18Bound='yes';
      ['input','change','keyup','search','paste','blur'].forEach(ev=>el.addEventListener(ev,()=>setTimeout(()=>{ if(window.resetPage) resetPage('cards'); window.renderCards(); },0)));
    });
    ['jobSearch','jobDivision','jobOps','jobsDateSort'].forEach(id=>{
      const el=document.getElementById(id); if(!el || el.dataset.v18Bound==='yes') return; el.dataset.v18Bound='yes';
      ['input','change','keyup','search','paste','blur'].forEach(ev=>el.addEventListener(ev,()=>setTimeout(()=>{ if(window.resetPage) resetPage('jobs'); window.renderJobs(); },0)));
    });
    try{window.renderCards();window.renderJobs();}catch(e){console.warn('B4 v18 filter render failed',e)}
  }
  const oldLoad=window.loadAll;
  if(typeof oldLoad==='function'){
    window.loadAll=async function(){const r=await oldLoad.apply(this,arguments); setTimeout(bind,0); return r;};
  }
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',()=>setTimeout(bind,500)); else setTimeout(bind,500);
})();
</script>`;

self.addEventListener('fetch', event => {
  const req = event.request;
  if (req.mode === 'navigate' || new URL(req.url).pathname === '/' || new URL(req.url).pathname.endsWith('/index.html')) {
    event.respondWith((async () => {
      const res = await fetch(req, { cache: 'no-store' });
      const type = res.headers.get('content-type') || '';
      if (!type.includes('text/html')) return res;
      let html = await res.text();
      if (!html.includes('__B4_FILTER_HARDFIX_V18__')) {
        html = html.replace('</body>', HOTFIX_SCRIPT + '\n</body>');
      }
      return new Response(html, { headers: { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store' } });
    })());
    return;
  }
  event.respondWith(fetch(req, { cache: 'no-store' }).catch(() => caches.match(req)));
});
