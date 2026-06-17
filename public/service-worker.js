const CACHE_NAME='b4-ops-cache-v19-filter-pipeline';

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
  if(window.__B4_FILTER_PIPELINE_V19__) return;
  window.__B4_FILTER_PIPELINE_V19__ = true;

  const VALID_OPS = ['Jolanda','Eugen','Carmon','Bresler','Elliotte'];

  function g(name, fallback){
    try{
      const value = Function('try{return typeof '+name+'!=="undefined"?'+name+':undefined}catch(e){return undefined}')();
      return value == null ? fallback : value;
    }catch(e){ return fallback; }
  }
  function call(name, args, fallback){
    try{
      const fn = g(name, null);
      return typeof fn === 'function' ? fn.apply(window, args || []) : fallback;
    }catch(e){ return fallback; }
  }
  function nf(v){
    return String(v ?? '')
      .normalize('NFD').replace(/[\\u0300-\\u036f]/g,'')
      .trim().toLowerCase().replace(/\\s+/g,' ');
  }
  window.normFilter = nf;
  function escSafe(v){return call('esc',[v],String(v ?? '').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])))}
  function moneySafe(v){return call('money',[v],'N$ '+Number(v||0).toFixed(2))}
  function dateSafe(v){return call('date',[v],v?String(v).slice(0,10):'')}
  function numSafe(v){return call('num',[v],Number(v||0).toLocaleString())}
  function rows(name){const value=g(name,[]); return Array.isArray(value) ? value : []}
  function sortRows(list,key,dir){return call('sortByDate',[list,key,dir],[...list].sort((a,b)=>((new Date(a[key]||0))-(new Date(b[key]||0)))*(dir==='asc'?1:-1)))}
  function pageRows(name,list){return call('pageSlice',[name,list],list.slice(0,25))}
  function pager(name,fn,len){return call('makePager',[name,fn,len],'')}
  function reset(name){call('resetPage',[name],null); try{const tp=g('tablePages',null); if(tp) tp[name]=1;}catch(e){} }
  function setOptionsStrict(id, first, values){
    const el=document.getElementById(id); if(!el) return;
    const old=el.value;
    const opts=[...new Set((values||[]).map(v=>String(v||'').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
    el.innerHTML='<option value="">'+escSafe(first)+'</option>'+opts.map(v=>'<option value="'+escSafe(v)+'">'+escSafe(v)+'</option>').join('');
    el.value = opts.includes(old) ? old : '';
  }
  function cardCompletedBy(r){return call('cardCompletedByLabel',[r],r?.completed_by_name || (nf(r?.source).includes('manual')||nf(r?.source).includes('offshore')?'Crew / App':'Imported Excel'))}
  function cardDays(r){return call('cardDaysLabel',[r],1)}
  function rowTotal(r){return Number(r.total_cost||0) || Number(r.labour_cost||0)+Number(r.equipment_cost||0)+Number(r.workshop_cost||0)+Number(r.material_cost||0)+Number(r.other_cost||0)}
  function regText(r){return nf([r.job_number,r.job_date,r.ops_manager,r.division,r.client_name,r.description,r.quote_number,r.po_number,r.invoice_number,r.value_incl_vat,r.value_excl_vat,r.client_feedback].join(' '))}

  window.setJobCardFilterOptions = function(){
    setOptionsStrict('cardsOpsFilter','All OPS managers',VALID_OPS);
    setOptionsStrict('cardsVesselFilter','All vessels',rows('cards').map(c=>c.vessel_name));
    setOptionsStrict('cardsActivityFilter','All activities',rows('cards').map(c=>c.activity_type));
  };

  window.setFinancialJobFilterOptionsV19 = function(){
    setOptionsStrict('jobDivision','All divisions',rows('jobs').map(j=>j.division));
    setOptionsStrict('jobOps','All OPS managers',VALID_OPS);
  };

  window.renderCards = function(){
    const q=nf(document.getElementById('cardsSearch')?.value||'');
    const src=nf(document.getElementById('cardsSource')?.value||'');
    const ops=nf(document.getElementById('cardsOpsFilter')?.value||'');
    const vessel=nf(document.getElementById('cardsVesselFilter')?.value||'');
    const activity=nf(document.getElementById('cardsActivityFilter')?.value||'');
    const sort=document.getElementById('cardsDateSort')?.value||'desc';
    let filteredRows=rows('cards').filter(r=>{
      const source=nf(r.source||'imported');
      const isManual=source.includes('manual')||source.includes('offshore');
      const hay=nf([r.job_number,r.client_name,r.location,r.vessel_name,r.ops_manager,r.activity_type,r.staff_names,r.equipment_used,r.description,r.division,cardCompletedBy(r)].join(' '));
      return (!q || hay.includes(q))
        && (!src || (src==='manual' ? isManual : !isManual))
        && (!ops || nf(r.ops_manager)===ops)
        && (!vessel || nf(r.vessel_name)===vessel)
        && (!activity || nf(r.activity_type)===activity);
    });
    filteredRows=sortRows(filteredRows,'job_date',sort);
    call('renderPendingReviews',[],null);
    const pagerEl=document.getElementById('cardsPager'); if(pagerEl) pagerEl.innerHTML=pager('cards','renderCards',filteredRows.length);
    const body=document.getElementById('cardRows'); if(!body) return;
    body.innerHTML=pageRows('cards',filteredRows).map(r=>'<tr>'+
      '<td>'+escSafe(r.source||'import')+'</td><td>'+escSafe(cardCompletedBy(r))+'</td><td>'+escSafe(r.job_number)+'</td><td>'+dateSafe(r.job_date)+'</td><td>'+numSafe(cardDays(r))+'</td><td>'+escSafe(r.client_name)+'</td><td>'+escSafe(r.location)+'</td><td>'+escSafe(r.vessel_name)+'</td><td>'+escSafe(r.ops_manager)+'</td><td>'+escSafe(r.activity_type)+'</td><td>'+escSafe(r.staff_names)+'</td><td>'+escSafe(r.equipment_used)+'</td><td>'+escSafe(r.description)+'</td><td>'+Number(r.hours||0).toFixed(2)+'</td><td>'+moneySafe(r.labour_cost)+'</td><td>'+moneySafe(r.equipment_cost)+'</td><td>'+moneySafe(r.workshop_cost)+'</td><td>'+moneySafe(r.material_cost)+'</td><td>'+moneySafe(r.other_cost)+'</td><td class="warn">'+moneySafe(rowTotal(r))+'</td><td>'+escSafe(r.division)+'</td><td class="profit-only"><div class="row-actions"><button class="action-btn edit" onclick="editJobCardRow('+Number(r.id||0)+')">✏️ Edit</button><button class="action-btn delete" onclick="deleteJobCardRow('+Number(r.id||0)+')">🗑 Delete</button></div></td></tr>'
    ).join('') || '<tr><td colspan="22" class="muted">No job cards match the current search/filter selection.</td></tr>';
  };

  window.renderJobs = function(){
    const q=nf(document.getElementById('jobSearch')?.value||'');
    const d=nf(document.getElementById('jobDivision')?.value||'');
    const o=nf(document.getElementById('jobOps')?.value||'');
    const sort=document.getElementById('jobsDateSort')?.value||'desc';
    let filteredRows=rows('jobs').filter(j=>{
      const hay=nf([j.job_number,j.invoice_number,j.division,j.ops_manager,j.client_name,j.description].join(' '));
      return (!q || hay.includes(q)) && (!d || nf(j.division)===d) && (!o || nf(j.ops_manager)===o);
    });
    filteredRows=sortRows(filteredRows,'job_date',sort);
    const pagerEl=document.getElementById('jobsPager'); if(pagerEl) pagerEl.innerHTML=pager('jobs','renderJobs',filteredRows.length);
    const body=document.getElementById('jobsRows'); if(!body) return;
    body.innerHTML=pageRows('jobs',filteredRows).map(j=>call('jobRow',[j],'<tr><td>'+escSafe(j.job_number)+'</td><td>'+dateSafe(j.job_date)+'</td><td>'+escSafe(j.invoice_number)+'</td><td>'+escSafe(j.division)+'</td><td>'+escSafe(j.ops_manager)+'</td><td>'+Number(j.hours||0).toFixed(1)+'</td><td>'+moneySafe(j.revenue)+'</td><td>'+moneySafe(j.total_cost)+'</td><td>'+moneySafe(j.gross_profit)+'</td></tr>')).join('') || '<tr><td colspan="9" class="muted">No financial jobs match the current filters.</td></tr>';
  };

  window.renderRegister = function(){
    const q=nf(document.getElementById('registerSearch')?.value||'');
    const sort=document.getElementById('registerDateSort')?.value||'desc';
    let base=rows('registerSearchResults');
    if(!base.length) base=rows('registers');
    let filteredRows=base.filter(r=>!q || regText(r).includes(q));
    filteredRows=sortRows(filteredRows,'job_date',sort);
    const pagerEl=document.getElementById('registerPager'); if(pagerEl) pagerEl.innerHTML=pager('register','renderRegister',filteredRows.length);
    const body=document.getElementById('registerRows'); if(!body) return;
    body.innerHTML=pageRows('register',filteredRows).map(r=>'<tr><td>'+escSafe(r.job_number)+'</td><td>'+dateSafe(r.job_date)+'</td><td>'+escSafe(r.ops_manager)+'</td><td>'+escSafe(r.division)+'</td><td>'+escSafe(r.client_name)+'</td><td>'+escSafe(r.description)+'</td><td>'+escSafe(r.quote_number)+'</td><td>'+escSafe(r.po_number)+'</td><td>'+escSafe(r.invoice_number)+'</td><td>'+moneySafe(r.value_incl_vat)+'</td><td>'+moneySafe(r.value_excl_vat)+'</td></tr>').join('') || '<tr><td colspan="11" class="muted">No Job Register rows match the current search/filter selection.</td></tr>';
  };

  function bind(ids,page,renderFn){
    ids.forEach(id=>{
      const el=document.getElementById(id); if(!el || el.dataset.v19Bound==='yes') return;
      el.dataset.v19Bound='yes';
      ['input','change','keyup','search','paste','blur'].forEach(ev=>el.addEventListener(ev,()=>setTimeout(()=>{reset(page); renderFn();},0)));
    });
  }
  function apply(){
    try{ window.setFinancialJobFilterOptionsV19(); window.setJobCardFilterOptions(); }catch(e){console.warn('B4 v19 option setup failed',e)}
    bind(['cardsSearch','cardsSource','cardsOpsFilter','cardsVesselFilter','cardsActivityFilter','cardsDateSort'],'cards',window.renderCards);
    bind(['jobSearch','jobDivision','jobOps','jobsDateSort'],'jobs',window.renderJobs);
    bind(['registerSearch','registerDateSort'],'register',window.renderRegister);
    try{window.renderCards();window.renderJobs();window.renderRegister();}catch(e){console.warn('B4 v19 render failed',e)}
  }
  const oldLoad=g('loadAll',null);
  if(typeof oldLoad==='function'){
    window.loadAll=async function(){const out=await oldLoad.apply(this,arguments); setTimeout(apply,0); return out;};
  }
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',()=>setTimeout(apply,700)); else setTimeout(apply,700);
})();
</script>`;

self.addEventListener('fetch', event => {
  const req = event.request;
  const url = new URL(req.url);
  if (req.mode === 'navigate' || url.pathname === '/' || url.pathname.endsWith('/index.html')) {
    event.respondWith((async () => {
      const res = await fetch(req, { cache: 'no-store' });
      const type = res.headers.get('content-type') || '';
      if (!type.includes('text/html')) return res;
      let html = await res.text();
      if (!html.includes('__B4_FILTER_PIPELINE_V19__')) {
        html = html.replace('</body>', HOTFIX_SCRIPT + '\n</body>');
      }
      return new Response(html, { headers: { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store' } });
    })());
    return;
  }
  event.respondWith(fetch(req, { cache: 'no-store' }).catch(() => caches.match(req)));
});
