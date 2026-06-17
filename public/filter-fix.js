(function(){
  if(window.__B4_STABLE_FILTER_FIX__) return;
  window.__B4_STABLE_FILTER_FIX__ = true;

  const VALID_OPS = ['Jolanda','Eugen','Carmon','Bresler','Elliotte'];
  const clean = v => String(v || '').trim();
  const uniq = values => [...new Set((values || []).map(clean).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  const escapeText = v => {
    try { return typeof esc === 'function' ? esc(v) : String(v || ''); }
    catch(e) { return String(v || ''); }
  };

  function setSelect(id, firstLabel, values, fixedList){
    const el = document.getElementById(id);
    if(!el) return;
    const old = el.value;
    const opts = uniq(values);
    el.innerHTML = '<option value="">' + escapeText(firstLabel) + '</option>' + opts.map(v => '<option>' + escapeText(v) + '</option>').join('');
    el.value = opts.includes(old) ? old : '';
    if(fixedList && old && !opts.includes(old)) el.value = '';
  }

  window.setJobCardFilterOptions = function(){
    setSelect('cardsOpsFilter', 'All OPS managers', VALID_OPS, true);
    setSelect('cardsVesselFilter', 'All vessels', (window.cards || []).map(c => c.vessel_name));
    setSelect('cardsActivityFilter', 'All activities', (window.cards || []).map(c => c.activity_type));
  };

  window.setFinancialJobFilterOptions = function(){
    setSelect('jobDivision', 'All divisions', (window.jobs || []).map(j => j.division));
    setSelect('jobOps', 'All OPS managers', VALID_OPS, true);
  };

  function resetTable(page){
    try { if(typeof resetPage === 'function') resetPage(page); else if(window.tablePages) window.tablePages[page] = 1; }
    catch(e) {}
  }

  function safeCall(fn){
    try { if(typeof fn === 'function') fn(); } catch(e) { console.warn('Filter refresh failed', e); }
  }

  function bind(ids, page, renderName){
    ids.forEach(id => {
      const el = document.getElementById(id);
      if(!el || el.dataset.stableFilterBound === 'yes') return;
      el.dataset.stableFilterBound = 'yes';
      const refresh = () => setTimeout(() => { resetTable(page); safeCall(window[renderName]); }, 0);
      ['input','keyup','search','change','paste','blur'].forEach(ev => el.addEventListener(ev, refresh));
    });
  }

  function applyFilterFix(){
    window.setJobCardFilterOptions();
    window.setFinancialJobFilterOptions();
    bind(['cardsSearch','cardsSource','cardsOpsFilter','cardsVesselFilter','cardsActivityFilter','cardsDateSort'], 'cards', 'renderCards');
    bind(['jobSearch','jobDivision','jobOps','jobsDateSort'], 'jobs', 'renderJobs');
    bind(['registerSearch','registerDateSort'], 'register', 'renderRegister');
    safeCall(window.renderCards);
    safeCall(window.renderJobs);
    safeCall(window.renderRegister);
  }

  const originalLoadAll = window.loadAll;
  if(typeof originalLoadAll === 'function' && !originalLoadAll.__stableFilterWrapped){
    const wrapped = async function(){
      const result = await originalLoadAll.apply(this, arguments);
      setTimeout(applyFilterFix, 0);
      return result;
    };
    wrapped.__stableFilterWrapped = true;
    window.loadAll = wrapped;
  }

  if(document.readyState === 'loading') document.addEventListener('DOMContentLoaded', () => setTimeout(applyFilterFix, 500));
  else setTimeout(applyFilterFix, 500);
})();
