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
    try { 
      if(typeof tablePages === 'object') tablePages[page] = 1;
    }
    catch(e) { console.warn('Reset table page failed:', e); }
  }

  function safeCall(fn){
    try { if(typeof fn === 'function') fn(); } catch(e) { console.warn('Filter refresh failed', e); }
  }

  function bindFilterElement(id, resetPage, renderFn){
    const el = document.getElementById(id);
    if(!el || el.dataset.stableFilterBound === 'yes') return;
    el.dataset.stableFilterBound = 'yes';
    
    const refresh = () => {
      resetTable(resetPage);
      setTimeout(() => { safeCall(window[renderFn]); }, 0);
    };
    
    // Bind multiple event types for robust filter triggering
    ['input','keyup','search','change','paste','blur','click'].forEach(ev => {
      el.addEventListener(ev, refresh, false);
    });
  }

  function applyFilterFix(){
    // Set up filter options
    window.setJobCardFilterOptions();
    window.setFinancialJobFilterOptions();
    
    // Bind Job Cards filters
    bindFilterElement('cardsSearch', 'cards', 'renderCards');
    bindFilterElement('cardsSource', 'cards', 'renderCards');
    bindFilterElement('cardsOpsFilter', 'cards', 'renderCards');
    bindFilterElement('cardsVesselFilter', 'cards', 'renderCards');
    bindFilterElement('cardsActivityFilter', 'cards', 'renderCards');
    bindFilterElement('cardsDateSort', 'cards', 'renderCards');
    
    // Bind Financial Jobs filters
    bindFilterElement('jobSearch', 'jobs', 'renderJobs');
    bindFilterElement('jobDivision', 'jobs', 'renderJobs');
    bindFilterElement('jobOps', 'jobs', 'renderJobs');
    bindFilterElement('jobsDateSort', 'jobs', 'renderJobs');
    
    // Bind Job Register filters
    bindFilterElement('registerSearch', 'register', 'renderRegister');
    bindFilterElement('registerDateSort', 'register', 'renderRegister');
    
    // Bind Expense filters
    bindFilterElement('expenseSearch', 'expenses', 'renderExpenses');
    bindFilterElement('expenseCategory', 'expenses', 'renderExpenses');
    bindFilterElement('expenseSupplier', 'expenses', 'renderExpenses');
    bindFilterElement('expenseVessel', 'expenses', 'renderExpenses');
    bindFilterElement('expenseDivision', 'expenses', 'renderExpenses');
    bindFilterElement('expenseYear', 'expenses', 'renderExpenses');
    bindFilterElement('expenseMonth', 'expenses', 'renderExpenses');
    bindFilterElement('expenseDateSort', 'expenses', 'renderExpenses');
    
    // Bind Vessel filters
    bindFilterElement('vesselSearch', 'vessels', 'renderVessels');
    bindFilterElement('vesselFilter', 'vessels', 'renderVessels');
    bindFilterElement('vesselClient', 'vessels', 'renderVessels');
    bindFilterElement('vesselOpsFilter', 'vessels', 'renderVessels');
    bindFilterElement('vesselDivision', 'vessels', 'renderVessels');
    bindFilterElement('vesselYear', 'vessels', 'renderVessels');
    bindFilterElement('vesselMonthFilter', 'vessels', 'renderVessels');
    bindFilterElement('vesselProfitStatus', 'vessels', 'renderVessels');
    bindFilterElement('vesselDateSort', 'vessels', 'renderVessels');
    
    // Bind Invoice filters
    bindFilterElement('invoiceSearch', 'invoice', 'renderInvoice');
    bindFilterElement('invoiceStatus', 'invoice', 'renderInvoice');
    bindFilterElement('invoiceOps', 'invoice', 'renderInvoice');
    bindFilterElement('invoiceDivision', 'invoice', 'renderInvoice');
    bindFilterElement('invoiceDateSort', 'invoice', 'renderInvoice');
    
    // Initial render calls
    safeCall(window.renderCards);
    safeCall(window.renderJobs);
    safeCall(window.renderRegister);
  }

  // Wrap loadAll to ensure filter fix is applied after data loads
  const originalLoadAll = window.loadAll;
  if(typeof originalLoadAll === 'function' && !originalLoadAll.__stableFilterWrapped){
    const wrapped = async function(){
      const result = await originalLoadAll.apply(this, arguments);
      setTimeout(applyFilterFix, 100);
      return result;
    };
    wrapped.__stableFilterWrapped = true;
    window.loadAll = wrapped;
  }

  // Apply on DOMContentLoaded and periodically re-check
  if(document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => setTimeout(applyFilterFix, 500));
  } else {
    setTimeout(applyFilterFix, 500);
  }
  
  // Re-apply filter bindings after a delay to catch dynamically added elements
  setTimeout(applyFilterFix, 1500);
})();
