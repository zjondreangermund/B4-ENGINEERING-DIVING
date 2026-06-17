(function(){
  // Prevent multiple loads
  if(window.__B4_FILTER_FIX_LOADED__) return;
  window.__B4_FILTER_FIX_LOADED__ = true;

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

  // Expose global functions for filter setup
  window.setJobCardFilterOptions = function(){
    setSelect('cardsOpsFilter', 'All OPS managers', VALID_OPS, true);
    setSelect('cardsVesselFilter', 'All vessels', (window.cards || []).map(c => c.vessel_name));
    setSelect('cardsActivityFilter', 'All activities', (window.cards || []).map(c => c.activity_type));
  };

  window.setFinancialJobFilterOptions = function(){
    setSelect('jobDivision', 'All divisions', (window.jobs || []).map(j => j.division));
    setSelect('jobOps', 'All OPS managers', VALID_OPS, true);
  };

  function applyFilterFix(){
    try {
      // Set up filter dropdowns
      if(typeof window.setJobCardFilterOptions === 'function') window.setJobCardFilterOptions();
      if(typeof window.setFinancialJobFilterOptions === 'function') window.setFinancialJobFilterOptions();
      
      // Trigger initial renders to ensure data is displayed
      if(typeof window.renderCards === 'function') window.renderCards();
      if(typeof window.renderJobs === 'function') window.renderJobs();
      if(typeof window.renderRegister === 'function') window.renderRegister();
    } catch(e) {
      console.warn('Filter setup error:', e);
    }
  }

  // Hook into loadAll to apply filters after data loads
  const originalLoadAll = window.loadAll;
  if(typeof originalLoadAll === 'function' && !originalLoadAll.__filterFixWrapped){
    window.loadAll = async function(){
      const result = await originalLoadAll.apply(this, arguments);
      setTimeout(applyFilterFix, 50);
      return result;
    };
    window.loadAll.__filterFixWrapped = true;
  }

  // Apply on page load
  if(document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => setTimeout(applyFilterFix, 300));
  } else {
    setTimeout(applyFilterFix, 300);
  }
})();
