function showSection(sectionId) {
  document.querySelectorAll('.section').forEach(section => {
    section.classList.remove('active');
  });
  document.querySelectorAll('.nav-btn').forEach(btn => {
    btn.classList.remove('active');
  });
  
  const targetSection = document.getElementById(sectionId);
  if (targetSection) {
    targetSection.classList.add('active');
  }
  
  document.querySelectorAll('.nav-btn').forEach(btn => {
    var oc = btn.getAttribute('onclick') || '';
    if (oc.indexOf("'" + sectionId + "'") !== -1) {
      btn.classList.add('active');
    }
  });
  
  if (sectionId !== 'dashboard' && typeof stopDashReportRefresh === 'function') {
    stopDashReportRefresh();
  }

  const loaderMap = {
    'dashboard': 'loadDashboard',
    'products': 'loadProducts',
    'pricerate': 'loadPriceRate',
    'sell': 'loadSells',
    'tradein': 'loadTradeins',
    'exchange': 'loadExchanges',
    'switch': 'loadSwitches',
    'freeexchange': 'loadFreeExchanges',
    'buyback': 'loadBuybacks',
    'withdraw': 'loadWithdraws',
    'inventory': 'loadInventory',
    'cashbank': 'loadCashBank',
    'accounting': 'loadAccounting',
    'diff': 'loadDiff',
    'reports': 'loadReports',
    'stockold': 'loadStockOld',
    'stocknew': 'loadStockNew',
    'wac': 'loadWAC'
  };
  const fn = loaderMap[sectionId];
  if (fn && typeof window[fn] === 'function') window[fn]();
}

function showLoading() {
  const loader = document.getElementById('loader');
  if (loader) loader.style.display = 'flex';
}

function hideLoading() {
  const loader = document.getElementById('loader');
  if (loader) loader.style.display = 'none';
}

document.querySelectorAll('.modal').forEach(modal => {
  modal.addEventListener('click', function(e) {
    if (e.target === modal) {
      modal.classList.remove('active');
    }
  });
});