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
  
  const sectionToTabMap = {
    'dashboard': 'dashboard',
    'products': 'products',
    'pricerate': 'price rate',
    'sell': 'sell',
    'tradein': 'trade-in',
    'exchange': 'exchange',
    'switch': 'switch',
    'freeexchange': 'free ex',
    'buyback': 'buyback',
    'withdraw': 'withdraw',
    'inventory': 'inventory',
    'cashbank': 'cash/bank',
    'accounting': 'accounting',
    'reports': 'reports'
  };
  
  const tabName = sectionToTabMap[sectionId];
  if (tabName) {
    document.querySelectorAll('.nav-btn').forEach(btn => {
      if (btn.textContent.toLowerCase().includes(tabName)) {
        btn.classList.add('active');
      }
    });
  }
  
  if (sectionId === 'dashboard') loadDashboard();
  else if (sectionId === 'products') loadProducts();
  else if (sectionId === 'pricerate') loadPriceRate();
  else if (sectionId === 'sell') {
    loadSells();
  }
  else if (sectionId === 'tradein') loadTradeins();
  else if (sectionId === 'exchange') loadExchanges();
  else if (sectionId === 'switch') loadSwitches();
  else if (sectionId === 'freeexchange') loadFreeExchanges();
  else if (sectionId === 'buyback') loadBuybacks();
  else if (sectionId === 'withdraw') loadWithdraws();
  else if (sectionId === 'inventory') loadInventory();
  else if (sectionId === 'cashbank') loadCashBank();
  else if (sectionId === 'accounting') loadAccounting();
  else if (sectionId === 'reports') loadReports();
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