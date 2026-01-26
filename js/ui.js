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
    if (btn.textContent.toLowerCase().includes(sectionId.replace('tradein', 'trade-in'))) {
      btn.classList.add('active');
    }
  });
  
  if (sectionId === 'dashboard') loadDashboard();
  else if (sectionId === 'products') loadProducts();
  else if (sectionId === 'pricerate') loadPriceRate();
  else if (sectionId === 'sell') {
    loadSells();
  }
  else if (sectionId === 'tradein') loadTradeins();
  else if (sectionId === 'exchange') loadExchanges();
  else if (sectionId === 'buyback') loadBuybacks();
  else if (sectionId === 'withdraw') loadWithdraws();
  else if (sectionId === 'inventory') {
    loadInventory();
    loadStock();
  }
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
