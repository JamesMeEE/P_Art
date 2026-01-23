function showSection(sectionId) {
  document.querySelectorAll('.section').forEach(section => {
    section.classList.remove('active');
  });
  document.querySelectorAll('.nav-btn').forEach(btn => {
    btn.classList.remove('active');
  });
  document.getElementById(sectionId).classList.add('active');
  event.target.classList.add('active');
  
  if (sectionId === 'dashboard') loadDashboard();
  else if (sectionId === 'products') loadProducts();
  else if (sectionId === 'pricerate') loadPriceRate();
  else if (sectionId === 'sell') loadSells();
  else if (sectionId === 'tradein') loadTradeins();
  else if (sectionId === 'buyback') loadBuybacks();
  else if (sectionId === 'exchange') loadExchanges();
  else if (sectionId === 'withdraw') loadWithdraw();
  else if (sectionId === 'inventory') loadInventory();
  else if (sectionId === 'cashbank') loadCashBank();
  else if (sectionId === 'accounting') loadAccounting();
  else if (sectionId === 'reports') loadReports();
}

document.querySelectorAll('.modal').forEach(modal => {
  modal.addEventListener('click', function(e) {
    if (e.target === modal) {
      modal.classList.remove('active');
    }
  });
});
