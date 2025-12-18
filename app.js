const CONFIG = {
  SHEET_ID: '1FF4odviKZ2LnRvPf8ltM0o_jxM0ZHuJHBlkQCjC3sxA',
  API_KEY: 'AIzaSyAB6yjxTB0TNbEk2C68aOP5u0IkdmK12tg',
  APPS_SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbziDXIkJa_VIXVJpRnwv5aYDq425OU5O1vkDvMXEDmzj5KAzg80PJQFtN5DKOmlv0qp/exec'
};

function showLoading() {
  document.getElementById('loading').classList.add('active');
}

function hideLoading() {
  document.getElementById('loading').classList.remove('active');
}

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
  else if (sectionId === 'pricing') loadPricing();
  else if (sectionId === 'sales') loadSales();
  else if (sectionId === 'purchases') loadPurchases();
  else if (sectionId === 'inventory') loadInventory();
  else if (sectionId === 'audit') loadAudits();
  else if (sectionId === 'logs') loadLogs();
  else if (sectionId === 'reports') loadReports();
}

function openModal(modalId) {
  document.getElementById(modalId).classList.add('active');
}

function closeModal(modalId) {
  document.getElementById(modalId).classList.remove('active');
}

async function fetchSheetData(range) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.SHEET_ID}/values/${range}?key=${CONFIG.API_KEY}`;
  const response = await fetch(url);
  if (!response.ok) throw new Error('Failed to fetch data');
  const data = await response.json();
  return data.values || [];
}

async function callAppsScript(action, params) {
  const url = `${CONFIG.APPS_SCRIPT_URL}?action=${action}&${new URLSearchParams(params).toString()}`;
  const response = await fetch(url);
  if (!response.ok) throw new Error('Failed to call API');
  return await response.json();
}

async function loadDashboard() {
  try {
    showLoading();
    const [salesData, purchasesData, productsData] = await Promise.all([
      fetchSheetData('Sales!A:D'),
      fetchSheetData('Purchases!A:D'),
      fetchSheetData('Products!A:G')
    ]);

    const totalSales = salesData.slice(1).reduce((sum, row) => sum + (parseFloat(row[3]) || 0), 0);
    const totalPurchases = purchasesData.slice(1).reduce((sum, row) => sum + (parseFloat(row[3]) || 0), 0);
    const totalProducts = productsData.length - 1;
    const netProfit = totalSales - totalPurchases;

    document.getElementById('totalSales').textContent = formatNumber(totalSales);
    document.getElementById('totalPurchases').textContent = formatNumber(totalPurchases);
    document.getElementById('totalProducts').textContent = totalProducts;
    document.getElementById('netProfit').textContent = formatNumber(netProfit);

    const logs = await fetchSheetData('Logs!A:E');
    const recentLogs = logs.slice(-5).reverse();
    document.getElementById('recentActivity').innerHTML = recentLogs.length > 0
      ? recentLogs.map(log => `<p style="color: var(--text-secondary); margin-bottom: 10px; padding: 10px; border-left: 2px solid var(--gold-primary);">${log[1]} - ${log[3]}</p>`).join('')
      : '<p style="color: var(--text-secondary); text-align: center; padding: 20px;">No recent activity</p>';

    hideLoading();
  } catch (error) {
    console.error('Error loading dashboard:', error);
    hideLoading();
  }
}

async function loadProducts() {
  try {
    showLoading();
    const data = await fetchSheetData('Products!A:G');
    const tbody = document.getElementById('productsTable');
    
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="7" style="text-align: center; padding: 40px;">No products available</td></tr>';
    } else {
      tbody.innerHTML = data.slice(1).map(row => `
        <tr>
          <td>${row[0]}</td>
          <td>${row[1]}</td>
          <td>${row[2]}</td>
          <td>${row[3]}</td>
          <td>${row[4]}</td>
          <td>${row[5]}</td>
          <td>${formatDate(row[6])}</td>
        </tr>
      `).join('');
    }
    hideLoading();
  } catch (error) {
    console.error('Error loading products:', error);
    hideLoading();
  }
}

async function loadPricing() {
  try {
    showLoading();
    const data = await fetchSheetData('PriceConfig!A:D');
    if (data.length > 1) {
      document.getElementById('buyPrice').value = data[1][0] || '';
      document.getElementById('sellPrice').value = data[1][1] || '';
      document.getElementById('exchangeRate').value = data[1][2] || '';
    }
    hideLoading();
  } catch (error) {
    console.error('Error loading pricing:', error);
    hideLoading();
  }
}

async function loadSales() {
  try {
    showLoading();
    const data = await fetchSheetData('Sales!A:G');
    const tbody = document.getElementById('salesTable');
    
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="6" style="text-align: center; padding: 40px;">No sales records</td></tr>';
    } else {
      tbody.innerHTML = data.slice(1).map(row => `
        <tr>
          <td>${row[0]}</td>
          <td>${row[1]}</td>
          <td>${formatNumber(row[3])}</td>
          <td><span class="status-badge status-${row[4].toLowerCase()}">${row[4]}</span></td>
          <td>${formatDate(row[5])}</td>
          <td>${row[6]}</td>
        </tr>
      `).join('');
    }
    hideLoading();
  } catch (error) {
    console.error('Error loading sales:', error);
    hideLoading();
  }
}

async function loadPurchases() {
  try {
    showLoading();
    const data = await fetchSheetData('Purchases!A:G');
    const tbody = document.getElementById('purchasesTable');
    
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="6" style="text-align: center; padding: 40px;">No purchase records</td></tr>';
    } else {
      tbody.innerHTML = data.slice(1).map(row => `
        <tr>
          <td>${row[0]}</td>
          <td>${row[1]}</td>
          <td>${formatNumber(row[3])}</td>
          <td><span class="status-badge status-${row[4].toLowerCase()}">${row[4]}</span></td>
          <td>${formatDate(row[5])}</td>
          <td>${row[6]}</td>
        </tr>
      `).join('');
    }
    hideLoading();
  } catch (error) {
    console.error('Error loading purchases:', error);
    hideLoading();
  }
}

async function loadInventory() {
  try {
    showLoading();
    const data = await fetchSheetData('Inventory!A:G');
    const tbody = document.getElementById('inventoryTable');
    
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="7" style="text-align: center; padding: 40px;">No inventory transactions</td></tr>';
    } else {
      tbody.innerHTML = data.slice(1).map(row => `
        <tr>
          <td>${row[0]}</td>
          <td>${row[1]}</td>
          <td>${row[2]}</td>
          <td>${row[3]}</td>
          <td>${formatDate(row[4])}</td>
          <td>${row[5]}</td>
          <td>${row[6]}</td>
        </tr>
      `).join('');
    }
    hideLoading();
  } catch (error) {
    console.error('Error loading inventory:', error);
    hideLoading();
  }
}

async function loadAudits() {
  try {
    showLoading();
    const data = await fetchSheetData('Audits!A:I');
    const tbody = document.getElementById('auditTable');
    
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 40px;">No audit records</td></tr>';
    } else {
      tbody.innerHTML = data.slice(1).map(row => `
        <tr>
          <td>${row[0]}</td>
          <td>${row[1]}</td>
          <td>${formatNumber(row[2])}</td>
          <td>${formatNumber(row[3])}</td>
          <td>${row[4]}</td>
          <td>${row[5]}</td>
          <td>${formatDate(row[6])}</td>
          <td>${row[7]}</td>
        </tr>
      `).join('');
    }
    hideLoading();
  } catch (error) {
    console.error('Error loading audits:', error);
    hideLoading();
  }
}

async function loadLogs() {
  try {
    showLoading();
    const data = await fetchSheetData('Logs!A:E');
    const tbody = document.getElementById('logsTable');
    
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="5" style="text-align: center; padding: 40px;">No logs available</td></tr>';
    } else {
      tbody.innerHTML = data.slice(1).reverse().map(row => `
        <tr>
          <td>${row[0]}</td>
          <td>${row[1]}</td>
          <td>${row[2]}</td>
          <td>${formatDate(row[3])}</td>
          <td>${row[4]}</td>
        </tr>
      `).join('');
    }
    hideLoading();
  } catch (error) {
    console.error('Error loading logs:', error);
    hideLoading();
  }
}

async function loadReports() {
  try {
    showLoading();
    const [salesData, purchasesData, productsData] = await Promise.all([
      fetchSheetData('Sales!A:D'),
      fetchSheetData('Purchases!A:D'),
      fetchSheetData('Products!A:G')
    ]);

    const totalSales = salesData.slice(1).reduce((sum, row) => sum + (parseFloat(row[3]) || 0), 0);
    const totalPurchases = purchasesData.slice(1).reduce((sum, row) => sum + (parseFloat(row[3]) || 0), 0);
    const totalProducts = productsData.length - 1;
    const netProfit = totalSales - totalPurchases;

    document.getElementById('reportContent').innerHTML = `
      <div style="margin: 20px 0;">
        <div style="margin-bottom: 15px;">
          <span style="color: var(--text-secondary);">Total Sales:</span>
          <span style="color: var(--gold-primary); font-weight: 600; float: right;">${formatNumber(totalSales)} THB</span>
        </div>
        <div style="margin-bottom: 15px;">
          <span style="color: var(--text-secondary);">Total Purchases:</span>
          <span style="color: var(--gold-primary); font-weight: 600; float: right;">${formatNumber(totalPurchases)} THB</span>
        </div>
        <div style="margin-bottom: 15px;">
          <span style="color: var(--text-secondary);">Total Products:</span>
          <span style="color: var(--gold-primary); font-weight: 600; float: right;">${totalProducts}</span>
        </div>
        <div style="margin-bottom: 15px; padding-top: 15px; border-top: 1px solid var(--border-color);">
          <span style="color: var(--text-secondary);">Net Profit:</span>
          <span style="color: ${netProfit >= 0 ? 'var(--success)' : 'var(--danger)'}; font-weight: 700; float: right;">${formatNumber(netProfit)} THB</span>
        </div>
      </div>
    `;
    hideLoading();
  } catch (error) {
    console.error('Error loading reports:', error);
    hideLoading();
  }
}

async function addProduct() {
  const name = document.getElementById('productName').value;
  const category = document.getElementById('productCategory').value;
  const weight = document.getElementById('productWeight').value;
  const purity = document.getElementById('productPurity').value;
  const stock = document.getElementById('productStock').value;

  if (!name || !weight || !stock) {
    alert('Please fill all required fields');
    return;
  }

  try {
    showLoading();
    const result = await callAppsScript('ADD_PRODUCT', {
      name, category, weight, purity, stock
    });

    if (result.success) {
      alert('Product added successfully!');
      closeModal('productModal');
      document.getElementById('productName').value = '';
      document.getElementById('productWeight').value = '';
      document.getElementById('productStock').value = '';
      loadProducts();
    } else {
      alert('Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('Error adding product: ' + error.message);
    hideLoading();
  }
}

async function updatePricing() {
  const buyPrice = document.getElementById('buyPrice').value;
  const sellPrice = document.getElementById('sellPrice').value;
  const exchangeRate = document.getElementById('exchangeRate').value;

  if (!buyPrice || !sellPrice || !exchangeRate) {
    alert('Please fill all fields');
    return;
  }

  try {
    showLoading();
    const result = await callAppsScript('UPDATE_PRICING', {
      buyPrice, sellPrice, exchangeRate
    });

    if (result.success) {
      alert('Pricing updated successfully!');
    } else {
      alert('Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('Error updating pricing: ' + error.message);
    hideLoading();
  }
}

async function addSale() {
  const customerName = document.getElementById('saleCustomer').value;
  const totalAmount = document.getElementById('saleAmount').value;

  if (!customerName || !totalAmount) {
    alert('Please fill all fields');
    return;
  }

  try {
    showLoading();
    const result = await callAppsScript('ADD_SALE', {
      customerName, totalAmount, products: '[]'
    });

    if (result.success) {
      alert('Sale created successfully!');
      closeModal('saleModal');
      document.getElementById('saleCustomer').value = '';
      document.getElementById('saleAmount').value = '';
      loadSales();
      loadDashboard();
    } else {
      alert('Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('Error creating sale: ' + error.message);
    hideLoading();
  }
}

async function addPurchase() {
  const supplierName = document.getElementById('purchaseSupplier').value;
  const totalAmount = document.getElementById('purchaseAmount').value;

  if (!supplierName || !totalAmount) {
    alert('Please fill all fields');
    return;
  }

  try {
    showLoading();
    const result = await callAppsScript('ADD_PURCHASE', {
      supplierName, totalAmount, products: '[]'
    });

    if (result.success) {
      alert('Purchase created successfully!');
      closeModal('purchaseModal');
      document.getElementById('purchaseSupplier').value = '';
      document.getElementById('purchaseAmount').value = '';
      loadPurchases();
      loadDashboard();
    } else {
      alert('Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('Error creating purchase: ' + error.message);
    hideLoading();
  }
}

async function addInventoryTransaction() {
  const productId = document.getElementById('invProductId').value;
  const type = document.getElementById('invType').value;
  const quantity = document.getElementById('invQuantity').value;
  const note = document.getElementById('invNote').value;

  if (!productId || !quantity) {
    alert('Please fill all required fields');
    return;
  }

  try {
    showLoading();
    const result = await callAppsScript('ADD_INVENTORY', {
      productId, type, quantity, note
    });

    if (result.success) {
      alert('Transaction created successfully!');
      closeModal('inventoryModal');
      document.getElementById('invProductId').value = '';
      document.getElementById('invQuantity').value = '';
      document.getElementById('invNote').value = '';
      loadInventory();
    } else {
      alert('Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('Error creating transaction: ' + error.message);
    hideLoading();
  }
}

async function addAudit() {
  const type = document.getElementById('auditType').value;
  const expectedCash = document.getElementById('expectedCash').value;
  const actualCash = document.getElementById('actualCash').value;
  const expectedGold = document.getElementById('expectedGold').value;
  const actualGold = document.getElementById('actualGold').value;
  const note = document.getElementById('auditNote').value;

  if (!expectedCash || !actualCash || !expectedGold || !actualGold) {
    alert('Please fill all required fields');
    return;
  }

  try {
    showLoading();
    const result = await callAppsScript('ADD_AUDIT', {
      type, expectedCash, actualCash, expectedGold, actualGold, note
    });

    if (result.success) {
      alert('Audit record created successfully!');
      closeModal('auditModal');
      document.getElementById('expectedCash').value = '';
      document.getElementById('actualCash').value = '';
      document.getElementById('expectedGold').value = '';
      document.getElementById('actualGold').value = '';
      document.getElementById('auditNote').value = '';
      loadAudits();
    } else {
      alert('Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('Error creating audit: ' + error.message);
    hideLoading();
  }
}

function formatNumber(num) {
  return new Intl.NumberFormat('th-TH', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  }).format(num);
}

function formatDate(date) {
  if (!date) return '-';
  return new Date(date).toLocaleString('th-TH');
}

window.addEventListener('load', loadDashboard);

document.querySelectorAll('.modal').forEach(modal => {
  modal.addEventListener('click', function(e) {
    if (e.target === modal) {
      modal.classList.remove('active');
    }
  });
});
