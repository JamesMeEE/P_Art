const CONFIG = {
  SHEET_ID: '1FF4odviKZ2LnRvPf8ltM0o_jxM0ZHuJHBlkQCjC3sxA',
  API_KEY: 'AIzaSyAB6yjxTB0TNbEk2C68aOP5u0IkdmK12tg',
  APPS_SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbziDXIkJa_VIXVJpRnwv5aYDq425OU5O1vkDvMXEDmzj5KAzg80PJQFtN5DKOmlv0qp/exec'
};

const FIXED_PRODUCTS = [
  { id: 'G01', name: '10 บาท', unit: '10B' },
  { id: 'G02', name: '5 บาท', unit: '5B' },
  { id: 'G03', name: '2 บาท', unit: '2B' },
  { id: 'G04', name: '1 บาท', unit: '1B' },
  { id: 'G05', name: '2 สลึง', unit: '0.5B' },
  { id: 'G06', name: '1 สลึง', unit: '0.25B' },
  { id: 'G07', name: '1 กรัม', unit: '1g' }
];

const EXCHANGE_RATES = {
  'G01': 169000, 'G02': 169000, 'G03': 169000, 'G04': 169000,
  'G05': 99000, 'G06': 99000, 'G07': 99000
};

const USERS = {
  manager: { password: 'manager123', role: 'Manager' },
  teller: { password: 'teller123', role: 'Teller' },
  accountant: { password: 'accountant123', role: 'Accountant' }
};

let currentUser = null;
let currentApprovalItem = null;
let currentPricing = { sell1Baht: 0 };

function roundTo1000(value) {
  return Math.round(value / 1000) * 1000;
}

function calculateSellPrice(productId, sell1Baht) {
  let price = 0;
  switch(productId) {
    case 'G01': price = sell1Baht * 10; break;
    case 'G02': price = sell1Baht * 5; break;
    case 'G03': price = sell1Baht * 2; break;
    case 'G04': price = sell1Baht; break;
    case 'G05': price = (sell1Baht / 2) + 120000; break;
    case 'G06': price = (sell1Baht / 4) + 120000; break;
    case 'G07': price = (sell1Baht / 15) + 120000; break;
  }
  return roundTo1000(price);
}

function calculateBuybackPrice(productId, sell1Baht) {
  const buyback1B = sell1Baht - 530000;
  const sell1g = calculateSellPrice('G07', sell1Baht);
  const buyback1g = sell1g - 155000;
  
  let price = 0;
  switch(productId) {
    case 'G01': price = buyback1B * 10; break;
    case 'G02': price = buyback1B * 5; break;
    case 'G03': price = buyback1B * 2; break;
    case 'G04': price = buyback1B; break;
    case 'G05': price = buyback1B / 2; break;
    case 'G06': price = buyback1B / 4; break;
    case 'G07': price = buyback1g; break;
  }
  return roundTo1000(price);
}

function login() {
  const username = document.getElementById('loginUsername').value;
  const password = document.getElementById('loginPassword').value;

  if (USERS[username] && USERS[username].password === password) {
    currentUser = { username, ...USERS[username] };
    document.getElementById('loginScreen').classList.remove('active');
    document.getElementById('mainHeader').style.display = 'block';
    document.getElementById('mainContainer').style.display = 'block';
    document.getElementById('userName').textContent = currentUser.role;
    document.getElementById('userRole').textContent = currentUser.role;
    document.getElementById('userAvatar').textContent = currentUser.role[0];
    
    document.body.className = 'role-' + username;
    
    if (currentUser.role === 'Accountant') {
      document.body.classList.add('accountant-readonly');
      document.getElementById('quickSales').style.display = 'none';
      document.getElementById('quickBuyback').style.display = 'none';
      document.getElementById('addSaleBtn').style.display = 'none';
      document.getElementById('addBuybackBtn').style.display = 'none';
      document.getElementById('withdrawBtn').style.display = 'none';
    }
    
    loadDashboard();
  } else {
    alert('Invalid username or password');
  }
}

function logout() {
  currentUser = null;
  document.getElementById('loginScreen').classList.add('active');
  document.getElementById('mainHeader').style.display = 'none';
  document.getElementById('mainContainer').style.display = 'none';
  document.getElementById('loginUsername').value = '';
  document.getElementById('loginPassword').value = '';
  document.body.className = '';
}

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
  else if (sectionId === 'sales') loadSales();
  else if (sectionId === 'buyback') loadBuyback();
  else if (sectionId === 'inventory') loadInventory();
  else if (sectionId === 'reports') loadReports();
}

function openModal(modalId) {
  if (modalId === 'saleModal' || modalId === 'buybackModal' || modalId === 'withdrawModal' || 
      modalId === 'stockInModal' || modalId === 'stockOutModal') {
    loadProductOptions();
  }
  if (modalId === 'pricingModal') {
    document.getElementById('sell1BahtPrice').value = currentPricing.sell1Baht || '';
  }
  document.getElementById(modalId).classList.add('active');
}

function closeModal(modalId) {
  document.getElementById(modalId).classList.remove('active');
}

async function fetchSheetData(range) {
  try {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.SHEET_ID}/values/${range}?key=${CONFIG.API_KEY}`;
    const response = await fetch(url);
    if (!response.ok) throw new Error('Failed to fetch data');
    const data = await response.json();
    return data.values || [];
  } catch (error) {
    console.error('fetchSheetData error:', error);
    throw error;
  }
}

async function callAppsScript(action, params) {
  try {
    params.user = currentUser.username;
    params.role = currentUser.role;
    const url = `${CONFIG.APPS_SCRIPT_URL}?action=${action}&${new URLSearchParams(params).toString()}`;
    const response = await fetch(url);
    if (!response.ok) throw new Error('Failed to call API');
    const result = await response.json();
    return result;
  } catch (error) {
    console.error('callAppsScript error:', error);
    throw error;
  }
}

async function loadDashboard() {
  try {
    showLoading();
    const [salesData, buybackData] = await Promise.all([
      fetchSheetData('Sales!A:H'),
      fetchSheetData('Buyback!A:H')
    ]);

    let gp = 0;
    let fxGainLoss = 0;

    salesData.slice(1).forEach(row => {
      if (row[6] === 'COMPLETED') gp += parseFloat(row[4]) || 0;
    });

    buybackData.slice(1).forEach(row => {
      if (row[6] === 'COMPLETED') gp -= parseFloat(row[4]) || 0;
    });

    const netProfit = gp + fxGainLoss;

    document.getElementById('grossProfit').textContent = formatNumber(gp);
    document.getElementById('fxGainLoss').textContent = formatNumber(fxGainLoss);
    document.getElementById('netProfit').textContent = formatNumber(netProfit);

    document.getElementById('summaryContent').innerHTML = `
      <div style="margin: 20px 0;">
        <div style="margin-bottom: 15px;">
          <span style="color: var(--text-secondary);">Total Sales:</span>
          <span style="color: var(--gold-primary); font-weight: 600; float: right;">${salesData.length - 1}</span>
        </div>
        <div style="margin-bottom: 15px;">
          <span style="color: var(--text-secondary);">Total Buyback:</span>
          <span style="color: var(--gold-primary); font-weight: 600; float: right;">${buybackData.length - 1}</span>
        </div>
        <div style="margin-bottom: 15px;">
          <span style="color: var(--text-secondary);">Total SKUs:</span>
          <span style="color: var(--gold-primary); font-weight: 600; float: right;">7</span>
        </div>
      </div>
    `;

    hideLoading();
  } catch (error) {
    console.error('Error loading dashboard:', error);
    alert('Error loading dashboard: ' + error.message);
    hideLoading();
  }
}

async function loadProducts() {
  try {
    showLoading();
    const pricingData = await fetchSheetData('Pricing!A:B');
    const stockData = await fetchSheetData('Stock!A:B');
    
    if (pricingData.length > 1) {
      currentPricing.sell1Baht = parseFloat(pricingData[1][0]) || 0;
    }
    
    const stockMap = {};
    stockData.slice(1).forEach(row => {
      stockMap[row[0]] = parseInt(row[1]) || 0;
    });
    
    const tbody = document.getElementById('productsTable');
    tbody.innerHTML = FIXED_PRODUCTS.map(product => {
      const sellPrice = calculateSellPrice(product.id, currentPricing.sell1Baht);
      const buybackPrice = calculateBuybackPrice(product.id, currentPricing.sell1Baht);
      const exchangeRate = EXCHANGE_RATES[product.id];
      const stock = stockMap[product.id] || 0;
      
      return `
        <tr>
          <td>${product.id}</td>
          <td>${product.name}</td>
          <td>${product.unit}</td>
          <td>${formatNumber(sellPrice)}</td>
          <td>${formatNumber(buybackPrice)}</td>
          <td>${formatNumber(exchangeRate)}</td>
          <td>${stock}</td>
        </tr>
      `;
    }).join('');
    
    hideLoading();
  } catch (error) {
    console.error('Error loading products:', error);
    hideLoading();
  }
}

async function loadProductOptions() {
  const options = FIXED_PRODUCTS.map(p => 
    `<option value="${p.id}">${p.name} (${p.unit})</option>`
  ).join('');
  
  document.getElementById('saleProduct').innerHTML = '<option value="">Select product...</option>' + options;
  document.getElementById('buybackProduct').innerHTML = '<option value="">Select product...</option>' + options;
  document.getElementById('withdrawProduct').innerHTML = '<option value="">Select product...</option>' + options;
  document.getElementById('stockInProduct').innerHTML = '<option value="">Select product...</option>' + options;
  document.getElementById('stockOutProduct').innerHTML = '<option value="">Select product...</option>' + options;
}

async function updatePricing() {
  const sell1Baht = parseFloat(document.getElementById('sell1BahtPrice').value);
  
  if (!sell1Baht || sell1Baht <= 0) {
    alert('Please enter a valid price');
    return;
  }

  try {
    showLoading();
    const result = await callAppsScript('UPDATE_PRICING', { sell1Baht });

    if (result.success) {
      alert('✅ Pricing updated successfully!');
      closeModal('pricingModal');
      currentPricing.sell1Baht = sell1Baht;
      loadProducts();
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}

async function loadSales() {
  try {
    showLoading();
    const data = await fetchSheetData('Sales!A:H');
    const tbody = document.getElementById('salesTable');
    
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 40px;">No sales records</td></tr>';
    } else {
      tbody.innerHTML = data.slice(1).map(row => {
        let actions = '';
        if (row[6] === 'PENDING' && currentUser.role === 'Manager') {
          actions = `<button class="btn-action" onclick="reviewSale('${row[0]}')">ตรวจสอบ</button>`;
        } else if (row[6] === 'READY' && (currentUser.role === 'Manager' || currentUser.role === 'Teller')) {
          actions = `<button class="btn-action" onclick="confirmSale('${row[0]}')">ยืนยัน</button>`;
        }
        
        return `
          <tr>
            <td>${row[0]}</td>
            <td>${row[1]}</td>
            <td>${row[2]}</td>
            <td>${row[3]}</td>
            <td>${formatNumber(row[4])}</td>
            <td>${formatDate(row[5])}</td>
            <td><span class="status-badge status-${row[6].toLowerCase()}">${row[6]}</span></td>
            <td>${actions}</td>
          </tr>
        `;
      }).join('');
    }
    hideLoading();
  } catch (error) {
    console.error('Error loading sales:', error);
    hideLoading();
  }
}

async function addSale() {
  const customer = document.getElementById('saleCustomer').value;
  const product = document.getElementById('saleProduct').value;
  const amount = document.getElementById('saleAmount').value;
  const price = document.getElementById('salePrice').value;

  if (!customer || !product || !amount || !price) {
    alert('Please fill all fields');
    return;
  }

  try {
    showLoading();
    const result = await callAppsScript('ADD_SALE', {
      customer, product, amount, price
    });

    if (result.success) {
      alert('✅ Sale created successfully!');
      closeModal('saleModal');
      document.getElementById('saleCustomer').value = '';
      document.getElementById('saleAmount').value = '';
      document.getElementById('salePrice').value = '';
      loadSales();
      loadDashboard();
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}

function reviewSale(id) {
  currentApprovalItem = { id, type: 'SALE' };
  document.getElementById('approvalDetails').innerHTML = `<p>รหัสรายการ: ${id}</p><p>ประเภท: Sales</p>`;
  openModal('approvalModal');
}

async function approveItem() {
  if (!currentApprovalItem) return;
  
  try {
    showLoading();
    const result = await callAppsScript('APPROVE_ITEM', {
      id: currentApprovalItem.id,
      type: currentApprovalItem.type
    });

    if (result.success) {
      alert('✅ Approved successfully!');
      closeModal('approvalModal');
      if (currentApprovalItem.type === 'SALE') loadSales();
      else if (currentApprovalItem.type === 'BUYBACK') loadBuyback();
      else if (currentApprovalItem.type === 'WITHDRAW') loadInventory();
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}

async function rejectItem() {
  const reason = document.getElementById('approvalReason').value;
  if (!reason) {
    alert('Please enter reason');
    return;
  }
  
  try {
    showLoading();
    const result = await callAppsScript('REJECT_ITEM', {
      id: currentApprovalItem.id,
      type: currentApprovalItem.type,
      reason
    });

    if (result.success) {
      alert('✅ Rejected successfully!');
      closeModal('approvalModal');
      if (currentApprovalItem.type === 'SALE') loadSales();
      else if (currentApprovalItem.type === 'BUYBACK') loadBuyback();
      else if (currentApprovalItem.type === 'WITHDRAW') loadInventory();
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}

async function confirmSale(id) {
  if (!confirm('Confirm this sale?')) return;
  
  try {
    showLoading();
    const result = await callAppsScript('CONFIRM_SALE', { id });

    if (result.success) {
      alert('✅ Sale confirmed!');
      loadSales();
      loadDashboard();
      loadProducts();
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}

async function loadBuyback() {
  try {
    showLoading();
    const data = await fetchSheetData('Buyback!A:H');
    const tbody = document.getElementById('buybackTable');
    
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 40px;">No buyback records</td></tr>';
    } else {
      tbody.innerHTML = data.slice(1).map(row => {
        let actions = '';
        if (row[6] === 'PENDING' && currentUser.role === 'Manager') {
          actions = `<button class="btn-action" onclick="reviewBuyback('${row[0]}')">ตรวจสอบ</button>`;
        }
        
        return `
          <tr>
            <td>${row[0]}</td>
            <td>${row[1]}</td>
            <td>${row[2]}</td>
            <td>${row[3]}</td>
            <td>${formatNumber(row[4])}</td>
            <td>${formatDate(row[5])}</td>
            <td><span class="status-badge status-${row[6].toLowerCase()}">${row[6]}</span></td>
            <td>${actions}</td>
          </tr>
        `;
      }).join('');
    }
    hideLoading();
  } catch (error) {
    console.error('Error loading buyback:', error);
    hideLoading();
  }
}

async function addBuyback() {
  const customer = document.getElementById('buybackCustomer').value;
  const product = document.getElementById('buybackProduct').value;
  const amount = document.getElementById('buybackAmount').value;
  const price = document.getElementById('buybackPrice').value;

  if (!customer || !product || !amount || !price) {
    alert('Please fill all fields');
    return;
  }

  try {
    showLoading();
    const result = await callAppsScript('ADD_BUYBACK', {
      customer, product, amount, price
    });

    if (result.success) {
      alert('✅ Buyback created successfully!');
      closeModal('buybackModal');
      document.getElementById('buybackCustomer').value = '';
      document.getElementById('buybackAmount').value = '';
      document.getElementById('buybackPrice').value = '';
      loadBuyback();
      loadDashboard();
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}

function reviewBuyback(id) {
  currentApprovalItem = { id, type: 'BUYBACK' };
  document.getElementById('approvalDetails').innerHTML = `<p>รหัสรายการ: ${id}</p><p>ประเภท: Buyback</p>`;
  openModal('approvalModal');
}

async function loadInventory() {
  try {
    showLoading();
    const data = await fetchSheetData('Inventory!A:H');
    const tbody = document.getElementById('inventoryTable');
    
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 40px;">No inventory records</td></tr>';
    } else {
      tbody.innerHTML = data.slice(1).map(row => {
        let actions = '';
        if (row[6] === 'PENDING' && row[2] === 'WITHDRAW' && currentUser.role === 'Manager') {
          actions = `<button class="btn-action" onclick="reviewWithdraw('${row[0]}')">ตรวจสอบ</button>`;
        }
        
        return `
          <tr>
            <td>${row[0]}</td>
            <td>${row[1]}</td>
            <td>${row[2]}</td>
            <td>${row[3]}</td>
            <td>${row[4] ? formatNumber(row[4]) : '-'}</td>
            <td>${formatDate(row[5])}</td>
            <td><span class="status-badge status-${row[6].toLowerCase()}">${row[6]}</span></td>
            <td>${actions}</td>
          </tr>
        `;
      }).join('');
    }
    hideLoading();
  } catch (error) {
    console.error('Error loading inventory:', error);
    hideLoading();
  }
}

async function addWithdraw() {
  const product = document.getElementById('withdrawProduct').value;
  const quantity = document.getElementById('withdrawQuantity').value;

  if (!product || !quantity) {
    alert('Please fill all fields');
    return;
  }

  try {
    showLoading();
    const result = await callAppsScript('ADD_WITHDRAW', {
      product, quantity
    });

    if (result.success) {
      alert('✅ Withdraw request created!');
      closeModal('withdrawModal');
      document.getElementById('withdrawQuantity').value = '';
      loadInventory();
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}

function reviewWithdraw(id) {
  currentApprovalItem = { id, type: 'WITHDRAW' };
  document.getElementById('approvalDetails').innerHTML = `<p>รหัสรายการ: ${id}</p><p>ประเภท: Withdraw</p>`;
  openModal('approvalModal');
}

async function addStockIn() {
  const product = document.getElementById('stockInProduct').value;
  const quantity = document.getElementById('stockInQuantity').value;
  const cost = document.getElementById('stockInCost').value;

  if (!product || !quantity || !cost) {
    alert('Please fill all fields');
    return;
  }

  try {
    showLoading();
    const result = await callAppsScript('ADD_STOCK_IN', {
      product, quantity, cost
    });

    if (result.success) {
      alert('✅ Stock added successfully!');
      closeModal('stockInModal');
      document.getElementById('stockInQuantity').value = '';
      document.getElementById('stockInCost').value = '';
      loadInventory();
      loadProducts();
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}

async function addStockOut() {
  const product = document.getElementById('stockOutProduct').value;
  const quantity = document.getElementById('stockOutQuantity').value;

  if (!product || !quantity) {
    alert('Please fill all fields');
    return;
  }

  try {
    showLoading();
    const result = await callAppsScript('ADD_STOCK_OUT', {
      product, quantity
    });

    if (result.success) {
      alert('✅ Stock out recorded!');
      closeModal('stockOutModal');
      document.getElementById('stockOutQuantity').value = '';
      loadInventory();
      loadProducts();
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}

async function loadReports() {
  try {
    showLoading();
    const [salesData, buybackData] = await Promise.all([
      fetchSheetData('Sales!A:H'),
      fetchSheetData('Buyback!A:H')
    ]);

    let gp = 0;
    salesData.slice(1).forEach(row => {
      if (row[6] === 'COMPLETED') gp += parseFloat(row[4]) || 0;
    });
    buybackData.slice(1).forEach(row => {
      if (row[6] === 'COMPLETED') gp -= parseFloat(row[4]) || 0;
    });

    document.getElementById('reportContent').innerHTML = `
      <div style="margin: 20px 0;">
        <div style="margin-bottom: 15px;">
          <span style="color: var(--text-secondary);">Gross Profit:</span>
          <span style="color: var(--gold-primary); font-weight: 600; float: right;">${formatNumber(gp)} LAK</span>
        </div>
        <div style="margin-bottom: 15px;">
          <span style="color: var(--text-secondary);">FX Gain/Loss:</span>
          <span style="color: var(--gold-primary); font-weight: 600; float: right;">0 LAK</span>
        </div>
        <div style="margin-bottom: 15px; padding-top: 15px; border-top: 1px solid var(--border-color);">
          <span style="color: var(--text-secondary);">Net Profit:</span>
          <span style="color: ${gp >= 0 ? 'var(--success)' : 'var(--danger)'}; font-weight: 700; float: right;">${formatNumber(gp)} LAK</span>
        </div>
      </div>
    `;
    hideLoading();
  } catch (error) {
    console.error('Error loading reports:', error);
    hideLoading();
  }
}

function formatNumber(num) {
  return new Intl.NumberFormat('en-US', {
    minimumFractionDigits: 0,
    maximumFractionDigits: 0
  }).format(num);
}

function formatDate(date) {
  if (!date) return '-';
  return new Date(date).toLocaleString('en-US');
}

document.querySelectorAll('.modal').forEach(modal => {
  modal.addEventListener('click', function(e) {
    if (e.target === modal) {
      modal.classList.remove('active');
    }
  });
});
