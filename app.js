const CONFIG = {
  SHEET_ID: '1FF4odviKZ2LnRvPf8ltM0o_jxM0ZHuJHBlkQCjC3sxA',
  API_KEY: 'AIzaSyAB6yjxTB0TNbEk2C68aOP5u0IkdmK12tg',
  APPS_SCRIPT_URL: 'https://script.google.com/macros/s/1axvE4dvC-l4wfS8I69kfxg8nGVD_qzHx5hVuD9CmNZ6jsEBIopC58Ln9/exec'
};

const USERS = {
  manager: { password: 'manager123', role: 'Manager' },
  teller: { password: 'teller123', role: 'Teller' },
  accountant: { password: 'accountant123', role: 'Accountant' }
};

let currentUser = null;
let currentApprovalItem = null;

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
      document.getElementById('addProductBtn').style.display = 'none';
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
  document.getElementById(modalId).classList.add('active');
}

function closeModal(modalId) {
  document.getElementById(modalId).classList.remove('active');
}

async function fetchSheetData(range) {
  try {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.SHEET_ID}/values/${range}?key=${CONFIG.API_KEY}`;
    console.log('Fetching:', url);
    const response = await fetch(url);
    if (!response.ok) {
      const errorText = await response.text();
      console.error('Fetch error:', errorText);
      throw new Error('Failed to fetch data: ' + errorText);
    }
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
    console.log('Calling Apps Script:', url);
    const response = await fetch(url);
    if (!response.ok) {
      const errorText = await response.text();
      console.error('Apps Script error:', errorText);
      throw new Error('Failed to call API: ' + errorText);
    }
    const result = await response.json();
    console.log('Apps Script response:', result);
    return result;
  } catch (error) {
    console.error('callAppsScript error:', error);
    throw error;
  }
}

async function loadDashboard() {
  try {
    showLoading();
    const [salesData, buybackData, productsData] = await Promise.all([
      fetchSheetData('Sales!A:H'),
      fetchSheetData('Buyback!A:H'),
      fetchSheetData('Products!A:I')
    ]);

    let gp = 0;
    let fxGainLoss = 0;

    salesData.slice(1).forEach(row => {
      if (row[6] === 'COMPLETED') {
        const price = parseFloat(row[4]) || 0;
        gp += price;
      }
    });

    buybackData.slice(1).forEach(row => {
      if (row[6] === 'COMPLETED') {
        const price = parseFloat(row[4]) || 0;
        gp -= price;
      }
    });

    const netProfit = gp + fxGainLoss;

    document.getElementById('grossProfit').textContent = formatNumber(gp);
    document.getElementById('fxGainLoss').textContent = formatNumber(fxGainLoss);
    document.getElementById('netProfit').textContent = formatNumber(netProfit);
    document.getElementById('totalProducts').textContent = productsData.length - 1;

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
          <span style="color: var(--text-secondary);">Total Products:</span>
          <span style="color: var(--gold-primary); font-weight: 600; float: right;">${productsData.length - 1}</span>
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
    const data = await fetchSheetData('Products!A:I');
    const tbody = document.getElementById('productsTable');
    
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="9" style="text-align: center; padding: 40px;">No products available</td></tr>';
    } else {
      tbody.innerHTML = data.slice(1).map(row => `
        <tr>
          <td>${row[0]}</td>
          <td>${row[1]}</td>
          <td>${row[2]}</td>
          <td>${row[3] || '-'}</td>
          <td>${formatNumber(row[4])}</td>
          <td>${formatNumber(row[5])}</td>
          <td>${formatNumber(row[6])}</td>
          <td>${row[7]}</td>
          <td class="manager-only">
            <button class="btn-action" onclick="editProduct('${row[0]}')">Edit</button>
          </td>
        </tr>
      `).join('');
    }
    hideLoading();
  } catch (error) {
    console.error('Error loading products:', error);
    alert('Error loading products: ' + error.message);
    hideLoading();
  }
}

async function loadProductOptions() {
  try {
    const data = await fetchSheetData('Products!A:I');
    const options = data.slice(1).map(row => 
      `<option value="${row[0]}">${row[1]} (${row[2]}g)</option>`
    ).join('');
    
    document.getElementById('saleProduct').innerHTML = '<option value="">Select product...</option>' + options;
    document.getElementById('buybackProduct').innerHTML = '<option value="">Select product...</option>' + options;
    document.getElementById('withdrawProduct').innerHTML = '<option value="">Select product...</option>' + options;
    document.getElementById('stockInProduct').innerHTML = '<option value="">Select product...</option>' + options;
    document.getElementById('stockOutProduct').innerHTML = '<option value="">Select product...</option>' + options;
  } catch (error) {
    console.error('Error loading product options:', error);
  }
}

async function addProduct() {
  const name = document.getElementById('productName').value;
  const weight = document.getElementById('productWeight').value;
  const requirement = document.getElementById('productRequirement').value;
  const buyPrice = document.getElementById('productBuyPrice').value;
  const sellPrice = document.getElementById('productSellPrice').value;
  const exchangeRate = document.getElementById('productExchangeRate').value;
  const stock = document.getElementById('productStock').value;

  if (!name || !weight) {
    alert('Please fill required fields (Product Name and Weight)');
    return;
  }

  try {
    showLoading();
    const result = await callAppsScript('ADD_PRODUCT', {
      name, weight, requirement, buyPrice, sellPrice, exchangeRate, stock
    });

    if (result.success) {
      alert('✅ Product added successfully!');
      closeModal('productModal');
      document.getElementById('productName').value = '';
      document.getElementById('productWeight').value = '';
      document.getElementById('productRequirement').value = '';
      document.getElementById('productBuyPrice').value = '';
      document.getElementById('productSellPrice').value = '';
      document.getElementById('productExchangeRate').value = '';
      document.getElementById('productStock').value = '';
      loadProducts();
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    console.error('Add product error:', error);
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
    alert('Error loading sales: ' + error.message);
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
      alert('✅ Sale confirmed! Generating receipt...');
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
    alert('Error loading buyback: ' + error.message);
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
    alert('Error loading inventory: ' + error.message);
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
    const [salesData, buybackData, productsData] = await Promise.all([
      fetchSheetData('Sales!A:H'),
      fetchSheetData('Buyback!A:H'),
      fetchSheetData('Products!A:I')
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
          <span style="color: var(--gold-primary); font-weight: 600; float: right;">0.00 LAK</span>
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
    alert('Error loading reports: ' + error.message);
    hideLoading();
  }
}

function formatNumber(num) {
  return new Intl.NumberFormat('en-US', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
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
