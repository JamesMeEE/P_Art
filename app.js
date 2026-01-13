const CONFIG = {
  SHEET_ID: '1FF4odviKZ2LnRvPf8ltM0o_jxM0ZHuJHBlkQCjC3sxA',
  API_KEY: 'AIzaSyAB6yjxTB0TNbEk2C68aOP5u0IkdmK12tg',
  APPS_SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbziDXIkJa_VIXVJpRnwv5aYDq425OU5O1vkDvMXEDmzj5KAzg80PJQFtN5DKOmlv0qp/exec',
  EXCHANGE_API: 'https://api.exchangerate-api.com/v4/latest/LAK'
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

const EXCHANGE_FEES = {
  'G01': 169000, 'G02': 169000, 'G03': 169000, 'G04': 169000,
  'G05': 99000, 'G06': 99000, 'G07': 99000
};

const PREMIUM_PRODUCTS = ['G05', 'G06', 'G07'];
const PREMIUM_PER_PIECE = 120000;

const USERS = {
  manager: { password: 'manager123', role: 'Manager' },
  teller: { password: 'teller123', role: 'Teller' },
  accountant: { password: 'accountant123', role: 'Accountant' }
};

let currentUser = null;
let currentApprovalItem = null;
let currentPricing = { sell1Baht: 0 };
let currentExchangeRates = { THB: 0, USD: 0 };
let sellProductCounter = 0;
let buybackProductCounter = 0;
let tradeinOldCounter = 0;
let tradeinNewCounter = 0;
let exchangeOldCounter = 0;
let exchangeNewCounter = 0;
let currentTradeinData = null;
let currentExchangeData = null;

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

function formatItemsForDisplay(items) {
  const parsed = JSON.parse(items);
  return parsed.map(item => {
    const product = FIXED_PRODUCTS.find(p => p.id === item.productId);
    return `${product.name}: ${item.qty} unit`;
  }).join('\n');
}

function formatItemsForTable(items) {
  const lines = formatItemsForDisplay(items).split('\n');
  return lines.join('<br>');
}

async function fetchExchangeRates() {
  try {
    const response = await fetch(CONFIG.EXCHANGE_API);
    const data = await response.json();
    currentExchangeRates.THB = data.rates.THB;
    currentExchangeRates.USD = data.rates.USD;
  } catch (error) {
    console.error('Failed to fetch exchange rates:', error);
  }
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
      const readonlyBtns = ['quickSales', 'quickTradein', 'quickBuyback', 'addSellBtn', 'addTradeinBtn', 'addBuybackBtn', 'addExchangeBtn', 'withdrawBtn'];
      readonlyBtns.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.style.display = 'none';
      });
    }
    
    fetchExchangeRates();
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
  else if (sectionId === 'sell') loadSells();
  else if (sectionId === 'tradein') loadTradeins();
  else if (sectionId === 'buyback') loadBuybacks();
  else if (sectionId === 'exchange') loadExchanges();
  else if (sectionId === 'inventory') loadInventory();
  else if (sectionId === 'reports') loadReports();
}

function openModal(modalId) {
  if (modalId === 'sellModal') {
    sellProductCounter = 0;
    document.getElementById('sellProducts').innerHTML = '';
    document.getElementById('sellCustomer').value = '';
    document.getElementById('sellPrice').value = '';
    addSellProduct();
  }
  if (modalId === 'buybackModal') {
    buybackProductCounter = 0;
    document.getElementById('buybackProducts').innerHTML = '';
    document.getElementById('buybackCustomer').value = '';
    document.getElementById('buybackPrice').value = '';
    addBuybackProduct();
  }
  if (modalId === 'tradeinModal') {
    tradeinOldCounter = 0;
    tradeinNewCounter = 0;
    document.getElementById('tradeinOldGold').innerHTML = '';
    document.getElementById('tradeinNewGold').innerHTML = '';
    document.getElementById('tradeinCustomer').value = '';
    addTradeinOldGold();
    addTradeinNewGold();
  }
  if (modalId === 'exchangeModal') {
    exchangeOldCounter = 0;
    exchangeNewCounter = 0;
    document.getElementById('exchangeOldGold').innerHTML = '';
    document.getElementById('exchangeNewGold').innerHTML = '';
    document.getElementById('exchangeCustomer').value = '';
    addExchangeOldGold();
    addExchangeNewGold();
  }
  if (modalId === 'withdrawModal' || modalId === 'stockInModal' || modalId === 'stockOutModal') {
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

function backToTradein() {
  closeModal('tradeinResultModal');
  openModal('tradeinModal');
}

function backToExchange() {
  closeModal('exchangeResultModal');
  openModal('exchangeModal');
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
    const response = await fetch(url, { method: 'POST' });
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
    const [sellData, tradeinData, buybackData, exchangeData] = await Promise.all([
      fetchSheetData('Sells!A:I'),
      fetchSheetData('Tradeins!A:K'),
      fetchSheetData('Buybacks!A:F'),
      fetchSheetData('Exchanges!A:J')
    ]);

    let gp = 0;
    let fxGainLoss = 0;

    sellData.slice(1).forEach(row => {
      if (row[7] === 'COMPLETED') gp += parseFloat(row[3]) || 0;
    });

    buybackData.slice(1).forEach(row => {
      if (row[5] === 'COMPLETED') gp -= parseFloat(row[3]) || 0;
    });

    const netProfit = gp + fxGainLoss;

    document.getElementById('grossProfit').textContent = formatNumber(gp);
    document.getElementById('fxGainLoss').textContent = formatNumber(fxGainLoss);
    document.getElementById('netProfit').textContent = formatNumber(netProfit);

    document.getElementById('summaryContent').innerHTML = `
      <div style="margin: 20px 0;">
        <div style="margin-bottom: 15px;">
          <span style="color: var(--text-secondary);">Total Sells:</span>
          <span style="color: var(--gold-primary); font-weight: 600; float: right;">${sellData.length - 1}</span>
        </div>
        <div style="margin-bottom: 15px;">
          <span style="color: var(--text-secondary);">Total Trade-ins:</span>
          <span style="color: var(--gold-primary); font-weight: 600; float: right;">${tradeinData.length - 1}</span>
        </div>
        <div style="margin-bottom: 15px;">
          <span style="color: var(--text-secondary);">Total Buybacks:</span>
          <span style="color: var(--gold-primary); font-weight: 600; float: right;">${buybackData.length - 1}</span>
        </div>
        <div style="margin-bottom: 15px;">
          <span style="color: var(--text-secondary);">Total Exchanges:</span>
          <span style="color: var(--gold-primary); font-weight: 600; float: right;">${exchangeData.length - 1}</span>
        </div>
      </div>
    `;

    hideLoading();
  } catch (error) {
    console.error('Error loading dashboard:', error);
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
    
    document.getElementById('currentPriceDisplay').textContent = formatNumber(currentPricing.sell1Baht) + ' LAK';
    
    const stockMap = {};
    stockData.slice(1).forEach(row => {
      stockMap[row[0]] = parseInt(row[1]) || 0;
    });
    
    const tbody = document.getElementById('productsTable');
    tbody.innerHTML = FIXED_PRODUCTS.map(product => {
      const sellPrice = calculateSellPrice(product.id, currentPricing.sell1Baht);
      const buybackPrice = calculateBuybackPrice(product.id, currentPricing.sell1Baht);
      const exchangeFee = EXCHANGE_FEES[product.id];
      const stock = stockMap[product.id] || 0;
      
      return `
        <tr>
          <td>${product.id}</td>
          <td>${product.name}</td>
          <td>${product.unit}</td>
          <td>${formatNumber(sellPrice)}</td>
          <td>${formatNumber(buybackPrice)}</td>
          <td>${formatNumber(exchangeFee)}</td>
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
  
  const selects = ['withdrawProduct', 'stockInProduct', 'stockOutProduct'];
  selects.forEach(id => {
    const el = document.getElementById(id);
    if (el) el.innerHTML = '<option value="">Select product...</option>' + options;
  });
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

function addSellProduct() {
  sellProductCounter++;
  const html = `
    <div class="product-row" id="sellProduct${sellProductCounter}">
      <select class="form-select" onchange="calculateSellTotal()">
        <option value="">Select product...</option>
        ${FIXED_PRODUCTS.map(p => `<option value="${p.id}">${p.name}</option>`).join('')}
      </select>
      <input type="number" class="form-input" placeholder="Qty" min="1" value="1" onchange="calculateSellTotal()">
      <button onclick="removeSellProduct(${sellProductCounter})">×</button>
    </div>
  `;
  document.getElementById('sellProducts').insertAdjacentHTML('beforeend', html);
  calculateSellTotal();
}

function removeSellProduct(id) {
  const el = document.getElementById(`sellProduct${id}`);
  if (el) el.remove();
  calculateSellTotal();
}

function calculateSellTotal() {
  let total = 0;
  const products = document.querySelectorAll('#sellProducts .product-row');
  products.forEach(row => {
    const productId = row.querySelector('select').value;
    const qty = parseInt(row.querySelector('input').value) || 0;
    if (productId && qty > 0) {
      const price = calculateSellPrice(productId, currentPricing.sell1Baht);
      total += price * qty;
    }
  });
  document.getElementById('sellPrice').value = total;
}

async function submitSell() {
  const customer = document.getElementById('sellCustomer').value;
  const currency = document.getElementById('sellCurrency').value;
  const total = parseFloat(document.getElementById('sellPrice').value);
  
  const products = [];
  document.querySelectorAll('#sellProducts .product-row').forEach(row => {
    const productId = row.querySelector('select').value;
    const qty = parseInt(row.querySelector('input').value) || 0;
    if (productId && qty > 0) {
      products.push({ productId, qty });
    }
  });

  if (!customer || products.length === 0 || !total) {
    alert('Please fill all fields');
    return;
  }

  try {
    showLoading();
    await fetchExchangeRates();
    const exchangeRate = currency === 'LAK' ? 1 : (currency === 'THB' ? currentExchangeRates.THB : currentExchangeRates.USD);
    
    const result = await callAppsScript('ADD_SELL', {
      customer,
      items: JSON.stringify(products),
      total,
      currency,
      exchangeRate
    });

    if (result.success) {
      alert('✅ Sell created successfully!');
      closeModal('sellModal');
      loadSells();
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

async function loadSells() {
  try {
    showLoading();
    const data = await fetchSheetData('Sells!A:I');
    const tbody = document.getElementById('sellTable');
    
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="9" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = data.slice(1).map(row => {
        let actions = '';
        if (row[7] === 'PENDING' && currentUser.role === 'Manager') {
          actions = `<button class="btn-action" onclick="reviewTransaction('${row[0]}', 'SELL')">ตรวจสอบ</button>`;
        } else if (row[7] === 'READY' && (currentUser.role === 'Manager' || currentUser.role === 'Teller')) {
          actions = `<button class="btn-action" onclick="confirmTransaction('${row[0]}', 'SELL')">ยืนยัน</button>`;
        }
        
        return `
          <tr>
            <td>${row[0]}</td>
            <td>${row[1]}</td>
            <td>${formatItemsForTable(row[2])}</td>
            <td>${formatNumber(row[3])}</td>
            <td>${row[4]}</td>
            <td>${row[5]}</td>
            <td>${formatDateOnly(row[6])}</td>
            <td><span class="status-badge status-${row[7].toLowerCase()}">${row[7]}</span></td>
            <td>${actions}</td>
          </tr>
        `;
      }).join('');
    }
    hideLoading();
  } catch (error) {
    console.error('Error loading sells:', error);
    hideLoading();
  }
}

function addTradeinOldGold() {
  tradeinOldCounter++;
  const html = `
    <div class="product-row" id="tradeinOld${tradeinOldCounter}">
      <select class="form-select">
        <option value="">Select product...</option>
        ${FIXED_PRODUCTS.map(p => `<option value="${p.id}">${p.name}</option>`).join('')}
      </select>
      <input type="number" class="form-input" placeholder="Qty" min="1" value="1">
      <button onclick="removeTradeinOld(${tradeinOldCounter})">×</button>
    </div>
  `;
  document.getElementById('tradeinOldGold').insertAdjacentHTML('beforeend', html);
}

function removeTradeinOld(id) {
  const el = document.getElementById(`tradeinOld${id}`);
  if (el) el.remove();
}

function addTradeinNewGold() {
  tradeinNewCounter++;
  const html = `
    <div class="product-row" id="tradeinNew${tradeinNewCounter}">
      <select class="form-select">
        <option value="">Select product...</option>
        ${FIXED_PRODUCTS.map(p => `<option value="${p.id}">${p.name}</option>`).join('')}
      </select>
      <input type="number" class="form-input" placeholder="Qty" min="1" value="1">
      <button onclick="removeTradeinNew(${tradeinNewCounter})">×</button>
    </div>
  `;
  document.getElementById('tradeinNewGold').insertAdjacentHTML('beforeend', html);
}

function removeTradeinNew(id) {
  const el = document.getElementById(`tradeinNew${id}`);
  if (el) el.remove();
}

function calculateTradein() {
  const customer = document.getElementById('tradeinCustomer').value;
  if (!customer) {
    alert('Please enter customer name');
    return;
  }

  const oldGold = [];
  document.querySelectorAll('#tradeinOldGold .product-row').forEach(row => {
    const productId = row.querySelector('select').value;
    const qty = parseInt(row.querySelector('input').value) || 0;
    if (productId && qty > 0) {
      oldGold.push({ productId, qty });
    }
  });

  const newGold = [];
  document.querySelectorAll('#tradeinNewGold .product-row').forEach(row => {
    const productId = row.querySelector('select').value;
    const qty = parseInt(row.querySelector('input').value) || 0;
    if (productId && qty > 0) {
      newGold.push({ productId, qty });
    }
  });

  if (oldGold.length === 0 || newGold.length === 0) {
    alert('Please add both old and new gold');
    return;
  }

  let oldValue = 0;
  oldGold.forEach(item => {
    oldValue += calculateBuybackPrice(item.productId, currentPricing.sell1Baht) * item.qty;
  });

  let newValue = 0;
  newGold.forEach(item => {
    newValue += calculateSellPrice(item.productId, currentPricing.sell1Baht) * item.qty;
  });

  let exchangeFee = 0;
  newGold.forEach(item => {
    exchangeFee += EXCHANGE_FEES[item.productId] * item.qty;
  });

  let premium = 0;
  newGold.forEach(item => {
    if (PREMIUM_PRODUCTS.includes(item.productId)) {
      premium += PREMIUM_PER_PIECE * item.qty;
    }
  });

  const difference = newValue - oldValue;

  currentTradeinData = {
    customer,
    oldGold: JSON.stringify(oldGold),
    newGold: JSON.stringify(newGold),
    difference,
    exchangeFee,
    premium
  };

  const oldItems = oldGold.map(i => `${FIXED_PRODUCTS.find(p => p.id === i.productId).name} (${i.qty})`).join(', ');
  const newItems = newGold.map(i => `${FIXED_PRODUCTS.find(p => p.id === i.productId).name} (${i.qty})`).join(', ');

  document.getElementById('tradeinResult').innerHTML = `
    <div class="summary-box">
      <div class="summary-row">
        <span class="summary-label">Old Gold:</span>
        <span class="summary-value">${oldItems}</span>
      </div>
      <div class="summary-row">
        <span class="summary-label">Old Value:</span>
        <span class="summary-value">${formatNumber(oldValue)} LAK</span>
      </div>
      <div class="summary-row">
        <span class="summary-label">New Gold:</span>
        <span class="summary-value">${newItems}</span>
      </div>
      <div class="summary-row">
        <span class="summary-label">New Value:</span>
        <span class="summary-value">${formatNumber(newValue)} LAK</span>
      </div>
      <div class="summary-row">
        <span class="summary-label">Difference:</span>
        <span class="summary-value">${formatNumber(difference)} LAK</span>
      </div>
      <div class="summary-row">
        <span class="summary-label">Exchange Fee:</span>
        <span class="summary-value">${formatNumber(exchangeFee)} LAK</span>
      </div>
      <div class="summary-row">
        <span class="summary-label">Premium:</span>
        <span class="summary-value">${formatNumber(premium)} LAK</span>
      </div>
      <div class="summary-row">
        <span class="summary-label">Total to Pay:</span>
        <span class="summary-value">${formatNumber(difference + exchangeFee + premium)} LAK</span>
      </div>
    </div>
  `;

  closeModal('tradeinModal');
  openModal('tradeinResultModal');
}

async function submitTradein() {
  if (!currentTradeinData) return;

  try {
    showLoading();
    const result = await callAppsScript('ADD_TRADEIN', currentTradeinData);

    if (result.success) {
      alert('✅ Trade-in created successfully!');
      closeModal('tradeinResultModal');
      currentTradeinData = null;
      loadTradeins();
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

async function loadTradeins() {
  try {
    showLoading();
    const data = await fetchSheetData('Tradeins!A:K');
    const tbody = document.getElementById('tradeinTable');
    
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="10" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = data.slice(1).map(row => {
        let actions = '';
        if (row[8] === 'PENDING' && currentUser.role === 'Manager') {
          actions = `<button class="btn-action" onclick="reviewTransaction('${row[0]}', 'TRADEIN')">ตรวจสอบ</button>`;
        } else if (row[8] === 'READY' && (currentUser.role === 'Manager' || currentUser.role === 'Teller')) {
          actions = `<button class="btn-action" onclick="confirmTransaction('${row[0]}', 'TRADEIN')">ยืนยัน</button>`;
        }
        
        return `
          <tr>
            <td>${row[0]}</td>
            <td>${row[1]}</td>
            <td>${formatItemsForTable(row[2])}</td>
            <td>${formatItemsForTable(row[3])}</td>
            <td>${formatNumber(row[4])}</td>
            <td>${formatNumber(row[5])}</td>
            <td>${formatNumber(row[6])}</td>
            <td>${formatDateOnly(row[7])}</td>
            <td><span class="status-badge status-${row[8].toLowerCase()}">${row[8]}</span></td>
            <td>${actions}</td>
          </tr>
        `;
      }).join('');
    }
    hideLoading();
  } catch (error) {
    console.error('Error loading tradeins:', error);
    hideLoading();
  }
}

function addExchangeOldGold() {
  exchangeOldCounter++;
  const html = `
    <div class="product-row" id="exchangeOld${exchangeOldCounter}">
      <select class="form-select">
        <option value="">Select product...</option>
        ${FIXED_PRODUCTS.map(p => `<option value="${p.id}">${p.name}</option>`).join('')}
      </select>
      <input type="number" class="form-input" placeholder="Qty" min="1" value="1">
      <button onclick="removeExchangeOld(${exchangeOldCounter})">×</button>
    </div>
  `;
  document.getElementById('exchangeOldGold').insertAdjacentHTML('beforeend', html);
}

function removeExchangeOld(id) {
  const el = document.getElementById(`exchangeOld${id}`);
  if (el) el.remove();
}

function addExchangeNewGold() {
  exchangeNewCounter++;
  const html = `
    <div class="product-row" id="exchangeNew${exchangeNewCounter}">
      <select class="form-select">
        <option value="">Select product...</option>
        ${FIXED_PRODUCTS.map(p => `<option value="${p.id}">${p.name}</option>`).join('')}
      </select>
      <input type="number" class="form-input" placeholder="Qty" min="1" value="1">
      <button onclick="removeExchangeNew(${exchangeNewCounter})">×</button>
    </div>
  `;
  document.getElementById('exchangeNewGold').insertAdjacentHTML('beforeend', html);
}

function removeExchangeNew(id) {
  const el = document.getElementById(`exchangeNew${id}`);
  if (el) el.remove();
}

function calculateExchange() {
  const customer = document.getElementById('exchangeCustomer').value;
  if (!customer) {
    alert('Please enter customer name');
    return;
  }

  const oldGold = [];
  document.querySelectorAll('#exchangeOldGold .product-row').forEach(row => {
    const productId = row.querySelector('select').value;
    const qty = parseInt(row.querySelector('input').value) || 0;
    if (productId && qty > 0) {
      oldGold.push({ productId, qty });
    }
  });

  const newGold = [];
  document.querySelectorAll('#exchangeNewGold .product-row').forEach(row => {
    const productId = row.querySelector('select').value;
    const qty = parseInt(row.querySelector('input').value) || 0;
    if (productId && qty > 0) {
      newGold.push({ productId, qty });
    }
  });

  if (oldGold.length === 0 || newGold.length === 0) {
    alert('Please add both old and new gold');
    return;
  }

  let exchangeFee = 0;
  newGold.forEach(item => {
    exchangeFee += EXCHANGE_FEES[item.productId] * item.qty;
  });

  let premium = 0;
  newGold.forEach(item => {
    if (PREMIUM_PRODUCTS.includes(item.productId)) {
      premium += PREMIUM_PER_PIECE * item.qty;
    }
  });

  currentExchangeData = {
    customer,
    oldGold: JSON.stringify(oldGold),
    newGold: JSON.stringify(newGold),
    exchangeFee,
    premium
  };

  const oldItems = oldGold.map(i => `${FIXED_PRODUCTS.find(p => p.id === i.productId).name} (${i.qty})`).join(', ');
  const newItems = newGold.map(i => `${FIXED_PRODUCTS.find(p => p.id === i.productId).name} (${i.qty})`).join(', ');

  document.getElementById('exchangeResult').innerHTML = `
    <div class="summary-box">
      <div class="summary-row">
        <span class="summary-label">Old Gold:</span>
        <span class="summary-value">${oldItems}</span>
      </div>
      <div class="summary-row">
        <span class="summary-label">New Gold:</span>
        <span class="summary-value">${newItems}</span>
      </div>
      <div class="summary-row">
        <span class="summary-label">Exchange Fee:</span>
        <span class="summary-value">${formatNumber(exchangeFee)} LAK</span>
      </div>
      <div class="summary-row">
        <span class="summary-label">Premium:</span>
        <span class="summary-value">${formatNumber(premium)} LAK</span>
      </div>
      <div class="summary-row">
        <span class="summary-label">Total to Pay:</span>
        <span class="summary-value">${formatNumber(exchangeFee + premium)} LAK</span>
      </div>
    </div>
  `;

  closeModal('exchangeModal');
  openModal('exchangeResultModal');
}

async function submitExchange() {
  if (!currentExchangeData) return;

  try {
    showLoading();
    const result = await callAppsScript('ADD_EXCHANGE', currentExchangeData);

    if (result.success) {
      alert('✅ Exchange created successfully!');
      closeModal('exchangeResultModal');
      currentExchangeData = null;
      loadExchanges();
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

async function loadExchanges() {
  try {
    showLoading();
    const data = await fetchSheetData('Exchanges!A:J');
    const tbody = document.getElementById('exchangeTable');
    
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="9" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = data.slice(1).map(row => {
        let actions = '';
        if (row[7] === 'PENDING' && currentUser.role === 'Manager') {
          actions = `<button class="btn-action" onclick="reviewTransaction('${row[0]}', 'EXCHANGE')">ตรวจสอบ</button>`;
        } else if (row[7] === 'READY' && (currentUser.role === 'Manager' || currentUser.role === 'Teller')) {
          actions = `<button class="btn-action" onclick="confirmTransaction('${row[0]}', 'EXCHANGE')">ยืนยัน</button>`;
        }
        
        return `
          <tr>
            <td>${row[0]}</td>
            <td>${row[1]}</td>
            <td>${formatItemsForTable(row[2])}</td>
            <td>${formatItemsForTable(row[3])}</td>
            <td>${formatNumber(row[4])}</td>
            <td>${formatNumber(row[5])}</td>
            <td>${formatDateOnly(row[6])}</td>
            <td><span class="status-badge status-${row[7].toLowerCase()}">${row[7]}</span></td>
            <td>${actions}</td>
          </tr>
        `;
      }).join('');
    }
    hideLoading();
  } catch (error) {
    console.error('Error loading exchanges:', error);
    hideLoading();
  }
}

function addBuybackProduct() {
  buybackProductCounter++;
  const html = `
    <div class="product-row" id="buybackProduct${buybackProductCounter}">
      <select class="form-select" onchange="calculateBuybackTotal()">
        <option value="">Select product...</option>
        ${FIXED_PRODUCTS.map(p => `<option value="${p.id}">${p.name}</option>`).join('')}
      </select>
      <input type="number" class="form-input" placeholder="Qty" min="1" value="1" onchange="calculateBuybackTotal()">
      <button onclick="removeBuybackProduct(${buybackProductCounter})">×</button>
    </div>
  `;
  document.getElementById('buybackProducts').insertAdjacentHTML('beforeend', html);
  calculateBuybackTotal();
}

function removeBuybackProduct(id) {
  const el = document.getElementById(`buybackProduct${id}`);
  if (el) el.remove();
  calculateBuybackTotal();
}

function calculateBuybackTotal() {
  let total = 0;
  const products = document.querySelectorAll('#buybackProducts .product-row');
  products.forEach(row => {
    const productId = row.querySelector('select').value;
    const qty = parseInt(row.querySelector('input').value) || 0;
    if (productId && qty > 0) {
      const price = calculateBuybackPrice(productId, currentPricing.sell1Baht);
      total += price * qty;
    }
  });
  document.getElementById('buybackPrice').value = total;
}

async function submitBuyback() {
  const customer = document.getElementById('buybackCustomer').value;
  const total = parseFloat(document.getElementById('buybackPrice').value);
  
  const products = [];
  document.querySelectorAll('#buybackProducts .product-row').forEach(row => {
    const productId = row.querySelector('select').value;
    const qty = parseInt(row.querySelector('input').value) || 0;
    if (productId && qty > 0) {
      products.push({ productId, qty });
    }
  });

  if (!customer || products.length === 0 || !total) {
    alert('Please fill all fields');
    return;
  }

  try {
    showLoading();
    const result = await callAppsScript('ADD_BUYBACK', {
      customer,
      items: JSON.stringify(products),
      total
    });

    if (result.success) {
      alert('✅ Buyback created successfully!');
      closeModal('buybackModal');
      loadBuybacks();
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

async function loadBuybacks() {
  try {
    showLoading();
    const data = await fetchSheetData('Buybacks!A:G');
    const tbody = document.getElementById('buybackTable');
    
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="7" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = data.slice(1).map(row => {
        let actions = '';
        if (row[5] === 'PENDING' && currentUser.role === 'Manager') {
          actions = `<button class="btn-action" onclick="reviewTransaction('${row[0]}', 'BUYBACK')">ตรวจสอบ</button>`;
        }
        
        return `
          <tr>
            <td>${row[0]}</td>
            <td>${row[1]}</td>
            <td>${formatItemsForTable(row[2])}</td>
            <td>${formatNumber(row[3])}</td>
            <td>${formatDateOnly(row[4])}</td>
            <td><span class="status-badge status-${row[5].toLowerCase()}">${row[5]}</span></td>
            <td>${actions}</td>
          </tr>
        `;
      }).join('');
    }
    hideLoading();
  } catch (error) {
    console.error('Error loading buybacks:', error);
    hideLoading();
  }
}

function reviewTransaction(id, type) {
  currentApprovalItem = { id, type };
  document.getElementById('approvalDetails').innerHTML = `<p>รหัสรายการ: ${id}</p><p>ประเภท: ${type}</p>`;
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
      if (currentApprovalItem.type === 'SELL') loadSells();
      else if (currentApprovalItem.type === 'TRADEIN') loadTradeins();
      else if (currentApprovalItem.type === 'BUYBACK') loadBuybacks();
      else if (currentApprovalItem.type === 'EXCHANGE') loadExchanges();
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
      if (currentApprovalItem.type === 'SELL') loadSells();
      else if (currentApprovalItem.type === 'TRADEIN') loadTradeins();
      else if (currentApprovalItem.type === 'BUYBACK') loadBuybacks();
      else if (currentApprovalItem.type === 'EXCHANGE') loadExchanges();
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

async function confirmTransaction(id, type) {
  if (!confirm('Confirm this transaction?')) return;
  
  try {
    showLoading();
    const result = await callAppsScript('CONFIRM_TRANSACTION', { id, type });

    if (result.success) {
      alert('✅ Transaction confirmed!');
      if (type === 'SELL') loadSells();
      else if (type === 'TRADEIN') loadTradeins();
      else if (type === 'EXCHANGE') loadExchanges();
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
          actions = `<button class="btn-action" onclick="reviewTransaction('${row[0]}', 'WITHDRAW')">ตรวจสอบ</button>`;
        }
        
        return `
          <tr>
            <td>${row[0]}</td>
            <td>${row[1]}</td>
            <td>${row[2]}</td>
            <td>${row[3]}</td>
            <td>${row[4] ? formatNumber(row[4]) : '-'}</td>
            <td>${formatDateOnly(row[5])}</td>
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
    const result = await callAppsScript('ADD_WITHDRAW', { product, quantity });

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
    const result = await callAppsScript('ADD_STOCK_IN', { product, quantity, cost });

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
    const result = await callAppsScript('ADD_STOCK_OUT', { product, quantity });

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
    const [sellData, tradeinData, buybackData] = await Promise.all([
      fetchSheetData('Sells!A:I'),
      fetchSheetData('Tradeins!A:K'),
      fetchSheetData('Buybacks!A:G')
    ]);

    let gp = 0;
    sellData.slice(1).forEach(row => {
      if (row[7] === 'COMPLETED') gp += parseFloat(row[3]) || 0;
    });
    buybackData.slice(1).forEach(row => {
      if (row[5] === 'COMPLETED') gp -= parseFloat(row[3]) || 0;
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

function formatDateOnly(date) {
  if (!date) return '-';
  const d = new Date(date);
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  return `${day}/${month}/${year}`;
}

function formatDateTime(date) {
  if (!date) return '-';
  const d = new Date(date);
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  const hour = String(d.getHours()).padStart(2, '0');
  const minute = String(d.getMinutes()).padStart(2, '0');
  return `${day}/${month}/${year} ${hour}:${minute}`;
}

document.querySelectorAll('.modal').forEach(modal => {
  modal.addEventListener('click', function(e) {
    if (e.target === modal) {
      modal.classList.remove('active');
    }
  });
});
