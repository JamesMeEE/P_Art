var stockNewDateFrom = null;
var stockNewDateTo = null;

async function loadStockNew() {
  try {
    showLoading();
    var isFiltered = stockNewDateFrom && stockNewDateTo;
    var isToday = false;
    if (isFiltered) {
      var td = getTodayDateString();
      isToday = (stockNewDateFrom === td && stockNewDateTo === td);
    }

    var stockData = await safeFetch('Stock_New!A:D');

    if (!isFiltered || isToday) {
      var carry = {}, qtyIn = {}, qtyOut = {};
      FIXED_PRODUCTS.forEach(function(p) { carry[p.id] = 0; qtyIn[p.id] = 0; qtyOut[p.id] = 0; });
      if (stockData.length > 1) {
        var lastRow = stockData[stockData.length - 1];
        try { carry = JSON.parse(lastRow[1] || '{}'); } catch(e) {}
        try { qtyIn = JSON.parse(lastRow[2] || '{}'); } catch(e) {}
        try { qtyOut = JSON.parse(lastRow[3] || '{}'); } catch(e) {}
      }
      renderStockNewSummary(carry, qtyIn, qtyOut);

      var moveResult = await callAppsScript('GET_STOCK_MOVES', { sheet: 'StockMove_New' });
      var prevW = moveResult.data ? moveResult.data.prevW || 0 : 0;
      var prevC = moveResult.data ? moveResult.data.prevC || 0 : 0;
      var moves = moveResult.data ? moveResult.data.moves || [] : [];
      renderStockNewMovements(moves, prevW, prevC);
    } else {
      var from = new Date(stockNewDateFrom); from.setHours(0,0,0,0);
      var to = new Date(stockNewDateTo); to.setHours(23,59,59,999);
      var carry = {}, qtyIn = {}, qtyOut = {};
      FIXED_PRODUCTS.forEach(function(p) { carry[p.id] = 0; qtyIn[p.id] = 0; qtyOut[p.id] = 0; });
      var foundCarry = false;
      for (var r = 1; r < stockData.length; r++) {
        var rd = new Date(stockData[r][0]);
        if (isNaN(rd.getTime())) continue;
        rd.setHours(0,0,0,0);
        if (rd >= from && rd <= to) {
          if (!foundCarry) {
            try { carry = JSON.parse(stockData[r][1] || '{}'); } catch(e) {}
            foundCarry = true;
          }
          var ri = {}, ro = {};
          try { ri = JSON.parse(stockData[r][2] || '{}'); } catch(e) {}
          try { ro = JSON.parse(stockData[r][3] || '{}'); } catch(e) {}
          FIXED_PRODUCTS.forEach(function(p) { qtyIn[p.id] += (ri[p.id] || 0); qtyOut[p.id] += (ro[p.id] || 0); });
        }
      }
      renderStockNewSummary(carry, qtyIn, qtyOut);

      var moveResult = await callAppsScript('GET_STOCK_MOVES_RANGE', { sheet: 'StockMove_New', dateFrom: stockNewDateFrom, dateTo: stockNewDateTo });
      var moves = moveResult.data ? moveResult.data.moves || [] : [];
      var filtered = moves.filter(function(m) { return m.type === 'TRANSFER' || m.type === 'STOCK-IN'; });
      renderFilteredMoves('stockNewMovementTable', filtered, stockNewDateFrom, stockNewDateTo);
      document.getElementById('stockNewGoldG').textContent = '-';
      document.getElementById('stockNewCostValue').textContent = '-';
    }

    hideLoading();
  } catch(error) {
    console.error('Error loading stock new:', error);
    hideLoading();
  }
}

function renderStockNewSummary(carry, qtyIn, qtyOut) {
  document.getElementById('stockNewSummaryTable').innerHTML = FIXED_PRODUCTS.map(function(p) {
    var c = carry[p.id] || 0;
    var i = qtyIn[p.id] || 0;
    var o = qtyOut[p.id] || 0;
    return '<tr><td>' + p.id + '</td><td>' + p.name + '</td><td>' + c + '</td>' +
      '<td style="color:#4caf50;">' + i + '</td>' +
      '<td style="color:#f44336;">' + o + '</td>' +
      '<td style="font-weight:bold;">' + (c + i - o) + '</td></tr>';
  }).join('');
}

function renderStockNewMovements(moves, prevW, prevC) {
  var w = prevW, c = prevC;
  var todayMovements = [];
  moves.forEach(function(m) {
    var gIn = m.dir === 'IN' ? m.goldG : 0;
    var gOut = m.dir === 'OUT' ? m.goldG : 0;
    var pIn = m.dir === 'IN' ? m.price : 0;
    var pOut = m.dir === 'OUT' ? m.price : 0;
    w += gIn - gOut;
    c += pIn - pOut;
    if (m.type === 'TRANSFER' || m.type === 'STOCK-IN') {
      todayMovements.push({ id: m.id, type: m.type, goldIn: gIn, goldOut: gOut, priceIn: pIn, priceOut: pOut, w: w, c: c });
    }
  });

  document.getElementById('stockNewGoldG').textContent = formatWeight(w) + ' g';
  document.getElementById('stockNewCostValue').textContent = formatNumber(Math.round(c / 1000) * 1000) + ' LAK';
  window._stockNewLatest = { goldG: w, cost: c };

  var movBody = document.getElementById('stockNewMovementTable');
  var rows = '';

  if (prevW !== 0 || prevC !== 0) {
    rows += '<tr style="background:rgba(212,175,55,0.06);">' +
      '<td colspan="4" style="font-style:italic;color:var(--gold-primary);">📌 ยกมา</td>' +
      '<td style="font-weight:bold;">' + formatWeight(prevW) + '</td>' +
      '<td colspan="2"></td>' +
      '<td style="font-weight:bold;">' + formatNumber(Math.round(prevC / 1000) * 1000) + '</td>' +
      '<td></td></tr>';
  }

  if (todayMovements.length === 0 && rows === '') {
    movBody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:40px;">ไม่มีรายการวันนี้</td></tr>';
  } else {
    rows += todayMovements.map(function(m) { return '<tr>' +
      '<td>' + m.id + '</td>' +
      '<td><span class="status-badge">' + m.type + '</span></td>' +
      '<td style="color:#4caf50;">' + (m.goldIn > 0 ? formatWeight(m.goldIn) : '-') + '</td>' +
      '<td style="color:#f44336;">' + (m.goldOut > 0 ? formatWeight(m.goldOut) : '-') + '</td>' +
      '<td style="font-weight:bold;">' + formatWeight(m.w) + '</td>' +
      '<td style="color:#4caf50;">' + (m.priceIn > 0 ? formatNumber(m.priceIn) : '-') + '</td>' +
      '<td style="color:#f44336;">' + (m.priceOut > 0 ? formatNumber(m.priceOut) : '-') + '</td>' +
      '<td style="font-weight:bold;">' + formatNumber(Math.round(m.c / 1000) * 1000) + '</td>' +
      '<td><button class="btn-action" onclick="viewBillDetail(\'' + m.id + '\',\'' + m.type + '\')">📋</button></td>' +
      '</tr>'; }).join('');
    movBody.innerHTML = rows;
  }
}

function resetStockNewFilter() {
  var today = getTodayDateString();
  document.getElementById('stockNewDateFrom').value = today;
  document.getElementById('stockNewDateTo').value = today;
  stockNewDateFrom = today;
  stockNewDateTo = today;
  loadStockNew();
}

document.addEventListener('DOMContentLoaded', function() {
  var f = document.getElementById('stockNewDateFrom');
  var t = document.getElementById('stockNewDateTo');
  if (f && t) {
    f.addEventListener('change', function() { stockNewDateFrom = this.value; if (stockNewDateFrom && stockNewDateTo) loadStockNew(); });
    t.addEventListener('change', function() { stockNewDateTo = this.value; if (stockNewDateFrom && stockNewDateTo) loadStockNew(); });
  }
});

function loadPendingTransferCount() {}

function addStockNewProduct(containerId) {
  var container = document.getElementById(containerId);
  var row = document.createElement('div');
  row.className = 'product-row';
  row.style.cssText = 'display:flex;gap:10px;margin-bottom:10px;align-items:center;';
  row.innerHTML = '<select class="form-input" style="flex:1;">' +
    FIXED_PRODUCTS.map(function(p) { return '<option value="' + p.id + '">' + p.name + '</option>'; }).join('') +
    '</select>' +
    '<input type="number" class="form-input" placeholder="Quantity" min="1" style="width:150px;">' +
    '<button class="btn-danger" onclick="this.parentElement.remove()" style="padding:8px 15px;">Remove</button>';
  container.appendChild(row);
}

var _stockInPayments = { cash: [], bank: [] };

function openStockInNewModal() {
  document.getElementById('stockInNewProducts').innerHTML = '';
  document.getElementById('stockInNewNote').value = '';
  _stockInPayments = { cash: [], bank: [] };
  document.getElementById('stockInCashList').innerHTML = '';
  document.getElementById('stockInBankList').innerHTML = '';
  document.getElementById('stockInTotalCost').textContent = '0 LAK';
  addStockNewProduct('stockInNewProducts');
  openModal('stockInNewModal');
}

function addStockInCash() {
  _stockInPayments.cash.push({ id: Date.now(), currency: 'LAK', amount: 0 });
  renderStockInCash();
}

function addStockInBank() {
  _stockInPayments.bank.push({ id: Date.now(), bank: 'BCEL', currency: 'LAK', amount: 0 });
  renderStockInBank();
}

function renderStockInCash() {
  var container = document.getElementById('stockInCashList');
  container.innerHTML = _stockInPayments.cash.map(function(item) {
    return '<div data-id="' + item.id + '" style="margin-bottom:8px;padding:10px;background:var(--bg-light);border-radius:8px;display:flex;gap:8px;align-items:center;">' +
      '<select class="form-select" style="width:85px;" onchange="updateStockInCash(' + item.id + ',\'currency\',this.value)">' +
      '<option value="LAK"' + (item.currency === 'LAK' ? ' selected' : '') + '>LAK</option>' +
      '<option value="THB"' + (item.currency === 'THB' ? ' selected' : '') + '>THB</option>' +
      '<option value="USD"' + (item.currency === 'USD' ? ' selected' : '') + '>USD</option>' +
      '</select>' +
      '<input type="number" class="form-input" placeholder="Amount" value="' + (item.amount || '') + '" style="flex:1;" oninput="updateStockInCash(' + item.id + ',\'amount\',this.value)">' +
      '<button class="btn-secondary" style="padding:6px 10px;background:#f44336;color:#fff;" onclick="removeStockInCash(' + item.id + ')">✕</button>' +
      '</div>';
  }).join('');
  updateStockInTotal();
}

function renderStockInBank() {
  var banks = ['BCEL', 'JDB', 'LDB', 'OTHER'];
  var container = document.getElementById('stockInBankList');
  container.innerHTML = _stockInPayments.bank.map(function(item) {
    return '<div data-id="' + item.id + '" style="margin-bottom:8px;padding:10px;background:var(--bg-light);border-radius:8px;">' +
      '<div style="display:flex;gap:8px;align-items:center;">' +
      '<select class="form-select" style="width:90px;" onchange="updateStockInBank(' + item.id + ',\'bank\',this.value)">' +
      banks.map(function(b) { return '<option value="' + b + '"' + (item.bank === b ? ' selected' : '') + '>' + b + '</option>'; }).join('') +
      '</select>' +
      '<select class="form-select" style="width:85px;" onchange="updateStockInBank(' + item.id + ',\'currency\',this.value)">' +
      '<option value="LAK"' + (item.currency === 'LAK' ? ' selected' : '') + '>LAK</option>' +
      '<option value="THB"' + (item.currency === 'THB' ? ' selected' : '') + '>THB</option>' +
      '<option value="USD"' + (item.currency === 'USD' ? ' selected' : '') + '>USD</option>' +
      '</select>' +
      '<input type="number" class="form-input" placeholder="Amount" value="' + (item.amount || '') + '" style="flex:1;" oninput="updateStockInBank(' + item.id + ',\'amount\',this.value)">' +
      '<button class="btn-secondary" style="padding:6px 10px;background:#f44336;color:#fff;" onclick="removeStockInBank(' + item.id + ')">✕</button>' +
      '</div></div>';
  }).join('');
  updateStockInTotal();
}

function updateStockInCash(id, field, value) {
  var item = _stockInPayments.cash.find(function(i) { return i.id === id; });
  if (!item) return;
  if (field === 'amount') item.amount = parseFloat(value) || 0;
  else item[field] = value;
  updateStockInTotal();
}

function updateStockInBank(id, field, value) {
  var item = _stockInPayments.bank.find(function(i) { return i.id === id; });
  if (!item) return;
  if (field === 'amount') item.amount = parseFloat(value) || 0;
  else item[field] = value;
  updateStockInTotal();
}

function removeStockInCash(id) {
  _stockInPayments.cash = _stockInPayments.cash.filter(function(i) { return i.id !== id; });
  renderStockInCash();
}

function removeStockInBank(id) {
  _stockInPayments.bank = _stockInPayments.bank.filter(function(i) { return i.id !== id; });
  renderStockInBank();
}

function updateStockInTotal() {
  var total = 0;
  var rates = window._exchangeRates || { THB: 1, USD: 1 };
  _stockInPayments.cash.forEach(function(i) {
    if (i.currency === 'LAK') total += i.amount;
    else if (i.currency === 'THB') total += i.amount * (rates.THB || 1);
    else if (i.currency === 'USD') total += i.amount * (rates.USD || 1);
  });
  _stockInPayments.bank.forEach(function(i) {
    if (i.currency === 'LAK') total += i.amount;
    else if (i.currency === 'THB') total += i.amount * (rates.THB || 1);
    else if (i.currency === 'USD') total += i.amount * (rates.USD || 1);
  });
  document.getElementById('stockInTotalCost').textContent = formatNumber(Math.round(total)) + ' LAK';
}

async function confirmStockInNew() {
  try {
    var rows = document.querySelectorAll('#stockInNewProducts .product-row');
    var items = [];
    for (var i = 0; i < rows.length; i++) {
      var qty = parseInt(rows[i].querySelector('input').value);
      if (!qty || qty <= 0) { alert('กรุณากรอกจำนวนให้ถูกต้อง'); return; }
      items.push({ productId: rows[i].querySelector('select').value, qty: qty });
    }
    if (items.length === 0) { alert('กรุณาเพิ่มสินค้า'); return; }

    var payments = [];
    _stockInPayments.cash.forEach(function(c) {
      if (c.amount > 0) payments.push({ method: 'Cash', bank: '', currency: c.currency, amount: c.amount });
    });
    _stockInPayments.bank.forEach(function(b) {
      if (b.amount > 0) payments.push({ method: 'Bank', bank: b.bank, currency: b.currency, amount: b.amount });
    });
    if (payments.length === 0) { alert('กรุณาเพิ่มการชำระเงิน'); return; }

    var totalCost = 0;
    var rates = window._exchangeRates || { THB: 1, USD: 1 };
    payments.forEach(function(p) {
      if (p.currency === 'LAK') totalCost += p.amount;
      else if (p.currency === 'THB') totalCost += p.amount * (rates.THB || 1);
      else if (p.currency === 'USD') totalCost += p.amount * (rates.USD || 1);
    });

    if (!confirm('ยืนยัน Stock In (NEW) ' + items.length + ' รายการ ต้นทุน ' + formatNumber(Math.round(totalCost)) + ' LAK?')) return;

    var note = document.getElementById('stockInNewNote').value.trim();
    showLoading();
    var result = await callAppsScript('STOCK_IN_NEW', {
      items: JSON.stringify(items),
      note: note,
      cost: Math.round(totalCost),
      payments: JSON.stringify(payments),
      user: currentUser.nickname
    });
    hideLoading();
    if (result.success) {
      alert('✅ ' + result.message);
      closeModal('stockInNewModal');
      await loadStockNew();
    } else { alert('❌ ' + result.message); }
  } catch(e) { hideLoading(); alert('❌ ' + e.message); }
}