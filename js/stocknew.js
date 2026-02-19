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

function getStockInRate(currency) {
  if (currency === 'LAK') return 1;
  if (currency === 'THB') return currentExchangeRates.THB || 1;
  if (currency === 'USD') return currentExchangeRates.USD || 1;
  return 1;
}

function addStockInCash() {
  _stockInPayments.cash.push({ id: Date.now(), currency: 'LAK', amount: 0, rate: 1 });
  renderStockInCash();
}

function addStockInBank() {
  _stockInPayments.bank.push({ id: Date.now(), bank: 'BCEL', currency: 'LAK', amount: 0, rate: 1 });
  renderStockInBank();
}

function renderStockInCash() {
  var container = document.getElementById('stockInCashList');
  container.innerHTML = _stockInPayments.cash.map(function(item) {
    var lakAmount = item.amount * item.rate;
    var rateHtml = '';
    if (item.currency !== 'LAK') {
      rateHtml = '<div style="display:flex;justify-content:space-between;align-items:center;margin-top:8px;padding-top:8px;border-top:1px dashed var(--border-color);">' +
        '<span style="font-size:11px;color:var(--text-secondary);">Rate: 1 ' + item.currency + ' = ' + formatNumber(item.rate) + ' LAK</span>' +
        '<span class="lak-display" style="color:var(--gold-primary);font-weight:bold;">= ' + formatNumber(lakAmount) + ' LAK</span>' +
        '</div>';
    }
    return '<div class="payment-item" data-id="' + item.id + '" style="margin-bottom:10px;padding:12px;background:var(--bg-light);border-radius:8px;">' +
      '<div style="display:flex;gap:10px;align-items:center;">' +
      '<select class="form-select" style="width:90px;" onchange="updateStockInCashCurrency(' + item.id + ',this.value)">' +
      '<option value="LAK"' + (item.currency === 'LAK' ? ' selected' : '') + '>LAK</option>' +
      '<option value="THB"' + (item.currency === 'THB' ? ' selected' : '') + '>THB</option>' +
      '<option value="USD"' + (item.currency === 'USD' ? ' selected' : '') + '>USD</option>' +
      '</select>' +
      '<input type="number" class="form-input" placeholder="Amount" value="' + (item.amount || '') + '" style="flex:1;" oninput="updateStockInCashAmount(' + item.id + ',this.value)">' +
      '<button class="btn-secondary" style="padding:8px 12px;background:#f44336;color:white;" onclick="removeStockInCash(' + item.id + ')">✕</button>' +
      '</div>' + rateHtml + '</div>';
  }).join('');
  updateStockInTotal();
}

function renderStockInBank() {
  var container = document.getElementById('stockInBankList');
  container.innerHTML = _stockInPayments.bank.map(function(item) {
    var lakAmount = item.amount * item.rate;
    var rateHtml = '';
    if (item.currency !== 'LAK') {
      rateHtml = '<div style="display:flex;justify-content:space-between;align-items:center;margin-top:8px;padding-top:8px;border-top:1px dashed var(--border-color);">' +
        '<span style="font-size:11px;color:var(--text-secondary);">Rate: 1 ' + item.currency + ' = ' + formatNumber(item.rate) + ' LAK</span>' +
        '<span class="lak-display" style="color:var(--gold-primary);font-weight:bold;">= ' + formatNumber(lakAmount) + ' LAK</span>' +
        '</div>';
    }
    return '<div class="payment-item" data-id="' + item.id + '" style="margin-bottom:10px;padding:12px;background:var(--bg-light);border-radius:8px;">' +
      '<div style="display:flex;gap:10px;align-items:center;margin-bottom:8px;">' +
      '<select class="form-select" style="flex:1;" onchange="updateStockInBankName(' + item.id + ',this.value)">' +
      '<option value="BCEL"' + (item.bank === 'BCEL' ? ' selected' : '') + '>BCEL</option>' +
      '<option value="LDB"' + (item.bank === 'LDB' ? ' selected' : '') + '>LDB</option>' +
      '<option value="OTHER"' + (item.bank === 'OTHER' ? ' selected' : '') + '>อื่น ๆ</option>' +
      '</select>' +
      '<select class="form-select" style="flex:1;" onchange="updateStockInBankCurrency(' + item.id + ',this.value)">' +
      '<option value="LAK"' + (item.currency === 'LAK' ? ' selected' : '') + '>LAK</option>' +
      '<option value="THB"' + (item.currency === 'THB' ? ' selected' : '') + '>THB</option>' +
      '<option value="USD"' + (item.currency === 'USD' ? ' selected' : '') + '>USD</option>' +
      '</select>' +
      '<button class="btn-secondary" style="padding:8px 12px;background:#f44336;color:white;" onclick="removeStockInBank(' + item.id + ')">✕</button>' +
      '</div>' +
      '<div style="display:flex;gap:10px;align-items:center;">' +
      '<input type="number" class="form-input" placeholder="Amount" value="' + (item.amount || '') + '" style="flex:1;" oninput="updateStockInBankAmount(' + item.id + ',this.value)">' +
      '</div>' + rateHtml + '</div>';
  }).join('');
  updateStockInTotal();
}

function updateStockInCashCurrency(id, value) {
  var item = _stockInPayments.cash.find(function(i) { return i.id === id; });
  if (!item) return;
  item.currency = value;
  item.rate = getStockInRate(value);
  renderStockInCash();
}

function updateStockInCashAmount(id, value) {
  var item = _stockInPayments.cash.find(function(i) { return i.id === id; });
  if (!item) return;
  item.amount = parseFloat(value) || 0;
  var lakAmount = item.amount * item.rate;
  var container = document.querySelector('#stockInCashList .payment-item[data-id="' + id + '"]');
  if (container && item.currency !== 'LAK') {
    var lakSpan = container.querySelector('.lak-display');
    if (lakSpan) lakSpan.textContent = '= ' + formatNumber(lakAmount) + ' LAK';
  }
  updateStockInTotal();
}

function updateStockInBankName(id, value) {
  var item = _stockInPayments.bank.find(function(i) { return i.id === id; });
  if (item) item.bank = value;
}

function updateStockInBankCurrency(id, value) {
  var item = _stockInPayments.bank.find(function(i) { return i.id === id; });
  if (!item) return;
  item.currency = value;
  item.rate = getStockInRate(value);
  renderStockInBank();
}

function updateStockInBankAmount(id, value) {
  var item = _stockInPayments.bank.find(function(i) { return i.id === id; });
  if (!item) return;
  item.amount = parseFloat(value) || 0;
  var lakAmount = item.amount * item.rate;
  var container = document.querySelector('#stockInBankList .payment-item[data-id="' + id + '"]');
  if (container && item.currency !== 'LAK') {
    var lakSpan = container.querySelector('.lak-display');
    if (lakSpan) lakSpan.textContent = '= ' + formatNumber(lakAmount) + ' LAK';
  }
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
  _stockInPayments.cash.forEach(function(i) { total += i.amount * i.rate; });
  _stockInPayments.bank.forEach(function(i) { total += i.amount * i.rate; });
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
    var totalCost = 0;
    _stockInPayments.cash.forEach(function(c) {
      if (c.amount > 0) {
        payments.push({ method: 'Cash', bank: '', currency: c.currency, amount: c.amount });
        totalCost += c.amount * c.rate;
      }
    });
    _stockInPayments.bank.forEach(function(b) {
      if (b.amount > 0) {
        payments.push({ method: 'Bank', bank: b.bank, currency: b.currency, amount: b.amount });
        totalCost += b.amount * b.rate;
      }
    });
    if (payments.length === 0) { alert('กรุณาเพิ่มการชำระเงิน'); return; }

    showLoading();
    var dbData = await fetchSheetData('_database!A1:G31');
    hideLoading();

    var shopBalance = {
      cash: { LAK: 0, THB: 0, USD: 0 },
      BCEL: { LAK: 0, THB: 0, USD: 0 },
      LDB: { LAK: 0, THB: 0, USD: 0 },
      OTHER: { LAK: 0, THB: 0, USD: 0 }
    };
    if (dbData.length >= 17) {
      shopBalance.cash.LAK = parseFloat(dbData[16][0]) || 0;
      shopBalance.cash.THB = parseFloat(dbData[16][1]) || 0;
      shopBalance.cash.USD = parseFloat(dbData[16][2]) || 0;
      shopBalance.OTHER.LAK = parseFloat(dbData[16][4]) || 0;
      shopBalance.OTHER.THB = parseFloat(dbData[16][5]) || 0;
      shopBalance.OTHER.USD = parseFloat(dbData[16][6]) || 0;
    }
    if (dbData.length >= 20) {
      shopBalance.BCEL.LAK = parseFloat(dbData[19][0]) || 0;
      shopBalance.BCEL.THB = parseFloat(dbData[19][1]) || 0;
      shopBalance.BCEL.USD = parseFloat(dbData[19][2]) || 0;
    }
    if (dbData.length >= 23) {
      shopBalance.LDB.LAK = parseFloat(dbData[22][0]) || 0;
      shopBalance.LDB.THB = parseFloat(dbData[22][1]) || 0;
      shopBalance.LDB.USD = parseFloat(dbData[22][2]) || 0;
    }

    var usageSummary = {};
    payments.forEach(function(p) {
      var source = p.method === 'Cash' ? 'cash' : (p.bank || 'OTHER');
      var cur = p.currency || 'LAK';
      var key = source + '_' + cur;
      if (!usageSummary[key]) usageSummary[key] = { source: source, currency: cur, total: 0 };
      usageSummary[key].total += p.amount;
    });

    var insufficientList = [];
    Object.keys(usageSummary).forEach(function(key) {
      var u = usageSummary[key];
      var available = (shopBalance[u.source] && shopBalance[u.source][u.currency]) ? shopBalance[u.source][u.currency] : 0;
      if (u.total > available) {
        var label = u.source === 'cash' ? 'เงินสด' : u.source;
        insufficientList.push(label + ' ' + u.currency + ': ต้องการ ' + formatNumber(u.total) + ' แต่มี ' + formatNumber(available));
      }
    });

    if (insufficientList.length > 0) {
      alert('❌ ยอดเงินร้านไม่เพียงพอ\n\n' + insufficientList.join('\n'));
      return;
    }

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