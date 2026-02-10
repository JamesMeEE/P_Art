async function loadStockNew() {
  try {
    showLoading();

    const [stockData, moveResult] = await Promise.all([
      safeFetch('Stock_New!A:D'),
      callAppsScript('GET_STOCK_MOVES', { sheet: 'StockMove_New' })
    ]);

    let carry = {};
    FIXED_PRODUCTS.forEach(p => { carry[p.id] = 0; });
    let qtyIn = {}, qtyOut = {};
    FIXED_PRODUCTS.forEach(p => { qtyIn[p.id] = 0; qtyOut[p.id] = 0; });

    if (stockData.length > 1) {
      const lastRow = stockData[stockData.length - 1];
      try { carry = JSON.parse(lastRow[1] || '{}'); } catch(e) {}
      try { qtyIn = JSON.parse(lastRow[2] || '{}'); } catch(e) {}
      try { qtyOut = JSON.parse(lastRow[3] || '{}'); } catch(e) {}
    }

    document.getElementById('stockNewSummaryTable').innerHTML = FIXED_PRODUCTS.map(p => {
      const c = carry[p.id] || 0;
      const i = qtyIn[p.id] || 0;
      const o = qtyOut[p.id] || 0;
      return '<tr><td>' + p.id + '</td><td>' + p.name + '</td><td>' + c + '</td>' +
        '<td style="color:#4caf50;">' + (i > 0 ? '+' + i : '0') + '</td>' +
        '<td style="color:#f44336;">' + (o > 0 ? '-' + o : '0') + '</td>' +
        '<td style="font-weight:bold;">' + (c + i - o) + '</td></tr>';
    }).join('');

    let carryWeightG = 0;
    FIXED_PRODUCTS.forEach(p => { carryWeightG += (carry[p.id] || 0) * getGoldWeight(p.id); });

    var prevW = carryWeightG + (moveResult.data ? moveResult.data.prevW || 0 : 0);
    var prevC = moveResult.data ? moveResult.data.prevC || 0 : 0;
    var moves = moveResult.data ? moveResult.data.moves || [] : [];

    var todayMovements = moves.map(function(m) {
      var gIn = m.dir === 'IN' ? m.goldG : 0;
      var gOut = m.dir === 'OUT' ? m.goldG : 0;
      var pIn = m.dir === 'IN' ? m.price : 0;
      var pOut = m.dir === 'OUT' ? m.price : 0;
      return { id: m.id, type: m.type, goldIn: gIn, goldOut: gOut, priceIn: pIn, priceOut: pOut };
    });

    let w = prevW, c = prevC;
    todayMovements.forEach(m => {
      w += m.goldIn - m.goldOut;
      c += m.priceIn - m.priceOut;
      m.w = w; m.c = c;
    });

    const latestW = todayMovements.length > 0 ? todayMovements[todayMovements.length - 1].w : prevW;
    const latestC = todayMovements.length > 0 ? todayMovements[todayMovements.length - 1].c : prevC;

    document.getElementById('stockNewGoldG').textContent = formatWeight(latestW) + ' g';
    document.getElementById('stockNewCostValue').textContent = formatNumber(Math.round(latestC / 1000) * 1000) + ' LAK';
    window._stockNewLatest = { goldG: latestW, cost: latestC };

    const movBody = document.getElementById('stockNewMovementTable');
    let rows = '';

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
      rows += todayMovements.map(m => '<tr>' +
        '<td>' + m.id + '</td>' +
        '<td><span class="status-badge">' + m.type + '</span></td>' +
        '<td style="color:#4caf50;">' + (m.goldIn > 0 ? formatWeight(m.goldIn) : '-') + '</td>' +
        '<td style="color:#f44336;">' + (m.goldOut > 0 ? formatWeight(m.goldOut) : '-') + '</td>' +
        '<td style="font-weight:bold;">' + formatWeight(m.w) + '</td>' +
        '<td style="color:#4caf50;">' + (m.priceIn > 0 ? formatNumber(m.priceIn) : '-') + '</td>' +
        '<td style="color:#f44336;">' + (m.priceOut > 0 ? formatNumber(m.priceOut) : '-') + '</td>' +
        '<td style="font-weight:bold;">' + formatNumber(Math.round(m.c / 1000) * 1000) + '</td>' +
        '<td><button class="btn-action" onclick="viewBillDetail(\'' + m.id + '\',\'' + m.type + '\')">📋</button></td>' +
        '</tr>').join('');
      movBody.innerHTML = rows;
    }

    hideLoading();
    loadPendingTransferCount();
  } catch(error) {
    console.error('Error loading stock new:', error);
    hideLoading();
  }
}

async function loadPendingTransferCount() {
  try {
    var result = await callAppsScript('GET_PENDING_TRANSFERS');
    var pending = result.data ? result.data.pending || [] : [];
    const badge = document.getElementById('pendingTransferBadge');
    if (badge) {
      if (pending.length > 0) { badge.textContent = pending.length; badge.style.display = 'block'; }
      else { badge.style.display = 'none'; }
    }
  } catch(e) {}
}

var _pendingTransfers = [];

async function openPendingTransferModal() {
  openModal('pendingTransferModal');
  const body = document.getElementById('pendingTransferBody');
  body.innerHTML = '<p style="text-align:center;padding:20px;">Loading...</p>';

  try {
    var result = await callAppsScript('GET_PENDING_TRANSFERS');
    _pendingTransfers = result.data ? result.data.pending || [] : [];

    if (_pendingTransfers.length === 0) {
      body.innerHTML = '<p style="text-align:center;padding:40px;color:var(--text-secondary);">ไม่มีรายการรอรับทอง</p>';
      return;
    }

    body.innerHTML = _pendingTransfers.map(function(row, idx) {
      var items = [];
      try { items = JSON.parse(row.items); } catch(e) {}
      var totalQty = 0;
      items.forEach(function(i) { totalQty += i.qty; });
      var chips = items.map(function(i) {
        var p = FIXED_PRODUCTS.find(function(x) { return x.id === i.productId; });
        return '<span style="display:inline-block;background:rgba(212,175,55,0.15);padding:3px 8px;border-radius:10px;margin:2px;font-size:12px;">' + (p ? p.name : i.productId) + ' x' + i.qty + '</span>';
      }).join('');

      return '<div onclick="openTransferDetail(' + idx + ')" style="border:1px solid rgba(212,175,55,0.3);border-radius:10px;padding:15px;margin-bottom:10px;cursor:pointer;transition:background 0.2s;" onmouseover="this.style.background=\'rgba(212,175,55,0.08)\'" onmouseout="this.style.background=\'transparent\'">' +
        '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;">' +
        '<span style="font-weight:bold;color:var(--gold-primary);font-size:16px;">' + row.id + '</span>' +
        '<span style="font-size:12px;color:var(--text-secondary);">โดย ' + (row.user || '-') + '</span></div>' +
        '<div>' + chips + '</div>' +
        '<div style="text-align:right;margin-top:6px;font-size:12px;color:var(--text-secondary);">รวม ' + totalQty + ' ชิ้น →</div>' +
        '</div>';
    }).join('');
  } catch(e) {
    body.innerHTML = '<p style="color:#f44336;">Error: ' + e.message + '</p>';
  }
}

async function openTransferDetail(idx) {
  var row = _pendingTransfers[idx];
  if (!row) return;
  closeModal('pendingTransferModal');

  var titleEl = document.getElementById('transferDetailTitle');
  var bodyEl = document.getElementById('transferDetailBody');
  var footerEl = document.getElementById('transferDetailFooter');
  titleEl.textContent = '📋 ' + row.id;
  bodyEl.innerHTML = '<p style="text-align:center;padding:20px;">กำลังคำนวณ...</p>';
  footerEl.innerHTML = '';
  openModal('transferDetailModal');

  var items = [];
  try { items = JSON.parse(row.items); } catch(e) {}

  var wacPerG = 0;
  try {
    var wacResult = await callAppsScript('GET_WAC');
    if (wacResult.data) wacPerG = wacResult.data.wacPerG || 0;
  } catch(e) {}

  var tableRows = '';
  var totalWeight = 0;
  items.forEach(function(i) {
    var p = FIXED_PRODUCTS.find(function(x) { return x.id === i.productId; });
    var w = getGoldWeight(i.productId) * i.qty;
    totalWeight += w;
    tableRows += '<tr><td>' + (p ? p.name : i.productId) + '</td><td style="text-align:center;">' + i.qty + '</td><td style="text-align:right;">' + formatWeight(w) + '</td></tr>';
  });

  var totalCost = Math.round(totalWeight * wacPerG / 1000) * 1000;

  var html = '<div class="table-container"><table style="width:100%;font-size:14px;"><thead><tr><th>สินค้า</th><th style="text-align:center;">จำนวน</th><th style="text-align:right;">น้ำหนัก (g)</th></tr></thead><tbody>' + tableRows +
    '<tr style="font-weight:bold;background:rgba(212,175,55,0.1);"><td>รวม</td><td></td><td style="text-align:right;">' + formatWeight(totalWeight) + ' g</td></tr></tbody></table></div>';

  html += '<div style="margin-top:20px;display:grid;grid-template-columns:1fr 1fr;gap:12px;">';
  html += '<div style="padding:12px;background:rgba(212,175,55,0.08);border-radius:8px;border:1px solid rgba(212,175,55,0.2);"><div style="font-size:11px;color:var(--text-secondary);margin-bottom:4px;">WAC /g ปัจจุบัน</div><div style="font-size:18px;font-weight:bold;color:var(--gold-primary);">' + formatNumber(wacPerG) + ' LAK</div></div>';
  html += '<div style="padding:12px;background:rgba(212,175,55,0.08);border-radius:8px;border:1px solid rgba(212,175,55,0.2);"><div style="font-size:11px;color:var(--text-secondary);margin-bottom:4px;">PRICE ต้นทุน</div><div style="font-size:18px;font-weight:bold;color:var(--gold-primary);">' + formatNumber(totalCost) + ' LAK</div></div>';
  html += '</div>';

  html += '<div style="margin-top:8px;font-size:11px;color:var(--text-secondary);">= ' + formatWeight(totalWeight) + ' g × ' + formatNumber(wacPerG) + ' LAK/g</div>';

  html += '<div style="margin-top:15px;"><label style="font-size:12px;color:var(--text-secondary);display:block;margin-bottom:6px;">Note</label><textarea id="transferNote" class="form-input" rows="2" placeholder="หมายเหตุ (ถ้ามี)" style="font-size:13px;"></textarea></div>';

  html += '<div style="margin-top:10px;font-size:11px;color:var(--text-secondary);">สร้างโดย: ' + (row.user || '-') + '</div>';

  bodyEl.innerHTML = html;
  footerEl.innerHTML = '<button class="btn-secondary" onclick="backToPendingList()">ย้อนกลับ</button>' +
    '<button class="btn-danger" onclick="rejectTransfer(\'' + row.id + '\')" style="margin-left:auto;">ปฏิเสธ</button>' +
    '<button class="btn-primary" onclick="confirmTransferReceive(\'' + row.id + '\')">✅ ยืนยันรับทอง</button>';
}

function backToPendingList() {
  closeModal('transferDetailModal');
  openPendingTransferModal();
}

async function confirmTransferReceive(id) {
  if (!confirm('ยืนยันรับทอง ' + id + ' เข้า Stock (NEW)?')) return;
  try {
    showLoading();
    const result = await callAppsScript('CONFIRM_TRANSFER', { id: id });
    hideLoading();
    if (result.success) {
      alert('✅ ' + result.message);
      closeModal('transferDetailModal');
      await loadStockNew();
    } else {
      alert('❌ ' + result.message);
    }
  } catch(e) { hideLoading(); alert('❌ ' + e.message); }
}

async function rejectTransfer(id) {
  if (!confirm('ปฏิเสธรายการโอน ' + id + '? ทองจะถูกคืนกลับ Stock (OLD)')) return;
  try {
    showLoading();
    const result = await callAppsScript('REJECT_TRANSFER', { id: id });
    hideLoading();
    if (result.success) {
      alert('✅ ' + result.message);
      closeModal('transferDetailModal');
      loadPendingTransferCount();
    } else {
      alert('❌ ' + result.message);
    }
  } catch(e) { hideLoading(); alert('❌ ' + e.message); }
}

function addStockNewProduct(containerId) {
  const container = document.getElementById(containerId);
  const row = document.createElement('div');
  row.className = 'product-row';
  row.style.cssText = 'display:flex;gap:10px;margin-bottom:10px;align-items:center;';
  row.innerHTML = '<select class="form-input" style="flex:1;">' +
    FIXED_PRODUCTS.map(p => '<option value="' + p.id + '">' + p.name + '</option>').join('') +
    '</select>' +
    '<input type="number" class="form-input" placeholder="Quantity" min="1" style="width:150px;">' +
    '<button class="btn-danger" onclick="this.parentElement.remove()" style="padding:8px 15px;">Remove</button>';
  container.appendChild(row);
}

function openStockInNewModal() {
  document.getElementById('stockInNewProducts').innerHTML = '';
  document.getElementById('stockInNewNote').value = '';
  addStockNewProduct('stockInNewProducts');
  openModal('stockInNewModal');
}

function openStockOutNewModal() {
  document.getElementById('stockOutNewProducts').innerHTML = '';
  document.getElementById('stockOutNewNote').value = '';
  addStockNewProduct('stockOutNewProducts');
  openModal('stockOutNewModal');
}

async function confirmStockInNew() {
  try {
    const rows = document.querySelectorAll('#stockInNewProducts .product-row');
    const items = [];
    for (const row of rows) {
      const qty = parseInt(row.querySelector('input').value);
      if (!qty || qty <= 0) { alert('กรุณากรอกจำนวนให้ถูกต้อง'); return; }
      items.push({ productId: row.querySelector('select').value, qty: qty });
    }
    if (items.length === 0) { alert('กรุณาเพิ่มสินค้า'); return; }
    const note = document.getElementById('stockInNewNote').value.trim();
    if (!confirm('ยืนยัน Stock In (NEW) ' + items.length + ' รายการ?')) return;
    showLoading();
    const result = await callAppsScript('STOCK_IN_NEW', { items: JSON.stringify(items), note: note });
    hideLoading();
    if (result.success) {
      alert('✅ ' + result.message);
      closeModal('stockInNewModal');
      await loadStockNew();
    } else { alert('❌ ' + result.message); }
  } catch(e) { hideLoading(); alert('❌ ' + e.message); }
}

async function confirmStockOutNew() {
  try {
    const rows = document.querySelectorAll('#stockOutNewProducts .product-row');
    const items = [];
    for (const row of rows) {
      const qty = parseInt(row.querySelector('input').value);
      if (!qty || qty <= 0) { alert('กรุณากรอกจำนวนให้ถูกต้อง'); return; }
      items.push({ productId: row.querySelector('select').value, qty: qty });
    }
    if (items.length === 0) { alert('กรุณาเพิ่มสินค้า'); return; }
    const note = document.getElementById('stockOutNewNote').value.trim();
    if (!confirm('ยืนยัน Stock Out (NEW) ' + items.length + ' รายการ?')) return;
    showLoading();
    const result = await callAppsScript('STOCK_OUT_NEW', { items: JSON.stringify(items), note: note });
    hideLoading();
    if (result.success) {
      alert('✅ ' + result.message);
      closeModal('stockOutNewModal');
      await loadStockNew();
    } else { alert('❌ ' + result.message); }
  } catch(e) { hideLoading(); alert('❌ ' + e.message); }
}