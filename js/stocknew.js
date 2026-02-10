async function loadStockNew() {
  try {
    showLoading();

    const [stockData, moveData] = await Promise.all([
      safeFetch('Stock_New!A:D'),
      safeFetch('StockMove_New!A:J')
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

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayEnd = new Date(today);
    todayEnd.setHours(23, 59, 59, 999);

    let prevW = carryWeightG;
    let prevC = 0;
    const todayMovements = [];

    if (moveData.length > 1) {
      moveData.slice(1).forEach(row => {
        const d = parseSheetDate(row[0]);
        if (!d) return;
        const goldG = parseFloat(row[4]) || 0;
        const price = parseFloat(row[6]) || 0;
        const dir = String(row[5] || '');
        const gIn = dir === 'IN' ? goldG : 0;
        const gOut = dir === 'OUT' ? goldG : 0;
        const pIn = dir === 'IN' ? price : 0;
        const pOut = dir === 'OUT' ? price : 0;

        if (d < today) {
          prevW += gIn - gOut;
          prevC += pIn - pOut;
        } else if (d <= todayEnd) {
          todayMovements.push({ id: row[1], type: row[2], goldIn: gIn, goldOut: gOut, priceIn: pIn, priceOut: pOut });
        }
      });
    }

    let w = prevW, c = prevC;
    todayMovements.forEach(m => {
      w += m.goldIn - m.goldOut;
      c += m.priceIn - m.priceOut;
      m.w = w; m.c = c;
    });

    const latestW = todayMovements.length > 0 ? Math.abs(todayMovements[todayMovements.length - 1].w) : Math.abs(prevW);
    const latestC = todayMovements.length > 0 ? Math.abs(todayMovements[todayMovements.length - 1].c) : Math.abs(prevC);

    document.getElementById('stockNewGoldG').textContent = formatNumber(latestW.toFixed(2)) + ' g';
    document.getElementById('stockNewCostValue').textContent = formatNumber(Math.round(latestC / 1000) * 1000) + ' LAK';
    window._stockNewLatest = { goldG: latestW, cost: latestC };

    const movBody = document.getElementById('stockNewMovementTable');
    if (todayMovements.length === 0) {
      movBody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:40px;">ไม่มีรายการวันนี้</td></tr>';
    } else {
      movBody.innerHTML = todayMovements.map(m => '<tr>' +
        '<td>' + m.id + '</td>' +
        '<td><span class="status-badge">' + m.type + '</span></td>' +
        '<td style="color:#4caf50;">' + (m.goldIn > 0 ? formatNumber(m.goldIn.toFixed(2)) : '-') + '</td>' +
        '<td style="color:#f44336;">' + (m.goldOut > 0 ? formatNumber(m.goldOut.toFixed(2)) : '-') + '</td>' +
        '<td style="font-weight:bold;">' + formatNumber(Math.abs(m.w).toFixed(2)) + '</td>' +
        '<td style="color:#4caf50;">' + (m.priceIn > 0 ? formatNumber(m.priceIn) : '-') + '</td>' +
        '<td style="color:#f44336;">' + (m.priceOut > 0 ? formatNumber(m.priceOut) : '-') + '</td>' +
        '<td style="font-weight:bold;">' + formatNumber(Math.round(Math.abs(m.c) / 1000) * 1000) + '</td>' +
        '<td><button class="btn-action" onclick="viewBillDetail(\'' + m.id + '\',\'' + m.type + '\')">📋</button></td>' +
        '</tr>').join('');
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
    const data = await safeFetch('StockTransfer!A:F');
    let count = 0;
    if (data.length > 1) {
      data.slice(1).forEach(row => { if (row[3] === 'PENDING') count++; });
    }
    const badge = document.getElementById('pendingTransferBadge');
    if (badge) {
      if (count > 0) { badge.textContent = count; badge.style.display = 'block'; }
      else { badge.style.display = 'none'; }
    }
  } catch(e) { console.warn('loadPendingTransferCount error:', e); }
}

async function openPendingTransferModal() {
  openModal('pendingTransferModal');
  const body = document.getElementById('pendingTransferBody');
  body.innerHTML = '<p style="text-align:center;padding:20px;">Loading...</p>';

  try {
    const data = await safeFetch('StockTransfer!A:F');
    const pending = [];
    if (data.length > 1) {
      data.slice(1).forEach(row => { if (row[3] === 'PENDING') pending.push(row); });
    }

    if (pending.length === 0) {
      body.innerHTML = '<p style="text-align:center;padding:40px;color:var(--text-secondary);">ไม่มีรายการรอรับทอง</p>';
      return;
    }

    body.innerHTML = pending.map(row => {
      const id = row[0];
      let itemsHtml = '';
      try {
        const items = JSON.parse(row[1]);
        itemsHtml = items.map(i => {
          const p = FIXED_PRODUCTS.find(x => x.id === i.productId);
          return '<span style="display:inline-block;background:rgba(212,175,55,0.15);padding:4px 10px;border-radius:12px;margin:2px;font-size:13px;">' + (p ? p.name : i.productId) + ' x' + i.qty + '</span>';
        }).join('');
      } catch(e) {}

      return '<div style="border:1px solid rgba(212,175,55,0.3);border-radius:10px;padding:15px;margin-bottom:12px;">' +
        '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">' +
        '<span style="font-weight:bold;color:var(--gold-primary);font-size:16px;">' + id + '</span>' +
        '<span style="font-size:12px;color:var(--text-secondary);">โดย ' + (row[4] || '-') + '</span></div>' +
        '<div style="margin-bottom:12px;">' + itemsHtml + '</div>' +
        '<div style="display:flex;gap:10px;justify-content:flex-end;">' +
        '<button class="btn-danger" onclick="rejectTransfer(\'' + id + '\')" style="padding:6px 16px;font-size:13px;">ปฏิเสธ</button>' +
        '<button class="btn-primary" onclick="confirmTransferReceive(\'' + id + '\')" style="padding:6px 16px;font-size:13px;">✅ ยืนยันรับทอง</button>' +
        '</div></div>';
    }).join('');
  } catch(e) {
    body.innerHTML = '<p style="color:#f44336;">Error: ' + e.message + '</p>';
  }
}

async function confirmTransferReceive(id) {
  if (!confirm('ยืนยันรับทอง ' + id + ' เข้า Stock (NEW)?')) return;
  try {
    showLoading();
    const result = await executeGoogleScript('CONFIRM_TRANSFER', { id: id });
    hideLoading();
    if (result.success) {
      alert('✅ ' + result.message);
      await openPendingTransferModal();
      await loadStockNew();
      loadPendingTransferCount();
    } else {
      alert('❌ ' + result.message);
    }
  } catch(e) { hideLoading(); alert('❌ ' + e.message); }
}

async function rejectTransfer(id) {
  if (!confirm('ปฏิเสธรายการโอน ' + id + '? ทองจะถูกคืนกลับ Stock (OLD)')) return;
  try {
    showLoading();
    const result = await executeGoogleScript('REJECT_TRANSFER', { id: id });
    hideLoading();
    if (result.success) {
      alert('✅ ' + result.message);
      await openPendingTransferModal();
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
    const result = await executeGoogleScript('STOCK_IN_NEW', { items: JSON.stringify(items), note: note });
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
    const result = await executeGoogleScript('STOCK_OUT_NEW', { items: JSON.stringify(items), note: note });
    hideLoading();
    if (result.success) {
      alert('✅ ' + result.message);
      closeModal('stockOutNewModal');
      await loadStockNew();
    } else { alert('❌ ' + result.message); }
  } catch(e) { hideLoading(); alert('❌ ' + e.message); }
}