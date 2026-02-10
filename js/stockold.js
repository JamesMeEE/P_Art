async function safeFetch(range) {
  try {
    return await fetchSheetData(range);
  } catch(e) {
    console.warn('Sheet not found: ' + range);
    return [];
  }
}

async function loadStockOld() {
  try {
    showLoading();

    const [stockData, moveData] = await Promise.all([
      safeFetch('Stock_Old!A:D'),
      safeFetch('StockMove_Old!A:J')
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

    document.getElementById('stockOldSummaryTable').innerHTML = FIXED_PRODUCTS.map(p => {
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

    const latestW = todayMovements.length > 0 ? todayMovements[todayMovements.length - 1].w : prevW;
    const latestC = todayMovements.length > 0 ? todayMovements[todayMovements.length - 1].c : prevC;

    document.getElementById('stockOldGoldG').textContent = formatWeight(latestW) + ' g';
    document.getElementById('stockOldCostValue').textContent = formatNumber(Math.round(latestC / 1000) * 1000) + ' LAK';
    window._stockOldLatest = { goldG: latestW, cost: latestC };

    const movBody = document.getElementById('stockOldMovementTable');
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
  } catch(error) {
    console.error('Error loading stock old:', error);
    hideLoading();
  }
}

async function viewBillDetail(id, type) {
  try {
    showLoading();
    let sheetRange = null;

    if (type === 'BUYBACK') sheetRange = 'Buybacks!A:L';
    else if (type === 'TRADE-IN') sheetRange = 'Tradeins!A:N';
    else if (type === 'EXCHANGE') sheetRange = 'Exchanges!A:N';
    else if (type === 'SWITCH') sheetRange = 'Switches!A:N';
    else if (type === 'FREE-EX') sheetRange = 'FreeExchanges!A:J';
    else if (type === 'SELL') sheetRange = 'Sells!A:L';
    else if (type === 'WITHDRAW') sheetRange = 'Withdraws!A:J';

    const fetches = [safeFetch('StockMove_Old!A:J'), safeFetch('StockMove_New!A:J')];
    if (sheetRange) fetches.push(fetchSheetData(sheetRange));
    const results = await Promise.all(fetches);
    const moveOld = results[0];
    const moveNew = results[1];
    const txData = results.length > 2 ? results[2] : [];

    const findMove = (data) => {
      if (data.length <= 1) return null;
      for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][1] === id && data[i][2] === type) return data[i];
      }
      for (let i = data.length - 1; i >= 1; i--) {
        if (data[i][1] === id) return data[i];
      }
      return null;
    };
    let moveRow = findMove(moveOld) || findMove(moveNew);

    const row = txData.length > 1 ? txData.slice(1).find(r => r[0] === id) : null;
    hideLoading();

    const fmtItems = (json) => {
      try {
        return JSON.parse(json).map(i => {
          const p = FIXED_PRODUCTS.find(x => x.id === i.productId);
          return '<tr><td>' + (p ? p.name : i.productId) + '</td><td style="text-align:right;">' + i.qty + ' ชิ้น</td></tr>';
        }).join('');
      } catch(e) { return '<tr><td colspan="2">' + json + '</td></tr>'; }
    };

    let html = '<div style="margin-bottom:15px;"><span style="font-size:12px;color:var(--text-secondary);">Transaction ID</span><br><span style="font-size:18px;font-weight:bold;color:var(--gold-primary);">' + id + '</span></div>';
    html += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:15px;">';
    html += '<div><span style="color:var(--text-secondary);font-size:12px;">ประเภท</span><br><span class="status-badge">' + type + '</span></div>';

    if (row) {
      if (type === 'BUYBACK') {
        html += '<div><span style="color:var(--text-secondary);font-size:12px;">สถานะ</span><br><span class="status-badge">' + (row[10] || '-') + '</span></div></div>';
        html += '<div style="margin-bottom:15px;"><table class="data-table" style="width:100%;"><thead><tr><th>สินค้า</th><th>จำนวน</th></tr></thead><tbody>' + fmtItems(row[2]) + '</tbody></table></div>';
        html += '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;">';
        html += '<div class="stat-card" style="padding:10px;"><div style="color:var(--text-secondary);font-size:11px;">ราคา</div><div style="font-weight:bold;">' + formatNumber(row[3]) + '</div></div>';
        html += '<div class="stat-card" style="padding:10px;"><div style="color:var(--text-secondary);font-size:11px;">ค่าธรรมเนียม</div><div style="font-weight:bold;">' + formatNumber(row[5]) + '</div></div>';
        html += '<div class="stat-card" style="padding:10px;"><div style="color:var(--text-secondary);font-size:11px;">รวม</div><div style="font-weight:bold;color:var(--gold-primary);">' + formatNumber(row[6]) + '</div></div></div>';
      } else if (type === 'TRADE-IN' || type === 'EXCHANGE' || type === 'SWITCH') {
        html += '<div><span style="color:var(--text-secondary);font-size:12px;">สถานะ</span><br><span class="status-badge">' + (row[12] || '-') + '</span></div></div>';
        html += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:15px;margin-bottom:15px;">';
        html += '<div><span style="color:#ff9800;font-size:12px;">🥇 ทองเก่า (รับเข้า)</span><table class="data-table" style="width:100%;margin-top:5px;"><tbody>' + fmtItems(row[2]) + '</tbody></table></div>';
        html += '<div><span style="color:#2196f3;font-size:12px;">💎 ทองใหม่ (จ่ายออก)</span><table class="data-table" style="width:100%;margin-top:5px;"><tbody>' + fmtItems(row[3]) + '</tbody></table></div></div>';
        html += '<div class="stat-card" style="padding:10px;text-align:center;"><div style="color:var(--text-secondary);font-size:11px;">ยอดรวม</div><div style="font-weight:bold;font-size:18px;color:var(--gold-primary);">' + formatNumber(row[6]) + ' LAK</div></div>';
      } else if (type === 'FREE-EX') {
        html += '<div><span style="color:var(--text-secondary);font-size:12px;">สถานะ</span><br><span class="status-badge">' + (row[8] || '-') + '</span></div></div>';
        html += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:15px;margin-bottom:15px;">';
        html += '<div><span style="color:#ff9800;font-size:12px;">🥇 ทองเก่า</span><table class="data-table" style="width:100%;margin-top:5px;"><tbody>' + fmtItems(row[2]) + '</tbody></table></div>';
        html += '<div><span style="color:#2196f3;font-size:12px;">💎 ทองใหม่</span><table class="data-table" style="width:100%;margin-top:5px;"><tbody>' + fmtItems(row[3]) + '</tbody></table></div></div>';
        html += '<div class="stat-card" style="padding:10px;text-align:center;"><div style="color:var(--text-secondary);font-size:11px;">Premium</div><div style="font-weight:bold;font-size:18px;color:var(--gold-primary);">' + formatNumber(row[4]) + ' LAK</div></div>';
      } else if (type === 'SELL') {
        html += '<div><span style="color:var(--text-secondary);font-size:12px;">สถานะ</span><br><span class="status-badge">' + (row[10] || '-') + '</span></div></div>';
        html += '<div style="margin-bottom:15px;"><table class="data-table" style="width:100%;"><thead><tr><th>สินค้า</th><th>จำนวน</th></tr></thead><tbody>' + fmtItems(row[2]) + '</tbody></table></div>';
        html += '<div class="stat-card" style="padding:10px;text-align:center;"><div style="color:var(--text-secondary);font-size:11px;">ยอดรวม</div><div style="font-weight:bold;font-size:18px;color:var(--gold-primary);">' + formatNumber(row[3]) + ' LAK</div></div>';
      } else if (type === 'WITHDRAW') {
        html += '<div><span style="color:var(--text-secondary);font-size:12px;">สถานะ</span><br><span class="status-badge">' + (row[7] || '-') + '</span></div></div>';
        html += '<div style="margin-bottom:15px;"><table class="data-table" style="width:100%;"><thead><tr><th>สินค้า</th><th>จำนวน</th></tr></thead><tbody>' + fmtItems(row[2]) + '</tbody></table></div>';
        html += '<div class="stat-card" style="padding:10px;text-align:center;"><div style="color:var(--text-secondary);font-size:11px;">ยอดรวม</div><div style="font-weight:bold;font-size:18px;color:var(--gold-primary);">' + formatNumber(row[4]) + ' LAK</div></div>';
      }
    } else if (moveRow) {
      html += '<div><span style="color:var(--text-secondary);font-size:12px;">Direction</span><br><span class="status-badge">' + (moveRow[5] || '') + '</span></div></div>';
      html += '<div style="margin-bottom:15px;"><table class="data-table" style="width:100%;"><thead><tr><th>สินค้า</th><th>จำนวน</th></tr></thead><tbody>' + fmtItems(moveRow[3]) + '</tbody></table></div>';
      html += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">';
      html += '<div class="stat-card" style="padding:10px;"><div style="color:var(--text-secondary);font-size:11px;">น้ำหนัก</div><div style="font-weight:bold;">' + formatWeight(parseFloat(moveRow[4] || 0)) + ' g</div></div>';
      html += '<div class="stat-card" style="padding:10px;"><div style="color:var(--text-secondary);font-size:11px;">มูลค่า</div><div style="font-weight:bold;color:var(--gold-primary);">' + formatNumber(moveRow[6]) + ' LAK</div></div></div>';
    } else {
      html += '<div></div></div><p style="text-align:center;color:var(--text-secondary);">ไม่พบข้อมูลเพิ่มเติม</p>';
    }

    if (moveRow) {
      const wG = parseFloat(moveRow[8]) || 0;
      const wB = parseFloat(moveRow[9]) || 0;
      if (wG > 0 || wB > 0) {
        html += '<div style="margin-top:15px;padding:12px;background:rgba(212,175,55,0.08);border-radius:8px;border:1px solid rgba(212,175,55,0.2);">';
        html += '<div style="font-size:12px;color:var(--gold-primary);margin-bottom:8px;font-weight:bold;">📊 WAC ณ เวลาทำรายการ</div>';
        html += '<div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;">';
        html += '<div><span style="color:var(--text-secondary);font-size:11px;">WAC/g</span><br><span style="font-weight:bold;">' + formatNumber(wG) + ' LAK</span></div>';
        html += '<div><span style="color:var(--text-secondary);font-size:11px;">WAC/บาท</span><br><span style="font-weight:bold;">' + formatNumber(wB) + ' LAK</span></div>';
        html += '</div></div>';
      }
    }

    showBillModal(id, type, html);
  } catch(e) {
    hideLoading();
    showBillModal('Error', '', '<p style="color:#f44336;">' + e.message + '</p>');
  }
}

function showBillModal(id, type, contentHtml) {
  let modal = document.getElementById('billDetailModal');
  if (!modal) {
    modal = document.createElement('div');
    modal.id = 'billDetailModal';
    modal.className = 'modal';
    modal.innerHTML = '<div class="modal-content" style="max-width:520px;"><div class="modal-header"><h3 id="billDetailTitle"></h3><button class="close-btn" onclick="closeModal(\'billDetailModal\')">&times;</button></div><div class="modal-body" id="billDetailBody" style="max-height:70vh;overflow-y:auto;"></div><div class="modal-footer"><button class="btn-secondary" onclick="closeModal(\'billDetailModal\')">ปิด</button></div></div>';
    document.body.appendChild(modal);
  }
  document.getElementById('billDetailTitle').textContent = '📋 ' + type + ' - ' + id;
  document.getElementById('billDetailBody').innerHTML = contentHtml;
  openModal('billDetailModal');
}