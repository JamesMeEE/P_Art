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
  } catch(error) {
    console.error('Error loading stock new:', error);
    hideLoading();
  }
}