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

    const summaryBody = document.getElementById('stockOldSummaryTable');
    summaryBody.innerHTML = FIXED_PRODUCTS.map(p => {
      const c = carry[p.id] || 0;
      const i = qtyIn[p.id] || 0;
      const o = qtyOut[p.id] || 0;
      const bal = c + i - o;
      return '<tr>' +
        '<td>' + p.id + '</td><td>' + p.name + '</td>' +
        '<td>' + c + '</td>' +
        '<td style="color:#4caf50;">' + (i > 0 ? '+' + i : '0') + '</td>' +
        '<td style="color:#f44336;">' + (o > 0 ? '-' + o : '0') + '</td>' +
        '<td style="font-weight:bold;">' + bal + '</td>' +
        '</tr>';
    }).join('');

    let carryWeightG = 0;
    FIXED_PRODUCTS.forEach(p => {
      carryWeightG += (carry[p.id] || 0) * getGoldWeight(p.id);
    });

    const today = new Date();
    today.setHours(0,0,0,0);
    const todayEnd = new Date(today);
    todayEnd.setHours(23,59,59,999);

    const movements = [];
    if (moveData.length > 1) {
      moveData.slice(1).forEach(row => {
        const d = parseSheetDate(row[0]);
        if (!d || d < today || d > todayEnd) return;
        const goldG = parseFloat(row[4]) || 0;
        const price = parseFloat(row[6]) || 0;
        const direction = String(row[5] || '');
        const wacG = parseFloat(row[8]) || 0;
        const wacBaht = parseFloat(row[9]) || 0;
        movements.push({
          id: row[1],
          type: row[2],
          goldIn: direction === 'IN' ? goldG : 0,
          goldOut: direction === 'OUT' ? goldG : 0,
          priceIn: direction === 'IN' ? price : 0,
          priceOut: direction === 'OUT' ? price : 0,
          wacG: wacG,
          wacBaht: wacBaht,
          date: row[0]
        });
      });
    }

    let runningWeight = carryWeightG;
    let runningCost = 0;
    movements.forEach(m => {
      runningWeight += m.goldIn - m.goldOut;
      runningCost += m.priceIn - m.priceOut;
      m.runningWeight = runningWeight;
      m.runningCost = runningCost;
    });

    const latestWeight = movements.length > 0 ? movements[movements.length - 1].runningWeight : carryWeightG;
    const latestCost = movements.length > 0 ? movements[movements.length - 1].runningCost : 0;

    document.getElementById('stockOldGoldG').textContent = formatNumber(latestWeight.toFixed(2)) + ' g';
    document.getElementById('stockOldCostValue').textContent = formatNumber(Math.round(latestCost / 1000) * 1000) + ' LAK';

    window._stockOldLatest = { goldG: latestWeight, cost: latestCost };

    const movBody = document.getElementById('stockOldMovementTable');
    if (movements.length === 0) {
      movBody.innerHTML = '<tr><td colspan="11" style="text-align:center;padding:40px;">No records today</td></tr>';
    } else {
      movBody.innerHTML = movements.map(m => '<tr>' +
        '<td>' + m.id + '</td>' +
        '<td><span class="status-badge">' + m.type + '</span></td>' +
        '<td style="color:#4caf50;">' + (m.goldIn > 0 ? formatNumber(m.goldIn.toFixed(2)) : '-') + '</td>' +
        '<td style="color:#f44336;">' + (m.goldOut > 0 ? formatNumber(m.goldOut.toFixed(2)) : '-') + '</td>' +
        '<td>' + formatNumber(m.runningWeight.toFixed(2)) + ' g</td>' +
        '<td>' + (m.priceIn > 0 ? formatNumber(m.priceIn) : '-') + '</td>' +
        '<td>' + (m.priceOut > 0 ? formatNumber(m.priceOut) : '-') + '</td>' +
        '<td>' + formatNumber(Math.round(m.runningCost / 1000) * 1000) + '</td>' +
        '<td>' + (m.wacG > 0 ? formatNumber(m.wacG) : '-') + '</td>' +
        '<td>' + (m.wacBaht > 0 ? formatNumber(m.wacBaht) : '-') + '</td>' +
        '<td><button class="btn-action" onclick="viewBillDetail(\'' + m.id + '\',\'' + m.type + '\')">📋</button></td>' +
        '</tr>').join('');
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
    let sheetName;
    if (type === 'BUYBACK') sheetName = 'Buybacks!A:L';
    else if (type === 'TRADE-IN') sheetName = 'Tradeins!A:N';
    else if (type === 'EXCHANGE') sheetName = 'Exchanges!A:N';
    else if (type === 'SWITCH') sheetName = 'Switch!A:N';
    else if (type === 'FREE-EX') sheetName = 'FreeExchanges!A:J';
    else if (type === 'SELL') sheetName = 'Sells!A:L';
    else if (type === 'WITHDRAW') sheetName = 'Withdraws!A:J';
    else sheetName = 'Inventory!A:R';

    const data = await fetchSheetData(sheetName);
    const row = data.slice(1).find(r => r[0] === id);
    hideLoading();

    if (!row) { alert('ไม่พบข้อมูลบิล: ' + id); return; }

    const formatItems = (json) => {
      try {
        return JSON.parse(json).map(i => (FIXED_PRODUCTS.find(p => p.id === i.productId) || {}).name + ' x' + i.qty).join(', ');
      } catch(e) { return json; }
    };

    let detail = 'Transaction ID: ' + row[0] + '\nPhone: ' + (row[1] || '-') + '\n';
    if (type === 'BUYBACK') {
      detail += 'Items: ' + formatItems(row[2]) + '\nPrice: ' + formatNumber(row[3]) + ' LAK\nFee: ' + formatNumber(row[5]) + ' LAK\nTotal: ' + formatNumber(row[6]) + ' LAK\nStatus: ' + row[10];
    } else if (type === 'TRADE-IN' || type === 'EXCHANGE' || type === 'SWITCH') {
      detail += 'Old Gold: ' + formatItems(row[2]) + '\nNew Gold: ' + formatItems(row[3]) + '\nTotal: ' + formatNumber(row[6]) + ' LAK\nStatus: ' + row[12];
    } else if (type === 'FREE-EX') {
      detail += 'Old Gold: ' + formatItems(row[2]) + '\nNew Gold: ' + formatItems(row[3]) + '\nPremium: ' + formatNumber(row[4]) + ' LAK\nStatus: ' + row[8];
    } else if (type === 'SELL') {
      detail += 'Items: ' + formatItems(row[2]) + '\nTotal: ' + formatNumber(row[3]) + ' LAK\nStatus: ' + row[10];
    } else if (type === 'WITHDRAW') {
      detail += 'Items: ' + formatItems(row[2]) + '\nTotal: ' + formatNumber(row[4]) + ' LAK\nStatus: ' + row[7];
    }
    alert(detail);
  } catch(e) {
    hideLoading();
    alert('Error: ' + e.message);
  }
}