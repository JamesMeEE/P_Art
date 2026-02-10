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

    const summaryBody = document.getElementById('stockNewSummaryTable');
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
        const d = new Date(row[0]);
        if (isNaN(d.getTime()) || d < today || d > todayEnd) return;
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

    const latestWeight = movements.length > 0 ? Math.abs(movements[movements.length - 1].runningWeight) : carryWeightG;
    const latestCost = movements.length > 0 ? Math.abs(movements[movements.length - 1].runningCost) : 0;

    document.getElementById('stockNewGoldG').textContent = formatNumber(latestWeight.toFixed(2)) + ' g';
    document.getElementById('stockNewCostValue').textContent = formatNumber(Math.round(latestCost / 1000) * 1000) + ' LAK';

    window._stockNewLatest = { goldG: latestWeight, cost: latestCost };

    const movBody = document.getElementById('stockNewMovementTable');
    if (movements.length === 0) {
      movBody.innerHTML = '<tr><td colspan="11" style="text-align:center;padding:40px;">No records today</td></tr>';
    } else {
      movBody.innerHTML = movements.map(m => '<tr>' +
        '<td>' + m.id + '</td>' +
        '<td><span class="status-badge">' + m.type + '</span></td>' +
        '<td style="color:#4caf50;">' + (m.goldIn > 0 ? formatNumber(m.goldIn.toFixed(2)) : '-') + '</td>' +
        '<td style="color:#f44336;">' + (m.goldOut > 0 ? formatNumber(m.goldOut.toFixed(2)) : '-') + '</td>' +
        '<td>' + formatNumber(Math.abs(m.runningWeight).toFixed(2)) + ' g</td>' +
        '<td>' + (m.priceIn > 0 ? formatNumber(m.priceIn) : '-') + '</td>' +
        '<td>' + (m.priceOut > 0 ? formatNumber(m.priceOut) : '-') + '</td>' +
        '<td>' + formatNumber(Math.round(Math.abs(m.runningCost) / 1000) * 1000) + '</td>' +
        '<td>' + (m.wacG > 0 ? formatNumber(m.wacG) : '-') + '</td>' +
        '<td>' + (m.wacBaht > 0 ? formatNumber(m.wacBaht) : '-') + '</td>' +
        '<td><button class="btn-action" onclick="viewBillDetail(\'' + m.id + '\',\'' + m.type + '\')">📋</button></td>' +
        '</tr>').join('');
    }

    hideLoading();
  } catch(error) {
    console.error('Error loading stock new:', error);
    hideLoading();
  }
}