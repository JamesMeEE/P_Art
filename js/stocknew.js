async function loadStockNew() {
  try {
    showLoading();

    const [stockData, sellData, tradeinData, exchangeData, switchData, freeExData, withdrawData, pricingData] = await Promise.all([
      fetchSheetData('Stock_New!A:D'),
      fetchSheetData('Sells!A:L'),
      fetchSheetData('Tradeins!A:N'),
      fetchSheetData('Exchanges!A:N'),
      fetchSheetData('Switch!A:N'),
      fetchSheetData('FreeExchanges!A:J'),
      fetchSheetData('Withdraws!A:J'),
      fetchSheetData('Pricing!A:B')
    ]);

    const sell1Baht = pricingData.length > 1 ? parseFloat(pricingData[pricingData.length - 1][1]) || 0 : 0;
    const sellPerG = sell1Baht / 15;

    let carry = {}, qtyIn = {}, qtyOut = {};
    FIXED_PRODUCTS.forEach(p => { carry[p.id] = 0; qtyIn[p.id] = 0; qtyOut[p.id] = 0; });

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
      return `<tr>
        <td>${p.id}</td><td>${p.name}</td>
        <td>${c}</td>
        <td style="color:#4caf50;">${i > 0 ? '+' + i : '0'}</td>
        <td style="color:#f44336;">${o > 0 ? '-' + o : '0'}</td>
        <td style="font-weight:bold;">${bal}</td>
      </tr>`;
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

    const collectNewGoldOut = (data, type, itemsCol, dateCol, statusCol) => {
      data.slice(1).forEach(row => {
        if (row[statusCol] !== 'COMPLETED') return;
        const d = new Date(row[dateCol]);
        if (d < today || d > todayEnd) return;
        try {
          const items = JSON.parse(row[itemsCol]);
          let totalWeightG = 0;
          items.forEach(item => { totalWeightG += getGoldWeight(item.productId) * item.qty; });
          const priceOut = totalWeightG * sellPerG;
          movements.push({ id: row[0], type, goldIn: 0, goldOut: totalWeightG, priceIn: 0, priceOut: Math.round(priceOut / 1000) * 1000, date: row[dateCol] });
        } catch(e) {}
      });
    };

    collectNewGoldOut(sellData, 'SELL', 2, 9, 10);
    collectNewGoldOut(withdrawData, 'WITHDRAW', 2, 6, 7);

    const collectExNewOut = (data, type, dateCol, statusCol) => {
      data.slice(1).forEach(row => {
        if (row[statusCol] !== 'COMPLETED') return;
        const d = new Date(row[dateCol]);
        if (d < today || d > todayEnd) return;
        try {
          const newItems = JSON.parse(row[3]);
          let totalWeightG = 0;
          newItems.forEach(item => { totalWeightG += getGoldWeight(item.productId) * item.qty; });
          const priceOut = totalWeightG * sellPerG;
          movements.push({ id: row[0], type, goldIn: 0, goldOut: totalWeightG, priceIn: 0, priceOut: Math.round(priceOut / 1000) * 1000, date: row[dateCol] });
        } catch(e) {}
      });
    };

    collectExNewOut(tradeinData, 'TRADE-IN', 11, 12);
    collectExNewOut(exchangeData, 'EXCHANGE', 11, 12);
    collectExNewOut(switchData, 'SWITCH', 11, 12);
    collectExNewOut(freeExData, 'FREE-EX', 7, 8);

    movements.sort((a, b) => new Date(a.date) - new Date(b.date));

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
      movBody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:40px;">No records today</td></tr>';
    } else {
      movBody.innerHTML = movements.slice().reverse().map(m => `<tr>
        <td>${m.id}</td>
        <td><span class="status-badge">${m.type}</span></td>
        <td style="color:#4caf50;">${m.goldIn > 0 ? formatNumber(m.goldIn.toFixed(2)) : '-'}</td>
        <td style="color:#f44336;">${m.goldOut > 0 ? formatNumber(m.goldOut.toFixed(2)) : '-'}</td>
        <td>${formatNumber(Math.abs(m.runningWeight).toFixed(2))} g</td>
        <td>${m.priceIn > 0 ? formatNumber(m.priceIn) : '-'}</td>
        <td>${m.priceOut > 0 ? formatNumber(m.priceOut) : '-'}</td>
        <td>${formatNumber(Math.round(Math.abs(m.runningCost) / 1000) * 1000)}</td>
        <td><button class="btn-action" onclick="viewBillDetail('${m.id}','${m.type}')">📋</button></td>
      </tr>`).join('');
    }

    hideLoading();
  } catch(error) {
    console.error('Error loading stock new:', error);
    hideLoading();
  }
}