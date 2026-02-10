let stockNewDateFrom = null;
let stockNewDateTo = null;

function clearStockNewDateFilter() {
  const today = getTodayDateString();
  document.getElementById('stockNewDateFrom').value = today;
  document.getElementById('stockNewDateTo').value = today;
  stockNewDateFrom = today;
  stockNewDateTo = today;
  loadStockNew();
}

async function loadStockNew() {
  try {
    showLoading();

    const [sellData, tradeinData, exchangeData, switchData, freeExData, withdrawData, pricingData] = await Promise.all([
      fetchSheetData('Sells!A:L'),
      fetchSheetData('Tradeins!A:N'),
      fetchSheetData('Exchanges!A:N'),
      fetchSheetData('Switch!A:N'),
      fetchSheetData('FreeExchanges!A:L'),
      fetchSheetData('Withdraws!A:J'),
      fetchSheetData('Pricing!A:B')
    ]);

    const sell1Baht = pricingData.length > 1 ? parseFloat(pricingData[pricingData.length - 1][1]) || 0 : 0;
    const sellPerG = sell1Baht / 15;

    const today = new Date();
    today.setHours(0,0,0,0);
    const filterFrom = stockNewDateFrom ? new Date(stockNewDateFrom) : today;
    const filterTo = stockNewDateTo ? new Date(stockNewDateTo) : today;
    filterFrom.setHours(0,0,0,0);
    filterTo.setHours(23,59,59,999);

    const cfBefore = new Date(filterFrom);
    cfBefore.setMilliseconds(-1);

    const cfQty = {};
    const qtyIn = {};
    const qtyOut = {};
    FIXED_PRODUCTS.forEach(p => { cfQty[p.id] = 0; qtyIn[p.id] = 0; qtyOut[p.id] = 0; });

    const movements = [];

    const collectSellOut = (data, type, itemsCol, dateCol, statusCol) => {
      data.slice(1).forEach(row => {
        if (row[statusCol] !== 'COMPLETED') return;
        const d = new Date(row[dateCol]);
        try {
          const items = JSON.parse(row[itemsCol]);
          if (d <= cfBefore) {
            items.forEach(item => { cfQty[item.productId] = (cfQty[item.productId] || 0) - item.qty; });
            return;
          }
          if (d < filterFrom || d > filterTo) return;
          let totalWeightG = 0;
          items.forEach(item => {
            totalWeightG += getGoldWeight(item.productId) * item.qty;
            qtyOut[item.productId] = (qtyOut[item.productId] || 0) + item.qty;
          });
          const priceOut = totalWeightG * sellPerG;
          movements.push({ id: row[0], type, goldIn: 0, goldOut: totalWeightG, weight: totalWeightG, priceIn: 0, priceOut: Math.round(priceOut / 1000) * 1000, date: row[dateCol] });
        } catch(e) {}
      });
    };

    collectSellOut(sellData, 'SELL', 2, 9, 10);
    collectSellOut(withdrawData, 'WITHDRAW', 2, 7, 8);

    const collectExchangeNewOut = (data, type, dateCol, statusCol) => {
      data.slice(1).forEach(row => {
        if (row[statusCol] !== 'COMPLETED') return;
        const d = new Date(row[dateCol]);
        try {
          const newItems = JSON.parse(row[3]);
          if (d <= cfBefore) {
            newItems.forEach(item => { cfQty[item.productId] = (cfQty[item.productId] || 0) - item.qty; });
            return;
          }
          if (d < filterFrom || d > filterTo) return;
          let totalWeightG = 0;
          newItems.forEach(item => {
            totalWeightG += getGoldWeight(item.productId) * item.qty;
            qtyOut[item.productId] = (qtyOut[item.productId] || 0) + item.qty;
          });
          const priceOut = totalWeightG * sellPerG;
          movements.push({ id: row[0], type, goldIn: 0, goldOut: totalWeightG, weight: totalWeightG, priceIn: 0, priceOut: Math.round(priceOut / 1000) * 1000, date: row[dateCol] });
        } catch(e) {}
      });
    };

    collectExchangeNewOut(tradeinData, 'TRADE-IN', 11, 12);
    collectExchangeNewOut(exchangeData, 'EXCHANGE', 11, 12);
    collectExchangeNewOut(switchData, 'SWITCH', 11, 12);
    collectExchangeNewOut(freeExData, 'FREE-EX', 7, 8);

    movements.sort((a, b) => new Date(a.date) - new Date(b.date));

    let runningCost = 0;
    let runningWeight = 0;
    movements.forEach(m => {
      runningWeight += m.goldIn - m.goldOut;
      runningCost += m.priceIn - m.priceOut;
      m.runningCost = runningCost;
      m.runningWeight = runningWeight;
    });

    const summaryBody = document.getElementById('stockNewSummaryTable');
    summaryBody.innerHTML = FIXED_PRODUCTS.map(p => {
      const cf = cfQty[p.id] || 0;
      const inQ = qtyIn[p.id] || 0;
      const outQ = qtyOut[p.id] || 0;
      const balance = cf + inQ - outQ;
      return `<tr>
        <td>${p.id}</td><td>${p.name}</td>
        <td>${cf}</td>
        <td style="color:#4caf50;">${inQ > 0 ? '+' + inQ : '0'}</td>
        <td style="color:#f44336;">${outQ > 0 ? '-' + outQ : '0'}</td>
        <td style="font-weight:bold;">${balance}</td>
      </tr>`;
    }).join('');

    const absWeight = Math.abs(runningWeight);
    const absCost = Math.abs(runningCost);
    document.getElementById('stockNewGoldG').textContent = formatNumber(absWeight.toFixed(2)) + ' g';
    document.getElementById('stockNewCostValue').textContent = formatNumber(Math.round(absCost / 1000) * 1000) + ' LAK';

    window._stockNewLatest = { goldG: absWeight, cost: absCost };

    const movBody = document.getElementById('stockNewMovementTable');
    if (movements.length === 0) {
      movBody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:40px;">No records</td></tr>';
    } else {
      movBody.innerHTML = movements.slice().reverse().map(m => `<tr>
        <td>${m.id}</td>
        <td><span class="status-badge">${m.type}</span></td>
        <td style="color:#4caf50;">${m.goldIn > 0 ? formatNumber(m.goldIn.toFixed(2)) : '-'}</td>
        <td style="color:#f44336;">${m.goldOut > 0 ? formatNumber(m.goldOut.toFixed(2)) : '-'}</td>
        <td>${formatNumber((m.goldIn || m.goldOut).toFixed(2))} g</td>
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

document.addEventListener('DOMContentLoaded', function() {
  const f = document.getElementById('stockNewDateFrom');
  const t = document.getElementById('stockNewDateTo');
  if (f && t) {
    f.addEventListener('change', function() { stockNewDateFrom = this.value; loadStockNew(); });
    t.addEventListener('change', function() { stockNewDateTo = this.value; loadStockNew(); });
  }
});
