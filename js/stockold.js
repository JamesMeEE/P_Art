let stockOldDateFrom = null;
let stockOldDateTo = null;

function clearStockOldDateFilter() {
  const today = getTodayDateString();
  document.getElementById('stockOldDateFrom').value = today;
  document.getElementById('stockOldDateTo').value = today;
  stockOldDateFrom = today;
  stockOldDateTo = today;
  loadStockOld();
}

async function loadStockOld() {
  try {
    showLoading();

    const [buybackData, tradeinData, exchangeData, switchData, freeExData, pricingData] = await Promise.all([
      fetchSheetData('Buybacks!A:L'),
      fetchSheetData('Tradeins!A:N'),
      fetchSheetData('Exchanges!A:N'),
      fetchSheetData('Switch!A:N'),
      fetchSheetData('FreeExchanges!A:L'),
      fetchSheetData('Pricing!A:B')
    ]);

    const sell1Baht = pricingData.length > 1 ? parseFloat(pricingData[pricingData.length - 1][1]) || 0 : 0;
    const buyback1Baht = sell1Baht - 530000;
    const sellPerG = sell1Baht / 15;
    const buybackPerG = buyback1Baht / 15;

    const today = new Date();
    today.setHours(0,0,0,0);
    const filterFrom = stockOldDateFrom ? new Date(stockOldDateFrom) : today;
    const filterTo = stockOldDateTo ? new Date(stockOldDateTo) : today;
    filterFrom.setHours(0,0,0,0);
    filterTo.setHours(23,59,59,999);

    const cfBefore = new Date(filterFrom);
    cfBefore.setMilliseconds(-1);

    const cfQty = {};
    const qtyIn = {};
    const qtyOut = {};
    FIXED_PRODUCTS.forEach(p => { cfQty[p.id] = 0; qtyIn[p.id] = 0; qtyOut[p.id] = 0; });

    const movements = [];

    const collectOldGoldIn = (data, type, dateCol, statusCol, useBuybackPrice) => {
      data.slice(1).forEach(row => {
        if (row[statusCol] !== 'COMPLETED' && row[statusCol] !== 'PARTIAL') return;
        const d = new Date(row[dateCol]);
        try {
          const items = JSON.parse(row[2]);
          if (d <= cfBefore) {
            items.forEach(item => { cfQty[item.productId] = (cfQty[item.productId] || 0) + item.qty; });
            return;
          }
          if (d < filterFrom || d > filterTo) return;
          let totalWeightG = 0;
          items.forEach(item => {
            totalWeightG += getGoldWeight(item.productId) * item.qty;
            qtyIn[item.productId] = (qtyIn[item.productId] || 0) + item.qty;
          });
          const perG = useBuybackPrice ? buybackPerG : sellPerG;
          const priceIn = totalWeightG * perG;
          movements.push({ id: row[0], type, goldIn: totalWeightG, goldOut: 0, weight: totalWeightG, priceIn: Math.round(priceIn / 1000) * 1000, priceOut: 0, date: row[dateCol] });
        } catch(e) {}
      });
    };

    collectOldGoldIn(buybackData, 'BUYBACK', 9, 10, true);
    collectOldGoldIn(tradeinData, 'TRADE-IN', 11, 12, false);
    collectOldGoldIn(exchangeData, 'EXCHANGE', 11, 12, false);
    collectOldGoldIn(switchData, 'SWITCH', 11, 12, false);
    collectOldGoldIn(freeExData, 'FREE-EX', 7, 8, false);

    movements.sort((a, b) => new Date(a.date) - new Date(b.date));

    let runningCost = 0;
    let runningWeight = 0;
    movements.forEach(m => {
      runningWeight += m.goldIn - m.goldOut;
      runningCost += m.priceIn - m.priceOut;
      m.runningCost = runningCost;
      m.runningWeight = runningWeight;
    });

    const summaryBody = document.getElementById('stockOldSummaryTable');
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

    document.getElementById('stockOldGoldG').textContent = formatNumber(runningWeight.toFixed(2)) + ' g';
    document.getElementById('stockOldCostValue').textContent = formatNumber(Math.round(runningCost / 1000) * 1000) + ' LAK';

    window._stockOldLatest = { goldG: runningWeight, cost: runningCost };

    const movBody = document.getElementById('stockOldMovementTable');
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
        <td>${formatNumber(Math.round(m.runningCost / 1000) * 1000)}</td>
        <td><button class="btn-action" onclick="viewBillDetail('${m.id}','${m.type}')">📋</button></td>
      </tr>`).join('');
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
    else if (type === 'FREE-EX') sheetName = 'FreeExchanges!A:L';
    else if (type === 'SELL') sheetName = 'Sells!A:L';
    else if (type === 'WITHDRAW') sheetName = 'Withdraws!A:J';
    else sheetName = 'Inventory!A:R';

    const data = await fetchSheetData(sheetName);
    const row = data.slice(1).find(r => r[0] === id);
    hideLoading();

    if (!row) { alert('ไม่พบข้อมูลบิล: ' + id); return; }

    let detail = 'Transaction ID: ' + row[0] + '\nPhone: ' + (row[1] || '-') + '\n';
    if (type === 'BUYBACK') {
      detail += 'Items: ' + formatItemsForDisplay(row[2]) + '\nPrice: ' + formatNumber(row[3]) + ' LAK\nFee: ' + formatNumber(row[5]) + ' LAK\nTotal: ' + formatNumber(row[6]) + ' LAK\nStatus: ' + row[10];
    } else if (type === 'TRADE-IN' || type === 'EXCHANGE' || type === 'SWITCH') {
      detail += 'Old Gold: ' + formatItemsForDisplay(row[2]) + '\nNew Gold: ' + formatItemsForDisplay(row[3]) + '\nTotal: ' + formatNumber(row[6]) + ' LAK\nStatus: ' + row[12];
    } else if (type === 'FREE-EX') {
      detail += 'Old Gold: ' + formatItemsForDisplay(row[2]) + '\nNew Gold: ' + formatItemsForDisplay(row[3]) + '\nPremium: ' + formatNumber(row[4]) + ' LAK\nStatus: ' + row[8];
    } else if (type === 'SELL') {
      detail += 'Items: ' + formatItemsForDisplay(row[2]) + '\nTotal: ' + formatNumber(row[3]) + ' LAK\nStatus: ' + row[10];
    } else if (type === 'WITHDRAW') {
      detail += 'Items: ' + formatItemsForDisplay(row[2]) + '\nTotal: ' + formatNumber(row[4]) + ' LAK\nStatus: ' + row[8];
    } else {
      detail += 'Type: ' + row[1] + '\nProduct: ' + row[3] + '\nQty IN: ' + row[7] + '\nQty OUT: ' + row[8];
    }
    alert(detail);
  } catch(e) {
    hideLoading();
    alert('Error: ' + e.message);
  }
}

document.addEventListener('DOMContentLoaded', function() {
  const f = document.getElementById('stockOldDateFrom');
  const t = document.getElementById('stockOldDateTo');
  if (f && t) {
    f.addEventListener('change', function() { stockOldDateFrom = this.value; loadStockOld(); });
    t.addEventListener('change', function() { stockOldDateTo = this.value; loadStockOld(); });
  }
});
