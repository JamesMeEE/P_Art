async function loadWAC() {
  try {
    showLoading();

    await callAppsScript('INIT_STOCK');

    const [stockOldData, stockNewData, pricingData, buybackData, tradeinData, exchangeData, switchData, freeExData, sellData, withdrawData] = await Promise.all([
      fetchSheetData('Stock_Old!A:D'),
      fetchSheetData('Stock_New!A:D'),
      fetchSheetData('Pricing!A:B'),
      fetchSheetData('Buybacks!A:L'),
      fetchSheetData('Tradeins!A:N'),
      fetchSheetData('Exchanges!A:N'),
      fetchSheetData('Switch!A:N'),
      fetchSheetData('FreeExchanges!A:J'),
      fetchSheetData('Sells!A:L'),
      fetchSheetData('Withdraws!A:J')
    ]);

    const sell1Baht = pricingData.length > 1 ? parseFloat(pricingData[pricingData.length - 1][1]) || 0 : 0;
    const buyback1Baht = sell1Baht - 530000;
    const sellPerG = sell1Baht / 15;
    const buybackPerG = buyback1Baht / 15;

    let oldCarry = {};
    if (stockOldData.length > 1) {
      try { oldCarry = JSON.parse(stockOldData[stockOldData.length - 1][1] || '{}'); } catch(e) {}
    }
    let oldCarryWeight = 0;
    FIXED_PRODUCTS.forEach(p => { oldCarryWeight += (oldCarry[p.id] || 0) * getGoldWeight(p.id); });

    let newCarry = {};
    if (stockNewData.length > 1) {
      try { newCarry = JSON.parse(stockNewData[stockNewData.length - 1][1] || '{}'); } catch(e) {}
    }
    let newCarryWeight = 0;
    FIXED_PRODUCTS.forEach(p => { newCarryWeight += (newCarry[p.id] || 0) * getGoldWeight(p.id); });

    const today = new Date();
    today.setHours(0,0,0,0);
    const todayEnd = new Date(today);
    todayEnd.setHours(23,59,59,999);

    const isToday = (dateStr) => { const d = new Date(dateStr); return d >= today && d <= todayEnd; };

    let oldWeight = oldCarryWeight;
    let oldCost = 0;

    const addOldIn = (data, dateCol, statusCol, useBuyback) => {
      data.slice(1).forEach(row => {
        if (row[statusCol] !== 'COMPLETED' && row[statusCol] !== 'PARTIAL') return;
        if (!isToday(row[dateCol])) return;
        try {
          const items = JSON.parse(row[2]);
          let w = 0;
          items.forEach(item => { w += getGoldWeight(item.productId) * item.qty; });
          oldWeight += w;
          oldCost += w * (useBuyback ? buybackPerG : sellPerG);
        } catch(e) {}
      });
    };

    addOldIn(buybackData, 9, 10, true);
    addOldIn(tradeinData, 11, 12, false);
    addOldIn(exchangeData, 11, 12, false);
    addOldIn(switchData, 11, 12, false);
    addOldIn(freeExData, 7, 8, false);

    let newWeight = newCarryWeight;
    let newCost = 0;

    const addNewOut = (data, itemsCol, dateCol, statusCol) => {
      data.slice(1).forEach(row => {
        if (row[statusCol] !== 'COMPLETED') return;
        if (!isToday(row[dateCol])) return;
        try {
          const items = JSON.parse(row[itemsCol]);
          let w = 0;
          items.forEach(item => { w += getGoldWeight(item.productId) * item.qty; });
          newWeight -= w;
          newCost += w * sellPerG;
        } catch(e) {}
      });
    };

    addNewOut(sellData, 2, 9, 10);
    addNewOut(withdrawData, 2, 6, 7);
    const addExNewOut = (data, dateCol, statusCol) => {
      data.slice(1).forEach(row => {
        if (row[statusCol] !== 'COMPLETED') return;
        if (!isToday(row[dateCol])) return;
        try {
          const items = JSON.parse(row[3]);
          let w = 0;
          items.forEach(item => { w += getGoldWeight(item.productId) * item.qty; });
          newWeight -= w;
          newCost += w * sellPerG;
        } catch(e) {}
      });
    };
    addExNewOut(tradeinData, 11, 12);
    addExNewOut(exchangeData, 11, 12);
    addExNewOut(switchData, 11, 12);
    addExNewOut(freeExData, 7, 8);

    const absNewWeight = Math.abs(newWeight);
    const absNewCost = Math.abs(newCost);
    const totalGoldG = oldWeight + absNewWeight;
    const totalCost = Math.round(oldCost / 1000) * 1000 + Math.round(absNewCost / 1000) * 1000;

    document.getElementById('wacSummaryTable').innerHTML = `
      <tr><td>Stock (NEW)</td><td>${formatNumber(absNewWeight.toFixed(2))} g</td><td>${formatNumber(Math.round(absNewCost / 1000) * 1000)} LAK</td></tr>
      <tr><td>Stock (OLD)</td><td>${formatNumber(oldWeight.toFixed(2))} g</td><td>${formatNumber(Math.round(oldCost / 1000) * 1000)} LAK</td></tr>
      <tr style="font-weight:bold;background:rgba(212,175,55,0.1);"><td>ผลรวม</td><td>${formatNumber(totalGoldG.toFixed(2))} g</td><td>${formatNumber(totalCost)} LAK</td></tr>
    `;

    const wacPerG = totalGoldG > 0 ? totalCost / totalGoldG : 0;
    const wacPerBaht = wacPerG * 15;

    document.getElementById('wacCalcTable').innerHTML = `
      <tr><td>ราคา /g</td><td style="font-weight:bold;color:var(--gold-primary);">${formatNumber(Math.round(wacPerG / 1000) * 1000)} LAK</td></tr>
      <tr><td>ราคา /บาท</td><td style="font-weight:bold;color:var(--gold-primary);">${formatNumber(Math.round(wacPerBaht / 1000) * 1000)} LAK</td></tr>
    `;

    hideLoading();
  } catch(error) {
    console.error('Error loading WAC:', error);
    hideLoading();
  }
}