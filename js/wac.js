async function loadWAC() {
  try {
    showLoading();

    const [stockOldData, stockNewData, moveOldData, moveNewData] = await Promise.all([
      safeFetch('Stock_Old!A:D'),
      safeFetch('Stock_New!A:D'),
      safeFetch('StockMove_Old!A:J'),
      safeFetch('StockMove_New!A:J')
    ]);

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

    const isToday = (dateStr) => { const d = parseSheetDate(dateStr); return d && d >= today && d <= todayEnd; };

    let oldWeight = oldCarryWeight;
    let oldCost = 0;
    if (moveOldData.length > 1) {
      moveOldData.slice(1).forEach(row => {
        if (!isToday(row[0])) return;
        const g = parseFloat(row[4]) || 0;
        const p = parseFloat(row[6]) || 0;
        if (String(row[5]) === 'IN') { oldWeight += g; oldCost += p; }
        else { oldWeight -= g; oldCost -= p; }
      });
    }

    let newWeight = newCarryWeight;
    let newCost = 0;
    if (moveNewData.length > 1) {
      moveNewData.slice(1).forEach(row => {
        if (!isToday(row[0])) return;
        const g = parseFloat(row[4]) || 0;
        const p = parseFloat(row[6]) || 0;
        if (String(row[5]) === 'IN') { newWeight += g; newCost += p; }
        else { newWeight -= g; newCost -= p; }
      });
    }

    const absNewWeight = Math.abs(newWeight);
    const absNewCost = Math.abs(newCost);
    const totalGoldG = oldWeight + absNewWeight;
    const totalCost = Math.round(oldCost / 1000) * 1000 + Math.round(absNewCost / 1000) * 1000;

    document.getElementById('wacSummaryTable').innerHTML =
      '<tr><td>Stock (NEW)</td><td>' + formatNumber(absNewWeight.toFixed(2)) + ' g</td><td>' + formatNumber(Math.round(absNewCost / 1000) * 1000) + ' LAK</td></tr>' +
      '<tr><td>Stock (OLD)</td><td>' + formatNumber(oldWeight.toFixed(2)) + ' g</td><td>' + formatNumber(Math.round(oldCost / 1000) * 1000) + ' LAK</td></tr>' +
      '<tr style="font-weight:bold;background:rgba(212,175,55,0.1);"><td>ผลรวม</td><td>' + formatNumber(totalGoldG.toFixed(2)) + ' g</td><td>' + formatNumber(totalCost) + ' LAK</td></tr>';

    const wacPerG = totalGoldG > 0 ? totalCost / totalGoldG : 0;
    const wacPerBaht = wacPerG * 15;

    document.getElementById('wacCalcTable').innerHTML =
      '<tr><td>ราคา /g</td><td style="font-weight:bold;color:var(--gold-primary);">' + formatNumber(Math.round(wacPerG / 1000) * 1000) + ' LAK</td></tr>' +
      '<tr><td>ราคา /บาท</td><td style="font-weight:bold;color:var(--gold-primary);">' + formatNumber(Math.round(wacPerBaht / 1000) * 1000) + ' LAK</td></tr>';

    hideLoading();
  } catch(error) {
    console.error('Error loading WAC:', error);
    hideLoading();
  }
}