async function loadWAC() {
  try {
    showLoading();

    const [stockOldData, stockNewData, moveOldResult, moveNewResult] = await Promise.all([
      safeFetch('Stock_Old!A:D'),
      safeFetch('Stock_New!A:D'),
      callAppsScript('GET_STOCK_MOVES', { sheet: 'StockMove_Old' }),
      callAppsScript('GET_STOCK_MOVES', { sheet: 'StockMove_New' })
    ]);

    let oldCarry = {};
    if (stockOldData.length > 1) {
      try { oldCarry = JSON.parse(stockOldData[stockOldData.length - 1][1] || '{}'); } catch(e) {}
    }
    let oldCarryW = 0;
    FIXED_PRODUCTS.forEach(p => { oldCarryW += (oldCarry[p.id] || 0) * getGoldWeight(p.id); });

    let newCarry = {};
    if (stockNewData.length > 1) {
      try { newCarry = JSON.parse(stockNewData[stockNewData.length - 1][1] || '{}'); } catch(e) {}
    }
    let newCarryW = 0;
    FIXED_PRODUCTS.forEach(p => { newCarryW += (newCarry[p.id] || 0) * getGoldWeight(p.id); });

    var oData = moveOldResult.data || { prevW: 0, prevC: 0, moves: [] };
    var nData = moveNewResult.data || { prevW: 0, prevC: 0, moves: [] };

    var oldCumW = oData.prevW, oldCumC = oData.prevC;
    (oData.moves || []).forEach(function(m) {
      if (m.dir === 'IN') { oldCumW += m.goldG; oldCumC += m.price; }
      else { oldCumW -= m.goldG; oldCumC -= m.price; }
    });

    var newCumW = nData.prevW, newCumC = nData.prevC;
    (nData.moves || []).forEach(function(m) {
      if (m.dir === 'IN') { newCumW += m.goldG; newCumC += m.price; }
      else { newCumW -= m.goldG; newCumC -= m.price; }
    });

    var totalOldW = oldCarryW + oldCumW;
    var absNewW = Math.abs(newCarryW + newCumW);
    var absNewC = Math.abs(newCumC);
    var totalGoldG = totalOldW + absNewW;
    var totalCost = Math.round(oldCumC / 1000) * 1000 + Math.round(absNewC / 1000) * 1000;

    document.getElementById('wacSummaryTable').innerHTML =
      '<tr><td>Stock (NEW)</td><td>' + formatWeight(absNewW) + ' g</td><td>' + formatNumber(Math.round(absNewC / 1000) * 1000) + ' LAK</td></tr>' +
      '<tr><td>Stock (OLD)</td><td>' + formatWeight(totalOldW) + ' g</td><td>' + formatNumber(Math.round(oldCumC / 1000) * 1000) + ' LAK</td></tr>' +
      '<tr style="font-weight:bold;background:rgba(212,175,55,0.1);"><td>ผลรวม</td><td>' + formatWeight(totalGoldG) + ' g</td><td>' + formatNumber(totalCost) + ' LAK</td></tr>';

    var wacPerG = totalGoldG > 0 ? totalCost / totalGoldG : 0;
    var wacPerBaht = wacPerG * 15;

    document.getElementById('wacCalcTable').innerHTML =
      '<tr><td>ราคา /g</td><td style="font-weight:bold;color:var(--gold-primary);">' + formatNumber(Math.round(wacPerG / 1000) * 1000) + ' LAK</td></tr>' +
      '<tr><td>ราคา /บาท</td><td style="font-weight:bold;color:var(--gold-primary);">' + formatNumber(Math.round(wacPerBaht / 1000) * 1000) + ' LAK</td></tr>';

    hideLoading();
  } catch(error) {
    console.error('Error loading WAC:', error);
    hideLoading();
  }
}