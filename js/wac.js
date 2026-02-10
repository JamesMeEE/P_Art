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
    let oldCarryW = 0;
    FIXED_PRODUCTS.forEach(p => { oldCarryW += (oldCarry[p.id] || 0) * getGoldWeight(p.id); });

    let newCarry = {};
    if (stockNewData.length > 1) {
      try { newCarry = JSON.parse(stockNewData[stockNewData.length - 1][1] || '{}'); } catch(e) {}
    }
    let newCarryW = 0;
    FIXED_PRODUCTS.forEach(p => { newCarryW += (newCarry[p.id] || 0) * getGoldWeight(p.id); });

    const accumulate = (data, startW) => {
      let w = startW, c = 0;
      if (data.length > 1) {
        data.slice(1).forEach(row => {
          const g = parseFloat(row[4]) || 0;
          const p = parseFloat(row[6]) || 0;
          if (String(row[5]) === 'IN') { w += g; c += p; }
          else { w -= g; c -= p; }
        });
      }
      return { weight: w, cost: c };
    };

    const old = accumulate(moveOldData, oldCarryW);
    const nw = accumulate(moveNewData, newCarryW);

    const absNewW = Math.abs(nw.weight);
    const absNewC = Math.abs(nw.cost);
    const totalGoldG = old.weight + absNewW;
    const totalCost = Math.round(old.cost / 1000) * 1000 + Math.round(absNewC / 1000) * 1000;

    document.getElementById('wacSummaryTable').innerHTML =
      '<tr><td>Stock (NEW)</td><td>' + formatNumber(absNewW.toFixed(2)) + ' g</td><td>' + formatNumber(Math.round(absNewC / 1000) * 1000) + ' LAK</td></tr>' +
      '<tr><td>Stock (OLD)</td><td>' + formatNumber(old.weight.toFixed(2)) + ' g</td><td>' + formatNumber(Math.round(old.cost / 1000) * 1000) + ' LAK</td></tr>' +
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