async function loadWAC() {
  try {
    showLoading();

    const stockNewData = window._stockNewLatest || { goldG: 0, cost: 0 };
    const stockOldData = window._stockOldLatest || { goldG: 0, cost: 0 };

    const newGoldG = stockNewData.goldG || 0;
    const newCost = stockNewData.cost || 0;
    const oldGoldG = stockOldData.goldG || 0;
    const oldCost = stockOldData.cost || 0;

    const totalGoldG = newGoldG + oldGoldG;
    const totalCost = newCost + oldCost;

    const summaryBody = document.getElementById('wacSummaryTable');
    summaryBody.innerHTML = `
      <tr>
        <td>Stock (NEW)</td>
        <td>${formatNumber(newGoldG.toFixed(2))} g</td>
        <td>${formatNumber(Math.round(newCost / 1000) * 1000)} LAK</td>
      </tr>
      <tr>
        <td>Stock (OLD)</td>
        <td>${formatNumber(oldGoldG.toFixed(2))} g</td>
        <td>${formatNumber(Math.round(oldCost / 1000) * 1000)} LAK</td>
      </tr>
      <tr style="font-weight:bold; background: rgba(212,175,55,0.1);">
        <td>ผลรวม</td>
        <td>${formatNumber(totalGoldG.toFixed(2))} g</td>
        <td>${formatNumber(Math.round(totalCost / 1000) * 1000)} LAK</td>
      </tr>
    `;

    const wacPerG = totalGoldG > 0 ? totalCost / totalGoldG : 0;
    const wacPerBaht = wacPerG * 15;

    const calcBody = document.getElementById('wacCalcTable');
    calcBody.innerHTML = `
      <tr>
        <td>ราคา /g</td>
        <td style="font-weight:bold; color: var(--gold-primary);">${formatNumber(Math.round(wacPerG / 1000) * 1000)} LAK</td>
      </tr>
      <tr>
        <td>ราคา /บาท</td>
        <td style="font-weight:bold; color: var(--gold-primary);">${formatNumber(Math.round(wacPerBaht / 1000) * 1000)} LAK</td>
      </tr>
    `;

    hideLoading();
  } catch(error) {
    console.error('Error loading WAC:', error);
    hideLoading();
  }
}
