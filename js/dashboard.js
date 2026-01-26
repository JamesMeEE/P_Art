async function loadDashboard() {
  try {
    showLoading();
    
    const today = new Date();
    const todayStr = formatDateOnly(today);
    document.getElementById('dashboardDate').textContent = todayStr;
    
    const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 1);
    const todayEnd = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59);
    
    const [sellData, tradeinData, buybackData, exchangeData, dbData] = await Promise.all([
      fetchSheetData('Sells!A:L'),
      fetchSheetData('Tradeins!A:K'),
      fetchSheetData('Buybacks!A:G'),
      fetchSheetData('Exchanges!A:J'),
      fetchSheetData('_database!A1:G20')
    ]);
    
    let sellCount = 0, buybackCount = 0, tradeinCount = 0, exchangeCount = 0;
    let goldFlowBaht = 0;
    let cashFlow = 0, bankFlow = 0;
    
    sellData.slice(1).forEach(row => {
      const date = parseSheetDate(row[9]);
      if (date && date >= todayStart && date <= todayEnd && row[10] === 'COMPLETED') {
        sellCount++;
        const items = JSON.parse(row[2]);
        items.forEach(item => {
          goldFlowBaht -= GOLD_WEIGHTS[item.productId] * item.qty;
        });
        const amount = parseFloat(row[3]) || 0;
        if (row[4] === 'LAK') cashFlow += amount;
        else bankFlow += amount;
      }
    });
    
    buybackData.slice(1).forEach(row => {
      const date = parseSheetDate(row[7]);
      if (date && date >= todayStart && date <= todayEnd && row[8] === 'COMPLETED') {
        buybackCount++;
        const items = JSON.parse(row[2]);
        items.forEach(item => {
          goldFlowBaht += GOLD_WEIGHTS[item.productId] * item.qty;
        });
        cashFlow -= parseFloat(row[6]) || 0;
      }
    });
    
    tradeinData.slice(1).forEach(row => {
      const date = parseSheetDate(row[11]);
      if (date && date >= todayStart && date <= todayEnd && row[12] === 'COMPLETED') {
        tradeinCount++;
        const diff = parseFloat(row[4]) || 0;
        const fee = parseFloat(row[5]) || 0;
        const premium = parseFloat(row[6]) || 0;
        cashFlow += (diff + fee + premium);
      }
    });
    
    exchangeData.slice(1).forEach(row => {
      const date = parseSheetDate(row[11]);
      if (date && date >= todayStart && date <= todayEnd && row[12] === 'COMPLETED') {
        exchangeCount++;
        const fee = parseFloat(row[4]) || 0;
        const premium = parseFloat(row[5]) || 0;
        cashFlow += (fee + premium);
      }
    });
    
    const goldFlowGrams = goldFlowBaht * 15;
    
    let totalGoldBaht = 0;
    if (dbData.length >= 7) {
      const newGoldRow = dbData[6];
      totalGoldBaht += (parseFloat(newGoldRow[0]) || 0) * GOLD_WEIGHTS['G01'];
      totalGoldBaht += (parseFloat(newGoldRow[1]) || 0) * GOLD_WEIGHTS['G02'];
      totalGoldBaht += (parseFloat(newGoldRow[2]) || 0) * GOLD_WEIGHTS['G03'];
      totalGoldBaht += (parseFloat(newGoldRow[3]) || 0) * GOLD_WEIGHTS['G04'];
      totalGoldBaht += (parseFloat(newGoldRow[4]) || 0) * GOLD_WEIGHTS['G05'];
      totalGoldBaht += (parseFloat(newGoldRow[5]) || 0) * GOLD_WEIGHTS['G06'];
      totalGoldBaht += (parseFloat(newGoldRow[6]) || 0) * GOLD_WEIGHTS['G07'];
    }
    
    let totalCashLAK = 0, totalBankLAK = 0;
    if (dbData.length >= 20) {
      totalCashLAK = parseFloat(dbData[13][0]) || 0;
      const bcelLAK = parseFloat(dbData[16][0]) || 0;
      const ldbLAK = parseFloat(dbData[19][0]) || 0;
      totalBankLAK = bcelLAK + ldbLAK;
    }
    
    const assetGrams = totalGoldBaht * 15;
    
    document.getElementById('dashSell').textContent = sellCount;
    document.getElementById('dashBuyback').textContent = buybackCount;
    document.getElementById('dashTradein').textContent = tradeinCount;
    document.getElementById('dashExchange').textContent = exchangeCount;
    document.getElementById('dashGoldFlow').textContent = goldFlowBaht.toFixed(2);
    document.getElementById('dashGoldFlowG').textContent = goldFlowGrams.toFixed(2);
    document.getElementById('dashCash').textContent = formatNumber(totalCashLAK);
    document.getElementById('dashBank').textContent = formatNumber(totalBankLAK);
    document.getElementById('dashAsset').textContent = assetGrams.toFixed(2);

    document.getElementById('summaryContent').innerHTML = `
      <div style="margin: 20px 0;">
        <div style="margin-bottom: 15px;">
          <span style="color: var(--text-secondary);">Total Sells:</span>
          <span style="color: var(--gold-primary); font-weight: 600; float: right;">${sellData.length - 1}</span>
        </div>
        <div style="margin-bottom: 15px;">
          <span style="color: var(--text-secondary);">Total Trade-ins:</span>
          <span style="color: var(--gold-primary); font-weight: 600; float: right;">${tradeinData.length - 1}</span>
        </div>
        <div style="margin-bottom: 15px;">
          <span style="color: var(--text-secondary);">Total Buybacks:</span>
          <span style="color: var(--gold-primary); font-weight: 600; float: right;">${buybackData.length - 1}</span>
        </div>
        <div style="margin-bottom: 15px;">
          <span style="color: var(--text-secondary);">Total Exchanges:</span>
          <span style="color: var(--gold-primary); font-weight: 600; float: right;">${exchangeData.length - 1}</span>
        </div>
      </div>
    `;

    hideLoading();
  } catch (error) {
    console.error('Error loading dashboard:', error);
    hideLoading();
  }
}