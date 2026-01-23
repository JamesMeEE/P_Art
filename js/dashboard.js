async function loadDashboard() {
  try {
    showLoading();
    
    const today = new Date();
    const todayStr = formatDateOnly(today);
    document.getElementById('dashboardDate').textContent = todayStr;
    
    const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 1);
    const todayEnd = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59);
    
    const [sellData, tradeinData, buybackData, exchangeData, cashbankData, stockData] = await Promise.all([
      fetchSheetData('Sells!A:I'),
      fetchSheetData('Tradeins!A:K'),
      fetchSheetData('Buybacks!A:G'),
      fetchSheetData('Exchanges!A:J'),
      fetchSheetData('CashBank!A:G'),
      fetchSheetData('Stock!A:F')
    ]);
    
    let sellCount = 0, buybackCount = 0, tradeinCount = 0, exchangeCount = 0;
    let goldFlowBaht = 0;
    let cashFlow = 0, bankFlow = 0;
    
    sellData.slice(1).forEach(row => {
      const date = new Date(row[6]);
      if (date >= todayStart && date <= todayEnd && row[7] === 'COMPLETED') {
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
      const date = new Date(row[4]);
      if (date >= todayStart && date <= todayEnd && row[5] === 'COMPLETED') {
        buybackCount++;
        const items = JSON.parse(row[2]);
        items.forEach(item => {
          goldFlowBaht += GOLD_WEIGHTS[item.productId] * item.qty;
        });
        cashFlow -= parseFloat(row[3]) || 0;
      }
    });
    
    tradeinData.slice(1).forEach(row => {
      const date = new Date(row[7]);
      if (date >= todayStart && date <= todayEnd && row[8] === 'COMPLETED') {
        tradeinCount++;
        const diff = parseFloat(row[4]) || 0;
        const fee = parseFloat(row[5]) || 0;
        const premium = parseFloat(row[6]) || 0;
        cashFlow += (diff + fee + premium);
      }
    });
    
    exchangeData.slice(1).forEach(row => {
      const date = new Date(row[6]);
      if (date >= todayStart && date <= todayEnd && row[7] === 'COMPLETED') {
        exchangeCount++;
        const fee = parseFloat(row[4]) || 0;
        const premium = parseFloat(row[5]) || 0;
        cashFlow += (fee + premium);
      }
    });
    
    cashbankData.slice(1).forEach(row => {
      const amount = parseFloat(row[2]) || 0;
      const method = row[3];
      const type = row[1];
      if (type === 'OWNER_DEPOSIT' || type === 'OTHER_INCOME') {
        if (method === 'CASH') cashFlow += amount;
        else bankFlow += amount;
      } else if (type === 'OTHER_EXPENSE') {
        if (method === 'CASH') cashFlow -= amount;
        else bankFlow -= amount;
      }
    });
    
    const goldFlowGrams = goldFlowBaht * 15;
    
    let totalGoldBaht = 0;
    const stockMap = {};
    stockData.slice(1).forEach(row => {
      const productId = row[0];
      const qty = parseInt(row[3]) || 0;
      if (!stockMap[productId]) stockMap[productId] = 0;
      stockMap[productId] += qty;
    });
    
    Object.keys(stockMap).forEach(productId => {
      totalGoldBaht += stockMap[productId] * GOLD_WEIGHTS[productId];
    });
    
    const assetGrams = totalGoldBaht * 15;
    
    document.getElementById('dashSell').textContent = sellCount;
    document.getElementById('dashBuyback').textContent = buybackCount;
    document.getElementById('dashTradein').textContent = tradeinCount;
    document.getElementById('dashExchange').textContent = exchangeCount;
    document.getElementById('dashGoldFlow').textContent = goldFlowBaht.toFixed(2);
    document.getElementById('dashGoldFlowG').textContent = goldFlowGrams.toFixed(2);
    document.getElementById('dashCash').textContent = formatNumber(cashFlow);
    document.getElementById('dashBank').textContent = formatNumber(bankFlow);
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
