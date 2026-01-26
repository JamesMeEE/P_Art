async function loadDashboard() {
  try {
    showLoading();
    
    const today = new Date();
    const todayStr = formatDateOnly(today);
    document.getElementById('dashboardDate').textContent = todayStr;
    
    const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const todayEnd = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);
    
    const [sellData, tradeinData, buybackData, exchangeData, withdrawData, dbData] = await Promise.all([
      fetchSheetData('Sells!A:L'),
      fetchSheetData('Tradeins!A:N'),
      fetchSheetData('Buybacks!A:J'),
      fetchSheetData('Exchanges!A:N'),
      fetchSheetData('Withdraws!A:J'),
      fetchSheetData('_database!A1:G20')
    ]);
    
    let sell = { money: 0, goldG: 0, pieces: 0 };
    let tradein = { money: 0, goldDiffG: 0, oldGoldG: 0, pieces: 0 };
    let exchange = { money: 0, totalGoldG: 0, oldGoldG: 0, newGoldG: 0 };
    let buyback = { money: 0, goldG: 0, pieces: 0 };
    let withdraw = { money: 0, goldG: 0, pieces: 0 };
    
    sellData.slice(1).forEach(row => {
      const date = parseSheetDate(row[9]);
      if (date && date >= todayStart && date <= todayEnd && row[10] === 'COMPLETED') {
        const totalLAK = parseFloat(row[3]) || 0;
        sell.money += totalLAK;
        
        try {
          const items = JSON.parse(row[2]);
          items.forEach(item => {
            sell.goldG += getGoldWeight(item.productId) * item.qty;
            sell.pieces += item.qty;
          });
        } catch (e) {}
      }
    });
    
    tradeinData.slice(1).forEach(row => {
      const date = parseSheetDate(row[11]);
      if (date && date >= todayStart && date <= todayEnd && row[12] === 'COMPLETED') {
        const diff = parseFloat(row[4]) || 0;
        const fee = parseFloat(row[5]) || 0;
        const premium = parseFloat(row[6]) || 0;
        tradein.money += (diff + fee + premium);
        
        try {
          const oldItems = JSON.parse(row[2]);
          const newItems = JSON.parse(row[3]);
          
          let oldG = 0, newG = 0;
          oldItems.forEach(item => {
            oldG += getGoldWeight(item.productId) * item.qty;
          });
          newItems.forEach(item => {
            newG += getGoldWeight(item.productId) * item.qty;
            tradein.pieces += item.qty;
          });
          
          tradein.oldGoldG += oldG;
          tradein.goldDiffG += (newG - oldG);
        } catch (e) {}
      }
    });
    
    exchangeData.slice(1).forEach(row => {
      const date = parseSheetDate(row[11]);
      if (date && date >= todayStart && date <= todayEnd && row[12] === 'COMPLETED') {
        const fee = parseFloat(row[4]) || 0;
        const premium = parseFloat(row[5]) || 0;
        exchange.money += (fee + premium);
        
        try {
          const oldItems = JSON.parse(row[2]);
          const newItems = JSON.parse(row[3]);
          
          oldItems.forEach(item => {
            exchange.oldGoldG += getGoldWeight(item.productId) * item.qty;
          });
          newItems.forEach(item => {
            exchange.newGoldG += getGoldWeight(item.productId) * item.qty;
          });
          exchange.totalGoldG = exchange.oldGoldG + exchange.newGoldG;
        } catch (e) {}
      }
    });
    
    buybackData.slice(1).forEach(row => {
      const date = parseSheetDate(row[7]);
      if (date && date >= todayStart && date <= todayEnd && row[8] === 'COMPLETED') {
        buyback.money += parseFloat(row[6]) || 0;
        
        try {
          const items = JSON.parse(row[2]);
          items.forEach(item => {
            buyback.goldG += getGoldWeight(item.productId) * item.qty;
            buyback.pieces += item.qty;
          });
        } catch (e) {}
      }
    });
    
    withdrawData.slice(1).forEach(row => {
      const date = parseSheetDate(row[6]);
      if (date && date >= todayStart && date <= todayEnd && row[7] === 'COMPLETED') {
        withdraw.money += parseFloat(row[3]) || 0;
        
        try {
          const items = JSON.parse(row[2]);
          items.forEach(item => {
            withdraw.goldG += getGoldWeight(item.productId) * item.qty;
            withdraw.pieces += item.qty;
          });
        } catch (e) {}
      }
    });
    
    let newStock = { pieces: 0, goldG: 0 };
    let oldStock = { goldG: 0 };
    let cash = { LAK: 0, THB: 0, USD: 0 };
    let bank = { LAK: 0, THB: 0, USD: 0 };
    
    if (dbData.length >= 7) {
      const newGoldRow = dbData[6];
      const products = ['G01', 'G02', 'G03', 'G04', 'G05', 'G06', 'G07'];
      products.forEach((p, i) => {
        const qty = parseFloat(newGoldRow[i]) || 0;
        newStock.pieces += qty;
        newStock.goldG += qty * getGoldWeight(p);
      });
    }
    
    if (dbData.length >= 10) {
      const oldGoldRow = dbData[9];
      const products = ['G01', 'G02', 'G03', 'G04', 'G05', 'G06', 'G07'];
      products.forEach((p, i) => {
        const qty = parseFloat(oldGoldRow[i]) || 0;
        oldStock.goldG += qty * getGoldWeight(p);
      });
    }
    
    if (dbData.length >= 14) {
      cash.LAK = parseFloat(dbData[13][0]) || 0;
      cash.THB = parseFloat(dbData[13][1]) || 0;
      cash.USD = parseFloat(dbData[13][2]) || 0;
    }
    
    if (dbData.length >= 20) {
      bank.LAK = (parseFloat(dbData[16][0]) || 0) + (parseFloat(dbData[19][0]) || 0);
      bank.THB = (parseFloat(dbData[16][1]) || 0) + (parseFloat(dbData[19][1]) || 0);
      bank.USD = (parseFloat(dbData[16][2]) || 0) + (parseFloat(dbData[19][2]) || 0);
    }
    
    const sellAvg = sell.goldG > 0 ? sell.money / sell.goldG : 0;
    document.getElementById('dashSellBox').innerHTML = `
      <h3 style="color: var(--gold-primary); margin-bottom: 10px;">SELL</h3>
      <p style="font-size: 18px; margin: 5px 0;">${formatNumber(sell.money)} LAK <span style="font-size: 12px; color: var(--text-secondary);">(${formatNumber(sellAvg)}/g)</span></p>
      <p style="font-size: 14px; color: var(--text-secondary);">Gold: ${sell.goldG.toFixed(2)} g</p>
      <p style="font-size: 14px; color: var(--text-secondary);">Pieces: ${sell.pieces}</p>
    `;
    
    const tradeinAvg = tradein.goldDiffG > 0 ? tradein.money / tradein.goldDiffG : 0;
    document.getElementById('dashTradeinBox').innerHTML = `
      <h3 style="color: var(--gold-primary); margin-bottom: 10px;">TRADE-IN</h3>
      <p style="font-size: 18px; margin: 5px 0;">${formatNumber(tradein.money)} LAK <span style="font-size: 12px; color: var(--text-secondary);">(${formatNumber(tradeinAvg)}/g)</span></p>
      <p style="font-size: 14px; color: var(--text-secondary);">Diff: ${tradein.goldDiffG.toFixed(2)} g</p>
      <p style="font-size: 14px; color: var(--text-secondary);">Old Gold: ${tradein.oldGoldG.toFixed(2)} g</p>
      <p style="font-size: 14px; color: var(--text-secondary);">Pieces Out: ${tradein.pieces}</p>
    `;
    
    document.getElementById('dashExchangeBox').innerHTML = `
      <h3 style="color: var(--gold-primary); margin-bottom: 10px;">EXCHANGE</h3>
      <p style="font-size: 18px; margin: 5px 0;">${formatNumber(exchange.money)} LAK</p>
      <p style="font-size: 14px; color: var(--text-secondary);">Total: ${exchange.totalGoldG.toFixed(2)} g</p>
      <p style="font-size: 14px; color: var(--text-secondary);">Old In: ${exchange.oldGoldG.toFixed(2)} g</p>
      <p style="font-size: 14px; color: var(--text-secondary);">New Out: ${exchange.newGoldG.toFixed(2)} g</p>
    `;
    
    const buybackAvg = buyback.goldG > 0 ? buyback.money / buyback.goldG : 0;
    document.getElementById('dashBuybackBox').innerHTML = `
      <h3 style="color: var(--gold-primary); margin-bottom: 10px;">BUYBACK</h3>
      <p style="font-size: 18px; margin: 5px 0;">${formatNumber(buyback.money)} LAK <span style="font-size: 12px; color: var(--text-secondary);">(${formatNumber(buybackAvg)}/g)</span></p>
      <p style="font-size: 14px; color: var(--text-secondary);">Gold: ${buyback.goldG.toFixed(2)} g</p>
      <p style="font-size: 14px; color: var(--text-secondary);">Pieces: ${buyback.pieces}</p>
    `;
    
    document.getElementById('dashWithdrawBox').innerHTML = `
      <h3 style="color: var(--gold-primary); margin-bottom: 10px;">WITHDRAW</h3>
      <p style="font-size: 18px; margin: 5px 0;">${formatNumber(withdraw.money)} LAK</p>
      <p style="font-size: 14px; color: var(--text-secondary);">Gold: ${withdraw.goldG.toFixed(2)} g</p>
      <p style="font-size: 14px; color: var(--text-secondary);">Pieces: ${withdraw.pieces}</p>
    `;
    
    document.getElementById('dashNewStockBox').innerHTML = `
      <h3 style="color: #4caf50; margin-bottom: 10px;">NEW STOCK</h3>
      <p style="font-size: 18px; margin: 5px 0;">${newStock.pieces} pcs</p>
      <p style="font-size: 14px; color: var(--text-secondary);">${newStock.goldG.toFixed(2)} g</p>
    `;
    
    document.getElementById('dashOldStockBox').innerHTML = `
      <h3 style="color: #ff9800; margin-bottom: 10px;">OLD STOCK</h3>
      <p style="font-size: 18px; margin: 5px 0;">${oldStock.goldG.toFixed(2)} g</p>
    `;
    
    document.getElementById('dashCashBox').innerHTML = `
      <h3 style="color: var(--gold-primary); margin-bottom: 10px;">CASH</h3>
      <p style="font-size: 16px; margin: 3px 0;">${formatNumber(cash.LAK)} LAK</p>
      <p style="font-size: 16px; margin: 3px 0;">${formatNumber(cash.THB)} THB</p>
      <p style="font-size: 16px; margin: 3px 0;">${formatNumber(cash.USD)} USD</p>
    `;
    
    document.getElementById('dashBankBox').innerHTML = `
      <h3 style="color: var(--gold-primary); margin-bottom: 10px;">BANK</h3>
      <p style="font-size: 16px; margin: 3px 0;">${formatNumber(bank.LAK)} LAK</p>
      <p style="font-size: 16px; margin: 3px 0;">${formatNumber(bank.THB)} THB</p>
      <p style="font-size: 16px; margin: 3px 0;">${formatNumber(bank.USD)} USD</p>
    `;
    
    const totalGold = newStock.goldG + oldStock.goldG;
    document.getElementById('dashTotalGoldBox').innerHTML = `
      <h3 style="color: var(--gold-primary); margin-bottom: 10px;">TOTAL GOLD</h3>
      <p style="font-size: 24px; font-weight: bold; color: var(--gold-primary);">${totalGold.toFixed(2)} g</p>
      <p style="font-size: 12px; color: var(--text-secondary);">NEW: ${newStock.goldG.toFixed(2)} g | OLD: ${oldStock.goldG.toFixed(2)} g</p>
    `;
    
    const totalCash = {
      LAK: cash.LAK + bank.LAK,
      THB: cash.THB + bank.THB,
      USD: cash.USD + bank.USD
    };
    document.getElementById('dashTotalCashBox').innerHTML = `
      <h3 style="color: var(--gold-primary); margin-bottom: 10px;">TOTAL CASH + BANK</h3>
      <p style="font-size: 18px; margin: 5px 0;">${formatNumber(totalCash.LAK)} LAK</p>
      <p style="font-size: 18px; margin: 5px 0;">${formatNumber(totalCash.THB)} THB</p>
      <p style="font-size: 18px; margin: 5px 0;">${formatNumber(totalCash.USD)} USD</p>
    `;

    hideLoading();
  } catch (error) {
    console.error('Error loading dashboard:', error);
    hideLoading();
  }
}

function getGoldWeight(productId) {
  const weights = {
    'G01': 150,
    'G02': 75,
    'G03': 30,
    'G04': 15,
    'G05': 7.5,
    'G06': 3.75,
    'G07': 1
  };
  return weights[productId] || 0;
}