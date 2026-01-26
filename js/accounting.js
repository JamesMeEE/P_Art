async function loadAccounting() {
  try {
    showLoading();
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    const accountingData = await fetchSheetData('Accounting!A:M');
    
    if (accountingData.length > 1) {
      const lastRecord = accountingData[accountingData.length - 1];
      const lastDate = lastRecord[0];
      let lastDateObj;
      
      if (lastDate instanceof Date) {
        lastDateObj = new Date(lastDate);
      } else if (typeof lastDate === 'string') {
        if (lastDate.includes('/')) {
          const parts = lastDate.split('/');
          lastDateObj = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
        } else {
          lastDateObj = new Date(lastDate);
        }
      } else {
        lastDateObj = new Date(lastDate);
      }
      
      lastDateObj.setHours(0, 0, 0, 0);
      
      const diffTime = today.getTime() - lastDateObj.getTime();
      const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
      
      if (diffDays > 0) {
        for (let i = 1; i <= diffDays; i++) {
          const targetDate = new Date(lastDateObj);
          targetDate.setDate(targetDate.getDate() + i);
          await saveAccountingForDate(targetDate);
        }
      }
    }
    
    await loadTodayStats();
    await loadAccountingHistory();
    
    hideLoading();
  } catch (error) {
    console.error('Error loading accounting:', error);
    hideLoading();
  }
}

async function saveAccountingForDate(targetDate) {
  try {
    const sells = await fetchSheetData('Sells!A:L');
    const tradeins = await fetchSheetData('Tradeins!A:N');
    const exchanges = await fetchSheetData('Exchanges!A:N');
    const buybacks = await fetchSheetData('Buybacks!A:J');
    const withdraws = await fetchSheetData('Withdraws!A:J');
    
    const dayStart = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate());
    const dayEnd = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 23, 59, 59);
    
    let stats = {
      sellMoney: 0, sellGold: 0,
      tradeinMoney: 0, tradeinGold: 0,
      exchangeMoney: 0, exchangeGold: 0,
      buybackMoney: 0, buybackGold: 0,
      withdrawMoney: 0, withdrawGold: 0,
      incompleteMoney: 0, incompleteGold: 0
    };
    
    sells.slice(1).forEach(row => {
      const date = new Date(row[9]);
      if (date >= dayStart && date <= dayEnd) {
        const status = row[10];
        const totalLAK = parseFloat(row[6]) || 0;
        const items = JSON.parse(row[2] || '[]');
        const goldG = calculateTotalGold(items);
        
        if (status === 'COMPLETED') {
          stats.sellMoney += totalLAK;
          stats.sellGold += goldG;
        } else {
          stats.incompleteMoney += totalLAK;
          stats.incompleteGold += goldG;
        }
      }
    });
    
    tradeins.slice(1).forEach(row => {
      const date = new Date(row[11]);
      if (date >= dayStart && date <= dayEnd) {
        const status = row[12];
        const totalLAK = parseFloat(row[8]) || 0;
        const items = JSON.parse(row[3] || '[]');
        const goldG = calculateTotalGold(items);
        
        if (status === 'COMPLETED') {
          stats.tradeinMoney += totalLAK;
          stats.tradeinGold += goldG;
        } else {
          stats.incompleteMoney += totalLAK;
          stats.incompleteGold += goldG;
        }
      }
    });
    
    exchanges.slice(1).forEach(row => {
      const date = new Date(row[11]);
      if (date >= dayStart && date <= dayEnd) {
        const status = row[12];
        const totalLAK = parseFloat(row[8]) || 0;
        const items = JSON.parse(row[3] || '[]');
        const goldG = calculateTotalGold(items);
        
        if (status === 'COMPLETED') {
          stats.exchangeMoney += totalLAK;
          stats.exchangeGold += goldG;
        } else {
          stats.incompleteMoney += totalLAK;
          stats.incompleteGold += goldG;
        }
      }
    });
    
    buybacks.slice(1).forEach(row => {
      const date = new Date(row[7]);
      if (date >= dayStart && date <= dayEnd) {
        const status = row[8];
        const totalLAK = parseFloat(row[6]) || 0;
        const items = JSON.parse(row[2] || '[]');
        const goldG = calculateTotalGold(items);
        
        if (status === 'COMPLETED') {
          stats.buybackMoney += totalLAK;
          stats.buybackGold += goldG;
        } else {
          stats.incompleteMoney += totalLAK;
          stats.incompleteGold += goldG;
        }
      }
    });
    
    withdraws.slice(1).forEach(row => {
      const date = new Date(row[6]);
      if (date >= dayStart && date <= dayEnd) {
        const status = row[7];
        const totalLAK = parseFloat(row[4]) || 0;
        const items = JSON.parse(row[2] || '[]');
        const goldG = calculateTotalGold(items);
        
        if (status === 'COMPLETED') {
          stats.withdrawMoney += totalLAK;
          stats.withdrawGold += goldG;
        } else {
          stats.incompleteMoney += totalLAK;
          stats.incompleteGold += goldG;
        }
      }
    });
    
    const dateStr = `${targetDate.getDate()}/${targetDate.getMonth() + 1}/${targetDate.getFullYear()}`;
    
    await executeGoogleScript({
      action: 'SAVE_ACCOUNTING',
      date: dateStr,
      sellMoney: stats.sellMoney,
      sellGold: stats.sellGold,
      tradeinMoney: stats.tradeinMoney,
      tradeinGold: stats.tradeinGold,
      exchangeMoney: stats.exchangeMoney,
      exchangeGold: stats.exchangeGold,
      buybackMoney: stats.buybackMoney,
      buybackGold: stats.buybackGold,
      withdrawMoney: stats.withdrawMoney,
      withdrawGold: stats.withdrawGold,
      incompleteMoney: stats.incompleteMoney,
      incompleteGold: stats.incompleteGold
    });
  } catch (error) {
    console.error('Error saving accounting for date:', error);
  }
}

async function loadTodayStats() {
  try {
    const sells = await fetchSheetData('Sells!A:L');
    const tradeins = await fetchSheetData('Tradeins!A:N');
    const exchanges = await fetchSheetData('Exchanges!A:N');
    const buybacks = await fetchSheetData('Buybacks!A:J');
    const withdraws = await fetchSheetData('Withdraws!A:J');
    
    const today = new Date();
    const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const todayEnd = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);
    
    let stats = {
      sellMoney: 0, sellGold: 0,
      tradeinMoney: 0, tradeinGold: 0,
      exchangeMoney: 0, exchangeGold: 0,
      buybackMoney: 0, buybackGold: 0,
      withdrawMoney: 0, withdrawGold: 0,
      incompleteMoney: 0, incompleteGold: 0
    };
    
    sells.slice(1).forEach(row => {
      const date = new Date(row[9]);
      if (date >= todayStart && date <= todayEnd) {
        const status = row[10];
        const totalLAK = parseFloat(row[6]) || 0;
        const items = JSON.parse(row[2] || '[]');
        const goldG = calculateTotalGold(items);
        
        if (status === 'COMPLETED') {
          stats.sellMoney += totalLAK;
          stats.sellGold += goldG;
        } else {
          stats.incompleteMoney += totalLAK;
          stats.incompleteGold += goldG;
        }
      }
    });
    
    tradeins.slice(1).forEach(row => {
      const date = new Date(row[11]);
      if (date >= todayStart && date <= todayEnd) {
        const status = row[12];
        const totalLAK = parseFloat(row[8]) || 0;
        const items = JSON.parse(row[3] || '[]');
        const goldG = calculateTotalGold(items);
        
        if (status === 'COMPLETED') {
          stats.tradeinMoney += totalLAK;
          stats.tradeinGold += goldG;
        } else {
          stats.incompleteMoney += totalLAK;
          stats.incompleteGold += goldG;
        }
      }
    });
    
    exchanges.slice(1).forEach(row => {
      const date = new Date(row[11]);
      if (date >= todayStart && date <= todayEnd) {
        const status = row[12];
        const totalLAK = parseFloat(row[8]) || 0;
        const items = JSON.parse(row[3] || '[]');
        const goldG = calculateTotalGold(items);
        
        if (status === 'COMPLETED') {
          stats.exchangeMoney += totalLAK;
          stats.exchangeGold += goldG;
        } else {
          stats.incompleteMoney += totalLAK;
          stats.incompleteGold += goldG;
        }
      }
    });
    
    buybacks.slice(1).forEach(row => {
      const date = new Date(row[7]);
      if (date >= todayStart && date <= todayEnd) {
        const status = row[8];
        const totalLAK = parseFloat(row[6]) || 0;
        const items = JSON.parse(row[2] || '[]');
        const goldG = calculateTotalGold(items);
        
        if (status === 'COMPLETED') {
          stats.buybackMoney += totalLAK;
          stats.buybackGold += goldG;
        } else {
          stats.incompleteMoney += totalLAK;
          stats.incompleteGold += goldG;
        }
      }
    });
    
    withdraws.slice(1).forEach(row => {
      const date = new Date(row[6]);
      if (date >= todayStart && date <= todayEnd) {
        const status = row[7];
        const totalLAK = parseFloat(row[4]) || 0;
        const items = JSON.parse(row[2] || '[]');
        const goldG = calculateTotalGold(items);
        
        if (status === 'COMPLETED') {
          stats.withdrawMoney += totalLAK;
          stats.withdrawGold += goldG;
        } else {
          stats.incompleteMoney += totalLAK;
          stats.incompleteGold += goldG;
        }
      }
    });
    
    document.getElementById('accountingStats').innerHTML = `
      <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px;">
        <div class="stat-card">
          <h3>SELL</h3>
          <p style="font-size: 20px; color: var(--gold-primary); margin: 10px 0;">${formatNumber(stats.sellMoney)} LAK</p>
          <p style="font-size: 16px;">${stats.sellGold.toFixed(2)} g</p>
        </div>
        <div class="stat-card">
          <h3>TRADE-IN</h3>
          <p style="font-size: 20px; color: var(--gold-primary); margin: 10px 0;">${formatNumber(stats.tradeinMoney)} LAK</p>
          <p style="font-size: 16px;">${stats.tradeinGold.toFixed(2)} g</p>
        </div>
        <div class="stat-card">
          <h3>EXCHANGE</h3>
          <p style="font-size: 20px; color: var(--gold-primary); margin: 10px 0;">${formatNumber(stats.exchangeMoney)} LAK</p>
          <p style="font-size: 16px;">${stats.exchangeGold.toFixed(2)} g</p>
        </div>
        <div class="stat-card">
          <h3>BUY BACK</h3>
          <p style="font-size: 20px; color: var(--gold-primary); margin: 10px 0;">${formatNumber(stats.buybackMoney)} LAK</p>
          <p style="font-size: 16px;">${stats.buybackGold.toFixed(2)} g</p>
        </div>
        <div class="stat-card">
          <h3>WITHDRAW</h3>
          <p style="font-size: 20px; color: var(--gold-primary); margin: 10px 0;">${formatNumber(stats.withdrawMoney)} LAK</p>
          <p style="font-size: 16px;">${stats.withdrawGold.toFixed(2)} g</p>
        </div>
        <div class="stat-card">
          <h3>INCOMPLETE</h3>
          <p style="font-size: 20px; color: var(--gold-primary); margin: 10px 0;">${formatNumber(stats.incompleteMoney)} LAK</p>
          <p style="font-size: 16px;">${stats.incompleteGold.toFixed(2)} g</p>
        </div>
      </div>
    `;
  } catch (error) {
    console.error('Error loading today stats:', error);
  }
}

async function loadAccountingHistory() {
  try {
    const data = await fetchSheetData('Accounting!A:M');
    const tbody = document.getElementById('accountingTable');
    
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="13" style="text-align: center; padding: 40px;">No records</td></tr>';
      return;
    }
    
    const history = data.slice(1).reverse();
    
    tbody.innerHTML = history.map(row => `
      <tr>
        <td>${row[0]}</td>
        <td>${formatNumber(row[1])} LAK</td>
        <td>${parseFloat(row[2] || 0).toFixed(2)} g</td>
        <td>${formatNumber(row[3])} LAK</td>
        <td>${parseFloat(row[4] || 0).toFixed(2)} g</td>
        <td>${formatNumber(row[5])} LAK</td>
        <td>${parseFloat(row[6] || 0).toFixed(2)} g</td>
        <td>${formatNumber(row[7])} LAK</td>
        <td>${parseFloat(row[8] || 0).toFixed(2)} g</td>
        <td>${formatNumber(row[9])} LAK</td>
        <td>${parseFloat(row[10] || 0).toFixed(2)} g</td>
        <td>${formatNumber(row[11])} LAK</td>
        <td>${parseFloat(row[12] || 0).toFixed(2)} g</td>
      </tr>
    `).join('');
  } catch (error) {
    console.error('Error loading accounting history:', error);
  }
}

async function filterAccountingHistory() {
  try {
    const startDate = document.getElementById('accountingStartDate').value;
    const endDate = document.getElementById('accountingEndDate').value;
    
    if (!startDate || !endDate) {
      return;
    }
    
    const data = await fetchSheetData('Accounting!A:M');
    const tbody = document.getElementById('accountingTable');
    
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="13" style="text-align: center; padding: 40px;">No records</td></tr>';
      return;
    }
    
    const filtered = data.slice(1).filter(row => {
      const rowDate = row[0];
      return rowDate >= startDate && rowDate <= endDate;
    }).reverse();
    
    if (filtered.length === 0) {
      tbody.innerHTML = '<tr><td colspan="13" style="text-align: center; padding: 40px;">No records in this date range</td></tr>';
      return;
    }
    
    tbody.innerHTML = filtered.map(row => `
      <tr>
        <td>${row[0]}</td>
        <td>${formatNumber(row[1])} LAK</td>
        <td>${parseFloat(row[2] || 0).toFixed(2)} g</td>
        <td>${formatNumber(row[3])} LAK</td>
        <td>${parseFloat(row[4] || 0).toFixed(2)} g</td>
        <td>${formatNumber(row[5])} LAK</td>
        <td>${parseFloat(row[6] || 0).toFixed(2)} g</td>
        <td>${formatNumber(row[7])} LAK</td>
        <td>${parseFloat(row[8] || 0).toFixed(2)} g</td>
        <td>${formatNumber(row[9])} LAK</td>
        <td>${parseFloat(row[10] || 0).toFixed(2)} g</td>
        <td>${formatNumber(row[11])} LAK</td>
        <td>${parseFloat(row[12] || 0).toFixed(2)} g</td>
      </tr>
    `).join('');
  } catch (error) {
    console.error('Error filtering accounting:', error);
  }
}

function checkAccountingFilter() {
  const startDate = document.getElementById('accountingStartDate').value;
  const endDate = document.getElementById('accountingEndDate').value;
  
  if (startDate && endDate) {
    filterAccountingHistory();
  }
}

async function showTodayAccounting() {
  const today = new Date().toISOString().split('T')[0];
  document.getElementById('accountingStartDate').value = today;
  document.getElementById('accountingEndDate').value = today;
  
  const data = await fetchSheetData('Accounting!A:M');
  const tbody = document.getElementById('accountingTable');
  
  if (data.length <= 1) {
    tbody.innerHTML = '<tr><td colspan="13" style="text-align: center; padding: 40px;">No records</td></tr>';
    return;
  }
  
  const filtered = data.slice(1).filter(row => row[0] === today).reverse();
  
  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="13" style="text-align: center; padding: 40px;">No records for today</td></tr>';
    return;
  }
  
  tbody.innerHTML = filtered.map(row => `
    <tr>
      <td>${row[0]}</td>
      <td>${formatNumber(row[1])} LAK</td>
      <td>${parseFloat(row[2] || 0).toFixed(2)} g</td>
      <td>${formatNumber(row[3])} LAK</td>
      <td>${parseFloat(row[4] || 0).toFixed(2)} g</td>
      <td>${formatNumber(row[5])} LAK</td>
      <td>${parseFloat(row[6] || 0).toFixed(2)} g</td>
      <td>${formatNumber(row[7])} LAK</td>
      <td>${parseFloat(row[8] || 0).toFixed(2)} g</td>
      <td>${formatNumber(row[9])} LAK</td>
      <td>${parseFloat(row[10] || 0).toFixed(2)} g</td>
      <td>${formatNumber(row[11])} LAK</td>
      <td>${parseFloat(row[12] || 0).toFixed(2)} g</td>
    </tr>
  `).join('');
}

function calculateTotalGold(items) {
  let totalG = 0;
  items.forEach(item => {
    const qty = item.qty || 0;
    switch(item.productId) {
      case 'G01': totalG += 10 * qty; break;
      case 'G02': totalG += 5 * qty; break;
      case 'G03': totalG += 2 * qty; break;
      case 'G04': totalG += 1 * qty; break;
      case 'G05': totalG += 0.5 * qty; break;
      case 'G06': totalG += 0.25 * qty; break;
      case 'G07': totalG += 0.0666 * qty; break;
    }
  });
  return totalG;
}