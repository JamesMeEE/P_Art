async function loadAccounting() {
  try {
    showLoading();
    const today = new Date();
    const todayStr = formatDateOnly(today);
    document.getElementById('accountingDate').textContent = todayStr;
    
    const [sellData, tradeinData, buybackData, exchangeData, cashbankData, stockData] = await Promise.all([
      fetchSheetData('Sells!A:I'),
      fetchSheetData('Tradeins!A:K'),
      fetchSheetData('Buybacks!A:G'),
      fetchSheetData('Exchanges!A:J'),
      fetchSheetData('CashBank!A:G'),
      fetchSheetData('Stock!A:F')
    ]);
    
    const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 1);
    const todayEnd = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59);
    
    let totalTransactions = 0;
    let cashbank = 0;
    let revenue = 0;
    let expense = 0;
    let cost = 0;
    
    sellData.slice(1).forEach(row => {
      const date = new Date(row[6]);
      if (date >= todayStart && date <= todayEnd && row[7] === 'COMPLETED') {
        totalTransactions++;
        revenue += parseFloat(row[3]) || 0;
      }
    });
    
    buybackData.slice(1).forEach(row => {
      const date = new Date(row[4]);
      if (date >= todayStart && date <= todayEnd && row[5] === 'COMPLETED') {
        totalTransactions++;
        expense += parseFloat(row[3]) || 0;
      }
    });
    
    tradeinData.slice(1).forEach(row => {
      const date = new Date(row[7]);
      if (date >= todayStart && date <= todayEnd && row[8] === 'COMPLETED') {
        totalTransactions++;
        revenue += parseFloat(row[4]) || 0;
        revenue += parseFloat(row[5]) || 0;
        revenue += parseFloat(row[6]) || 0;
      }
    });
    
    exchangeData.slice(1).forEach(row => {
      const date = new Date(row[6]);
      if (date >= todayStart && date <= todayEnd && row[7] === 'COMPLETED') {
        totalTransactions++;
        revenue += parseFloat(row[4]) || 0;
        revenue += parseFloat(row[5]) || 0;
      }
    });
    
    cashbankData.slice(1).forEach(row => {
      const amount = parseFloat(row[2]) || 0;
      const type = row[1];
      if (type === 'OWNER_DEPOSIT' || type === 'OTHER_INCOME') {
        cashbank += amount;
      } else if (type === 'OTHER_EXPENSE') {
        cashbank -= amount;
      }
    });
    
    const pl = revenue - expense;
    
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
    
    const assetBaht = ((cashbank / currentPricing.sell1Baht) || 0) + totalGoldBaht;
    const assetGrams = assetBaht * 15;
    
    document.getElementById('accTransactions').textContent = totalTransactions;
    document.getElementById('accCashBank').textContent = formatNumber(cashbank);
    document.getElementById('accRevenue').textContent = formatNumber(revenue - expense);
    document.getElementById('accPL').textContent = formatNumber(pl);
    document.getElementById('accCost').textContent = formatNumber(cost);
    document.getElementById('accAsset').textContent = assetGrams.toFixed(2);
    
    const reconcileData = await fetchSheetData('Reconcile!A:J');
    const tbody = document.getElementById('reconcileHistoryTable');
    if (reconcileData.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = reconcileData.slice(1).reverse().map(row => `
        <tr>
          <td>${formatDateOnly(row[0])}</td>
          <td>${row[1]} / ${row[2]}</td>
          <td>${formatNumber(row[3])} / ${formatNumber(row[4])}</td>
          <td>${formatNumber(row[5])} / ${formatNumber(row[6])}</td>
          <td>-</td>
          <td>-</td>
          <td>${row[7]} / ${row[8]}</td>
          <td>${row[9]}</td>
        </tr>
      `).join('');
    }
    
    hideLoading();
  } catch (error) {
    console.error('Error loading accounting:', error);
    hideLoading();
  }
}

function openReconcileModal(type) {
  currentReconcileType = type;
  const titles = {
    'transactions': 'Reconcile Transactions',
    'cashbank': 'Reconcile Cash/Bank',
    'revenue': 'Reconcile Revenue-Expense',
    'pl': 'Reconcile P/L',
    'cost': 'Reconcile Cost',
    'asset': 'Reconcile Asset (Gold)'
  };
  
  document.getElementById('reconcileModalTitle').textContent = titles[type];
  
  const systemValues = {
    'transactions': document.getElementById('accTransactions').textContent,
    'cashbank': document.getElementById('accCashBank').textContent,
    'revenue': document.getElementById('accRevenue').textContent,
    'pl': document.getElementById('accPL').textContent,
    'cost': document.getElementById('accCost').textContent,
    'asset': document.getElementById('accAsset').textContent
  };
  
  document.getElementById('reconcileSystemValue').value = systemValues[type];
  
  if (currentReconcileData[type]) {
    document.getElementById('reconcileActualValue').value = currentReconcileData[type].actual;
    document.getElementById('reconcileDifference').value = currentReconcileData[type].difference;
  } else {
    document.getElementById('reconcileActualValue').value = '';
    document.getElementById('reconcileDifference').value = '';
  }
  
  document.getElementById('reconcileActualValue').oninput = function() {
    const system = parseFloat(systemValues[type].replace(/,/g, '')) || 0;
    const actual = parseFloat(this.value) || 0;
    document.getElementById('reconcileDifference').value = (actual - system).toFixed(2);
  };
  
  openModal('reconcileModal');
}

async function submitReconcile() {
  const actualValue = document.getElementById('reconcileActualValue').value;
  if (!actualValue) {
    alert('Please enter actual value');
    return;
  }
  
  currentReconcileData[currentReconcileType] = {
    system: document.getElementById('reconcileSystemValue').value.replace(/,/g, ''),
    actual: actualValue,
    difference: document.getElementById('reconcileDifference').value
  };
  
  const cardIds = {
    'transactions': 'accTransCard',
    'cashbank': 'accCashCard',
    'revenue': 'accRevCard',
    'pl': 'accPLCard',
    'cost': 'accCostCard',
    'asset': 'accAssetCard'
  };
  
  const actualIds = {
    'transactions': 'accTransActual',
    'cashbank': 'accCashActual',
    'revenue': 'accRevActual',
    'pl': 'accPLActual',
    'cost': 'accCostActual',
    'asset': 'accAssetActual'
  };
  
  document.getElementById(cardIds[currentReconcileType]).classList.add('reconciled');
  document.getElementById(actualIds[currentReconcileType]).textContent = `Actual: ${actualValue}`;
  
  closeModal('reconcileModal');
}

async function submitDailyReconcile() {
  if (Object.keys(currentReconcileData).length === 0) {
    alert('Please reconcile at least one item before submitting');
    return;
  }
  
  if (!confirm('Submit daily reconciliation?')) return;
  
  try {
    showLoading();
    const result = await callAppsScript('ADD_RECONCILE', {
      data: JSON.stringify(currentReconcileData)
    });
    
    if (result.success) {
      alert('✅ Daily reconciliation submitted successfully!');
      currentReconcileData = {};
      
      document.querySelectorAll('.stat-card.reconciled').forEach(card => {
        card.classList.remove('reconciled');
      });
      document.getElementById('accTransActual').textContent = '';
      document.getElementById('accCashActual').textContent = '';
      document.getElementById('accRevActual').textContent = '';
      document.getElementById('accPLActual').textContent = '';
      document.getElementById('accCostActual').textContent = '';
      document.getElementById('accAssetActual').textContent = '';
      
      loadAccounting();
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}
