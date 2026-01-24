async function loadCashBank() {
  try {
    showLoading();
    const data = await fetchSheetData('CashBank!A:I');
    
    let balances = {
      cash: { LAK: 0, THB: 0, USD: 0 },
      bcel: { LAK: 0, THB: 0, USD: 0 },
      ldb: { LAK: 0, THB: 0, USD: 0 }
    };
    
    if (data.length > 1) {
      data.slice(1).forEach(row => {
        const type = row[1];
        const amount = parseFloat(row[2]) || 0;
        const currency = row[3] || 'LAK';
        const method = row[4];
        const bank = row[5];
        
        if (type === 'CASH_IN') {
          balances.cash[currency] += amount;
        } else if (type === 'CASH_OUT') {
          balances.cash[currency] -= amount;
        } else if (type === 'BANK_DEPOSIT') {
          if (bank === 'BCEL') balances.bcel[currency] += amount;
          else if (bank === 'LDB') balances.ldb[currency] += amount;
        } else if (type === 'BANK_WITHDRAW') {
          if (bank === 'BCEL') balances.bcel[currency] -= amount;
          else if (bank === 'LDB') balances.ldb[currency] -= amount;
        } else if (type === 'OTHER_INCOME') {
          if (method === 'CASH') {
            balances.cash[currency] += amount;
          } else if (method === 'BANK') {
            if (bank === 'BCEL') balances.bcel[currency] += amount;
            else if (bank === 'LDB') balances.ldb[currency] += amount;
          }
        } else if (type === 'OTHER_EXPENSE' || type === 'BUYBACK') {
          if (method === 'CASH' || method === 'Cash') {
            balances.cash[currency] -= amount;
          } else if (method === 'BANK' || method === 'Bank') {
            if (bank === 'BCEL') balances.bcel[currency] -= amount;
            else if (bank === 'LDB') balances.ldb[currency] -= amount;
          }
        } else if (type === 'SELL' || type === 'TRADEIN' || type === 'EXCHANGE') {
          if (method === 'Cash') {
            balances.cash[currency] += amount;
          } else if (method === 'Bank') {
            if (bank === 'BCEL') balances.bcel[currency] += amount;
            else if (bank === 'LDB') balances.ldb[currency] += amount;
          }
        } else if (type === 'SELL_CHANGE' || type === 'TRADEIN_CHANGE' || type === 'EXCHANGE_CHANGE') {
          balances.cash[currency] -= amount;
        }
      });
    }
    
    document.getElementById('cashLAK').textContent = formatNumber(balances.cash.LAK);
    document.getElementById('cashTHB').textContent = formatNumber(balances.cash.THB);
    document.getElementById('cashUSD').textContent = formatNumber(balances.cash.USD);
    
    document.getElementById('bcelLAK').textContent = formatNumber(balances.bcel.LAK);
    document.getElementById('bcelTHB').textContent = formatNumber(balances.bcel.THB);
    document.getElementById('bcelUSD').textContent = formatNumber(balances.bcel.USD);
    
    document.getElementById('ldbLAK').textContent = formatNumber(balances.ldb.LAK);
    document.getElementById('ldbTHB').textContent = formatNumber(balances.ldb.THB);
    document.getElementById('ldbUSD').textContent = formatNumber(balances.ldb.USD);
    
    const tbody = document.getElementById('cashbankTable');
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = data.slice(1).reverse().map(row => `
        <tr>
          <td>${row[0]}</td>
          <td>${row[1]}</td>
          <td>${formatNumber(row[2])}</td>
          <td>${row[3] || '-'}</td>
          <td>${row[4]}</td>
          <td>${row[5] || '-'}</td>
          <td>${row[6] || '-'}</td>
          <td>${formatDateTime(row[7])}</td>
        </tr>
      `).join('');
    }
    
    hideLoading();
  } catch (error) {
    console.error('Error loading cashbank:', error);
    hideLoading();
  }
}

function toggleOtherIncomeBank() {
  const method = document.getElementById('otherIncomeMethod').value;
  const bankGroup = document.getElementById('otherIncomeBankGroup');
  bankGroup.style.display = method === 'BANK' ? 'block' : 'none';
}

function toggleOtherExpenseBank() {
  const method = document.getElementById('otherExpenseMethod').value;
  const bankGroup = document.getElementById('otherExpenseBankGroup');
  bankGroup.style.display = method === 'BANK' ? 'block' : 'none';
}

async function submitCash() {
  const type = document.getElementById('cashType').value;
  const amount = document.getElementById('cashAmount').value;
  const currency = document.getElementById('cashCurrency').value;
  const note = document.getElementById('cashNote').value;
  
  if (!amount || amount <= 0) {
    alert('Please enter amount');
    return;
  }
  
  try {
    showLoading();
    const result = await callAppsScript('ADD_CASHBANK', {
      type: type === 'IN' ? 'CASH_IN' : 'CASH_OUT',
      amount,
      currency,
      method: 'CASH',
      bank: '',
      note
    });
    
    if (result.success) {
      alert('✅ Transaction added successfully!');
      closeModal('cashModal');
      document.getElementById('cashAmount').value = '';
      document.getElementById('cashNote').value = '';
      loadCashBank();
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}

async function submitBank() {
  const type = document.getElementById('bankType').value;
  const bank = document.getElementById('bankName').value;
  const amount = document.getElementById('bankAmount').value;
  const currency = document.getElementById('bankCurrency').value;
  const note = document.getElementById('bankNote').value;
  
  if (!amount || amount <= 0) {
    alert('Please enter amount');
    return;
  }
  
  try {
    showLoading();
    const result = await callAppsScript('ADD_CASHBANK', {
      type: type === 'DEPOSIT' ? 'BANK_DEPOSIT' : 'BANK_WITHDRAW',
      amount,
      currency,
      method: 'BANK',
      bank,
      note
    });
    
    if (result.success) {
      alert('✅ Transaction added successfully!');
      closeModal('bankModal');
      document.getElementById('bankAmount').value = '';
      document.getElementById('bankNote').value = '';
      loadCashBank();
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}

async function submitOtherIncome() {
  const method = document.getElementById('otherIncomeMethod').value;
  const bank = method === 'BANK' ? document.getElementById('otherIncomeBank').value : '';
  const amount = document.getElementById('otherIncomeAmount').value;
  const currency = document.getElementById('otherIncomeCurrency').value;
  const note = document.getElementById('otherIncomeNote').value;
  
  if (!amount || amount <= 0) {
    alert('Please enter amount');
    return;
  }
  
  try {
    showLoading();
    const result = await callAppsScript('ADD_CASHBANK', {
      type: 'OTHER_INCOME',
      amount,
      currency,
      method,
      bank,
      note
    });
    
    if (result.success) {
      alert('✅ Transaction added successfully!');
      closeModal('otherIncomeModal');
      document.getElementById('otherIncomeAmount').value = '';
      document.getElementById('otherIncomeNote').value = '';
      loadCashBank();
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}

async function submitOtherExpense() {
  const method = document.getElementById('otherExpenseMethod').value;
  const bank = method === 'BANK' ? document.getElementById('otherExpenseBank').value : '';
  const amount = document.getElementById('otherExpenseAmount').value;
  const currency = document.getElementById('otherExpenseCurrency').value;
  const note = document.getElementById('otherExpenseNote').value;
  
  if (!amount || amount <= 0) {
    alert('Please enter amount');
    return;
  }
  
  try {
    showLoading();
    const result = await callAppsScript('ADD_CASHBANK', {
      type: 'OTHER_EXPENSE',
      amount,
      currency,
      method,
      bank,
      note
    });
    
    if (result.success) {
      alert('✅ Transaction added successfully!');
      closeModal('otherExpenseModal');
      document.getElementById('otherExpenseAmount').value = '';
      document.getElementById('otherExpenseNote').value = '';
      loadCashBank();
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}
