async function loadCashBank() {
  try {
    showLoading();
    
    const [cashbankData, dbData] = await Promise.all([
      fetchSheetData('CashBank!A:I'),
      fetchSheetData('_database!A1:G23')
    ]);
    
    let balances = {
      cash: { LAK: 0, THB: 0, USD: 0 },
      bcel: { LAK: 0, THB: 0, USD: 0 },
      ldb: { LAK: 0, THB: 0, USD: 0 }
    };
    
    if (dbData.length >= 17) {
      balances.cash.LAK = parseFloat(dbData[16][0]) || 0;
      balances.cash.THB = parseFloat(dbData[16][1]) || 0;
      balances.cash.USD = parseFloat(dbData[16][2]) || 0;
    }
    
    if (dbData.length >= 20) {
      balances.bcel.LAK = parseFloat(dbData[19][0]) || 0;
      balances.bcel.THB = parseFloat(dbData[19][1]) || 0;
      balances.bcel.USD = parseFloat(dbData[19][2]) || 0;
    }
    
    if (dbData.length >= 23) {
      balances.ldb.LAK = parseFloat(dbData[22][0]) || 0;
      balances.ldb.THB = parseFloat(dbData[22][1]) || 0;
      balances.ldb.USD = parseFloat(dbData[22][2]) || 0;
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
    if (cashbankData.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = cashbankData.slice(1).reverse().map(row => `
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