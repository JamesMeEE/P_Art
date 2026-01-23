async function loadCashBank() {
  try {
    showLoading();
    const [cashbankData, sellData, buybackData, tradeinData, exchangeData] = await Promise.all([
      fetchSheetData('CashBank!A:G'),
      fetchSheetData('Sells!A:I'),
      fetchSheetData('Buybacks!A:G'),
      fetchSheetData('Tradeins!A:K'),
      fetchSheetData('Exchanges!A:J')
    ]);
    
    let cash = 0;
    let bank = 0;
    
    if (cashbankData.length > 1) {
      cashbankData.slice(1).forEach(row => {
        const amount = parseFloat(row[2]) || 0;
        const method = row[3];
        const type = row[1];
        
        if (type === 'OWNER_DEPOSIT' || type === 'OTHER_INCOME') {
          if (method === 'CASH') cash += amount;
          else if (method === 'BANK') bank += amount;
        } else if (type === 'BANK_DEPOSIT') {
          cash -= amount;
          bank += amount;
        } else if (type === 'OTHER_EXPENSE') {
          if (method === 'CASH') cash -= amount;
          else if (method === 'BANK') bank -= amount;
        }
      });
    }
    
    sellData.slice(1).forEach(row => {
      if (row[7] === 'COMPLETED') {
        const amount = parseFloat(row[3]) || 0;
        const currency = row[4];
        if (currency === 'LAK') cash += amount;
        else bank += amount;
      }
    });
    
    buybackData.slice(1).forEach(row => {
      if (row[5] === 'COMPLETED') {
        cash -= parseFloat(row[3]) || 0;
      }
    });
    
    tradeinData.slice(1).forEach(row => {
      if (row[8] === 'COMPLETED') {
        const diff = parseFloat(row[4]) || 0;
        const fee = parseFloat(row[5]) || 0;
        const premium = parseFloat(row[6]) || 0;
        cash += (diff + fee + premium);
      }
    });
    
    exchangeData.slice(1).forEach(row => {
      if (row[7] === 'COMPLETED') {
        const fee = parseFloat(row[4]) || 0;
        const premium = parseFloat(row[5]) || 0;
        cash += (fee + premium);
      }
    });
    
    document.getElementById('cashBalance').textContent = formatNumber(cash);
    document.getElementById('bankBalance').textContent = formatNumber(bank);
    document.getElementById('totalBalance').textContent = formatNumber(cash + bank);
    
    const tbody = document.getElementById('cashbankTable');
    if (cashbankData.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="7" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = cashbankData.slice(1).reverse().map(row => `
        <tr>
          <td>${row[0]}</td>
          <td>${row[1]}</td>
          <td>${formatNumber(row[2])}</td>
          <td>${row[3]}</td>
          <td>${row[4] || '-'}</td>
          <td>${formatDateOnly(row[5])}</td>
          <td>${row[6]}</td>
        </tr>
      `).join('');
    }
    
    hideLoading();
  } catch (error) {
    console.error('Error loading cash/bank:', error);
    hideLoading();
  }
}

async function submitOwnerDeposit() {
  const amount = document.getElementById('ownerDepositAmount').value;
  const method = document.getElementById('ownerDepositMethod').value;
  const note = document.getElementById('ownerDepositNote').value;
  
  if (!amount) {
    alert('Please enter amount');
    return;
  }
  
  try {
    showLoading();
    const result = await callAppsScript('ADD_CASHBANK', {
      type: 'OWNER_DEPOSIT',
      amount,
      method,
      note
    });
    
    if (result.success) {
      alert('✅ Owner deposit recorded!');
      closeModal('ownerDepositModal');
      document.getElementById('ownerDepositAmount').value = '';
      document.getElementById('ownerDepositNote').value = '';
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
  const amount = document.getElementById('otherIncomeAmount').value;
  const method = document.getElementById('otherIncomeMethod').value;
  const note = document.getElementById('otherIncomeNote').value;
  
  if (!amount) {
    alert('Please enter amount');
    return;
  }
  
  try {
    showLoading();
    const result = await callAppsScript('ADD_CASHBANK', {
      type: 'OTHER_INCOME',
      amount,
      method,
      note
    });
    
    if (result.success) {
      alert('✅ Other income recorded!');
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

async function submitBankDeposit() {
  const amount = document.getElementById('bankDepositAmount').value;
  const note = document.getElementById('bankDepositNote').value;
  
  if (!amount) {
    alert('Please enter amount');
    return;
  }
  
  try {
    showLoading();
    const result = await callAppsScript('ADD_CASHBANK', {
      type: 'BANK_DEPOSIT',
      amount,
      method: 'BANK',
      note
    });
    
    if (result.success) {
      alert('✅ Bank deposit recorded!');
      closeModal('bankDepositModal');
      document.getElementById('bankDepositAmount').value = '';
      document.getElementById('bankDepositNote').value = '';
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
  const amount = document.getElementById('otherExpenseAmount').value;
  const method = document.getElementById('otherExpenseMethod').value;
  const note = document.getElementById('otherExpenseNote').value;
  
  if (!amount) {
    alert('Please enter amount');
    return;
  }
  
  try {
    showLoading();
    const result = await callAppsScript('ADD_CASHBANK', {
      type: 'OTHER_EXPENSE',
      amount,
      method,
      note
    });
    
    if (result.success) {
      alert('✅ Other expense recorded!');
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
