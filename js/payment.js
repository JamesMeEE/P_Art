async function openPaymentModal(sellId) {
  try {
    showLoading();
    const data = await fetchSheetData('Sells!A:I');
    const sell = data.slice(1).find(row => row[0] === sellId);
    
    if (!sell) {
      alert('Transaction not found');
      hideLoading();
      return;
    }
    
    currentPaymentData = {
      id: sell[0],
      customer: sell[1],
      items: sell[2],
      total: parseFloat(sell[3]) || 0
    };
    
    const items = formatItemsForDisplay(sell[2]);
    
    document.getElementById('paymentDetails').innerHTML = `
      <div style="background: rgba(212, 175, 55, 0.1); border: 1px solid var(--gold-primary); border-radius: 8px; padding: 15px;">
        <div style="margin-bottom: 10px;">
          <strong>Transaction ID:</strong> ${sell[0]}
        </div>
        <div style="margin-bottom: 10px;">
          <strong>Customer:</strong> ${sell[1]}
        </div>
        <div style="margin-bottom: 10px;">
          <strong>Items:</strong><br>${items.replace(/\n/g, '<br>')}
        </div>
        <div>
          <strong>Total (LAK):</strong> ${formatNumber(sell[3])}
        </div>
      </div>
    `;
    
    document.getElementById('paymentMethod').value = 'CASH';
    document.getElementById('paymentCurrency').value = 'LAK';
    document.getElementById('paymentReceived').value = '';
    document.getElementById('paymentChange').value = '';
    
    updatePaymentMethod();
    calculatePayment();
    
    hideLoading();
    openModal('paymentModal');
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}

function updatePaymentMethod() {
  const method = document.getElementById('paymentMethod').value;
  const bankGroup = document.getElementById('bankSelectionGroup');
  const receivedGroup = document.getElementById('receivedGroup');
  const changeGroup = document.getElementById('changeGroup');
  
  if (method === 'BANK') {
    bankGroup.style.display = 'block';
    receivedGroup.style.display = 'none';
    changeGroup.style.display = 'none';
  } else {
    bankGroup.style.display = 'none';
    receivedGroup.style.display = 'block';
    changeGroup.style.display = 'block';
  }
}

function calculatePayment() {
  if (!currentPaymentData) return;
  
  const currency = document.getElementById('paymentCurrency').value;
  const totalLAK = currentPaymentData.total;
  let amountToPay = totalLAK;
  let rate = 1;
  
  const rateGroup = document.getElementById('exchangeRateGroup');
  
  if (currency === 'THB') {
    rate = currentPriceRates.thbSell;
    amountToPay = totalLAK / rate;
    rateGroup.style.display = 'block';
    document.getElementById('paymentExchangeRate').value = `1 THB = ${formatNumber(rate)} LAK`;
  } else if (currency === 'USD') {
    rate = currentPriceRates.usdSell;
    amountToPay = totalLAK / rate;
    rateGroup.style.display = 'block';
    document.getElementById('paymentExchangeRate').value = `1 USD = ${formatNumber(rate)} LAK`;
  } else {
    rateGroup.style.display = 'none';
  }
  
  document.getElementById('paymentAmount').value = `${formatNumber(amountToPay.toFixed(2))} ${currency}`;
  document.getElementById('paymentAmountLAK').value = formatNumber(totalLAK) + ' LAK';
}

function calculateChange() {
  if (!currentPaymentData) return;
  
  const received = parseFloat(document.getElementById('paymentReceived').value) || 0;
  const currency = document.getElementById('paymentCurrency').value;
  const totalLAK = currentPaymentData.total;
  
  let receivedInLAK = received;
  if (currency === 'THB') {
    receivedInLAK = received * currentPriceRates.thbSell;
  } else if (currency === 'USD') {
    receivedInLAK = received * currentPriceRates.usdSell;
  }
  
  const change = receivedInLAK - totalLAK;
  document.getElementById('paymentChange').value = formatNumber(Math.max(0, change)) + ' LAK';
}

async function confirmPayment() {
  if (!currentPaymentData) return;
  
  const method = document.getElementById('paymentMethod').value;
  const currency = document.getElementById('paymentCurrency').value;
  const received = parseFloat(document.getElementById('paymentReceived').value) || 0;
  
  if (method === 'CASH' && received <= 0) {
    alert('Please enter received amount');
    return;
  }
  
  if (method === 'CASH') {
    const totalLAK = currentPaymentData.total;
    let receivedInLAK = received;
    
    if (currency === 'THB') {
      receivedInLAK = received * currentPriceRates.thbSell;
    } else if (currency === 'USD') {
      receivedInLAK = received * currentPriceRates.usdSell;
    }
    
    if (receivedInLAK < totalLAK) {
      alert('Received amount is less than total');
      return;
    }
  }
  
  try {
    showLoading();
    
    const bank = method === 'BANK' ? document.getElementById('bankSelection').value : '';
    
    const result = await callAppsScript('CONFIRM_SELL_PAYMENT', {
      sellId: currentPaymentData.id,
      method,
      currency,
      bank
    });
    
    if (result.success) {
      alert('✅ Payment confirmed successfully!');
      closeModal('paymentModal');
      currentPaymentData = null;
      loadSells();
      loadDashboard();
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}
