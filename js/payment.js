async function openPaymentModal(sellId) {
  try {
    showLoading();
    
    // โหลดอัตราแลกเปลี่ยนก่อน
    if (currentPriceRates.thbSell === 0 || currentPriceRates.usdSell === 0) {
      const rateData = await fetchSheetData('PriceRate!A:F');
      if (rateData.length > 1) {
        const latestRate = rateData[rateData.length - 1];
        currentPriceRates = {
          thbSell: parseFloat(latestRate[1]) || 0,
          usdSell: parseFloat(latestRate[2]) || 0,
          thbBuy: parseFloat(latestRate[3]) || 0,
          usdBuy: parseFloat(latestRate[4]) || 0
        };
      }
    }
    
    const data = await fetchSheetData('Sells!A:L');
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
          <strong>Phone:</strong> ${sell[1]}
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
  const receivedLabel = document.getElementById('receivedAmountLabel');
  
  if (currency === 'THB') {
    rate = currentPriceRates.thbSell || 270;
    amountToPay = Math.ceil((totalLAK / rate) * 100) / 100;
    rateGroup.style.display = 'block';
    document.getElementById('paymentExchangeRate').value = `1 THB = ${formatNumber(rate)} LAK`;
    if (receivedLabel) receivedLabel.textContent = 'Received Amount (THB)';
  } else if (currency === 'USD') {
    rate = currentPriceRates.usdSell || 21500;
    amountToPay = Math.ceil((totalLAK / rate) * 100) / 100;
    rateGroup.style.display = 'block';
    document.getElementById('paymentExchangeRate').value = `1 USD = ${formatNumber(rate)} LAK`;
    if (receivedLabel) receivedLabel.textContent = 'Received Amount (USD)';
  } else {
    rateGroup.style.display = 'none';
    if (receivedLabel) receivedLabel.textContent = 'Received Amount (LAK)';
  }
  
  if (currency === 'LAK') {
    document.getElementById('paymentAmount').value = `${formatNumber(amountToPay)} ${currency}`;
  } else {
    document.getElementById('paymentAmount').value = `${amountToPay.toFixed(2)} ${currency}`;
  }
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
  
  let customerPaid = 0;
  let exchangeRate = 1;
  
  if (method === 'CASH') {
    if (received <= 0) {
      alert('Please enter received amount');
      return;
    }
    
    const totalLAK = currentPaymentData.total;
    let receivedInLAK = received;
    
    if (currency === 'THB') {
      exchangeRate = currentPriceRates.thbSell || 270;
      receivedInLAK = received * exchangeRate;
      customerPaid = received;
    } else if (currency === 'USD') {
      exchangeRate = currentPriceRates.usdSell || 21500;
      receivedInLAK = received * exchangeRate;
      customerPaid = received;
    } else {
      customerPaid = received;
      exchangeRate = 1;
    }
    
    if (receivedInLAK < totalLAK) {
      alert('Received amount is less than total');
      return;
    }
  } else {
    // BANK - ใช้ total เป็น customerPaid
    customerPaid = currentPaymentData.total;
    if (currency === 'THB') {
      exchangeRate = currentPriceRates.thbSell || 270;
      customerPaid = currentPaymentData.total / exchangeRate;
    } else if (currency === 'USD') {
      exchangeRate = currentPriceRates.usdSell || 21500;
      customerPaid = currentPaymentData.total / exchangeRate;
    }
  }
  
  try {
    showLoading();
    
    const bank = method === 'BANK' ? document.getElementById('bankSelection').value : '';
    
    const result = await callAppsScript('CONFIRM_SELL_PAYMENT', {
      sellId: currentPaymentData.id,
      method,
      currency,
      bank,
      customerPaid: customerPaid,
      customerCurrency: currency,
      exchangeRate: exchangeRate
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