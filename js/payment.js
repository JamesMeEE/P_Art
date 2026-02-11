let currentPaymentData = null;
let paymentItems = { cash: [], bank: [] };

function openMultiPaymentModal(type, id, total, phone, details) {
  currentPaymentData = { type, id, total, phone, details };
  paymentItems = { cash: [], bank: [] };
  
  const title = type === 'BUYBACK' ? 'Payment Confirmation' : 'Receive Confirmation';
  
  document.getElementById('multiPaymentTitle').textContent = title;
  document.getElementById('multiPaymentId').textContent = id;
  document.getElementById('multiPaymentPhone').textContent = phone;
  document.getElementById('multiPaymentDetails').innerHTML = details;
  document.getElementById('multiPaymentTotal').textContent = formatNumber(total) + ' LAK';
  
  document.getElementById('cashPaymentsList').innerHTML = '';
  document.getElementById('bankPaymentsList').innerHTML = '';
  
  document.getElementById('multiPaymentFeeGroup').style.display = 'none';
  document.getElementById('multiPaymentFeeInput').value = '0';
  
  updatePaymentSummary();
  
  openModal('multiPaymentModal');
}

function addCashPayment() {
  const id = Date.now();
  paymentItems.cash.push({ id, currency: 'LAK', amount: 0, rate: 1 });
  renderCashPayments();
}

function addBankPayment() {
  const id = Date.now();
  paymentItems.bank.push({ id, bank: 'BCEL', currency: 'LAK', amount: 0, rate: 1 });
  renderBankPayments();
}

function renderCashPayments() {
  const container = document.getElementById('cashPaymentsList');
  container.innerHTML = paymentItems.cash.map((item, idx) => {
    const lakAmount = item.amount * item.rate;
    return `
    <div class="payment-item" data-id="${item.id}" style="margin-bottom: 10px; padding: 12px; background: var(--bg-light); border-radius: 8px;">
      <div style="display: flex; gap: 10px; align-items: center;">
        <select class="form-select" style="width: 90px;" onchange="updateCashCurrency(${item.id}, this.value)">
          <option value="LAK" ${item.currency === 'LAK' ? 'selected' : ''}>LAK</option>
          <option value="THB" ${item.currency === 'THB' ? 'selected' : ''}>THB</option>
          <option value="USD" ${item.currency === 'USD' ? 'selected' : ''}>USD</option>
        </select>
        <input type="number" class="form-input cash-amount-input" data-id="${item.id}" placeholder="Amount" value="${item.amount || ''}" 
               style="flex: 1;" oninput="updateCashAmountOnly(${item.id}, this.value)">
        <button class="btn-secondary" style="padding: 8px 12px; background: #f44336; color: white;" onclick="removeCashPayment(${item.id})">✕</button>
      </div>
      ${item.currency !== 'LAK' ? `
        <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 8px; padding-top: 8px; border-top: 1px dashed var(--border-color);">
          <span style="font-size: 11px; color: var(--text-secondary);">Rate: 1 ${item.currency} = ${formatNumber(item.rate)} LAK</span>
          <span class="lak-display" style="color: var(--gold-primary); font-weight: bold;">= ${formatNumber(lakAmount)} LAK</span>
        </div>
      ` : ''}
    </div>
  `}).join('');
  updatePaymentSummary();
}

function updateCashAmountOnly(id, value) {
  const item = paymentItems.cash.find(i => i.id === id);
  if (item) {
    item.amount = parseFloat(value) || 0;
    const lakAmount = item.amount * item.rate;
    const container = document.querySelector(`#cashPaymentsList .payment-item[data-id="${id}"]`);
    if (container && item.currency !== 'LAK') {
      const lakSpan = container.querySelector('.lak-display');
      if (lakSpan) lakSpan.textContent = `= ${formatNumber(lakAmount)} LAK`;
    }
    updatePaymentSummary();
  }
}

function renderBankPayments() {
  const container = document.getElementById('bankPaymentsList');
  container.innerHTML = paymentItems.bank.map((item, idx) => {
    const lakAmount = item.amount * item.rate;
    return `
    <div class="payment-item" data-id="${item.id}" style="margin-bottom: 10px; padding: 12px; background: var(--bg-light); border-radius: 8px;">
      <div style="display: flex; gap: 10px; align-items: center; margin-bottom: 8px;">
        <select class="form-select" style="flex: 1;" onchange="updateBankName(${item.id}, this.value)">
          <option value="BCEL" ${item.bank === 'BCEL' ? 'selected' : ''}>BCEL</option>
          <option value="LDB" ${item.bank === 'LDB' ? 'selected' : ''}>LDB</option>
        </select>
        <select class="form-select" style="flex: 1;" onchange="updateBankCurrency(${item.id}, this.value)">
          <option value="LAK" ${item.currency === 'LAK' ? 'selected' : ''}>LAK</option>
          <option value="THB" ${item.currency === 'THB' ? 'selected' : ''}>THB</option>
          <option value="USD" ${item.currency === 'USD' ? 'selected' : ''}>USD</option>
        </select>
        <button class="btn-secondary" style="padding: 8px 12px; background: #f44336; color: white;" onclick="removeBankPayment(${item.id})">✕</button>
      </div>
      <div style="display: flex; gap: 10px; align-items: center;">
        <input type="number" class="form-input bank-amount-input" data-id="${item.id}" placeholder="Amount" value="${item.amount || ''}" 
               style="flex: 1;" oninput="updateBankAmountOnly(${item.id}, this.value)">
      </div>
      ${item.currency !== 'LAK' ? `
        <div style="display: flex; justify-content: space-between; align-items: center; margin-top: 8px; padding-top: 8px; border-top: 1px dashed var(--border-color);">
          <span style="font-size: 11px; color: var(--text-secondary);">Rate: 1 ${item.currency} = ${formatNumber(item.rate)} LAK</span>
          <span class="lak-display" style="color: var(--gold-primary); font-weight: bold;">= ${formatNumber(lakAmount)} LAK</span>
        </div>
      ` : ''}
    </div>
  `}).join('');
  updatePaymentSummary();
}

function updateBankAmountOnly(id, value) {
  const item = paymentItems.bank.find(i => i.id === id);
  if (item) {
    item.amount = parseFloat(value) || 0;
    const lakAmount = item.amount * item.rate;
    const container = document.querySelector(`#bankPaymentsList .payment-item[data-id="${id}"]`);
    if (container && item.currency !== 'LAK') {
      const lakSpan = container.querySelector('.lak-display');
      if (lakSpan) lakSpan.textContent = `= ${formatNumber(lakAmount)} LAK`;
    }
    updatePaymentSummary();
  }
}

function updateCashCurrency(id, value) {
  const item = paymentItems.cash.find(i => i.id === id);
  if (item) {
    item.currency = value;
    item.rate = value === 'LAK' ? 1 : (value === 'THB' ? (currentPaymentData?.type === 'BUYBACK' ? (currentExchangeRates?.THB_Buy || 0) : (currentExchangeRates?.THB_Sell || 0)) : (currentPaymentData?.type === 'BUYBACK' ? (currentExchangeRates?.USD_Buy || 0) : (currentExchangeRates?.USD_Sell || 0)));
    renderCashPayments();
  }
}

function updateCashAmount(id, value) {
  const item = paymentItems.cash.find(i => i.id === id);
  if (item) {
    item.amount = parseFloat(value) || 0;
    renderCashPayments();
  }
}

function removeCashPayment(id) {
  paymentItems.cash = paymentItems.cash.filter(i => i.id !== id);
  renderCashPayments();
}

function updateBankName(id, value) {
  const item = paymentItems.bank.find(i => i.id === id);
  if (item) item.bank = value;
}

function updateBankCurrency(id, value) {
  const item = paymentItems.bank.find(i => i.id === id);
  if (item) {
    item.currency = value;
    item.rate = value === 'LAK' ? 1 : (value === 'THB' ? (currentPaymentData?.type === 'BUYBACK' ? (currentExchangeRates?.THB_Buy || 0) : (currentExchangeRates?.THB_Sell || 0)) : (currentPaymentData?.type === 'BUYBACK' ? (currentExchangeRates?.USD_Buy || 0) : (currentExchangeRates?.USD_Sell || 0)));
    renderBankPayments();
  }
}

function updateBankAmount(id, value) {
  const item = paymentItems.bank.find(i => i.id === id);
  if (item) {
    item.amount = parseFloat(value) || 0;
    renderBankPayments();
  }
}

function removeBankPayment(id) {
  paymentItems.bank = paymentItems.bank.filter(i => i.id !== id);
  renderBankPayments();
}

function updatePaymentSummary() {
  let totalPaid = 0;
  
  paymentItems.cash.forEach(item => {
    totalPaid += item.amount * item.rate;
  });
  
  paymentItems.bank.forEach(item => {
    totalPaid += item.amount * item.rate;
  });
  
  const total = currentPaymentData?.total || 0;
  const remaining = total - totalPaid;
  const change = totalPaid - total;
  
  document.getElementById('multiPaymentPaidTotal').textContent = formatNumber(totalPaid) + ' LAK';
  
  if (remaining > 0) {
    document.getElementById('multiPaymentRemaining').textContent = formatNumber(remaining) + ' LAK';
    document.getElementById('multiPaymentRemaining').style.color = '#f44336';
    document.getElementById('multiPaymentChange').textContent = '0 LAK';
  } else {
    document.getElementById('multiPaymentRemaining').textContent = '0 LAK';
    document.getElementById('multiPaymentRemaining').style.color = '#4caf50';
    document.getElementById('multiPaymentChange').textContent = formatNumber(Math.max(0, change)) + ' LAK';
  }
}

async function confirmMultiPayment() {
  if (!currentPaymentData) return;
  
  let totalPaid = 0;
  paymentItems.cash.forEach(item => { totalPaid += item.amount * item.rate; });
  paymentItems.bank.forEach(item => { totalPaid += item.amount * item.rate; });
  
  const total = currentPaymentData.total;
  
  if (currentPaymentData.type === 'BUYBACK') {
    if (totalPaid <= 0) {
      alert('❌ กรุณากรอกจำนวนเงินที่จ่าย');
      return;
    }
    if (totalPaid > total) {
      alert('❌ จำนวนเงินที่จ่ายมากกว่ายอดคงค้าง!');
      return;
    }
  } else {
    if (totalPaid < total) {
      alert('❌ ยอดชำระยังไม่ครบ! ขาดอีก ' + formatNumber(total - totalPaid) + ' LAK');
      return;
    }
  }
  
  const change = Math.max(0, totalPaid - total);

  if (change > 0) {
    try {
      var dbData = await fetchSheetData('_database!A1:G23');
      var cashLAK = dbData.length >= 17 ? (parseFloat(dbData[16][0]) || 0) : 0;
      if (cashLAK < change) {
        alert('❌ เงินสด LAK ไม่พอทอน! มี ' + formatNumber(cashLAK) + ' LAK แต่ต้องทอน ' + formatNumber(change) + ' LAK');
        hideLoading();
        return;
      }
    } catch(e) {}
  }

  try {
    showLoading();
    
    const actionMap = {
      'SELL': 'CONFIRM_SELL_PAYMENT',
      'TRADEIN': 'CONFIRM_TRADEIN_PAYMENT',
      'EXCHANGE': 'CONFIRM_EXCHANGE_PAYMENT',
      'WITHDRAW': 'CONFIRM_WITHDRAW_PAYMENT',
      'BUYBACK': 'CONFIRM_BUYBACK_PAYMENT',
      'SWITCH': 'CONFIRM_SWITCH_PAYMENT',
      'FREE_EXCHANGE': 'CONFIRM_FREE_EXCHANGE_PAYMENT'
    };
    
    const action = actionMap[currentPaymentData.type];
    
    const idParamMap = {
      'SELL': 'sellId',
      'TRADEIN': 'tradeinId',
      'EXCHANGE': 'exchangeId',
      'WITHDRAW': 'id',
      'BUYBACK': 'buybackId',
      'SWITCH': 'switchId',
      'FREE_EXCHANGE': 'freeExId'
    };
    
    const params = {
      [idParamMap[currentPaymentData.type]]: currentPaymentData.id,
      payments: JSON.stringify({ cash: paymentItems.cash, bank: paymentItems.bank }),
      totalPaid: totalPaid,
      change: change,
      user: currentUser.nickname
    };
    if (currentPaymentData.type === 'BUYBACK') {
      params.fee = parseFloat(document.getElementById('multiPaymentFeeInput').value) || 0;
      params.items = currentPaymentData.items || '';
      params.total = currentPaymentData.total || 0;
    }
    const result = await callAppsScript(action, params);
    
    if (result.success) {
      alert('✅ ยืนยันการชำระเงินสำเร็จ!');
      closeModal('multiPaymentModal');
      currentPaymentData = null;
      paymentItems = { cash: [], bank: [] };
      
      loadDashboard();
      if (typeof loadSells === 'function') loadSells();
      if (typeof loadTradeins === 'function') loadTradeins();
      if (typeof loadExchanges === 'function') loadExchanges();
      if (typeof loadWithdraws === 'function') loadWithdraws();
      if (typeof loadBuybacks === 'function') loadBuybacks();
      if (typeof loadSwitches === 'function') loadSwitches();
      if (typeof loadFreeExchanges === 'function') loadFreeExchanges();
    } else {
      alert('❌ เกิดข้อผิดพลาด: ' + result.message);
    }
    
    hideLoading();
  } catch (error) {
    alert('❌ เกิดข้อผิดพลาด: ' + error.message);
    hideLoading();
  }
}

async function openSellPayment(sellId) {
  const data = await fetchSheetData('Sells!A:L');
  const sell = data.slice(1).find(row => row[0] === sellId);
  if (!sell) return;
  
  const items = formatItemsForPayment(sell[2]);
  const total = parseFloat(sell[3]) || 0;
  
  openMultiPaymentModal('SELL', sellId, total, sell[1], `<strong>Items:</strong> ${items}`);
}

async function openTradeinPaymentModal(tradeinId) {
  const data = await fetchSheetData('Tradeins!A:N');
  const tradein = data.slice(1).find(row => row[0] === tradeinId);
  if (!tradein) return;
  
  const oldGold = formatItemsForPayment(tradein[2]);
  const newGold = formatItemsForPayment(tradein[3]);
  const total = parseFloat(tradein[6]) || 0;
  
  openMultiPaymentModal('TRADEIN', tradeinId, total, tradein[1], 
    `<strong>Old Gold:</strong> ${oldGold}<br><strong>New Gold:</strong> ${newGold}`);
}

async function openExchangePaymentModal(exchangeId) {
  const data = await fetchSheetData('Exchanges!A:N');
  const exchange = data.slice(1).find(row => row[0] === exchangeId);
  if (!exchange) return;
  
  const oldGold = formatItemsForPayment(exchange[2]);
  const newGold = formatItemsForPayment(exchange[3]);
  const total = parseFloat(exchange[6]) || 0;
  
  openMultiPaymentModal('EXCHANGE', exchangeId, total, exchange[1], 
    `<strong>Old Gold:</strong> ${oldGold}<br><strong>New Gold:</strong> ${newGold}`);
}

async function openWithdrawPaymentModal(withdrawId) {
  const data = await fetchSheetData('Withdraws!A:J');
  const withdraw = data.slice(1).find(row => row[0] === withdrawId);
  if (!withdraw) return;
  
  const items = formatItemsForPayment(withdraw[2]);
  const total = parseFloat(withdraw[4]) || 0;
  
  openMultiPaymentModal('WITHDRAW', withdrawId, total, withdraw[1], `<strong>Items:</strong> ${items}`);
}

function formatItemsForPayment(itemsJson) {
  try {
    const items = typeof itemsJson === 'string' ? JSON.parse(itemsJson) : itemsJson;
    return items.map(item => {
      const product = FIXED_PRODUCTS.find(p => p.id === item.productId);
      return `${product?.name || item.productId} x${item.qty}`;
    }).join(', ');
  } catch (e) {
    return '-';
  }
}

async function openBuybackPaymentModalFromList(buybackId) {
  const data = await fetchSheetData('Buybacks!A:L');
  const buyback = data.slice(1).find(row => row[0] === buybackId);
  if (!buyback) return;
  
  const price = parseFloat(buyback[6]) || 0;
  const paid = parseFloat(buyback[7]) || 0;
  const balance = price - paid;
  
  currentPaymentData = { 
    type: 'BUYBACK', 
    id: buybackId, 
    total: balance,
    basePrice: price,
    paid: paid,
    phone: buyback[1],
    items: buyback[2],
    details: `<strong>Items:</strong> ${formatItemsForPayment(buyback[2])}<br>
              <strong>Buyback Price:</strong> ${formatNumber(price)} LAK<br>
              <strong>ยอดคงค้าง:</strong> ${formatNumber(balance)} LAK`,
    allowPartial: true
  };
  paymentItems = { cash: [], bank: [] };
  
  document.getElementById('multiPaymentTitle').textContent = 'Payment Confirmation (Buyback)';
  document.getElementById('multiPaymentId').textContent = buybackId;
  document.getElementById('multiPaymentPhone').textContent = buyback[1];
  document.getElementById('multiPaymentDetails').innerHTML = currentPaymentData.details;
  document.getElementById('multiPaymentTotal').textContent = formatNumber(balance) + ' LAK';
  
  document.getElementById('multiPaymentFeeGroup').style.display = 'block';
  document.getElementById('multiPaymentFeeInput').value = '';
  document.getElementById('multiPaymentFeeInput').placeholder = '0';
  
  document.getElementById('cashPaymentsList').innerHTML = '';
  document.getElementById('bankPaymentsList').innerHTML = '';
  
  updatePaymentSummary();
  
  openModal('multiPaymentModal');
}

function onBuybackFeeChange() {
}