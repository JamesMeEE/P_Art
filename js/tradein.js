async function loadTradeins() {
  try {
    showLoading();
    const data = await fetchSheetData('Tradeins!A:N');
    
    let filteredData = data.slice(1);
    
    if (currentUser.role === 'User' || currentUser.role === 'Manager') {
      filteredData = filterTodayData(filteredData, 11, 13);
    }
    
    const tbody = document.getElementById('tradeinTable');
    if (filteredData.length === 0) {
      tbody.innerHTML = '<tr><td colspan="10" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = filteredData.reverse().map(row => {
        const oldGold = formatItemsForTable(row[2]);
        const newGold = formatItemsForTable(row[3]);
        const premium = calculatePremiumFromItems(row[3]);
        const saleName = row[13];
        const status = row[12];
        
        let actions = '';
        
        if (status === 'PENDING') {
          if (currentUser.role === 'Manager') {
            actions = `<button class="btn-action" onclick="reviewTradein('${row[0]}')">Review</button>`;
          } else {
            actions = '<span style="color: var(--text-secondary);">Waiting for review</span>';
          }
        } else if (status === 'READY') {
          if (currentUser.role === 'User') {
            actions = `<button class="btn-action" onclick="openTradeinPaymentModal('${row[0]}')">Confirm</button>`;
          } else {
            actions = '<span style="color: var(--text-secondary);">Waiting for confirmation</span>';
          }
        } else {
          actions = '-';
        }
        
        return `
          <tr>
            <td>${row[0]}</td>
            <td>${row[1]}</td>
            <td>${oldGold}</td>
            <td>${newGold}</td>
            <td>${formatNumber(row[4])}</td>
            <td>${formatNumber(premium)}</td>
            <td>${formatNumber(row[6])}</td>
            <td><span class="status-badge status-${status.toLowerCase()}">${status}</span></td>
            <td>${saleName}</td>
            <td>${actions}</td>
          </tr>
        `;
      }).join('');
    }
    
    hideLoading();
  } catch (error) {
    hideLoading();
  }
}

function addTradeinOldGold() {
  tradeinOldCounter++;
  const container = document.getElementById('tradeinOldGold');
  const row = document.createElement('div');
  row.className = 'product-row';
  row.id = `tradeinOld${tradeinOldCounter}`;
  row.innerHTML = `
    <select class="form-select">
      <option value="">เลือกสินค้า...</option>
      ${FIXED_PRODUCTS.map(p => `<option value="${p.id}">${p.name}</option>`).join('')}
    </select>
    <input type="number" class="form-input" placeholder="จำนวน" min="1" step="1">
    <button type="button" class="btn-remove" onclick="removeTradeinOldGold(${tradeinOldCounter})">×</button>
  `;
  container.appendChild(row);
}

function removeTradeinOldGold(id) {
  const row = document.getElementById(`tradeinOld${id}`);
  if (row) row.remove();
}

function addTradeinNewGold() {
  tradeinNewCounter++;
  const container = document.getElementById('tradeinNewGold');
  const row = document.createElement('div');
  row.className = 'product-row';
  row.id = `tradeinNew${tradeinNewCounter}`;
  row.innerHTML = `
    <select class="form-select">
      <option value="">เลือกสินค้า...</option>
      ${FIXED_PRODUCTS.map(p => `<option value="${p.id}">${p.name}</option>`).join('')}
    </select>
    <input type="number" class="form-input" placeholder="จำนวน" min="1" step="1">
    <button type="button" class="btn-remove" onclick="removeTradeinNewGold(${tradeinNewCounter})">×</button>
  `;
  container.appendChild(row);
}

function removeTradeinNewGold(id) {
  const row = document.getElementById(`tradeinNew${id}`);
  if (row) row.remove();
}

function calculateTradein() {
  const customer = document.getElementById('tradeinCustomer').value;
  if (!customer) {
    alert('กรุณากรอกชื่อลูกค้า');
    return;
  }

  const oldGold = [];
  document.querySelectorAll('#tradeinOldGold .product-row').forEach(row => {
    const productId = row.querySelector('select').value;
    const qty = parseInt(row.querySelector('input').value) || 0;
    if (productId && qty > 0) {
      oldGold.push({ productId, qty });
    }
  });

  const newGold = [];
  document.querySelectorAll('#tradeinNewGold .product-row').forEach(row => {
    const productId = row.querySelector('select').value;
    const qty = parseInt(row.querySelector('input').value) || 0;
    if (productId && qty > 0) {
      newGold.push({ productId, qty });
    }
  });

  if (oldGold.length === 0 || newGold.length === 0) {
    alert('กรุณาเลือกทองเก่าและทองใหม่');
    return;
  }

  let oldValue = 0;
  oldGold.forEach(item => {
    const pricePerPiece = calculateBuybackPrice(item.productId, currentPricing.sell1Baht);
    oldValue += pricePerPiece * item.qty;
  });

  let newValue = 0;
  let exchangeFee = 0;
  let premium = 0;

  newGold.forEach(item => {
    const pricePerPiece = calculateSellPrice(item.productId, currentPricing.sell1Baht);
    newValue += pricePerPiece * item.qty;
    exchangeFee += EXCHANGE_FEES[item.productId] * item.qty;
    
    if (PREMIUM_PRODUCTS.includes(item.productId)) {
      premium += PREMIUM_PER_PIECE * item.qty;
    }
  });

  const difference = newValue - oldValue;
  const total = roundTo1000(difference + exchangeFee + premium);

  currentTradeinData = {
    customer,
    oldGold: JSON.stringify(oldGold),
    newGold: JSON.stringify(newGold),
    difference,
    exchangeFee,
    premium,
    total
  };

  const oldItems = oldGold.map(i => `${FIXED_PRODUCTS.find(p => p.id === i.productId).name} (${i.qty})`).join(', ');
  const newItems = newGold.map(i => `${FIXED_PRODUCTS.find(p => p.id === i.productId).name} (${i.qty})`).join(', ');

  document.getElementById('tradeinResult').innerHTML = `
    <div class="summary-box">
      <div class="summary-row">
        <span class="summary-label">ทองเก่า:</span>
        <span class="summary-value">${oldItems}</span>
      </div>
      <div class="summary-row">
        <span class="summary-label">มูลค่าทองเก่า:</span>
        <span class="summary-value">${formatNumber(oldValue)} LAK</span>
      </div>
      <div class="summary-row">
        <span class="summary-label">ทองใหม่:</span>
        <span class="summary-value">${newItems}</span>
      </div>
      <div class="summary-row">
        <span class="summary-label">มูลค่าทองใหม่:</span>
        <span class="summary-value">${formatNumber(newValue)} LAK</span>
      </div>
      <div class="summary-row">
        <span class="summary-label">ส่วนต่าง:</span>
        <span class="summary-value">${formatNumber(difference)} LAK</span>
      </div>
      <div class="summary-row">
        <span class="summary-label">ค่าธรรมเนียมแลก:</span>
        <span class="summary-value">${formatNumber(exchangeFee)} LAK</span>
      </div>
      <div class="summary-row">
        <span class="summary-label">Premium:</span>
        <span class="summary-value">${formatNumber(premium)} LAK</span>
      </div>
      <div class="summary-row" style="border-top: 2px solid var(--gold-primary); padding-top: 10px; margin-top: 10px;">
        <span class="summary-label" style="font-size: 18px; font-weight: 600;">รวมที่ต้องชำระ:</span>
        <span class="summary-value" style="font-size: 20px; font-weight: 600; color: var(--gold-primary);">${formatNumber(total)} LAK</span>
      </div>
    </div>
  `;

  closeModal('tradeinModal');
  openModal('tradeinResultModal');
}

function backToTradein() {
  closeModal('tradeinResultModal');
  openModal('tradeinModal');
  
  if (currentTradeinData) {
    document.getElementById('tradeinCustomer').value = currentTradeinData.customer;
    
    const oldGold = JSON.parse(currentTradeinData.oldGold);
    const newGold = JSON.parse(currentTradeinData.newGold);
    
    document.getElementById('tradeinOldGold').innerHTML = '';
    tradeinOldCounter = 0;
    oldGold.forEach(item => {
      addTradeinOldGold();
      const rows = document.querySelectorAll('#tradeinOldGold .product-row');
      const lastRow = rows[rows.length - 1];
      lastRow.querySelector('select').value = item.productId;
      lastRow.querySelector('input').value = item.qty;
    });
    
    document.getElementById('tradeinNewGold').innerHTML = '';
    tradeinNewCounter = 0;
    newGold.forEach(item => {
      addTradeinNewGold();
      const rows = document.querySelectorAll('#tradeinNewGold .product-row');
      const lastRow = rows[rows.length - 1];
      lastRow.querySelector('select').value = item.productId;
      lastRow.querySelector('input').value = item.qty;
    });
  }
}

async function submitTradein() {
  if (!currentTradeinData) return;

  const newGold = JSON.parse(currentTradeinData.newGold);
  
  showLoading();
  for (const item of newGold) {
    const hasStock = await checkStock(item.productId, item.qty);
    if (!hasStock) {
      const currentStock = await getCurrentStock(item.productId);
      const productName = FIXED_PRODUCTS.find(p => p.id === item.productId).name;
      hideLoading();
      alert(`❌ สินค้าไม่พอสำหรับ ${productName}!\nต้องการ: ${item.qty}, มีอยู่: ${currentStock}`);
      return;
    }
  }
  hideLoading();

  try {
    showLoading();
    const result = await callAppsScript('ADD_TRADEIN', currentTradeinData);

    if (result.success) {
      alert('✅ สร้างรายการแลกเปลี่ยนสำเร็จ!');
      closeModal('tradeinResultModal');
      currentTradeinData = null;
      loadTradeins();
      loadDashboard();
    } else {
      alert('❌ เกิดข้อผิดพลาด: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ เกิดข้อผิดพลาด: ' + error.message);
    hideLoading();
  }
}

async function reviewTradein(tradeinId) {
  if (!confirm('ยืนยันการ Review รายการแลกเปลี่ยนนี้?')) return;
  
  try {
    showLoading();
    const result = await callAppsScript('REVIEW_TRADEIN', { tradeinId });
    
    if (result.success) {
      alert('✅ Review สำเร็จ! รอ User ยืนยันชำระเงิน');
      loadTradeins();
    } else {
      alert('❌ เกิดข้อผิดพลาด: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ เกิดข้อผิดพลาด: ' + error.message);
    hideLoading();
  }
}

async function openTradeinPaymentModal(tradeinId) {
  try {
    showLoading();
    const data = await fetchSheetData('Tradeins!A:N');
    const tradein = data.slice(1).find(row => row[0] === tradeinId);
    
    if (!tradein) {
      alert('❌ ไม่พบรายการแลกเปลี่ยน');
      hideLoading();
      return;
    }
    
    currentTradeinPayment = {
      tradeinId: tradein[0],
      customer: tradein[1],
      oldGold: tradein[2],
      newGold: tradein[3],
      total: parseFloat(tradein[6]) || 0
    };
    
    document.getElementById('tradeinPaymentId').textContent = tradein[0];
    document.getElementById('tradeinPaymentCustomer').textContent = tradein[1];
    document.getElementById('tradeinPaymentOldGold').textContent = formatItemsForDisplay(tradein[2]);
    document.getElementById('tradeinPaymentNewGold').textContent = formatItemsForDisplay(tradein[3]);
    document.getElementById('tradeinPaymentTotal').textContent = formatNumber(tradein[6]) + ' LAK';
    
    document.getElementById('tradeinPaymentCurrency').value = 'LAK';
    document.getElementById('tradeinPaymentMethod').value = 'Cash';
    document.getElementById('tradeinPaymentBankGroup').style.display = 'none';
    
    calculateTradeinPayment();
    
    hideLoading();
    openModal('tradeinPaymentModal');
  } catch (error) {
    alert('❌ เกิดข้อผิดพลาด: ' + error.message);
    hideLoading();
  }
}

function calculateTradeinPayment() {
  if (!currentTradeinPayment) return;
  
  const currency = document.getElementById('tradeinPaymentCurrency').value;
  const totalLAK = currentTradeinPayment.total;
  
  let rate = 1;
  let amountToPay = totalLAK;
  
  const rateGroup = document.getElementById('tradeinPaymentRateGroup');
  const receivedLabel = document.querySelector('label[for="tradeinPaymentReceived"]');
  
  if (currency === 'THB') {
    rate = parseFloat(document.getElementById('tradeinPaymentRateTHB').value) || 270;
    amountToPay = totalLAK / rate;
    rateGroup.style.display = 'block';
    document.getElementById('tradeinPaymentExchangeRate').value = `1 THB = ${formatNumber(rate)} LAK`;
    if (receivedLabel) receivedLabel.textContent = 'Received Amount (THB)';
  } else if (currency === 'USD') {
    rate = parseFloat(document.getElementById('tradeinPaymentRateUSD').value) || 21500;
    amountToPay = totalLAK / rate;
    rateGroup.style.display = 'block';
    document.getElementById('tradeinPaymentExchangeRate').value = `1 USD = ${formatNumber(rate)} LAK`;
    if (receivedLabel) receivedLabel.textContent = 'Received Amount (USD)';
  } else {
    rateGroup.style.display = 'none';
    if (receivedLabel) receivedLabel.textContent = 'Received Amount (LAK)';
  }
  
  document.getElementById('tradeinPaymentAmount').value = `${formatNumber(amountToPay.toFixed(2))} ${currency}`;
  document.getElementById('tradeinPaymentAmountLAK').value = formatNumber(totalLAK) + ' LAK';
}

function calculateTradeinChange() {
  if (!currentTradeinPayment) return;
  
  const currency = document.getElementById('tradeinPaymentCurrency').value;
  const received = parseFloat(document.getElementById('tradeinPaymentReceived').value) || 0;
  const totalLAK = currentTradeinPayment.total;
  
  let rate = 1;
  if (currency === 'THB') {
    rate = parseFloat(document.getElementById('tradeinPaymentRateTHB').value) || 270;
  } else if (currency === 'USD') {
    rate = parseFloat(document.getElementById('tradeinPaymentRateUSD').value) || 21500;
  }
  
  let receivedLAK = received;
  if (currency === 'THB') {
    receivedLAK = received * rate;
  } else if (currency === 'USD') {
    receivedLAK = received * rate;
  }
  
  const change = Math.max(0, receivedLAK - totalLAK);
  
  document.getElementById('tradeinPaymentChange').value = formatNumber(change) + ' LAK';
}

function toggleTradeinPaymentBank() {
  const method = document.getElementById('tradeinPaymentMethod').value;
  const bankGroup = document.getElementById('tradeinPaymentBankGroup');
  bankGroup.style.display = method === 'Bank' ? 'block' : 'none';
}

async function confirmTradeinPayment() {
  if (!currentTradeinPayment) return;
  
  const method = document.getElementById('tradeinPaymentMethod').value;
  const bank = method === 'Bank' ? document.getElementById('tradeinPaymentBank').value : '';
  const currency = document.getElementById('tradeinPaymentCurrency').value;
  const received = parseFloat(document.getElementById('tradeinPaymentReceived').value) || 0;
  
  if (received <= 0) {
    alert('กรุณากรอกจำนวนเงินที่รับ');
    return;
  }
  
  let rate = 1;
  if (currency === 'THB') {
    rate = parseFloat(document.getElementById('tradeinPaymentRateTHB').value) || 270;
  } else if (currency === 'USD') {
    rate = parseFloat(document.getElementById('tradeinPaymentRateUSD').value) || 21500;
  }
  
  let receivedLAK = received;
  if (currency === 'THB') {
    receivedLAK = received * rate;
  } else if (currency === 'USD') {
    receivedLAK = received * rate;
  }
  
  if (receivedLAK < currentTradeinPayment.total) {
    alert('❌ จำนวนเงินที่รับไม่เพียงพอ!');
    return;
  }
  
  const changeLAK = receivedLAK - currentTradeinPayment.total;
  
  try {
    showLoading();
    const result = await callAppsScript('CONFIRM_TRADEIN_PAYMENT', {
      tradeinId: currentTradeinPayment.tradeinId,
      method,
      bank,
      customerPaid: received,
      customerCurrency: currency,
      exchangeRate: rate,
      changeLAK,
      user: currentUser.nickname
    });
    
    if (result.success) {
      alert('✅ ยืนยันชำระเงินสำเร็จ!');
      closeModal('tradeinPaymentModal');
      currentTradeinPayment = null;
      loadTradeins();
      loadDashboard();
      loadCashBank();
    } else {
      alert('❌ เกิดข้อผิดพลาด: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ เกิดข้อผิดพลาด: ' + error.message);
    hideLoading();
  }
}

let currentTradeinPayment = null;
