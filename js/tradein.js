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

async function calculateTradein() {
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

  let oldWeight = 0;
  oldGold.forEach(item => {
    const product = FIXED_PRODUCTS.find(p => p.id === item.productId);
    console.log('Old Gold:', product.name, 'weight:', product.weight, 'qty:', item.qty);
    oldWeight += product.weight * item.qty;
  });

  let newWeight = 0;
  let premium = 0;

  newGold.forEach(item => {
    const product = FIXED_PRODUCTS.find(p => p.id === item.productId);
    console.log('New Gold:', product.name, 'weight:', product.weight, 'qty:', item.qty);
    newWeight += product.weight * item.qty;
    
    if (PREMIUM_PRODUCTS.includes(item.productId)) {
      premium += PREMIUM_PER_PIECE * item.qty;
    }
  });

  console.log('=== TRADE-IN CALCULATION ===');
  console.log('Old Weight:', oldWeight, 'บาท');
  console.log('New Weight:', newWeight, 'บาท');
  console.log('Weight Difference:', newWeight - oldWeight, 'บาท');
  console.log('Sell 1 Baht:', currentPricing.sell1Baht, 'LAK');
  console.log('Difference Value:', (newWeight - oldWeight) * currentPricing.sell1Baht, 'LAK');
  console.log('Premium:', premium, 'LAK');

  if (newWeight <= oldWeight) {
    alert('❌ น้ำหนักทองใหม่ต้องมากกว่าทองเก่า!\nทองเก่า: ' + oldWeight.toFixed(3) + ' บาท\nทองใหม่: ' + newWeight.toFixed(3) + ' บาท');
    return;
  }
  
  const weightDifference = newWeight - oldWeight;
  const difference = weightDifference * currentPricing.sell1Baht;
  const total = roundTo1000(difference + premium);

  console.log('FINAL - difference:', difference, 'LAK');
  console.log('FINAL - premium:', premium, 'LAK');
  console.log('FINAL - total:', total, 'LAK');
  console.log('===========================');

  try {
    showLoading();
    const result = await callAppsScript('ADD_TRADEIN', {
      customer,
      oldGold: JSON.stringify(oldGold),
      newGold: JSON.stringify(newGold),
      difference,
      premium,
      total,
      user: currentUser.nickname
    });
    
    if (result.success) {
      alert('✅ สร้างรายการแลกเปลี่ยนสำเร็จ! รอ Manager Review');
      closeModal('tradeinModal');
      document.getElementById('tradeinCustomer').value = '';
      document.getElementById('tradeinOldGold').innerHTML = '';
      document.getElementById('tradeinNewGold').innerHTML = '';
      tradeinOldCounter = 0;
      tradeinNewCounter = 0;
      addTradeinOldGold();
      addTradeinNewGold();
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
      oldGold: currentTradeinPayment.oldGold,
      newGold: currentTradeinPayment.newGold,
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

async function loadCurrentPricing() {
  try {
    const pricingData = await fetchSheetData('Pricing!A:B');
    
    if (pricingData.length > 1) {
      const latestPricing = pricingData[pricingData.length - 1];
      currentPricing = {
        sell1Baht: parseFloat(latestPricing[1]) || 0,
        buyback1Baht: 0
      };
      
      console.log('Loaded currentPricing:', currentPricing);
      return true;
    }
    return false;
  } catch (error) {
    console.error('Error loading pricing:', error);
    return false;
  }
}

async function openTradeinModal() {
  const hasPrice = await loadCurrentPricing();
  
  if (!hasPrice || !currentPricing.sell1Baht || currentPricing.sell1Baht === 0) {
    alert('❌ ยังไม่มีราคาทองในระบบ! กรุณาไปที่หน้า Products → Set New Price ก่อน');
    return;
  }
  
  openModal('tradeinModal');
}

let currentTradeinPayment = null;
