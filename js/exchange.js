async function loadExchanges() {
  try {
    showLoading();
    const data = await fetchSheetData('Exchanges!A:N');
    
    let filteredData = data.slice(1);
    
    if (currentUser.role === 'User' || currentUser.role === 'Manager') {
      if (exchangeDateFrom || exchangeDateTo) {
        filteredData = filterByDateRange(filteredData, 11, 13, exchangeDateFrom, exchangeDateTo);
      } else {
        filteredData = filterTodayData(filteredData, 11, 13);
      }
    }
    
    if (exchangeSortOrder === 'asc') {
      filteredData.sort((a, b) => new Date(a[11]) - new Date(b[11]));
    } else {
      filteredData.sort((a, b) => new Date(b[11]) - new Date(a[11]));
    }
    
    const tbody = document.getElementById('exchangeTable');
    if (filteredData.length === 0) {
      tbody.innerHTML = '<tr><td colspan="10" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = filteredData.map(row => {
        const oldGold = formatItemsForTable(row[2]);
        const newGold = formatItemsForTable(row[3]);
        const premium = calculatePremiumFromItems(row[3]);
        const saleName = row[13];
        const status = row[12];
        
        let actions = '';
        
        if (status === 'PENDING') {
          if (currentUser.role === 'Manager') {
            actions = `<button class="btn-action" onclick="reviewExchange('${row[0]}')">Review</button>`;
          } else {
            actions = '<span style="color: var(--text-secondary);">Waiting for review</span>';
          }
        } else if (status === 'READY') {
          if (currentUser.role === 'User') {
            actions = `<button class="btn-action" onclick="openExchangePaymentModal('${row[0]}')">Confirm</button>`;
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

let exchangeOldCounter = 0;
let exchangeNewCounter = 0;

function addExchangeOldGold() {
  exchangeOldCounter++;
  const container = document.getElementById('exchangeOldGold');
  const productOptions = FIXED_PRODUCTS.filter(p => p.id !== 'G07').map(p => 
    `<option value="${p.id}">${p.name}</option>`
  ).join('');
  
  container.insertAdjacentHTML('beforeend', `
    <div class="product-row" id="oldGold${exchangeOldCounter}">
      <select class="form-select" style="flex: 2;">
        <option value="">Select Product</option>
        ${productOptions}
      </select>
      <input type="number" class="form-input" placeholder="Qty" min="1" style="flex: 1;">
      <button type="button" class="btn-remove" onclick="document.getElementById('oldGold${exchangeOldCounter}').remove()">×</button>
    </div>
  `);
}

function addExchangeNewGold() {
  exchangeNewCounter++;
  const container = document.getElementById('exchangeNewGold');
  const productOptions = FIXED_PRODUCTS.filter(p => p.id !== 'G07').map(p => 
    `<option value="${p.id}">${p.name}</option>`
  ).join('');
  
  container.insertAdjacentHTML('beforeend', `
    <div class="product-row" id="newGold${exchangeNewCounter}">
      <select class="form-select" style="flex: 2;">
        <option value="">Select Product</option>
        ${productOptions}
      </select>
      <input type="number" class="form-input" placeholder="Qty" min="1" style="flex: 1;">
      <button type="button" class="btn-remove" onclick="document.getElementById('newGold${exchangeNewCounter}').remove()">×</button>
    </div>
  `);
}

async function calculateExchange() {
  const phone = document.getElementById('exchangePhone').value;
  if (!phone) {
    alert('กรุณากรอกเบอร์โทร');
    return;
  }

  const oldGold = [];
  document.querySelectorAll('#exchangeOldGold .product-row').forEach(row => {
    const productId = row.querySelector('select').value;
    const qty = parseInt(row.querySelector('input').value) || 0;
    if (productId && qty > 0) {
      oldGold.push({ productId, qty });
    }
  });

  const newGold = [];
  document.querySelectorAll('#exchangeNewGold .product-row').forEach(row => {
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
    oldWeight += product.weight * item.qty;
  });

  let newWeight = 0;
  newGold.forEach(item => {
    const product = FIXED_PRODUCTS.find(p => p.id === item.productId);
    newWeight += product.weight * item.qty;
  });

  if (Math.abs(oldWeight - newWeight) > 0.001) {
    alert('❌ น้ำหนักทองเก่าและทองใหม่ต้องเท่ากัน!\nทองเก่า: ' + oldWeight.toFixed(3) + ' บาท\nทองใหม่: ' + newWeight.toFixed(3) + ' บาท');
    return;
  }

  let exchangeFee = 0;
  let premium = 0;

  newGold.forEach(item => {
    exchangeFee += EXCHANGE_FEES[item.productId] * item.qty;
    
    if (PREMIUM_PRODUCTS.includes(item.productId)) {
      premium += PREMIUM_PER_PIECE * item.qty;
    }
  });

  const total = roundTo1000(exchangeFee + premium);

  try {
    showLoading();
    const result = await callAppsScript('ADD_EXCHANGE', {
      phone,
      oldGold: JSON.stringify(oldGold),
      newGold: JSON.stringify(newGold),
      exchangeFee,
      premium,
      total,
      user: currentUser.nickname
    });
    
    if (result.success) {
      alert('✅ สร้างรายการแลกเปลี่ยนสำเร็จ! รอ Manager Review');
      closeModal('exchangeModal');
      document.getElementById('exchangePhone').value = '';
      document.getElementById('exchangeOldGold').innerHTML = '';
      document.getElementById('exchangeNewGold').innerHTML = '';
      exchangeOldCounter = 0;
      exchangeNewCounter = 0;
      addExchangeOldGold();
      addExchangeNewGold();
      loadExchanges();
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

async function openExchangePaymentModal(exchangeId) {
  try {
    showLoading();
    const data = await fetchSheetData('Exchanges!A:N');
    const exchange = data.slice(1).find(row => row[0] === exchangeId);
    
    if (!exchange) {
      alert('❌ ไม่พบรายการแลกเปลี่ยน');
      hideLoading();
      return;
    }
    
    currentExchangePayment = {
      exchangeId: exchange[0],
      phone: exchange[1],
      oldGold: exchange[2],
      newGold: exchange[3],
      total: parseFloat(exchange[6]) || 0
    };
    
    document.getElementById('exchangePaymentId').textContent = exchange[0];
    document.getElementById('exchangePaymentPhone').textContent = exchange[1];
    document.getElementById('exchangePaymentOldGold').textContent = formatItemsForDisplay(exchange[2]);
    document.getElementById('exchangePaymentNewGold').textContent = formatItemsForDisplay(exchange[3]);
    document.getElementById('exchangePaymentTotal').textContent = formatNumber(exchange[6]) + ' LAK';
    
    document.getElementById('exchangePaymentCurrency').value = 'LAK';
    document.getElementById('exchangePaymentMethod').value = 'Cash';
    document.getElementById('exchangePaymentBankGroup').style.display = 'none';
    document.getElementById('exchangePaymentReceived').value = '';
    document.getElementById('exchangePaymentChange').value = '';
    
    calculateExchangePayment();
    
    hideLoading();
    openModal('exchangePaymentModal');
  } catch (error) {
    alert('❌ เกิดข้อผิดพลาด: ' + error.message);
    hideLoading();
  }
}

function calculateExchangePayment() {
  if (!currentExchangePayment) return;
  
  const currency = document.getElementById('exchangePaymentCurrency').value;
  const totalLAK = currentExchangePayment.total;
  
  let rate = 1;
  let amountToPay = totalLAK;
  
  const rateGroup = document.getElementById('exchangePaymentRateGroup');
  const receivedLabel = document.querySelector('label[for="exchangePaymentReceived"]');
  
  if (currency === 'THB') {
    rate = parseFloat(document.getElementById('exchangePaymentRateTHB').value) || 270;
    amountToPay = totalLAK / rate;
    rateGroup.style.display = 'block';
    document.getElementById('exchangePaymentExchangeRate').value = `1 THB = ${formatNumber(rate)} LAK`;
    if (receivedLabel) receivedLabel.textContent = 'Received Amount (THB)';
  } else if (currency === 'USD') {
    rate = parseFloat(document.getElementById('exchangePaymentRateUSD').value) || 21500;
    amountToPay = totalLAK / rate;
    rateGroup.style.display = 'block';
    document.getElementById('exchangePaymentExchangeRate').value = `1 USD = ${formatNumber(rate)} LAK`;
    if (receivedLabel) receivedLabel.textContent = 'Received Amount (USD)';
  } else {
    rateGroup.style.display = 'none';
    if (receivedLabel) receivedLabel.textContent = 'Received Amount (LAK)';
  }
  
  document.getElementById('exchangePaymentAmount').value = `${formatNumber(amountToPay.toFixed(2))} ${currency}`;
  document.getElementById('exchangePaymentAmountLAK').value = formatNumber(totalLAK) + ' LAK';
}

function calculateExchangeChange() {
  if (!currentExchangePayment) return;
  
  const currency = document.getElementById('exchangePaymentCurrency').value;
  const received = parseFloat(document.getElementById('exchangePaymentReceived').value) || 0;
  const totalLAK = currentExchangePayment.total;
  
  let rate = 1;
  if (currency === 'THB') {
    rate = parseFloat(document.getElementById('exchangePaymentRateTHB').value) || 270;
  } else if (currency === 'USD') {
    rate = parseFloat(document.getElementById('exchangePaymentRateUSD').value) || 21500;
  }
  
  let receivedLAK = received;
  if (currency === 'THB') {
    receivedLAK = received * rate;
  } else if (currency === 'USD') {
    receivedLAK = received * rate;
  }
  
  const change = Math.max(0, receivedLAK - totalLAK);
  
  document.getElementById('exchangePaymentChange').value = formatNumber(change) + ' LAK';
}

function toggleExchangePaymentBank() {
  const method = document.getElementById('exchangePaymentMethod').value;
  const bankGroup = document.getElementById('exchangePaymentBankGroup');
  bankGroup.style.display = method === 'Bank' ? 'block' : 'none';
}

async function confirmExchangePayment() {
  if (!currentExchangePayment) return;
  
  const method = document.getElementById('exchangePaymentMethod').value;
  const bank = method === 'Bank' ? document.getElementById('exchangePaymentBank').value : '';
  const currency = document.getElementById('exchangePaymentCurrency').value;
  const received = parseFloat(document.getElementById('exchangePaymentReceived').value) || 0;
  
  if (received <= 0) {
    alert('กรุณากรอกจำนวนเงินที่รับ');
    return;
  }
  
  let rate = 1;
  if (currency === 'THB') {
    rate = parseFloat(document.getElementById('exchangePaymentRateTHB').value) || 270;
  } else if (currency === 'USD') {
    rate = parseFloat(document.getElementById('exchangePaymentRateUSD').value) || 21500;
  }
  
  let receivedLAK = received;
  if (currency === 'THB') {
    receivedLAK = received * rate;
  } else if (currency === 'USD') {
    receivedLAK = received * rate;
  }
  
  if (receivedLAK < currentExchangePayment.total) {
    alert('❌ จำนวนเงินที่รับไม่เพียงพอ!');
    return;
  }
  
  const changeLAK = receivedLAK - currentExchangePayment.total;
  
  try {
    showLoading();
    const result = await callAppsScript('CONFIRM_EXCHANGE_PAYMENT', {
      exchangeId: currentExchangePayment.exchangeId,
      oldGold: currentExchangePayment.oldGold,
      newGold: currentExchangePayment.newGold,
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
      closeModal('exchangePaymentModal');
      currentExchangePayment = null;
      loadExchanges();
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

async function loadCurrentPricingForExchange() {
  try {
    const pricingData = await fetchSheetData('Pricing!A:B');
    
    if (pricingData.length > 1) {
      const latestPricing = pricingData[pricingData.length - 1];
      currentPricing = {
        sell1Baht: parseFloat(latestPricing[1]) || 0,
        buyback1Baht: 0
      };
      
      console.log('Loaded currentPricing for Exchange:', currentPricing);
      return true;
    }
    return false;
  } catch (error) {
    console.error('Error loading pricing:', error);
    return false;
  }
}

async function openExchangeModal() {
  const hasPrice = await loadCurrentPricingForExchange();
  
  if (!hasPrice || !currentPricing.sell1Baht || currentPricing.sell1Baht === 0) {
    alert('❌ ยังไม่มีราคาทองในระบบ! กรุณาไปที่หน้า Products → Set New Price ก่อน');
    return;
  }
  
  openModal('exchangeModal');
}


function resetExchangeDateFilter() {
  const today = getTodayDateString();
  document.getElementById('exchangeDateFrom').value = today;
  document.getElementById('exchangeDateTo').value = today;
  exchangeDateFrom = today;
  exchangeDateTo = today;
  loadExchanges();
}

document.addEventListener('DOMContentLoaded', function() {
  const fromInput = document.getElementById('exchangeDateFrom');
  const toInput = document.getElementById('exchangeDateTo');
  
  if (fromInput && toInput) {
    fromInput.addEventListener('change', function() {
      exchangeDateFrom = this.value;
      if (exchangeDateFrom && exchangeDateTo) {
        loadExchanges();
      }
    });
    
    toInput.addEventListener('change', function() {
      exchangeDateTo = this.value;
      if (exchangeDateFrom && exchangeDateTo) {
        loadExchanges();
      }
    });
  }
});

let currentExchangePayment = null;