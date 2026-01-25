async function loadWithdraws() {
  try {
    showLoading();
    const data = await fetchSheetData('Withdraws!A:J');
    
    let filteredData = data.slice(1);
    
    if (withdrawSortOrder === 'asc') {
      filteredData.sort((a, b) => new Date(a[6]) - new Date(b[6]));
    } else {
      filteredData.sort((a, b) => new Date(b[6]) - new Date(a[6]));
    }
    
    const tbody = document.getElementById('withdrawTable');
    if (filteredData.length === 0) {
      tbody.innerHTML = '<tr><td colspan="7" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = filteredData.map(row => {
        const items = formatItemsForTable(row[2]);
        const premium = parseFloat(row[3]) || 0;
        const saleName = row[8];
        const status = row[7];
        
        let actions = '';
        
        if (status === 'PENDING') {
          if (currentUser.role === 'Manager') {
            actions = `<button class="btn-action" onclick="reviewWithdraw('${row[0]}')">Review</button>`;
          } else {
            actions = '<span style="color: var(--text-secondary);">Waiting for review</span>';
          }
        } else if (status === 'READY') {
          if (currentUser.role === 'User') {
            actions = `<button class="btn-action" onclick="openWithdrawPaymentModal('${row[0]}')">Confirm</button>`;
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
            <td>${items}</td>
            <td>${formatNumber(premium)}</td>
            <td><span class="status-badge status-${status.toLowerCase()}">${status}</span></td>
            <td>${saleName}</td>
            <td>${actions}</td>
          </tr>
        `;
      }).join('');
    }
    
    hideLoading();
  } catch (error) {
    console.error('Error loading withdraws:', error);
    hideLoading();
  }
}

let withdrawCounter = 0;

function addWithdrawProduct() {
  withdrawCounter++;
  const productOptions = FIXED_PRODUCTS.map(p => 
    `<option value="${p.id}">${p.name}</option>`
  ).join('');
  
  document.getElementById('withdrawProducts').insertAdjacentHTML('beforeend', `
    <div class="product-row" id="withdraw${withdrawCounter}">
      <select class="form-select" style="flex: 2;" onchange="calculateWithdrawPremium()">
        <option value="">Select Product</option>
        ${productOptions}
      </select>
      <input type="number" class="form-input" placeholder="Qty" min="1" style="flex: 1;" oninput="calculateWithdrawPremium()">
      <button type="button" class="btn-remove" onclick="document.getElementById('withdraw${withdrawCounter}').remove(); calculateWithdrawPremium();">×</button>
    </div>
  `);
}

function calculateWithdrawPremium() {
  const products = [];
  document.querySelectorAll('#withdrawProducts .product-row').forEach(row => {
    const productId = row.querySelector('select').value;
    const qty = parseInt(row.querySelector('input').value) || 0;
    if (productId && qty > 0) {
      products.push({ productId, qty });
    }
  });

  let premium = 0;
  products.forEach(item => {
    if (PREMIUM_PRODUCTS.includes(item.productId)) {
      premium += PREMIUM_PER_PIECE * item.qty;
    }
  });

  document.getElementById('withdrawPremium').value = formatNumber(premium);
}

async function calculateWithdraw() {
  const customer = document.getElementById('withdrawCustomer').value;
  if (!customer) {
    alert('กรุณากรอกชื่อลูกค้า');
    return;
  }

  const products = [];
  document.querySelectorAll('#withdrawProducts .product-row').forEach(row => {
    const productId = row.querySelector('select').value;
    const qty = parseInt(row.querySelector('input').value) || 0;
    if (productId && qty > 0) {
      products.push({ productId, qty });
    }
  });

  if (products.length === 0) {
    alert('กรุณาเลือกสินค้า');
    return;
  }

  let premium = 0;
  products.forEach(item => {
    if (PREMIUM_PRODUCTS.includes(item.productId)) {
      premium += PREMIUM_PER_PIECE * item.qty;
    }
  });

  const total = roundTo1000(premium);

  try {
    showLoading();
    const result = await callAppsScript('ADD_WITHDRAW', {
      customer,
      products: JSON.stringify(products),
      premium,
      total,
      user: currentUser.nickname
    });
    
    if (result.success) {
      alert('✅ สร้างรายการถอนทองสำเร็จ! รอ Manager Review');
      closeModal('withdrawModal');
      
      document.getElementById('withdrawCustomer').value = '';
      document.getElementById('withdrawProducts').innerHTML = '';
      withdrawCounter = 0;
      addWithdrawProduct();
      
      loadWithdraws();
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

async function reviewWithdraw(withdrawId) {
  if (!confirm('ยืนยันการ Review รายการถอนทองนี้?')) return;
  
  try {
    showLoading();
    const result = await callAppsScript('REVIEW_WITHDRAW', { withdrawId });
    
    if (result.success) {
      alert('✅ Review สำเร็จ! รอ User ยืนยัน');
      loadWithdraws();
    } else {
      alert('❌ เกิดข้อผิดพลาด: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ เกิดข้อผิดพลาด: ' + error.message);
    hideLoading();
  }
}

async function openWithdrawPaymentModal(withdrawId) {
  try {
    showLoading();
    const data = await fetchSheetData('Withdraws!A:J');
    const withdraw = data.slice(1).find(row => row[0] === withdrawId);
    
    if (!withdraw) {
      alert('❌ ไม่พบรายการถอนทอง');
      hideLoading();
      return;
    }
    
    currentWithdrawPayment = {
      withdrawId: withdraw[0],
      customer: withdraw[1],
      items: withdraw[2],
      premium: parseFloat(withdraw[3]) || 0,
      total: parseFloat(withdraw[4]) || 0
    };
    
    document.getElementById('withdrawPaymentId').textContent = withdraw[0];
    document.getElementById('withdrawPaymentCustomer').textContent = withdraw[1];
    document.getElementById('withdrawPaymentItems').textContent = formatItemsForDisplay(withdraw[2]);
    document.getElementById('withdrawPaymentPremium').textContent = formatNumber(withdraw[3]) + ' LAK';
    document.getElementById('withdrawPaymentTotal').textContent = formatNumber(withdraw[4]) + ' LAK';
    
    document.getElementById('withdrawPaymentCurrency').value = 'LAK';
    document.getElementById('withdrawPaymentMethod').value = 'Cash';
    document.getElementById('withdrawPaymentBankGroup').style.display = 'none';
    
    document.getElementById('withdrawPaymentRateTHB').value = currentExchangeRates.THB || 270;
    document.getElementById('withdrawPaymentRateUSD').value = currentExchangeRates.USD || 21500;
    
    calculateWithdrawPayment();
    
    hideLoading();
    openModal('withdrawPaymentModal');
  } catch (error) {
    alert('❌ เกิดข้อผิดพลาด: ' + error.message);
    hideLoading();
  }
}

function calculateWithdrawPayment() {
  if (!currentWithdrawPayment) return;
  
  const totalLAK = currentWithdrawPayment.total;
  const currency = document.getElementById('withdrawPaymentCurrency').value;
  
  let rate = 1;
  let amountToPay = totalLAK;
  
  const rateGroup = document.getElementById('withdrawPaymentRateGroup');
  
  if (currency === 'THB') {
    rate = parseFloat(document.getElementById('withdrawPaymentRateTHB').value) || 270;
    amountToPay = totalLAK / rate;
    rateGroup.style.display = 'block';
    document.getElementById('withdrawPaymentExchangeRate').value = `1 THB = ${formatNumber(rate)} LAK`;
  } else if (currency === 'USD') {
    rate = parseFloat(document.getElementById('withdrawPaymentRateUSD').value) || 21500;
    amountToPay = totalLAK / rate;
    rateGroup.style.display = 'block';
    document.getElementById('withdrawPaymentExchangeRate').value = `1 USD = ${formatNumber(rate)} LAK`;
  } else {
    rateGroup.style.display = 'none';
  }
  
  document.getElementById('withdrawPaymentAmount').value = `${formatNumber(amountToPay.toFixed(2))} ${currency}`;
  document.getElementById('withdrawPaymentAmountLAK').value = formatNumber(totalLAK) + ' LAK';
  
  const customerPaid = parseFloat(document.getElementById('withdrawPaymentCustomerPaid').value) || 0;
  const customerPaidLAK = customerPaid * rate;
  const changeLAK = customerPaidLAK - totalLAK;
  document.getElementById('withdrawPaymentChange').value = formatNumber(changeLAK) + ' LAK';
}

function toggleWithdrawPaymentBank() {
  const method = document.getElementById('withdrawPaymentMethod').value;
  const bankGroup = document.getElementById('withdrawPaymentBankGroup');
  bankGroup.style.display = method === 'Bank' ? 'block' : 'none';
}

async function confirmWithdrawPayment() {
  if (!currentWithdrawPayment) return;
  
  const method = document.getElementById('withdrawPaymentMethod').value;
  const bank = method === 'Bank' ? document.getElementById('withdrawPaymentBank').value : '';
  const currency = document.getElementById('withdrawPaymentCurrency').value;
  const customerPaid = parseFloat(document.getElementById('withdrawPaymentCustomerPaid').value) || 0;
  
  let rate = 1;
  if (currency === 'THB') {
    rate = parseFloat(document.getElementById('withdrawPaymentRateTHB').value) || 270;
  } else if (currency === 'USD') {
    rate = parseFloat(document.getElementById('withdrawPaymentRateUSD').value) || 21500;
  }
  
  try {
    showLoading();
    const result = await callAppsScript('CONFIRM_WITHDRAW_PAYMENT', {
      withdrawId: currentWithdrawPayment.withdrawId,
      items: currentWithdrawPayment.items,
      method,
      bank,
      currency,
      customerPaid,
      exchangeRate: rate,
      user: currentUser.nickname
    });
    
    if (result.success) {
      alert('✅ ยืนยันชำระเงินสำเร็จ!');
      closeModal('withdrawPaymentModal');
      currentWithdrawPayment = null;
      loadWithdraws();
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

async function loadCurrentPricingForWithdraw() {
  try {
    const pricingData = await fetchSheetData('Pricing!A:B');
    
    if (pricingData.length > 1) {
      const latestPricing = pricingData[pricingData.length - 1];
      currentPricing = {
        sell1Baht: parseFloat(latestPricing[1]) || 0,
        buyback1Baht: 0
      };
      
      console.log('Loaded currentPricing for Withdraw:', currentPricing);
      return true;
    }
    return false;
  } catch (error) {
    console.error('Error loading pricing:', error);
    return false;
  }
}

async function openWithdrawModal() {
  const hasPrice = await loadCurrentPricingForWithdraw();
  
  if (!hasPrice || !currentPricing.sell1Baht || currentPricing.sell1Baht === 0) {
    alert('❌ ยังไม่มีราคาทองในระบบ! กรุณาไปที่หน้า Products → Set New Price ก่อน');
    return;
  }
  
  openModal('withdrawModal');
}

function toggleWithdrawSort() {
  withdrawSortOrder = withdrawSortOrder === 'desc' ? 'asc' : 'desc';
  document.getElementById('withdrawSortBtn').textContent = 
    withdrawSortOrder === 'desc' ? 'Sort: Newest First' : 'Sort: Oldest First';
  loadWithdraws();
}

let currentWithdrawPayment = null;
