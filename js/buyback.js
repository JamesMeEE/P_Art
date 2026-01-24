async function loadBuybacks() {
  try {
    showLoading();
    const data = await fetchSheetData('Buybacks!A:H');
    
    let filteredData = data.slice(1);
    
    if (currentUser.role === 'User' || currentUser.role === 'Manager') {
      filteredData = filterTodayData(filteredData, 5, 7);
    }
    
    const tbody = document.getElementById('buybackTable');
    if (filteredData.length === 0) {
      tbody.innerHTML = '<tr><td colspan="7" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = filteredData.reverse().map(row => {
        const items = formatItemsForTable(row[2]);
        const saleName = row[7];
        const status = row[6];
        
        let actions = '';
        
        if (status === 'PENDING') {
          if (currentUser.role === 'Manager') {
            actions = `<button class="btn-action" onclick="reviewBuyback('${row[0]}')">Review</button>`;
          } else {
            actions = '<span style="color: var(--text-secondary);">Waiting for review</span>';
          }
        } else if (status === 'READY') {
          if (currentUser.role === 'User') {
            actions = `<button class="btn-action" onclick="openBuybackPaymentModal('${row[0]}')">Confirm</button>`;
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
            <td>${formatNumber(row[3])}</td>
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

let buybackCounter = 0;

function addBuybackProduct() {
  buybackCounter++;
  const container = document.getElementById('buybackProducts');
  const productOptions = FIXED_PRODUCTS.map(p => 
    `<option value="${p.id}">${p.name}</option>`
  ).join('');
  
  container.insertAdjacentHTML('beforeend', `
    <div class="product-row" id="buyback${buybackCounter}">
      <select class="form-select" style="flex: 2;" onchange="calculateBuybackTotal()">
        <option value="">Select Product</option>
        ${productOptions}
      </select>
      <input type="number" class="form-input" placeholder="Qty" min="1" style="flex: 1;" oninput="calculateBuybackTotal()">
      <button type="button" class="btn-remove" onclick="document.getElementById('buyback${buybackCounter}').remove(); calculateBuybackTotal();">×</button>
    </div>
  `);
}

function calculateBuybackTotal() {
  const products = [];
  document.querySelectorAll('#buybackProducts .product-row').forEach(row => {
    const productId = row.querySelector('select').value;
    const qty = parseInt(row.querySelector('input').value) || 0;
    if (productId && qty > 0) {
      products.push({ productId, qty });
    }
  });

  let totalPrice = 0;
  products.forEach(item => {
    const pricePerPiece = calculateBuybackPrice(item.productId, currentPricing.sell1Baht);
    totalPrice += pricePerPiece * item.qty;
  });

  const total = roundTo1000(totalPrice);
  document.getElementById('buybackPrice').value = formatNumber(total);
  
  return total;
}

function calculateBuyback() {
  const customer = document.getElementById('buybackCustomer').value;
  if (!customer) {
    alert('กรุณากรอกชื่อลูกค้า');
    return;
  }

  const products = [];
  document.querySelectorAll('#buybackProducts .product-row').forEach(row => {
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

  const total = calculateBuybackTotal();

  currentBuybackData = {
    customer,
    products: JSON.stringify(products),
    total
  };

  const itemsList = products.map(p => {
    const product = FIXED_PRODUCTS.find(pr => pr.id === p.productId);
    return `${product.name} (${p.qty})`;
  }).join(', ');

  document.getElementById('buybackPaymentId').textContent = 'NEW';
  document.getElementById('buybackPaymentCustomer').textContent = customer;
  document.getElementById('buybackPaymentItems').textContent = itemsList;
  document.getElementById('buybackPaymentTotal').textContent = formatNumber(total) + ' LAK';

  document.getElementById('buybackPaymentCurrency').value = 'LAK';
  document.getElementById('buybackPaymentMethod').value = 'Cash';
  document.getElementById('buybackPaymentBankGroup').style.display = 'none';

  calculateBuybackPayment();

  closeModal('buybackModal');
  openModal('buybackPaymentModal');
}

async function reviewBuyback(buybackId) {
  if (!confirm('ยืนยันการ Review รายการรับซื้อนี้?')) return;
  
  try {
    showLoading();
    const result = await callAppsScript('REVIEW_BUYBACK', { buybackId });
    
    if (result.success) {
      alert('✅ Review สำเร็จ! รอ User ยืนยันชำระเงิน');
      loadBuybacks();
    } else {
      alert('❌ เกิดข้อผิดพลาด: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ เกิดข้อผิดพลาด: ' + error.message);
    hideLoading();
  }
}

async function openBuybackPaymentModal(buybackId) {
  try {
    showLoading();
    const data = await fetchSheetData('Buybacks!A:H');
    const buyback = data.slice(1).find(row => row[0] === buybackId);
    
    if (!buyback) {
      alert('❌ ไม่พบรายการรับซื้อ');
      hideLoading();
      return;
    }
    
    currentBuybackPayment = {
      buybackId: buyback[0],
      customer: buyback[1],
      items: buyback[2],
      total: parseFloat(buyback[3]) || 0
    };
    
    document.getElementById('buybackPaymentId').textContent = buyback[0];
    document.getElementById('buybackPaymentCustomer').textContent = buyback[1];
    document.getElementById('buybackPaymentItems').textContent = formatItemsForDisplay(buyback[2]);
    document.getElementById('buybackPaymentTotal').textContent = formatNumber(buyback[3]) + ' LAK';
    
    document.getElementById('buybackPaymentCurrency').value = 'LAK';
    document.getElementById('buybackPaymentMethod').value = 'Cash';
    document.getElementById('buybackPaymentBankGroup').style.display = 'none';
    
    calculateBuybackPayment();
    
    hideLoading();
    openModal('buybackPaymentModal');
  } catch (error) {
    alert('❌ เกิดข้อผิดพลาด: ' + error.message);
    hideLoading();
  }
}

function calculateBuybackPayment() {
  const isNew = document.getElementById('buybackPaymentId').textContent === 'NEW';
  const totalLAK = isNew ? currentBuybackData.total : currentBuybackPayment.total;
  
  const currency = document.getElementById('buybackPaymentCurrency').value;
  
  let rate = 1;
  let amountToPay = totalLAK;
  
  const rateGroup = document.getElementById('buybackPaymentRateGroup');
  
  if (currency === 'THB') {
    rate = parseFloat(document.getElementById('buybackPaymentRateTHB').value) || 270;
    amountToPay = totalLAK / rate;
    rateGroup.style.display = 'block';
    document.getElementById('buybackPaymentExchangeRate').value = `1 THB = ${formatNumber(rate)} LAK`;
  } else if (currency === 'USD') {
    rate = parseFloat(document.getElementById('buybackPaymentRateUSD').value) || 21500;
    amountToPay = totalLAK / rate;
    rateGroup.style.display = 'block';
    document.getElementById('buybackPaymentExchangeRate').value = `1 USD = ${formatNumber(rate)} LAK`;
  } else {
    rateGroup.style.display = 'none';
  }
  
  document.getElementById('buybackPaymentAmount').value = `${formatNumber(amountToPay.toFixed(2))} ${currency}`;
  document.getElementById('buybackPaymentAmountLAK').value = formatNumber(totalLAK) + ' LAK';
}

function toggleBuybackPaymentBank() {
  const method = document.getElementById('buybackPaymentMethod').value;
  const bankGroup = document.getElementById('buybackPaymentBankGroup');
  bankGroup.style.display = method === 'Bank' ? 'block' : 'none';
}

async function confirmBuybackPayment() {
  const isNew = document.getElementById('buybackPaymentId').textContent === 'NEW';
  
  const method = document.getElementById('buybackPaymentMethod').value;
  const bank = method === 'Bank' ? document.getElementById('buybackPaymentBank').value : '';
  const currency = document.getElementById('buybackPaymentCurrency').value;
  
  let rate = 1;
  if (currency === 'THB') {
    rate = parseFloat(document.getElementById('buybackPaymentRateTHB').value) || 270;
  } else if (currency === 'USD') {
    rate = parseFloat(document.getElementById('buybackPaymentRateUSD').value) || 21500;
  }
  
  try {
    showLoading();
    
    let result;
    if (isNew) {
      result = await callAppsScript('CREATE_BUYBACK_WITH_PAYMENT', {
        ...currentBuybackData,
        method,
        bank,
        currency,
        exchangeRate: rate,
        user: currentUser.nickname
      });
    } else {
      result = await callAppsScript('CONFIRM_BUYBACK_PAYMENT', {
        buybackId: currentBuybackPayment.buybackId,
        items: currentBuybackPayment.items,
        method,
        bank,
        currency,
        exchangeRate: rate,
        user: currentUser.nickname
      });
    }
    
    if (result.success) {
      alert('✅ ยืนยันชำระเงินสำเร็จ!');
      closeModal('buybackPaymentModal');
      currentBuybackPayment = null;
      currentBuybackData = null;
      
      document.getElementById('buybackCustomer').value = '';
      document.getElementById('buybackProducts').innerHTML = '';
      document.getElementById('buybackPrice').value = '';
      buybackCounter = 0;
      addBuybackProduct();
      
      loadBuybacks();
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

let currentBuybackPayment = null;
let currentBuybackData = null;
