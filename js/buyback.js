async function loadBuybacks() {
  try {
    showLoading();
    const data = await fetchSheetData('Buybacks!A:J');
    
    let filteredData = data.slice(1);
    
    if (currentUser.role === 'User' || currentUser.role === 'Manager') {
      if (buybackDateFrom || buybackDateTo) {
        filteredData = filterByDateRange(filteredData, 7, 9, buybackDateFrom, buybackDateTo);
      } else {
        filteredData = filterTodayData(filteredData, 7, 9);
      }
    }
    
    if (buybackSortOrder === 'asc') {
      filteredData.sort((a, b) => new Date(a[7]) - new Date(b[7]));
    } else {
      filteredData.sort((a, b) => new Date(b[7]) - new Date(a[7]));
    }
    
    const tbody = document.getElementById('buybackTable');
    if (filteredData.length === 0) {
      tbody.innerHTML = '<tr><td colspan="7" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = filteredData.map(row => {
        const items = formatItemsForTable(row[2]);
        const total = row[6] || row[3];
        const saleName = row[9];
        const status = row[8];
        
        let actions = '';
        
        if (status === 'PENDING') {
          if (currentUser.role === 'Manager') {
            actions = `<button class="btn-action" onclick="openBuybackPaymentModalFromList('${row[0]}')">Payment</button>`;
          } else {
            actions = '<span style="color: var(--text-secondary);">Waiting for payment</span>';
          }
        } else {
          actions = '-';
        }
        
        return `
          <tr>
            <td>${row[0]}</td>
            <td>${row[1]}</td>
            <td>${items}</td>
            <td>${formatNumber(total)}</td>
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
  if (!currentPricing.sell1Baht || currentPricing.sell1Baht === 0) {
    console.log('currentPricing not loaded yet');
    return 0;
  }

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
    console.log('Buyback:', item.productId, 'price:', pricePerPiece, 'qty:', item.qty);
    totalPrice += pricePerPiece * item.qty;
  });

  const total = roundTo1000(totalPrice);
  console.log('Buyback Total:', total, 'LAK');
  document.getElementById('buybackPrice').value = formatNumber(total);
  
  return total;
}

async function calculateBuyback() {
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

  try {
    showLoading();
    const result = await callAppsScript('ADD_BUYBACK', {
      customer,
      products: JSON.stringify(products),
      total,
      user: currentUser.nickname
    });
    
    if (result.success) {
      alert('✅ สร้างรายการรับซื้อสำเร็จ! รอ Manager Review');
      closeModal('buybackModal');
      
      document.getElementById('buybackCustomer').value = '';
      document.getElementById('buybackProducts').innerHTML = '';
      document.getElementById('buybackPrice').value = '';
      buybackCounter = 0;
      addBuybackProduct();
      
      loadBuybacks();
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

async function openBuybackPaymentModalFromList(buybackId) {
  try {
    showLoading();
    const data = await fetchSheetData('Buybacks!A:J');
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
      baseTotal: parseFloat(buyback[3]) || 0
    };
    
    document.getElementById('buybackPaymentId').textContent = buyback[0];
    document.getElementById('buybackPaymentCustomer').textContent = buyback[1];
    document.getElementById('buybackPaymentItems').textContent = formatItemsForDisplay(buyback[2]);
    document.getElementById('buybackPaymentBaseTotal').textContent = formatNumber(buyback[3]) + ' LAK';
    
    document.getElementById('buybackPaymentMethod').value = 'Cash';
    document.getElementById('buybackPaymentBankGroup').style.display = 'none';
    document.getElementById('buybackPaymentFee').value = '0';
    
    calculateBuybackPaymentTotal();
    
    hideLoading();
    openModal('buybackPaymentModal');
  } catch (error) {
    alert('❌ เกิดข้อผิดพลาด: ' + error.message);
    hideLoading();
  }
}

function calculateBuybackPaymentTotal() {
  if (!currentBuybackPayment) return;
  
  const baseTotal = currentBuybackPayment.baseTotal || 0;
  const fee = parseFloat(document.getElementById('buybackPaymentFee').value) || 0;
  const total = baseTotal + fee;
  
  document.getElementById('buybackPaymentTotal').value = formatNumber(total) + ' LAK';
}

function toggleBuybackPaymentBank() {
  const method = document.getElementById('buybackPaymentMethod').value;
  const bankGroup = document.getElementById('buybackPaymentBankGroup');
  bankGroup.style.display = method === 'Bank' ? 'block' : 'none';
}

async function confirmBuybackPayment() {
  if (!currentBuybackPayment) return;
  
  const method = document.getElementById('buybackPaymentMethod').value;
  const bank = method === 'Bank' ? document.getElementById('buybackPaymentBank').value : '';
  const fee = parseFloat(document.getElementById('buybackPaymentFee').value) || 0;
  const total = currentBuybackPayment.baseTotal + fee;
  
  try {
    showLoading();
    
    const result = await callAppsScript('CONFIRM_BUYBACK_PAYMENT', {
      buybackId: currentBuybackPayment.buybackId,
      items: currentBuybackPayment.items,
      method,
      bank,
      fee,
      total,
      user: currentUser.nickname
    });
    
    if (result.success) {
      alert('✅ ยืนยันชำระเงินสำเร็จ!');
      closeModal('buybackPaymentModal');
      currentBuybackPayment = null;
      
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

async function loadCurrentPricingForBuyback() {
  try {
    const pricingData = await fetchSheetData('Pricing!A:B');
    
    if (pricingData.length > 1) {
      const latestPricing = pricingData[pricingData.length - 1];
      currentPricing = {
        sell1Baht: parseFloat(latestPricing[1]) || 0,
        buyback1Baht: 0
      };
      
      console.log('Loaded currentPricing for Buyback:', currentPricing);
      return true;
    }
    return false;
  } catch (error) {
    console.error('Error loading pricing:', error);
    return false;
  }
}

async function openBuybackModal() {
  const hasPrice = await loadCurrentPricingForBuyback();
  
  if (!hasPrice || !currentPricing.sell1Baht || currentPricing.sell1Baht === 0) {
    alert('❌ ยังไม่มีราคาทองในระบบ! กรุณาไปที่หน้า Products → Set New Price ก่อน');
    return;
  }
  
  openModal('buybackModal');
}

function toggleBuybackSort() {
  buybackSortOrder = buybackSortOrder === 'desc' ? 'asc' : 'desc';
  document.getElementById('buybackSortBtn').textContent = 
    buybackSortOrder === 'desc' ? 'Sort: Newest First' : 'Sort: Oldest First';
  loadBuybacks();
}

function filterBuybackByDate() {
  buybackDateFrom = document.getElementById('buybackDateFrom').value;
  buybackDateTo = document.getElementById('buybackDateTo').value;
  loadBuybacks();
}

function resetBuybackDateFilter() {
  const today = getTodayDateString();
  document.getElementById('buybackDateFrom').value = today;
  document.getElementById('buybackDateTo').value = today;
  buybackDateFrom = today;
  buybackDateTo = today;
  loadBuybacks();
}

let currentBuybackPayment = null;
