async function loadTradeins() {
  try {
    showLoading();
    const data = await fetchSheetData('Tradeins!A:N');
    
    let filteredData = data.slice(1);
    
    if (currentUser.role === 'User' || currentUser.role === 'Manager') {
      if (tradeinDateFrom || tradeinDateTo) {
        filteredData = filterByDateRange(filteredData, 11, 13, tradeinDateFrom, tradeinDateTo);
      } else {
        filteredData = filterTodayData(filteredData, 11, 13);
      }
    }
    
    if (tradeinSortOrder === 'asc') {
      filteredData.sort((a, b) => new Date(a[11]) - new Date(b[11]));
    } else {
      filteredData.sort((a, b) => new Date(b[11]) - new Date(a[11]));
    }
    
    const tbody = document.getElementById('tradeinTable');
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
  const phone = document.getElementById('tradeinPhone').value;
  if (!phone) {
    alert('กรุณากรอกเบอร์โทร');
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
      phone,
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
      document.getElementById('tradeinPhone').value = '';
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


function resetTradeinDateFilter() {
  const today = getTodayDateString();
  document.getElementById('tradeinDateFrom').value = today;
  document.getElementById('tradeinDateTo').value = today;
  tradeinDateFrom = today;
  tradeinDateTo = today;
  loadTradeins();
}

document.addEventListener('DOMContentLoaded', function() {
  const fromInput = document.getElementById('tradeinDateFrom');
  const toInput = document.getElementById('tradeinDateTo');
  
  if (fromInput && toInput) {
    fromInput.addEventListener('change', function() {
      tradeinDateFrom = this.value;
      if (tradeinDateFrom && tradeinDateTo) {
        loadTradeins();
      }
    });
    
    toInput.addEventListener('change', function() {
      tradeinDateTo = this.value;
      if (tradeinDateFrom && tradeinDateTo) {
        loadTradeins();
      }
    });
  }
});

let currentTradeinPayment = null;