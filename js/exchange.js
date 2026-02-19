async function loadExchanges() {
  try {
    showLoading();
    const data = await fetchSheetData('Exchanges!A:N');
    
    let filteredData = data.slice(1);
    
    if (currentUser.role === 'User' || isManager()) {
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
        const exchangeFee = parseFloat(row[4]) || 0;
        const premium = parseFloat(row[5]) || 0;
        const saleName = row[13];
        const status = row[12];
        
        let actions = '';
        
        if (status === 'PENDING') {
          if (isManager()) {
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
          var detail = encodeURIComponent(JSON.stringify([['Transaction ID', row[0]], ['Phone', row[1]], ['Old Gold', oldGold], ['New Gold', newGold], ['Exchange Fee', formatNumber(exchangeFee) + ' LAK'], ['Premium', formatNumber(premium) + ' LAK'], ['Total', formatNumber(row[6]) + ' LAK'], ['Date', formatDateTime(row[11])], ['Status', status], ['Sale', saleName]]));
          actions = '<button class="btn-action" onclick="viewTransactionDetail(\'Exchange\',\'' + detail + '\')" style="background:#555;">👁 View</button>';
        }
        
        return `
          <tr>
            <td>${row[0]}</td>
            <td>${row[1]}</td>
            <td>${oldGold}</td>
            <td>${newGold}</td>
            <td>${formatNumber(exchangeFee)}</td>
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
  const productOptions = FIXED_PRODUCTS.map(p => 
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
  const productOptions = FIXED_PRODUCTS.map(p => 
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
  if (_isSubmitting) return;
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
    _isSubmitting = true;
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
    endSubmit();
  } catch (error) {
    alert('❌ เกิดข้อผิดพลาด: ' + error.message);
    endSubmit();
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