async function loadTradeins() {
  try {
    showLoading();
    const data = await fetchSheetData('Tradeins!A:K');
    
    let filteredData = data.slice(1);
    
    if (currentUser.role === 'User' || currentUser.role === 'Manager') {
      filteredData = filterTodayData(filteredData, 7, 10);
    }
    
    const tbody = document.getElementById('tradeinTable');
    if (filteredData.length === 0) {
      tbody.innerHTML = '<tr><td colspan="9" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = filteredData.reverse().map(row => {
        const oldGold = formatItemsForTable(row[2]);
        const newGold = formatItemsForTable(row[3]);
        
        let actions = '';
        if (row[8] === 'PENDING') {
          if (currentUser.role === 'Manager' || currentUser.role === 'Teller') {
            actions = `<button class="btn-action" onclick="confirmTradein('${row[0]}')">Confirm</button>`;
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
            <td>${formatNumber(row[5])}</td>
            <td>${formatNumber(row[6])}</td>
            <td><span class="status-badge status-${row[8].toLowerCase()}">${row[8]}</span></td>
            <td>${actions}</td>
          </tr>
        `;
      }).join('');
    }
    
    hideLoading();
  } catch (error) {
    console.error('Error loading tradeins:', error);
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
    const weight = GOLD_WEIGHTS[item.productId];
    oldValue += weight * currentPricing.buyback1Baht * item.qty;
  });

  let newValue = 0;
  let exchangeFee = 0;
  let premium = 0;

  newGold.forEach(item => {
    const weight = GOLD_WEIGHTS[item.productId];
    newValue += weight * currentPricing.sell1Baht * item.qty;
    exchangeFee += EXCHANGE_FEES[item.productId] * item.qty;
    
    if (PREMIUM_PRODUCTS.includes(item.productId)) {
      premium += PREMIUM_PER_PIECE * item.qty;
    }
  });

  const difference = newValue - oldValue;

  currentTradeinData = {
    customer,
    oldGold: JSON.stringify(oldGold),
    newGold: JSON.stringify(newGold),
    difference,
    exchangeFee,
    premium
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
      <div class="summary-row">
        <span class="summary-label">รวมที่ต้องชำระ:</span>
        <span class="summary-value">${formatNumber(difference + exchangeFee + premium)} LAK</span>
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

async function confirmTradein(tradeinId) {
  if (!confirm('ยืนยันรายการแลกเปลี่ยน?')) return;
  
  try {
    showLoading();
    const result = await callAppsScript('CONFIRM_TRADEIN', { tradeinId });
    
    if (result.success) {
      alert('✅ ยืนยันสำเร็จ!');
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
