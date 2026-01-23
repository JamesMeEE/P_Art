async function loadExchanges() {
  try {
    showLoading();
    const data = await fetchSheetData('Exchanges!A:J');
    
    let filteredData = data.slice(1);
    
    if (currentUser.role === 'User' || currentUser.role === 'Manager') {
      filteredData = filterTodayData(filteredData, 6, 9);
    }
    
    const tbody = document.getElementById('exchangeTable');
    if (filteredData.length === 0) {
      tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = filteredData.reverse().map(row => {
        const oldGold = formatItemsForTable(row[2]);
        const newGold = formatItemsForTable(row[3]);
        
        let actions = '';
        if (row[7] === 'PENDING') {
          if (currentUser.role === 'Manager' || currentUser.role === 'Teller') {
            actions = `<button class="btn-action" onclick="confirmExchange('${row[0]}')">Confirm</button>`;
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
            <td><span class="status-badge status-${row[7].toLowerCase()}">${row[7]}</span></td>
            <td>${actions}</td>
          </tr>
        `;
      }).join('');
    }
    
    hideLoading();
  } catch (error) {
    console.error('Error loading exchanges:', error);
    hideLoading();
  }
}

function addExchangeOldGold() {
  exchangeOldCounter++;
  const container = document.getElementById('exchangeOldGold');
  const row = document.createElement('div');
  row.className = 'product-row';
  row.id = `exchangeOld${exchangeOldCounter}`;
  row.innerHTML = `
    <select class="form-select">
      <option value="">เลือกสินค้า...</option>
      ${FIXED_PRODUCTS.map(p => `<option value="${p.id}">${p.name}</option>`).join('')}
    </select>
    <input type="number" class="form-input" placeholder="จำนวน" min="1" step="1">
    <button type="button" class="btn-remove" onclick="removeExchangeOldGold(${exchangeOldCounter})">×</button>
  `;
  container.appendChild(row);
}

function removeExchangeOldGold(id) {
  const row = document.getElementById(`exchangeOld${id}`);
  if (row) row.remove();
}

function addExchangeNewGold() {
  exchangeNewCounter++;
  const container = document.getElementById('exchangeNewGold');
  const row = document.createElement('div');
  row.className = 'product-row';
  row.id = `exchangeNew${exchangeNewCounter}`;
  row.innerHTML = `
    <select class="form-select">
      <option value="">เลือกสินค้า...</option>
      ${FIXED_PRODUCTS.map(p => `<option value="${p.id}">${p.name}</option>`).join('')}
    </select>
    <input type="number" class="form-input" placeholder="จำนวน" min="1" step="1">
    <button type="button" class="btn-remove" onclick="removeExchangeNewGold(${exchangeNewCounter})">×</button>
  `;
  container.appendChild(row);
}

function removeExchangeNewGold(id) {
  const row = document.getElementById(`exchangeNew${id}`);
  if (row) row.remove();
}

function calculateExchange() {
  const customer = document.getElementById('exchangeCustomer').value;
  if (!customer) {
    alert('กรุณากรอกชื่อลูกค้า');
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
  let oldHas1g = false;
  oldGold.forEach(item => {
    oldWeight += GOLD_WEIGHTS[item.productId] * item.qty;
    if (item.productId === 'G07') oldHas1g = true;
  });

  let newWeight = 0;
  let newHas1g = false;
  newGold.forEach(item => {
    newWeight += GOLD_WEIGHTS[item.productId] * item.qty;
    if (item.productId === 'G07') newHas1g = true;
  });

  if (Math.abs(oldWeight - newWeight) > 0.01) {
    alert(`❌ Exchange ต้องมีน้ำหนักเท่ากัน!\nทองเก่า: ${oldWeight.toFixed(2)} บาท\nทองใหม่: ${newWeight.toFixed(2)} บาท`);
    return;
  }

  if (oldHas1g !== newHas1g) {
    alert('❌ 1 กรัม แลกได้แค่กับ 1 กรัม เท่านั้น!');
    return;
  }

  let exchangeFee = 0;
  newGold.forEach(item => {
    exchangeFee += EXCHANGE_FEES[item.productId] * item.qty;
  });

  let premium = 0;
  newGold.forEach(item => {
    if (PREMIUM_PRODUCTS.includes(item.productId)) {
      premium += PREMIUM_PER_PIECE * item.qty;
    }
  });

  currentExchangeData = {
    customer,
    oldGold: JSON.stringify(oldGold),
    newGold: JSON.stringify(newGold),
    exchangeFee,
    premium
  };

  const oldItems = oldGold.map(i => `${FIXED_PRODUCTS.find(p => p.id === i.productId).name} (${i.qty})`).join(', ');
  const newItems = newGold.map(i => `${FIXED_PRODUCTS.find(p => p.id === i.productId).name} (${i.qty})`).join(', ');

  document.getElementById('exchangeResult').innerHTML = `
    <div class="summary-box">
      <div class="summary-row">
        <span class="summary-label">ทองเก่า:</span>
        <span class="summary-value">${oldItems}</span>
      </div>
      <div class="summary-row">
        <span class="summary-label">ทองใหม่:</span>
        <span class="summary-value">${newItems}</span>
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
        <span class="summary-value">${formatNumber(exchangeFee + premium)} LAK</span>
      </div>
    </div>
  `;

  closeModal('exchangeModal');
  openModal('exchangeResultModal');
}

function backToExchange() {
  closeModal('exchangeResultModal');
  openModal('exchangeModal');
  
  if (currentExchangeData) {
    document.getElementById('exchangeCustomer').value = currentExchangeData.customer;
    
    const oldGold = JSON.parse(currentExchangeData.oldGold);
    const newGold = JSON.parse(currentExchangeData.newGold);
    
    document.getElementById('exchangeOldGold').innerHTML = '';
    exchangeOldCounter = 0;
    oldGold.forEach(item => {
      addExchangeOldGold();
      const rows = document.querySelectorAll('#exchangeOldGold .product-row');
      const lastRow = rows[rows.length - 1];
      lastRow.querySelector('select').value = item.productId;
      lastRow.querySelector('input').value = item.qty;
    });
    
    document.getElementById('exchangeNewGold').innerHTML = '';
    exchangeNewCounter = 0;
    newGold.forEach(item => {
      addExchangeNewGold();
      const rows = document.querySelectorAll('#exchangeNewGold .product-row');
      const lastRow = rows[rows.length - 1];
      lastRow.querySelector('select').value = item.productId;
      lastRow.querySelector('input').value = item.qty;
    });
  }
}

async function submitExchange() {
  if (!currentExchangeData) return;

  const newGold = JSON.parse(currentExchangeData.newGold);
  
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
    const result = await callAppsScript('ADD_EXCHANGE', currentExchangeData);

    if (result.success) {
      alert('✅ สร้างรายการแลกเปลี่ยนสำเร็จ!');
      closeModal('exchangeResultModal');
      currentExchangeData = null;
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

async function confirmExchange(exchangeId) {
  if (!confirm('ยืนยันรายการแลกเปลี่ยน?')) return;
  
  try {
    showLoading();
    const result = await callAppsScript('CONFIRM_EXCHANGE', { exchangeId });
    
    if (result.success) {
      alert('✅ ยืนยันสำเร็จ!');
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
