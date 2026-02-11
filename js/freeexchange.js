let freeExOldCounter = 0;
let freeExNewCounter = 0;

function openFreeExchangeModal() {
  document.getElementById('freeExPhone').value = '';
  document.getElementById('freeExOldGold').innerHTML = '';
  document.getElementById('freeExNewGold').innerHTML = '';
  freeExOldCounter = 0;
  freeExNewCounter = 0;
  addFreeExOldGold();
  addFreeExNewGold();
  openModal('freeExchangeModal');
}

async function loadFreeExchanges() {
  try {
    showLoading();
    const data = await fetchSheetData('FreeExchanges!A:L');
    
    let filteredData = data.slice(1);
    
    if (currentUser.role === 'User' || currentUser.role === 'Manager') {
      if (freeExchangeDateFrom || freeExchangeDateTo) {
        filteredData = filterByDateRange(filteredData, 7, 9, freeExchangeDateFrom, freeExchangeDateTo);
      } else {
        filteredData = filterTodayData(filteredData, 7, 9);
      }
    }
    
    if (freeExchangeSortOrder === 'asc') {
      filteredData.sort((a, b) => new Date(a[7]) - new Date(b[7]));
    } else {
      filteredData.sort((a, b) => new Date(b[7]) - new Date(a[7]));
    }
    
    const tbody = document.getElementById('freeExchangeTable');
    if (filteredData.length === 0) {
      tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = filteredData.map(row => {
        const oldGold = formatItemsForTable(row[2]);
        const newGold = formatItemsForTable(row[3]);
        const premium = parseFloat(row[4]) || 0;
        const saleName = row[9];
        const status = row[8];
        
        let actions = '';
        
        if (status === 'PENDING') {
          if (currentUser.role === 'Manager') {
            actions = `<button class="btn-action" onclick="reviewFreeExchange('${row[0]}')">Review</button>`;
          } else {
            actions = '<span style="color: var(--text-secondary);">Waiting for review</span>';
          }
        } else if (status === 'READY') {
          if (currentUser.role === 'User') {
            if (premium > 0) {
              actions = `<button class="btn-action" onclick="openFreeExchangePaymentModal('${row[0]}')">Confirm</button>`;
            } else {
              actions = `<button class="btn-action" onclick="confirmFreeExchangeNoPay('${row[0]}')">Confirm</button>`;
            }
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
            <td>${formatNumber(premium)}</td>
            <td>${formatNumber(row[5])}</td>
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

function addFreeExOldGold() {
  freeExOldCounter++;
  const container = document.getElementById('freeExOldGold');
  const productOptions = FIXED_PRODUCTS.map(p => 
    `<option value="${p.id}">${p.name}</option>`
  ).join('');
  
  container.insertAdjacentHTML('beforeend', `
    <div class="product-row" id="freeExOld${freeExOldCounter}">
      <select class="form-select" style="flex: 2;">
        <option value="">Select Product</option>
        ${productOptions}
      </select>
      <input type="number" class="form-input" placeholder="Qty" min="1" style="flex: 1;">
      <button type="button" class="btn-remove" onclick="document.getElementById('freeExOld${freeExOldCounter}').remove()">×</button>
    </div>
  `);
}

function addFreeExNewGold() {
  freeExNewCounter++;
  const container = document.getElementById('freeExNewGold');
  const productOptions = FIXED_PRODUCTS.map(p => 
    `<option value="${p.id}">${p.name}</option>`
  ).join('');
  
  container.insertAdjacentHTML('beforeend', `
    <div class="product-row" id="freeExNew${freeExNewCounter}">
      <select class="form-select" style="flex: 2;">
        <option value="">Select Product</option>
        ${productOptions}
      </select>
      <input type="number" class="form-input" placeholder="Qty" min="1" style="flex: 1;">
      <button type="button" class="btn-remove" onclick="document.getElementById('freeExNew${freeExNewCounter}').remove()">×</button>
    </div>
  `);
}

async function calculateFreeExchange() {
  const phone = document.getElementById('freeExPhone').value;
  if (!phone) {
    alert('กรุณากรอกเบอร์โทร');
    return;
  }

  const oldGold = [];
  document.querySelectorAll('#freeExOldGold .product-row').forEach(row => {
    const productId = row.querySelector('select').value;
    const qty = parseInt(row.querySelector('input').value) || 0;
    if (productId && qty > 0) {
      oldGold.push({ productId, qty });
    }
  });

  const newGold = [];
  document.querySelectorAll('#freeExNewGold .product-row').forEach(row => {
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

  let oldPremium = 0;
  oldGold.forEach(item => {
    if (PREMIUM_PRODUCTS.includes(item.productId)) {
      oldPremium += PREMIUM_PER_PIECE * item.qty;
    }
  });

  let newPremium = 0;
  newGold.forEach(item => {
    if (PREMIUM_PRODUCTS.includes(item.productId)) {
      newPremium += PREMIUM_PER_PIECE * item.qty;
    }
  });

  let premium = newPremium > oldPremium ? newPremium - oldPremium : 0;

  const total = roundTo1000(premium);

  try {
    showLoading();
    const result = await callAppsScript('ADD_FREE_EXCHANGE', {
      phone,
      oldGold: JSON.stringify(oldGold),
      newGold: JSON.stringify(newGold),
      premium,
      total,
      user: currentUser.nickname
    });

    if (result.success) {
      alert('✅ สร้างรายการ Free Exchange สำเร็จ!');
      closeModal('freeExchangeModal');
      document.getElementById('freeExPhone').value = '';
      document.getElementById('freeExOldGold').innerHTML = '';
      document.getElementById('freeExNewGold').innerHTML = '';
      freeExOldCounter = 0;
      freeExNewCounter = 0;
      addFreeExOldGold();
      addFreeExNewGold();
      loadFreeExchanges();
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

function toggleFreeExchangeSortOrder() {
  freeExchangeSortOrder = freeExchangeSortOrder === 'desc' ? 'asc' : 'desc';
  document.getElementById('freeExchangeSortBtn').textContent = freeExchangeSortOrder === 'desc' ? '↓ Latest' : '↑ Oldest';
  loadFreeExchanges();
}

function filterFreeExchangeByDate() {
  freeExchangeDateFrom = document.getElementById('freeExchangeDateFrom').value;
  freeExchangeDateTo = document.getElementById('freeExchangeDateTo').value;
  loadFreeExchanges();
}

function clearFreeExchangeDateFilter() {
  freeExchangeDateFrom = null;
  freeExchangeDateTo = null;
  document.getElementById('freeExchangeDateFrom').value = '';
  document.getElementById('freeExchangeDateTo').value = '';
  loadFreeExchanges();
}

async function openFreeExchangePaymentModal(freeExId) {
  const data = await fetchSheetData('FreeExchanges!A:L');
  const freeEx = data.slice(1).find(row => row[0] === freeExId);
  if (!freeEx) return;
  
  const oldGold = formatItemsForPayment(freeEx[2]);
  const newGold = formatItemsForPayment(freeEx[3]);
  const total = parseFloat(freeEx[5]) || 0;
  
  openMultiPaymentModal('FREE_EXCHANGE', freeExId, total, freeEx[1], 
    `<strong>Old Gold:</strong> ${oldGold}<br><strong>New Gold:</strong> ${newGold}`);
}

async function confirmFreeExchangeNoPay(freeExId) {
  if (!confirm('ยืนยันการแลกเปลี่ยน (ไม่มีค่าใช้จ่าย)?')) return;
  
  try {
    showLoading();
    const result = await callAppsScript('CONFIRM_FREE_EXCHANGE_PAYMENT', {
      freeExId,
      payments: JSON.stringify({ cash: [], bank: [] }),
      totalPaid: 0,
      change: 0,
      user: currentUser.nickname
    });
    
    if (result.success) {
      alert('✅ ยืนยันการแลกเปลี่ยนสำเร็จ!');
      loadFreeExchanges();
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

async function reviewFreeExchange(freeExId) {
  try {
    var data = await fetchSheetData('FreeExchanges!A:L');
    var fe = data.slice(1).find(function(row) { return row[0] === freeExId; });
    if (fe) {
      openReviewDecisionModal('FREE_EXCHANGE', freeExId, fe[3]);
    }
  } catch(e) { alert('❌ Error loading data'); }
}