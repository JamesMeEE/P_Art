let switchOldCounter = 0;
let switchNewCounter = 0;

function openSwitchModal() {
  document.getElementById('switchPhone').value = '';
  document.getElementById('switchOldGold').innerHTML = '';
  document.getElementById('switchNewGold').innerHTML = '';
  switchOldCounter = 0;
  switchNewCounter = 0;
  addSwitchOldGold();
  addSwitchNewGold();
  openModal('switchModal');
}

async function loadSwitches() {
  try {
    showLoading();
    const data = await fetchSheetData('Switches!A:N');
    
    let filteredData = data.slice(1);
    
    if (currentUser.role === 'User' || currentUser.role === 'Manager') {
      if (switchDateFrom || switchDateTo) {
        filteredData = filterByDateRange(filteredData, 11, 13, switchDateFrom, switchDateTo);
      } else {
        filteredData = filterTodayData(filteredData, 11, 13);
      }
    }
    
    if (switchSortOrder === 'asc') {
      filteredData.sort((a, b) => new Date(a[11]) - new Date(b[11]));
    } else {
      filteredData.sort((a, b) => new Date(b[11]) - new Date(a[11]));
    }
    
    const tbody = document.getElementById('switchTable');
    if (filteredData.length === 0) {
      tbody.innerHTML = '<tr><td colspan="10" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = filteredData.map(row => {
        const oldGold = formatItemsForTable(row[2]);
        const newGold = formatItemsForTable(row[3]);
        const switchFee = parseFloat(row[4]) || 0;
        const premium = parseFloat(row[5]) || 0;
        const saleName = row[13];
        const status = row[12];
        
        let actions = '';
        
        if (status === 'PENDING') {
          if (currentUser.role === 'Manager') {
            actions = `<button class="btn-action" onclick="reviewSwitch('${row[0]}')">Review</button>`;
          } else {
            actions = '<span style="color: var(--text-secondary);">Waiting for review</span>';
          }
        } else if (status === 'READY') {
          if (currentUser.role === 'User') {
            actions = `<button class="btn-action" onclick="openSwitchPaymentModal('${row[0]}')">Confirm</button>`;
          } else {
            actions = '<span style="color: var(--text-secondary);">Waiting for confirmation</span>';
          }
        } else {
          var detail = encodeURIComponent(JSON.stringify([['Transaction ID', row[0]], ['Phone', row[1]], ['Old Gold', oldGold], ['New Gold', newGold], ['Switch Fee', formatNumber(switchFee) + ' LAK'], ['Premium', formatNumber(premium) + ' LAK'], ['Total', formatNumber(row[6]) + ' LAK'], ['Date', formatDateTime(row[11])], ['Status', status], ['Sale', saleName]]));
          actions = '<button class="btn-action" onclick="viewTransactionDetail(\'Switch\',\'' + detail + '\')" style="background:#555;">👁 View</button>';
        }
        
        return `
          <tr>
            <td>${row[0]}</td>
            <td>${row[1]}</td>
            <td>${oldGold}</td>
            <td>${newGold}</td>
            <td>${formatNumber(switchFee)}</td>
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

function addSwitchOldGold() {
  switchOldCounter++;
  const container = document.getElementById('switchOldGold');
  const productOptions = FIXED_PRODUCTS.map(p => 
    `<option value="${p.id}">${p.name}</option>`
  ).join('');
  
  container.insertAdjacentHTML('beforeend', `
    <div class="product-row" id="switchOld${switchOldCounter}">
      <select class="form-select" style="flex: 2;">
        <option value="">Select Product</option>
        ${productOptions}
      </select>
      <input type="number" class="form-input" placeholder="Qty" min="1" style="flex: 1;">
      <button type="button" class="btn-remove" onclick="document.getElementById('switchOld${switchOldCounter}').remove()">×</button>
    </div>
  `);
}

function addSwitchNewGold() {
  switchNewCounter++;
  const container = document.getElementById('switchNewGold');
  const productOptions = FIXED_PRODUCTS.map(p => 
    `<option value="${p.id}">${p.name}</option>`
  ).join('');
  
  container.insertAdjacentHTML('beforeend', `
    <div class="product-row" id="switchNew${switchNewCounter}">
      <select class="form-select" style="flex: 2;">
        <option value="">Select Product</option>
        ${productOptions}
      </select>
      <input type="number" class="form-input" placeholder="Qty" min="1" style="flex: 1;">
      <button type="button" class="btn-remove" onclick="document.getElementById('switchNew${switchNewCounter}').remove()">×</button>
    </div>
  `);
}

async function calculateSwitch() {
  const phone = document.getElementById('switchPhone').value;
  if (!phone) {
    alert('กรุณากรอกเบอร์โทร');
    return;
  }

  const oldGold = [];
  document.querySelectorAll('#switchOldGold .product-row').forEach(row => {
    const productId = row.querySelector('select').value;
    const qty = parseInt(row.querySelector('input').value) || 0;
    if (productId && qty > 0) {
      oldGold.push({ productId, qty });
    }
  });

  const newGold = [];
  document.querySelectorAll('#switchNewGold .product-row').forEach(row => {
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

  let switchFee = 0;
  let premium = 0;

  newGold.forEach(item => {
    switchFee += EXCHANGE_FEES_SWITCH[item.productId] * item.qty;
    
    if (PREMIUM_PRODUCTS.includes(item.productId)) {
      premium += PREMIUM_PER_PIECE * item.qty;
    }
  });

  const total = roundTo1000(switchFee + premium);

  try {
    showLoading();
    const result = await callAppsScript('ADD_SWITCH', {
      phone,
      oldGold: JSON.stringify(oldGold),
      newGold: JSON.stringify(newGold),
      switchFee,
      premium,
      total,
      user: currentUser.nickname
    });

    if (result.success) {
      alert('✅ สร้างรายการ Switch สำเร็จ!');
      closeModal('switchModal');
      document.getElementById('switchPhone').value = '';
      document.getElementById('switchOldGold').innerHTML = '';
      document.getElementById('switchNewGold').innerHTML = '';
      switchOldCounter = 0;
      switchNewCounter = 0;
      addSwitchOldGold();
      addSwitchNewGold();
      loadSwitches();
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

function toggleSwitchSortOrder() {
  switchSortOrder = switchSortOrder === 'desc' ? 'asc' : 'desc';
  document.getElementById('switchSortBtn').textContent = switchSortOrder === 'desc' ? '↓ Latest' : '↑ Oldest';
  loadSwitches();
}

function filterSwitchByDate() {
  switchDateFrom = document.getElementById('switchDateFrom').value;
  switchDateTo = document.getElementById('switchDateTo').value;
  loadSwitches();
}

function clearSwitchDateFilter() {
  switchDateFrom = null;
  switchDateTo = null;
  document.getElementById('switchDateFrom').value = '';
  document.getElementById('switchDateTo').value = '';
  loadSwitches();
}

async function openSwitchPaymentModal(switchId) {
  const data = await fetchSheetData('Switches!A:N');
  const switchData = data.slice(1).find(row => row[0] === switchId);
  if (!switchData) return;
  
  const oldGold = formatItemsForPayment(switchData[2]);
  const newGold = formatItemsForPayment(switchData[3]);
  const total = parseFloat(switchData[6]) || 0;
  
  openMultiPaymentModal('SWITCH', switchId, total, switchData[1], 
    `<strong>Old Gold:</strong> ${oldGold}<br><strong>New Gold:</strong> ${newGold}`);
}

async function reviewSwitch(switchId) {
  try {
    var data = await fetchSheetData('Switches!A:N');
    var sw = data.slice(1).find(function(row) { return row[0] === switchId; });
    if (sw) {
      openReviewDecisionModal('SWITCH', switchId, sw[3]);
    }
  } catch(e) { alert('❌ Error loading data'); }
}