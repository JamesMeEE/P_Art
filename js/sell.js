async function loadSells() {
  try {
    showLoading();
    const data = await fetchSheetData('Sells!A:I');
    
    let filteredData = data.slice(1);
    
    if (currentUser.role === 'User' || currentUser.role === 'Manager') {
      filteredData = filterTodayData(filteredData, 6, 8);
    }
    
    if (sellSortOrder === 'asc') {
      filteredData.sort((a, b) => new Date(a[6]) - new Date(b[6]));
    } else {
      filteredData.sort((a, b) => new Date(b[6]) - new Date(a[6]));
    }
    
    const tbody = document.getElementById('sellTable');
    if (filteredData.length === 0) {
      tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = filteredData.map(row => {
        const items = formatItemsForTable(row[2]);
        const premium = calculatePremiumFromItems(row[2]);
        const saleName = row[8];
        
        let actions = '';
        if (row[7] === 'PENDING') {
          if (currentUser.role === 'User') {
            actions = `<button class="btn-action" onclick="openPaymentModal('${row[0]}')">Confirm</button>`;
          } else if (currentUser.role === 'Manager') {
            actions = `<button class="btn-secondary" onclick="viewSellDetails('${row[0]}')">View</button>`;
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
            <td>${formatNumber(row[3])}</td>
            <td><span class="status-badge status-${row[7].toLowerCase()}">${row[7]}</span></td>
            <td>${saleName}</td>
            <td>${actions}</td>
          </tr>
        `;
      }).join('');
    }
    
    hideLoading();
  } catch (error) {
    console.error('Error loading sells:', error);
    hideLoading();
  }
}

function toggleSellSort() {
  sellSortOrder = sellSortOrder === 'desc' ? 'asc' : 'desc';
  document.getElementById('sellSortBtn').textContent = 
    sellSortOrder === 'desc' ? 'Sort: Newest First' : 'Sort: Oldest First';
  loadSells();
}

let sellCounter = 0;

function addSellProduct() {
  sellCounter++;
  const container = document.getElementById('sellProducts');
  const row = document.createElement('div');
  row.className = 'product-row';
  row.id = `sellProduct${sellCounter}`;
  row.innerHTML = `
    <select class="form-select">
      <option value="">เลือกสินค้า...</option>
      ${FIXED_PRODUCTS.map(p => `<option value="${p.id}">${p.name}</option>`).join('')}
    </select>
    <input type="number" class="form-input" placeholder="จำนวน" min="1" step="1">
    <button type="button" class="btn-remove" onclick="removeSellProduct(${sellCounter})">×</button>
  `;
  container.appendChild(row);
}

function removeSellProduct(id) {
  const row = document.getElementById(`sellProduct${id}`);
  if (row) row.remove();
}

async function submitSell() {
  const customer = document.getElementById('sellCustomer').value;
  if (!customer) {
    alert('กรุณากรอกชื่อลูกค้า');
    return;
  }

  const items = [];
  document.querySelectorAll('#sellProducts .product-row').forEach(row => {
    const productId = row.querySelector('select').value;
    const qty = parseInt(row.querySelector('input').value) || 0;
    if (productId && qty > 0) {
      items.push({ productId, qty });
    }
  });

  if (items.length === 0) {
    alert('กรุณาเลือกสินค้าอย่างน้อย 1 รายการ');
    return;
  }

  showLoading();
  for (const item of items) {
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

  let totalPrice = 0;
  let premium = 0;

  items.forEach(item => {
    const weight = GOLD_WEIGHTS[item.productId];
    const price = weight * currentPricing.sell1Baht;
    totalPrice += price * item.qty;

    if (PREMIUM_PRODUCTS.includes(item.productId)) {
      premium += PREMIUM_PER_PIECE * item.qty;
    }
  });

  totalPrice += premium;

  try {
    showLoading();
    const result = await callAppsScript('ADD_SELL', {
      customer,
      items: JSON.stringify(items),
      total: totalPrice
    });

    if (result.success) {
      alert('✅ สร้างรายการขายสำเร็จ!');
      closeModal('sellModal');
      document.getElementById('sellCustomer').value = '';
      document.getElementById('sellProducts').innerHTML = '';
      sellCounter = 0;
      addSellProduct();
      loadSells();
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

async function checkStock(productId, requiredQty) {
  const currentStock = await getCurrentStock(productId);
  return currentStock >= requiredQty;
}

async function getCurrentStock(productId) {
  const data = await fetchSheetData('Stock!A:F');
  let total = 0;
  data.slice(1).forEach(row => {
    if (row[0] === productId) {
      total += parseInt(row[3]) || 0;
    }
  });
  return total;
}

function viewSellDetails(sellId) {
  alert(`View details for ${sellId} (Manager view only)`);
}
