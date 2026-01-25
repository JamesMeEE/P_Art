async function loadSells() {
  try {
    showLoading();
    const data = await fetchSheetData('Sells!A:L');
    
    let filteredData = data.slice(1);
    
    if (currentUser.role === 'User' || currentUser.role === 'Manager') {
      if (sellDateFrom || sellDateTo) {
        filteredData = filterByDateRange(filteredData, 9, 11, sellDateFrom, sellDateTo);
      } else {
        filteredData = filterTodayData(filteredData, 9, 11);
      }
    }
    
    if (sellSortOrder === 'asc') {
      filteredData.sort((a, b) => new Date(a[9]) - new Date(b[9]));
    } else {
      filteredData.sort((a, b) => new Date(b[9]) - new Date(a[9]));
    }
    
    const tbody = document.getElementById('sellTable');
    if (filteredData.length === 0) {
      tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = filteredData.map(row => {
        const items = formatItemsForTable(row[2]);
        const premium = calculatePremiumFromItems(row[2]);
        const saleName = row[11];
        const status = row[10];
        
        let actions = '';
        
        if (status === 'PENDING') {
          if (currentUser.role === 'Manager') {
            actions = `<button class="btn-action" onclick="reviewSell('${row[0]}')">Review</button>`;
          } else {
            actions = '<span style="color: var(--text-secondary);">Waiting for review</span>';
          }
        } else if (status === 'READY') {
          if (currentUser.role === 'User') {
            actions = `<button class="btn-action" onclick="openPaymentModal('${row[0]}')">Confirm</button>`;
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
    console.error('❌ Error loading sells:', error);
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
    <select class="form-select" onchange="calculateSellTotal()">
      <option value="">เลือกสินค้า...</option>
      ${FIXED_PRODUCTS.map(p => `<option value="${p.id}">${p.name}</option>`).join('')}
    </select>
    <input type="number" class="form-input" placeholder="จำนวน" min="1" step="1" oninput="calculateSellTotal()">
    <button type="button" class="btn-remove" onclick="removeSellProduct(${sellCounter})">×</button>
  `;
  container.appendChild(row);
}

function removeSellProduct(id) {
  const row = document.getElementById(`sellProduct${id}`);
  if (row) {
    row.remove();
    calculateSellTotal();
  }
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
  let totalPremium = 0;

  items.forEach(item => {
    const pricePerPiece = calculateSellPrice(item.productId, currentPricing.sell1Baht);
    totalPrice += pricePerPiece * item.qty;
    
    if (PREMIUM_PRODUCTS.includes(item.productId)) {
      totalPremium += PREMIUM_PER_PIECE * item.qty;
    }
  });
  
  totalPrice = roundTo1000(totalPrice + totalPremium);

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

async function reviewSell(sellId) {
  if (!confirm('Approve this sell transaction?')) return;
  
  try {
    showLoading();
    const result = await callAppsScript('REVIEW_SELL', { sellId });
    
    if (result.success) {
      alert('✅ Transaction reviewed and ready for confirmation!');
      loadSells();
      loadDashboard();
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}

function calculateSellTotal() {
  let totalPrice = 0;
  let totalPremium = 0;
  
  
  document.querySelectorAll('#sellProducts .product-row').forEach((row, index) => {
    const productId = row.querySelector('select').value;
    const qty = parseInt(row.querySelector('input').value) || 0;
    
    if (productId && qty > 0) {
      const productName = FIXED_PRODUCTS.find(p => p.id === productId)?.name || productId;
      const pricePerPiece = calculateSellPrice(productId, currentPricing.sell1Baht);
      const lineTotal = pricePerPiece * qty;
      
      
      totalPrice += lineTotal;
      
      if (PREMIUM_PRODUCTS.includes(productId)) {
        const premium = PREMIUM_PER_PIECE * qty;
        totalPremium += premium;
      }
    }
  });
  
  
  const finalTotal = roundTo1000(totalPrice + totalPremium);
  
  const priceElement = document.getElementById('sellPrice');
  if (priceElement) {
    priceElement.value = finalTotal;
  }
}

async function openSellModal() {
  if (currentPricing.sell1Baht === 0) {
    showLoading();
    const data = await fetchSheetData('Pricing!A:C');
    if (data.length > 1) {
      const latestPricing = data[data.length - 1];
      currentPricing = {
        sell1Baht: parseFloat(latestPricing[1]) || 0,
        buyback1Baht: parseFloat(latestPricing[2]) || 0
      };
    }
    hideLoading();
  }
  
  if (currentPricing.sell1Baht === 0) {
    alert('กรุณากำหนดราคาก่อนใช้งาน (Products → Set Pricing)');
    return;
  }
  
  openModal('sellModal');
  
  if (document.querySelectorAll('#sellProducts .product-row').length === 0) {
    addSellProduct();
  }
}

function filterSellByDate() {
  sellDateFrom = document.getElementById('sellDateFrom').value;
  sellDateTo = document.getElementById('sellDateTo').value;
  loadSells();
}

function resetSellDateFilter() {
  const today = getTodayDateString();
  document.getElementById('sellDateFrom').value = today;
  document.getElementById('sellDateTo').value = today;
  sellDateFrom = today;
  sellDateTo = today;
  loadSells();
}
