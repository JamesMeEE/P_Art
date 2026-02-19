async function loadWithdraws() {
  try {
    showLoading();
    const data = await fetchSheetData('Withdraws!A:J');
    
    let filteredData = data.slice(1);
    
    if (currentUser.role === 'User' || isManager()) {
      if (withdrawDateFrom || withdrawDateTo) {
        filteredData = filterByDateRange(filteredData, 6, 8, withdrawDateFrom, withdrawDateTo);
      } else {
        filteredData = filterTodayData(filteredData, 6, 8);
      }
    }
    
    if (withdrawSortOrder === 'asc') {
      filteredData.sort((a, b) => new Date(a[6]) - new Date(b[6]));
    } else {
      filteredData.sort((a, b) => new Date(b[6]) - new Date(a[6]));
    }
    
    const tbody = document.getElementById('withdrawTable');
    if (filteredData.length === 0) {
      tbody.innerHTML = '<tr><td colspan="8" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = filteredData.map(row => {
        const items = formatItemsForTable(row[2]);
        const premium = parseFloat(row[3]) || 0;
        const total = parseFloat(row[4]) || premium;
        const saleName = row[8];
        const status = row[7];
        
        let actions = '';
        
        if (status === 'PENDING') {
          if (isManager()) {
            actions = `<button class="btn-action" onclick="reviewWithdraw('${row[0]}')">Review</button>`;
          } else {
            actions = '<span style="color: var(--text-secondary);">Waiting for review</span>';
          }
        } else if (status === 'READY') {
          if (currentUser.role === 'User') {
            actions = `<button class="btn-action" onclick="openWithdrawPaymentModal('${row[0]}')">Confirm</button>`;
          } else {
            actions = '<span style="color: var(--text-secondary);">Waiting for confirmation</span>';
          }
        } else {
          var detail = encodeURIComponent(JSON.stringify([['Transaction ID', row[0]], ['Phone', row[1]], ['Items', items], ['Premium', formatNumber(premium) + ' LAK'], ['Total', formatNumber(total) + ' LAK'], ['Date', formatDateTime(row[6])], ['Status', status], ['Sale', saleName]]));
          actions = '<button class="btn-action" onclick="viewTransactionDetail(\'Withdraw\',\'' + detail + '\')" style="background:#555;">👁 View</button>';
        }
        
        return `
          <tr>
            <td>${row[0]}</td>
            <td>${row[1]}</td>
            <td>${items}</td>
            <td>${formatNumber(premium)}</td>
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
    console.error('Error loading withdraws:', error);
    hideLoading();
  }
}

let withdrawCounter = 0;

function addWithdrawProduct() {
  withdrawCounter++;
  const productOptions = FIXED_PRODUCTS.map(p => 
    `<option value="${p.id}">${p.name}</option>`
  ).join('');
  
  document.getElementById('withdrawProducts').insertAdjacentHTML('beforeend', `
    <div class="product-row" id="withdraw${withdrawCounter}">
      <select class="form-select" style="flex: 2;" onchange="calculateWithdrawPremium()">
        <option value="">Select Product</option>
        ${productOptions}
      </select>
      <input type="number" class="form-input" placeholder="Qty" min="1" style="flex: 1;" oninput="calculateWithdrawPremium()">
      <button type="button" class="btn-remove" onclick="document.getElementById('withdraw${withdrawCounter}').remove(); calculateWithdrawPremium();">×</button>
    </div>
  `);
}

function calculateWithdrawPremium() {
  const products = [];
  document.querySelectorAll('#withdrawProducts .product-row').forEach(row => {
    const productId = row.querySelector('select').value;
    const qty = parseInt(row.querySelector('input').value) || 0;
    if (productId && qty > 0) {
      products.push({ productId, qty });
    }
  });

  let premium = 0;
  products.forEach(item => {
    if (PREMIUM_PRODUCTS.includes(item.productId) || item.productId === 'G07') {
      premium += PREMIUM_PER_PIECE * item.qty;
    }
  });

  document.getElementById('withdrawPremium').value = formatNumber(premium);
}

async function calculateWithdraw() {
  if (_isSubmitting) return;
  const phone = document.getElementById('withdrawPhone').value;
  if (!phone) {
    alert('กรุณากรอกเบอร์โทร');
    return;
  }

  const products = [];
  document.querySelectorAll('#withdrawProducts .product-row').forEach(row => {
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

  let premium = 0;
  products.forEach(item => {
    if (PREMIUM_PRODUCTS.includes(item.productId) || item.productId === 'G07') {
      premium += PREMIUM_PER_PIECE * item.qty;
    }
  });

  const total = roundTo1000(premium);

  try {
    _isSubmitting = true;
    showLoading();
    const result = await callAppsScript('ADD_WITHDRAW', {
      phone,
      items: JSON.stringify(products),
      premium,
      total,
      user: currentUser.nickname
    });
    
    if (result.success) {
      alert('✅ สร้างรายการถอนทองสำเร็จ! รอ Manager Review');
      closeModal('withdrawModal');
      
      document.getElementById('withdrawPhone').value = '';
      document.getElementById('withdrawProducts').innerHTML = '';
      withdrawCounter = 0;
      addWithdrawProduct();
      
      loadWithdraws();
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

async function loadCurrentPricingForWithdraw() {
  try {
    const pricingData = await fetchSheetData('Pricing!A:B');
    
    if (pricingData.length > 1) {
      const latestPricing = pricingData[pricingData.length - 1];
      currentPricing = {
        sell1Baht: parseFloat(latestPricing[1]) || 0,
        buyback1Baht: 0
      };
      
      console.log('Loaded currentPricing for Withdraw:', currentPricing);
      return true;
    }
    return false;
  } catch (error) {
    console.error('Error loading pricing:', error);
    return false;
  }
}

async function openWithdrawModal() {
  const hasPrice = await loadCurrentPricingForWithdraw();
  
  if (!hasPrice || !currentPricing.sell1Baht || currentPricing.sell1Baht === 0) {
    alert('❌ ยังไม่มีราคาทองในระบบ! กรุณาไปที่หน้า Products → Set New Price ก่อน');
    return;
  }
  
  openModal('withdrawModal');
}


function resetWithdrawDateFilter() {
  const today = getTodayDateString();
  document.getElementById('withdrawDateFrom').value = today;
  document.getElementById('withdrawDateTo').value = today;
  withdrawDateFrom = today;
  withdrawDateTo = today;
  loadWithdraws();
}

document.addEventListener('DOMContentLoaded', function() {
  const fromInput = document.getElementById('withdrawDateFrom');
  const toInput = document.getElementById('withdrawDateTo');
  
  if (fromInput && toInput) {
    fromInput.addEventListener('change', function() {
      withdrawDateFrom = this.value;
      if (withdrawDateFrom && withdrawDateTo) {
        loadWithdraws();
      }
    });
    
    toInput.addEventListener('change', function() {
      withdrawDateTo = this.value;
      if (withdrawDateFrom && withdrawDateTo) {
        loadWithdraws();
      }
    });
  }
});

let currentWithdrawPayment = null;