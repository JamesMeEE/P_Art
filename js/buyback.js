async function loadBuybacks() {
  try {
    showLoading();
    const data = await fetchSheetData('Buybacks!A:G');
    
    let filteredData = data.slice(1);
    
    if (currentUser.role === 'User' || currentUser.role === 'Manager') {
      filteredData = filterTodayData(filteredData, 4, 6);
    }
    
    const tbody = document.getElementById('buybackTable');
    if (filteredData.length === 0) {
      tbody.innerHTML = '<tr><td colspan="6" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = filteredData.reverse().map(row => {
        const items = formatItemsForTable(row[2]);
        
        let actions = '';
        if (row[5] === 'PENDING') {
          if (currentUser.role === 'Manager' || currentUser.role === 'Teller') {
            actions = `<button class="btn-action" onclick="confirmBuyback('${row[0]}')">Confirm</button>`;
          }
        } else {
          actions = '-';
        }
        
        return `
          <tr>
            <td>${row[0]}</td>
            <td>${row[1]}</td>
            <td>${items}</td>
            <td>${formatNumber(row[3])}</td>
            <td><span class="status-badge status-${row[5].toLowerCase()}">${row[5]}</span></td>
            <td>${actions}</td>
          </tr>
        `;
      }).join('');
    }
    
    hideLoading();
  } catch (error) {
    console.error('Error loading buybacks:', error);
    hideLoading();
  }
}

let buybackCounter = 0;

function addBuybackProduct() {
  buybackCounter++;
  const container = document.getElementById('buybackProducts');
  const row = document.createElement('div');
  row.className = 'product-row';
  row.id = `buybackProduct${buybackCounter}`;
  row.innerHTML = `
    <select class="form-select">
      <option value="">เลือกสินค้า...</option>
      ${FIXED_PRODUCTS.map(p => `<option value="${p.id}">${p.name}</option>`).join('')}
    </select>
    <input type="number" class="form-input" placeholder="จำนวน" min="1" step="1">
    <button type="button" class="btn-remove" onclick="removeBuybackProduct(${buybackCounter})">×</button>
  `;
  container.appendChild(row);
}

function removeBuybackProduct(id) {
  const row = document.getElementById(`buybackProduct${id}`);
  if (row) row.remove();
}

async function submitBuyback() {
  const customer = document.getElementById('buybackCustomer').value;
  if (!customer) {
    alert('กรุณากรอกชื่อลูกค้า');
    return;
  }

  const items = [];
  document.querySelectorAll('#buybackProducts .product-row').forEach(row => {
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

  let totalPrice = 0;
  items.forEach(item => {
    const weight = GOLD_WEIGHTS[item.productId];
    const price = weight * currentPricing.buyback1Baht;
    totalPrice += price * item.qty;
  });

  try {
    showLoading();
    const result = await callAppsScript('ADD_BUYBACK', {
      customer,
      items: JSON.stringify(items),
      total: totalPrice
    });

    if (result.success) {
      alert('✅ สร้างรายการรับซื้อคืนสำเร็จ!');
      closeModal('buybackModal');
      document.getElementById('buybackCustomer').value = '';
      document.getElementById('buybackProducts').innerHTML = '';
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

async function confirmBuyback(buybackId) {
  if (!confirm('ยืนยันรายการรับซื้อคืน?')) return;
  
  try {
    showLoading();
    const result = await callAppsScript('CONFIRM_BUYBACK', { buybackId });
    
    if (result.success) {
      alert('✅ ยืนยันสำเร็จ!');
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
