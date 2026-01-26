async function loadInventory() {
  try {
    showLoading();
    const data = await fetchSheetData('Inventory!A:R');
    const tbody = document.getElementById('inventoryTable');
    
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="18" style="text-align: center; padding: 40px;">No inventory records</td></tr>';
      hideLoading();
      return;
    }
    
    const rows = data.slice(1).reverse();
    
    tbody.innerHTML = rows.map(row => `
      <tr>
        <td>${row[0]}</td>
        <td>${row[1]}</td>
        <td>${row[2]}</td>
        <td>${row[3]}</td>
        <td>${formatNumber(row[4])}</td>
        <td>${formatNumber(row[5])}</td>
        <td>${formatNumber(row[6])}</td>
        <td>${formatNumber(row[7])}</td>
        <td>${formatNumber(row[8])}</td>
        <td>${formatNumber(row[9])}</td>
        <td>${formatNumber(row[10])}</td>
        <td>${formatNumber(row[11])}</td>
        <td>${formatNumber(row[12])}</td>
        <td>${formatNumber(row[13])}</td>
        <td>${formatNumber(row[14])}</td>
        <td>${formatNumber(row[15])}</td>
        <td>${formatDate(row[16])}</td>
        <td>${row[17]}</td>
      </tr>
    `).join('');
    
    hideLoading();
  } catch (error) {
    console.error('Error loading inventory:', error);
    hideLoading();
  }
}

async function loadStock() {
  try {
    showLoading();
    const data = await fetchSheetData('Stock!A:G');
    
    if (data.length <= 1) {
      hideLoading();
      return;
    }
    
    const stockData = {};
    
    data.slice(1).forEach(row => {
      const productId = row[0];
      const type = row[1];
      const oldNew = row[2];
      const qty = parseFloat(row[4]) || 0;
      
      const key = `${productId}_${oldNew}`;
      
      if (!stockData[key]) {
        stockData[key] = {
          productId: productId,
          oldNew: oldNew,
          qty: 0
        };
      }
      
      if (type === 'IN') {
        stockData[key].qty += qty;
      } else {
        stockData[key].qty -= qty;
      }
    });
    
    const stockTbody = document.getElementById('stockSummaryTable');
    
    const stockRows = Object.values(stockData).map(item => {
      const product = FIXED_PRODUCTS.find(p => p.id === item.productId);
      return `
        <tr>
          <td>${product.name}</td>
          <td>${item.oldNew}</td>
          <td>${item.qty}</td>
        </tr>
      `;
    }).join('');
    
    stockTbody.innerHTML = stockRows || '<tr><td colspan="3" style="text-align: center; padding: 40px;">No stock data</td></tr>';
    
    hideLoading();
  } catch (error) {
    console.error('Error loading stock:', error);
    hideLoading();
  }
}

function openTransferModal() {
  document.getElementById('transferOldProducts').innerHTML = '';
  addTransferProduct();
  openModal('transferModal');
}

function addTransferProduct() {
  const container = document.getElementById('transferOldProducts');
  const row = document.createElement('div');
  row.className = 'product-row';
  row.style.cssText = 'display: flex; gap: 10px; margin-bottom: 10px; align-items: center;';
  
  row.innerHTML = `
    <select class="form-input" style="flex: 1;">
      ${FIXED_PRODUCTS.map(p => `<option value="${p.id}">${p.name}</option>`).join('')}
    </select>
    <input type="number" class="form-input" placeholder="Quantity" min="1" style="width: 150px;">
    <button class="btn-danger" onclick="this.parentElement.remove()" style="padding: 8px 15px;">Remove</button>
  `;
  
  container.appendChild(row);
}

async function confirmTransfer() {
  try {
    const rows = document.querySelectorAll('#transferOldProducts .product-row');
    const items = [];
    
    for (const row of rows) {
      const select = row.querySelector('select');
      const input = row.querySelector('input');
      const productId = select.value;
      const qty = parseInt(input.value);
      
      if (!qty || qty <= 0) {
        alert('กรุณากรอกจำนวนให้ถูกต้อง');
        return;
      }
      
      items.push({ productId, qty });
    }
    
    if (items.length === 0) {
      alert('กรุณาเพิ่มสินค้าอย่างน้อย 1 รายการ');
      return;
    }
    
    if (!confirm(`ยืนยันการโอนทองเก่าไปทองใหม่ ${items.length} รายการ?`)) {
      return;
    }
    
    showLoading();
    
    const result = await executeGoogleScript({
      action: 'TRANSFER_OLD_TO_NEW',
      items: JSON.stringify(items),
      user: currentUser.nickname
    });
    
    if (result.success) {
      alert('✅ ' + result.message);
      closeModal('transferModal');
      await loadInventory();
      await loadStock();
    } else {
      alert('❌ ' + result.message);
    }
    
    hideLoading();
  } catch (error) {
    console.error('Error transferring:', error);
    alert('❌ เกิดข้อผิดพลาด: ' + error.message);
    hideLoading();
  }
}
