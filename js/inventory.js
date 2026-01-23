async function loadInventory() {
  try {
    showLoading();
    const data = await fetchSheetData('Inventory!A:H');
    
    const tbody = document.getElementById('inventoryTable');
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="7" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = data.slice(1).reverse().map(row => {
        const productName = FIXED_PRODUCTS.find(p => p.id === row[1])?.name || row[1];
        return `
          <tr>
            <td>${row[0]}</td>
            <td>${productName}</td>
            <td>${row[2]}</td>
            <td>${row[3]}</td>
            <td>${formatDateTime(row[5])}</td>
            <td><span class="status-badge status-${row[6].toLowerCase()}">${row[6]}</span></td>
            <td>${row[7]}</td>
          </tr>
        `;
      }).join('');
    }
    
    hideLoading();
  } catch (error) {
    console.error('Error loading inventory:', error);
    hideLoading();
  }
}

async function addStockIn() {
  const product = document.getElementById('stockInProduct').value;
  const quantity = document.getElementById('stockInQuantity').value;
  const cost = document.getElementById('stockInCost').value;
  
  if (!product || !quantity || !cost) {
    alert('กรุณากรอกข้อมูลให้ครบ');
    return;
  }
  
  try {
    showLoading();
    const result = await callAppsScript('ADD_STOCK_IN', {
      product,
      quantity,
      cost
    });
    
    if (result.success) {
      alert('✅ บันทึกสินค้าเข้าสำเร็จ!');
      closeModal('stockInModal');
      document.getElementById('stockInProduct').value = '';
      document.getElementById('stockInQuantity').value = '';
      document.getElementById('stockInCost').value = '';
      loadInventory();
      loadProducts();
    } else {
      alert('❌ เกิดข้อผิดพลาด: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ เกิดข้อผิดพลาด: ' + error.message);
    hideLoading();
  }
}

async function addStockOut() {
  const product = document.getElementById('stockOutProduct').value;
  const quantity = document.getElementById('stockOutQuantity').value;
  
  if (!product || !quantity) {
    alert('กรุณากรอกข้อมูลให้ครบ');
    return;
  }
  
  const hasStock = await checkStock(product, parseInt(quantity));
  if (!hasStock) {
    const currentStock = await getCurrentStock(product);
    const productName = FIXED_PRODUCTS.find(p => p.id === product).name;
    alert(`❌ สินค้าไม่พอสำหรับ ${productName}!\nต้องการ: ${quantity}, มีอยู่: ${currentStock}`);
    return;
  }
  
  try {
    showLoading();
    const result = await callAppsScript('ADD_STOCK_OUT', {
      product,
      quantity
    });
    
    if (result.success) {
      alert('✅ บันทึกสินค้าสูญหายสำเร็จ!');
      closeModal('stockOutModal');
      document.getElementById('stockOutProduct').value = '';
      document.getElementById('stockOutQuantity').value = '';
      loadInventory();
      loadProducts();
    } else {
      alert('❌ เกิดข้อผิดพลาด: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ เกิดข้อผิดพลาด: ' + error.message);
    hideLoading();
  }
}
