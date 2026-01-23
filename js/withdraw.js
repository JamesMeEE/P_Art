async function loadWithdraw() {
  try {
    showLoading();
    const data = await fetchSheetData('Inventory!A:H');
    
    let filteredData = data.slice(1).filter(row => row[2] === 'WITHDRAW');
    
    if (currentUser.role === 'User' || currentUser.role === 'Manager') {
      filteredData = filterTodayData(filteredData, 5, 7);
    }
    
    const tbody = document.getElementById('withdrawTable');
    if (filteredData.length === 0) {
      tbody.innerHTML = '<tr><td colspan="7" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = filteredData.reverse().map(row => {
        const productName = FIXED_PRODUCTS.find(p => p.id === row[1])?.name || row[1];
        
        let actions = '';
        if (row[6] === 'PENDING' && currentUser.role === 'Manager') {
          actions = `<button class="btn-action" onclick="approveWithdraw('${row[0]}')">Approve</button>`;
        } else {
          actions = '-';
        }
        
        return `
          <tr>
            <td>${row[0]}</td>
            <td>${productName}</td>
            <td>${row[3]}</td>
            <td>${formatDateTime(row[5])}</td>
            <td><span class="status-badge status-${row[6].toLowerCase()}">${row[6]}</span></td>
            <td>${row[7]}</td>
            <td>${actions}</td>
          </tr>
        `;
      }).join('');
    }
    
    hideLoading();
  } catch (error) {
    console.error('Error loading withdraw:', error);
    hideLoading();
  }
}

async function addWithdraw() {
  const product = document.getElementById('withdrawProduct').value;
  const quantity = document.getElementById('withdrawQuantity').value;
  
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
    const result = await callAppsScript('ADD_WITHDRAW', {
      product,
      quantity
    });
    
    if (result.success) {
      alert('✅ สร้างคำขอเบิกสินค้าสำเร็จ!');
      closeModal('withdrawModal');
      document.getElementById('withdrawProduct').value = '';
      document.getElementById('withdrawQuantity').value = '';
      loadWithdraw();
    } else {
      alert('❌ เกิดข้อผิดพลาด: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ เกิดข้อผิดพลาด: ' + error.message);
    hideLoading();
  }
}

async function approveWithdraw(withdrawId) {
  if (!confirm('อนุมัติคำขอเบิกสินค้า?')) return;
  
  try {
    showLoading();
    const result = await callAppsScript('APPROVE_WITHDRAW', {
      withdrawId
    });
    
    if (result.success) {
      alert('✅ อนุมัติสำเร็จ!');
      loadWithdraw();
      loadInventory();
    } else {
      alert('❌ เกิดข้อผิดพลาด: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ เกิดข้อผิดพลาด: ' + error.message);
    hideLoading();
  }
}
