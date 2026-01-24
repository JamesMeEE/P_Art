async function loadProducts() {
  try {
    showLoading();
    
    const pricingData = await fetchSheetData('Pricing!A:B');
    
    if (pricingData.length > 1) {
      const latestPricing = pricingData[pricingData.length - 1];
      currentPricing = {
        sell1Baht: parseFloat(latestPricing[1]) || 0,
        buyback1Baht: 0
      };
      
      document.getElementById('currentPriceDisplay').textContent = formatNumber(currentPricing.sell1Baht) + ' LAK';
    }
    
    const unitLabels = {
      'G01': '10B',
      'G02': '5B',
      'G03': '2B',
      'G04': '1B',
      'G05': '0.5B',
      'G06': '0.25B',
      'G07': '1g'
    };
    
    const tbody = document.getElementById('productsTable');
    tbody.innerHTML = FIXED_PRODUCTS.map(product => {
      const sellPrice = calculateSellPrice(product.id, currentPricing.sell1Baht);
      const buybackPrice = calculateBuybackPrice(product.id, currentPricing.sell1Baht);
      const exchangeFee = EXCHANGE_FEES[product.id];
      const unit = unitLabels[product.id] || product.unit;
      
      return `
        <tr>
          <td>${product.id}</td>
          <td>${product.name}</td>
          <td>${unit}</td>
          <td>${formatNumber(sellPrice)}</td>
          <td>${formatNumber(buybackPrice)}</td>
          <td>${formatNumber(exchangeFee)}</td>
        </tr>
      `;
    }).join('');
    
    await loadPriceHistory();
    
    hideLoading();
  } catch (error) {
    hideLoading();
  }
}

async function loadPriceHistory() {
  try {
    const data = await fetchSheetData('Pricing!A:C');
    const tbody = document.getElementById('priceHistoryTable');
    
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="3" style="text-align: center; padding: 40px;">No records</td></tr>';
      return;
    }
    
    tbody.innerHTML = data.slice(1).reverse().map(row => `
      <tr>
        <td>${formatDateTime(row[0])}</td>
        <td>${formatNumber(row[1])}</td>
        <td>${row[2] || '-'}</td>
      </tr>
    `).join('');
  } catch (error) {
  }
}

async function updatePricing() {
  const sell1Baht = document.getElementById('sell1BahtPrice').value;
  
  if (!sell1Baht) {
    alert('กรุณากรอกราคา 1 บาท');
    return;
  }
  
  try {
    showLoading();
    const result = await callAppsScript('UPDATE_PRICING', {
      sell1Baht
    });
    
    if (result.success) {
      alert('✅ อัพเดตราคาสำเร็จ!');
      closeModal('pricingModal');
      document.getElementById('sell1BahtPrice').value = '';
      loadProducts();
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}
