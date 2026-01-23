async function loadProducts() {
  try {
    showLoading();
    
    const [pricingData, stockData] = await Promise.all([
      fetchSheetData('Pricing!A:C'),
      fetchSheetData('Stock!A:F')
    ]);
    
    if (pricingData.length > 1) {
      const latestPricing = pricingData[pricingData.length - 1];
      currentPricing = {
        sell1Baht: parseFloat(latestPricing[1]) || 0,
        buyback1Baht: parseFloat(latestPricing[2]) || 0
      };
      
      document.getElementById('currentPriceDisplay').textContent = formatNumber(currentPricing.sell1Baht) + ' LAK';
    }
    
    const stockMap = {};
    stockData.slice(1).forEach(row => {
      const productId = row[0];
      const qty = parseInt(row[3]) || 0;
      if (!stockMap[productId]) stockMap[productId] = 0;
      stockMap[productId] += qty;
    });
    
    const tbody = document.getElementById('productsTable');
    tbody.innerHTML = FIXED_PRODUCTS.map(product => {
      const weight = GOLD_WEIGHTS[product.id];
      const sellPrice = weight * currentPricing.sell1Baht;
      const buybackPrice = weight * currentPricing.buyback1Baht;
      const exchangeFee = EXCHANGE_FEES[product.id];
      const stock = stockMap[product.id] || 0;
      
      return `
        <tr>
          <td>${product.id}</td>
          <td>${product.name}</td>
          <td>${product.unit}</td>
          <td>${formatNumber(sellPrice)}</td>
          <td>${formatNumber(buybackPrice)}</td>
          <td>${formatNumber(exchangeFee)}</td>
          <td>${stock}</td>
        </tr>
      `;
    }).join('');
    
    hideLoading();
  } catch (error) {
    console.error('Error loading products:', error);
    hideLoading();
  }
}

async function updatePricing() {
  const sell1Baht = document.getElementById('pricingSell1Baht').value;
  const buyback1Baht = document.getElementById('pricingBuyback1Baht').value;
  const premium = document.getElementById('pricingPremium').value;
  
  if (!sell1Baht || !buyback1Baht || !premium) {
    alert('กรุณากรอกข้อมูลให้ครบ');
    return;
  }
  
  try {
    showLoading();
    const result = await callAppsScript('UPDATE_PRICING', {
      sell1Baht,
      buyback1Baht,
      premium
    });
    
    if (result.success) {
      alert('✅ อัพเดตราคาสำเร็จ!');
      PREMIUM_PER_PIECE = parseFloat(premium);
      closeModal('pricingModal');
      document.getElementById('pricingSell1Baht').value = '';
      document.getElementById('pricingBuyback1Baht').value = '';
      document.getElementById('pricingPremium').value = '';
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
