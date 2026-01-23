async function loadPriceRate() {
  try {
    showLoading();
    const data = await fetchSheetData('PriceRate!A:F');
    
    if (data.length > 1) {
      const latestRate = data[data.length - 1];
      currentPriceRates = {
        thbSell: parseFloat(latestRate[1]) || 0,
        usdSell: parseFloat(latestRate[2]) || 0,
        thbBuy: parseFloat(latestRate[3]) || 0,
        usdBuy: parseFloat(latestRate[4]) || 0
      };
      
      document.getElementById('rateTHBSell').textContent = formatNumber(currentPriceRates.thbSell);
      document.getElementById('rateUSDSell').textContent = formatNumber(currentPriceRates.usdSell);
      document.getElementById('rateTHBBuy').textContent = formatNumber(currentPriceRates.thbBuy);
      document.getElementById('rateUSDBuy').textContent = formatNumber(currentPriceRates.usdBuy);
    }
    
    const tbody = document.getElementById('priceRateTable');
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="6" style="text-align: center; padding: 40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = data.slice(1).reverse().map(row => `
        <tr>
          <td>${formatDateTime(row[0])}</td>
          <td>${formatNumber(row[1])}</td>
          <td>${formatNumber(row[2])}</td>
          <td>${formatNumber(row[3])}</td>
          <td>${formatNumber(row[4])}</td>
          <td>${row[5]}</td>
        </tr>
      `).join('');
    }
    
    hideLoading();
  } catch (error) {
    console.error('Error loading price rate:', error);
    hideLoading();
  }
}

async function submitPriceRate() {
  const thbSell = document.getElementById('rateTHBSellInput').value;
  const usdSell = document.getElementById('rateUSDSellInput').value;
  const thbBuy = document.getElementById('rateTHBBuyInput').value;
  const usdBuy = document.getElementById('rateUSDBuyInput').value;
  
  if (!thbSell || !usdSell || !thbBuy || !usdBuy) {
    alert('Please fill all exchange rates');
    return;
  }
  
  try {
    showLoading();
    const result = await callAppsScript('ADD_PRICE_RATE', {
      thbSell, usdSell, thbBuy, usdBuy
    });
    
    if (result.success) {
      alert('✅ Price rate updated successfully!');
      closeModal('priceRateModal');
      document.getElementById('rateTHBSellInput').value = '';
      document.getElementById('rateUSDSellInput').value = '';
      document.getElementById('rateTHBBuyInput').value = '';
      document.getElementById('rateUSDBuyInput').value = '';
      loadPriceRate();
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}
