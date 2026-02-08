let currentReviewData = null;

function openReviewDecisionModal(type, id, items) {
  currentReviewData = { type, id, items };
  
  const titles = {
    'SELL': 'Review Sell',
    'TRADEIN': 'Review Trade-in',
    'EXCHANGE': 'Review Exchange',
    'WITHDRAW': 'Review Withdraw'
  };
  
  document.getElementById('reviewDecisionTitle').textContent = titles[type] || 'Review';
  document.getElementById('reviewDecisionId').textContent = `Transaction ID: ${id}`;
  document.getElementById('reviewDecisionNote').value = '';
  
  let itemsHTML = '<table style="width: 100%; font-size: 14px;">';
  itemsHTML += '<tr style="border-bottom: 1px solid var(--border-color);"><th style="text-align: left; padding: 5px;">Product</th><th style="text-align: right; padding: 5px;">Qty</th></tr>';
  
  try {
    const itemsArray = typeof items === 'string' ? JSON.parse(items) : items;
    itemsArray.forEach(item => {
      const productName = FIXED_PRODUCTS.find(p => p.id === item.productId)?.name || item.productId;
      itemsHTML += `<tr><td style="padding: 5px;">${productName}</td><td style="text-align: right; padding: 5px;">${item.qty}</td></tr>`;
    });
  } catch (e) {
    itemsHTML += '<tr><td colspan="2" style="padding: 5px; color: var(--text-secondary);">No items</td></tr>';
  }
  
  itemsHTML += '</table>';
  document.getElementById('reviewDecisionItems').innerHTML = itemsHTML;
  
  openModal('reviewDecisionModal');
}

async function submitReviewDecision(decision) {
  if (!currentReviewData) return;
  
  const note = document.getElementById('reviewDecisionNote').value.trim();
  
  const actionMap = {
    'SELL': 'REVIEW_SELL',
    'TRADEIN': 'REVIEW_TRADEIN',
    'EXCHANGE': 'REVIEW_EXCHANGE',
    'WITHDRAW': 'REVIEW_WITHDRAW'
  };
  
  const action = actionMap[currentReviewData.type];
  if (!action) {
    alert('❌ Unknown transaction type');
    return;
  }
  
  try {
    showLoading();
    
    const result = await callAppsScript(action, {
      id: currentReviewData.id,
      decision: decision,
      approvedBy: currentUser.nickname,
      note: note
    });
    
    if (result.success) {
      const msg = decision === 'APPROVE' ? '✅ Approved!' : '❌ Rejected!';
      alert(msg);
      closeModal('reviewDecisionModal');
      
      const type = currentReviewData.type;
      currentReviewData = null;
      
      if (type === 'SELL') loadSells();
      else if (type === 'TRADEIN') loadTradeins();
      else if (type === 'EXCHANGE') loadExchanges();
      else if (type === 'WITHDRAW') loadWithdraws();
      
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

async function reviewSell(sellId) {
  try {
    const data = await fetchSheetData('Sells!A:L');
    const sell = data.slice(1).find(row => row[0] === sellId);
    if (sell) {
      openReviewDecisionModal('SELL', sellId, sell[2]);
    }
  } catch (e) {
    alert('❌ Error loading data');
  }
}

async function reviewTradein(tradeinId) {
  try {
    const data = await fetchSheetData('Tradeins!A:N');
    const tradein = data.slice(1).find(row => row[0] === tradeinId);
    if (tradein) {
      openReviewDecisionModal('TRADEIN', tradeinId, tradein[3]);
    }
  } catch (e) {
    alert('❌ Error loading data');
  }
}

async function reviewExchange(exchangeId) {
  try {
    const data = await fetchSheetData('Exchanges!A:N');
    const exchange = data.slice(1).find(row => row[0] === exchangeId);
    if (exchange) {
      openReviewDecisionModal('EXCHANGE', exchangeId, exchange[3]);
    }
  } catch (e) {
    alert('❌ Error loading data');
  }
}

async function reviewWithdraw(withdrawId) {
  try {
    const data = await fetchSheetData('Withdraws!A:J');
    const withdraw = data.slice(1).find(row => row[0] === withdrawId);
    if (withdraw) {
      openReviewDecisionModal('WITHDRAW', withdrawId, withdraw[2]);
    }
  } catch (e) {
    alert('❌ Error loading data');
  }
}