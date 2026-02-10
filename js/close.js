let currentCloseId = null;

async function openCloseWorkModal() {
  try {
    showLoading();
    
    const today = new Date();
    const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const todayEnd = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);
    const userName = currentUser.nickname;
    
    const [sellData, tradeinData, exchangeData, buybackData, withdrawData, switchData, freeExData] = await Promise.all([
      fetchSheetData('Sells!A:L'),
      fetchSheetData('Tradeins!A:N'),
      fetchSheetData('Exchanges!A:N'),
      fetchSheetData('Buybacks!A:L'),
      fetchSheetData('Withdraws!A:J'),
      fetchSheetData('Switches!A:N'),
      fetchSheetData('FreeExchanges!A:J')
    ]);
    
    let cashReceived = { LAK: 0, THB: 0, USD: 0 };
    let oldGoldReceived = {};
    let newGoldGiven = {};
    
    const addOldGold = (items) => {
      try {
        const parsed = JSON.parse(items);
        parsed.forEach(item => {
          if (!oldGoldReceived[item.productId]) oldGoldReceived[item.productId] = 0;
          oldGoldReceived[item.productId] += item.qty;
        });
      } catch(e) {}
    };
    
    const addNewGold = (items) => {
      try {
        const parsed = JSON.parse(items);
        parsed.forEach(item => {
          if (!newGoldGiven[item.productId]) newGoldGiven[item.productId] = 0;
          newGoldGiven[item.productId] += item.qty;
        });
      } catch(e) {}
    };
    
    const processCash = (currency, customerPaid, changeLAK) => {
      if (currency === 'LAK') {
        cashReceived.LAK += customerPaid - changeLAK;
      } else if (currency === 'THB') {
        cashReceived.THB += customerPaid;
        cashReceived.LAK -= changeLAK;
      } else if (currency === 'USD') {
        cashReceived.USD += customerPaid;
        cashReceived.LAK -= changeLAK;
      }
    };
    
    sellData.slice(1).forEach(row => {
      const date = parseSheetDate(row[9]);
      const status = row[10];
      const createdBy = row[11];
      const isToday = date && date >= todayStart && date <= todayEnd;
      if (isToday && status === 'COMPLETED' && createdBy === userName) {
        processCash(row[6] || 'LAK', parseFloat(row[5]) || 0, parseFloat(row[8]) || 0);
        addNewGold(row[2]);
      }
    });
    
    tradeinData.slice(1).forEach(row => {
      const date = parseSheetDate(row[11]);
      const status = row[12];
      const createdBy = row[13];
      const isToday = date && date >= todayStart && date <= todayEnd;
      if (isToday && status === 'COMPLETED' && createdBy === userName) {
        processCash(row[8] || 'LAK', parseFloat(row[7]) || 0, parseFloat(row[10]) || 0);
        addOldGold(row[2]);
        addNewGold(row[3]);
      }
    });
    
    exchangeData.slice(1).forEach(row => {
      const date = parseSheetDate(row[11]);
      const status = row[12];
      const createdBy = row[13];
      const isToday = date && date >= todayStart && date <= todayEnd;
      if (isToday && status === 'COMPLETED' && createdBy === userName) {
        processCash(row[8] || 'LAK', parseFloat(row[7]) || 0, parseFloat(row[10]) || 0);
        addOldGold(row[2]);
        addNewGold(row[3]);
      }
    });
    
    buybackData.slice(1).forEach(row => {
      const date = parseSheetDate(row[9]);
      const status = row[10];
      const createdBy = row[11];
      const isToday = date && date >= todayStart && date <= todayEnd;
      if (isToday && status === 'COMPLETED' && createdBy === userName) {
        addOldGold(row[2]);
      }
    });
    
    withdrawData.slice(1).forEach(row => {
      const date = parseSheetDate(row[6]);
      const status = row[7];
      const createdBy = row[8];
      const isToday = date && date >= todayStart && date <= todayEnd;
      if (isToday && status === 'COMPLETED' && createdBy === userName) {
        addNewGold(row[2]);
      }
    });
    
    switchData.slice(1).forEach(row => {
      const date = parseSheetDate(row[11]);
      const status = row[12];
      const createdBy = row[13];
      const isToday = date && date >= todayStart && date <= todayEnd;
      if (isToday && status === 'COMPLETED' && createdBy === userName) {
        processCash(row[8] || 'LAK', parseFloat(row[7]) || 0, parseFloat(row[10]) || 0);
        addOldGold(row[2]);
        addNewGold(row[3]);
      }
    });
    
    freeExData.slice(1).forEach(row => {
      const date = parseSheetDate(row[7]);
      const status = row[8];
      const createdBy = row[9];
      const isToday = date && date >= todayStart && date <= todayEnd;
      if (isToday && status === 'COMPLETED' && createdBy === userName) {
        addOldGold(row[2]);
        addNewGold(row[3]);
      }
    });
    
    const productNames = {
      'G01': '10 บาท', 'G02': '5 บาท', 'G03': '2 บาท', 'G04': '1 บาท',
      'G05': '2 สลึง', 'G06': '1 สลึง', 'G07': '1 กรัม'
    };
    
    const buildGoldHTML = (goldObj) => {
      let html = '';
      let hasData = false;
      Object.keys(goldObj).sort().forEach(pid => {
        const qty = goldObj[pid];
        if (qty > 0) {
          hasData = true;
          html += `<tr><td>${productNames[pid] || pid}</td><td style="text-align:right;font-weight:bold;">${qty} ชิ้น</td></tr>`;
        }
      });
      return hasData ? html : '<tr><td colspan="2" style="text-align:center;color:var(--text-secondary);">ไม่มี</td></tr>';
    };
    
    document.getElementById('closeWorkSummary').innerHTML = `
      <div style="text-align:center;margin-bottom:20px;">
        <p style="font-size:18px;color:var(--gold-primary);">สรุปยอดประจำวัน</p>
        <p style="color:var(--text-secondary);">${userName} - ${formatDateOnly(today)}</p>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:20px;">
        <div class="stat-card">
          <h3 style="color:var(--gold-primary);margin-bottom:15px;">💵 เงินสดที่ได้รับ</h3>
          <table style="width:100%;">
            <tr><td>LAK</td><td style="text-align:right;font-weight:bold;font-size:18px;">${formatNumber(cashReceived.LAK)}</td></tr>
            <tr><td>THB</td><td style="text-align:right;font-weight:bold;font-size:18px;">${formatNumber(cashReceived.THB)}</td></tr>
            <tr><td>USD</td><td style="text-align:right;font-weight:bold;font-size:18px;">${formatNumber(cashReceived.USD)}</td></tr>
          </table>
        </div>
        <div class="stat-card">
          <h3 style="color:#ff9800;margin-bottom:15px;">🥇 ทองเก่าที่ได้รับ (IN)</h3>
          <table style="width:100%;">${buildGoldHTML(oldGoldReceived)}</table>
        </div>
        <div class="stat-card">
          <h3 style="color:#2196f3;margin-bottom:15px;">💎 ทองใหม่ที่จ่ายออก (OUT)</h3>
          <table style="width:100%;">${buildGoldHTML(newGoldGiven)}</table>
        </div>
      </div>
    `;
    
    window.currentCloseSummary = {
      user: userName,
      date: today.toISOString(),
      cashLAK: cashReceived.LAK,
      cashTHB: cashReceived.THB,
      cashUSD: cashReceived.USD,
      oldGold: JSON.stringify(oldGoldReceived),
      newGold: JSON.stringify(newGoldGiven)
    };
    
    hideLoading();
    openModal('closeWorkModal');
  } catch (error) {
    console.error('Error opening close work modal:', error);
    hideLoading();
    alert('❌ Error: ' + error.message);
  }
}

async function submitCloseWork() {
  if (!window.currentCloseSummary) return;
  
  try {
    showLoading();
    const result = await callAppsScript('SUBMIT_CLOSE', window.currentCloseSummary);
    if (result.success) {
      alert('✅ ส่ง Close สำเร็จ! รอ Manager อนุมัติ');
      closeModal('closeWorkModal');
      window.currentCloseSummary = null;
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}

async function checkPendingClose() {
  const closeBtn = document.getElementById('closeWorkBtn');
  const reviewBtn = document.getElementById('reviewCloseBtn');
  
  if (!currentUser) {
    if (closeBtn) closeBtn.style.display = 'none';
    if (reviewBtn) reviewBtn.style.display = 'none';
    return;
  }
  
  if (currentUser.role === 'Manager') {
    if (closeBtn) closeBtn.style.display = 'none';
    
    try {
      const closeData = await fetchSheetData('Close!A:K');
      const pendingCount = closeData.slice(1).filter(row => row[8] === 'PENDING').length;
      
      if (pendingCount > 0) {
        reviewBtn.style.display = 'inline-block';
        reviewBtn.textContent = '📋 Review Close (' + pendingCount + ')';
      } else {
        reviewBtn.style.display = 'none';
      }
    } catch (error) {
      console.error('Error checking pending close:', error);
      reviewBtn.style.display = 'none';
    }
  } else {
    if (closeBtn) closeBtn.style.display = 'inline-block';
    if (reviewBtn) reviewBtn.style.display = 'none';
  }
}

async function openReviewCloseModal() {
  try {
    showLoading();
    
    const closeData = await fetchSheetData('Close!A:K');
    const pendingCloses = closeData.slice(1).filter(row => row[8] === 'PENDING');
    
    if (pendingCloses.length === 0) {
      document.getElementById('reviewCloseList').innerHTML = '<p style="text-align:center;color:var(--text-secondary);padding:40px;">ไม่มีรายการ Close ที่รอการอนุมัติ</p>';
    } else {
      document.getElementById('reviewCloseList').innerHTML = `
        <table class="data-table" style="width:100%;">
          <thead><tr><th>ID</th><th>User</th><th>Date</th><th>Time</th><th>Action</th></tr></thead>
          <tbody>
            ${pendingCloses.map(row => `<tr>
              <td>${row[0]}</td><td>${row[1]}</td>
              <td>${formatDateOnly(parseSheetDate(row[2]))}</td>
              <td>${formatDateTime(row[7])}</td>
              <td><button class="btn-primary" style="padding:5px 15px;" onclick="openCloseDetail('${row[0]}')">Review</button></td>
            </tr>`).join('')}
          </tbody>
        </table>`;
    }
    
    hideLoading();
    openModal('reviewCloseModal');
  } catch (error) {
    console.error('Error opening review close modal:', error);
    hideLoading();
    alert('❌ Error: ' + error.message);
  }
}

async function openCloseDetail(closeId) {
  try {
    showLoading();
    currentCloseId = closeId;
    
    const closeData = await fetchSheetData('Close!A:K');
    const closeRecord = closeData.slice(1).find(row => row[0] === closeId);
    
    if (!closeRecord) {
      alert('ไม่พบข้อมูล');
      hideLoading();
      return;
    }
    
    const productNames = {
      'G01': '10 บาท', 'G02': '5 บาท', 'G03': '2 บาท', 'G04': '1 บาท',
      'G05': '2 สลึง', 'G06': '1 สลึง', 'G07': '1 กรัม'
    };
    
    const buildGoldTable = (jsonStr) => {
      let html = '';
      let hasData = false;
      try {
        const obj = JSON.parse(jsonStr);
        Object.keys(obj).sort().forEach(pid => {
          if (obj[pid] > 0) {
            hasData = true;
            html += '<tr><td>' + (productNames[pid] || pid) + '</td><td style="text-align:right;font-weight:bold;">' + obj[pid] + ' ชิ้น</td></tr>';
          }
        });
      } catch(e) {}
      return hasData ? html : '<tr><td colspan="2" style="text-align:center;color:var(--text-secondary);">ไม่มี</td></tr>';
    };
    
    document.getElementById('closeDetailContent').innerHTML = `
      <div style="text-align:center;margin-bottom:20px;">
        <p style="font-size:20px;color:var(--gold-primary);font-weight:bold;">${closeRecord[1]}</p>
        <p style="color:var(--text-secondary);">Close ID: ${closeRecord[0]}</p>
        <p style="color:var(--text-secondary);">ส่งเมื่อ: ${formatDateTime(closeRecord[7])}</p>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:20px;">
        <div class="stat-card">
          <h3 style="color:var(--gold-primary);margin-bottom:15px;">💵 เงินสด</h3>
          <table style="width:100%;">
            <tr><td>LAK</td><td style="text-align:right;font-weight:bold;font-size:18px;">${formatNumber(closeRecord[3])}</td></tr>
            <tr><td>THB</td><td style="text-align:right;font-weight:bold;font-size:18px;">${formatNumber(closeRecord[4])}</td></tr>
            <tr><td>USD</td><td style="text-align:right;font-weight:bold;font-size:18px;">${formatNumber(closeRecord[5])}</td></tr>
          </table>
        </div>
        <div class="stat-card">
          <h3 style="color:#ff9800;margin-bottom:15px;">🥇 ทองเก่า (IN)</h3>
          <table style="width:100%;">${buildGoldTable(closeRecord[6])}</table>
        </div>
        <div class="stat-card">
          <h3 style="color:#2196f3;margin-bottom:15px;">💎 ทองใหม่ (OUT)</h3>
          <table style="width:100%;">${buildGoldTable(closeRecord[10])}</table>
        </div>
      </div>
    `;
    
    closeModal('reviewCloseModal');
    hideLoading();
    openModal('closeDetailModal');
  } catch (error) {
    console.error('Error opening close detail:', error);
    hideLoading();
    alert('❌ Error: ' + error.message);
  }
}

async function approveClose() {
  if (!currentCloseId) return;
  
  try {
    showLoading();
    const result = await callAppsScript('APPROVE_CLOSE', {
      closeId: currentCloseId,
      approvedBy: currentUser.nickname
    });
    if (result.success) {
      alert('✅ อนุมัติ Close สำเร็จ!');
      closeModal('closeDetailModal');
      currentCloseId = null;
      checkPendingClose();
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (error) {
    alert('❌ Error: ' + error.message);
    hideLoading();
  }
}