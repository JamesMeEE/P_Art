let currentCloseId = null;

async function openCloseWorkModal() {
  try {
    showLoading();
    
    const today = new Date();
    const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const todayEnd = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);
    const userName = currentUser.nickname;

    const closeHistory = await fetchSheetData('Close!A:K');
    const alreadyClosed = closeHistory.slice(1).find(row => {
      const d = parseSheetDate(row[2]);
      const isToday = d && d >= todayStart && d <= todayEnd;
      return isToday && row[1] === userName && (row[8] === 'PENDING' || row[8] === 'APPROVED');
    });
    if (alreadyClosed) {
      hideLoading();
      alert('❌ คุณได้ปิดงานวันนี้แล้ว (' + alreadyClosed[0] + ' - ' + alreadyClosed[8] + ')');
      return;
    }
    
    var userSheetData = await fetchSheetData("'" + userName + "'!A:I");
    
    let cashReceived = { LAK: 0, THB: 0, USD: 0 };
    
    if (userSheetData && userSheetData.length > 1) {
      for (var i = 1; i < userSheetData.length; i++) {
        var r = userSheetData[i];
        var method = String(r[4] || '').trim();
        var currency = String(r[3] || '').trim();
        var amount = parseFloat(r[2]) || 0;
        if (method === 'Cash' && cashReceived.hasOwnProperty(currency)) {
          cashReceived[currency] += amount;
        }
      }
    }

    var userGoldData = await fetchSheetData("'" + userName + "_Gold'!A:F");
    
    let oldGoldReceived = {};
    if (userGoldData && userGoldData.length > 1) {
      for (var gi = 1; gi < userGoldData.length; gi++) {
        var gr = userGoldData[gi];
        var pid = String(gr[0] || '').trim();
        var gqty = parseFloat(gr[1]) || 0;
        if (pid && gqty > 0) {
          if (!oldGoldReceived[pid]) oldGoldReceived[pid] = 0;
          oldGoldReceived[pid] += gqty;
        }
      }
    }

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
          html += `<tr><td style="padding:6px 0;white-space:nowrap;">${productNames[pid] || pid}</td><td style="text-align:right;font-weight:bold;padding:6px 0;white-space:nowrap;">${qty} ชิ้น</td></tr>`;
        }
      });
      return hasData ? html : '<tr><td colspan="2" style="text-align:center;color:var(--text-secondary);padding:15px 0;">ไม่มี</td></tr>';
    };
    
    document.getElementById('closeWorkSummary').innerHTML = `
      <div style="text-align:center;margin-bottom:20px;">
        <p style="font-size:18px;color:var(--gold-primary);">สรุปยอดประจำวัน</p>
        <p style="color:var(--text-secondary);">${userName} - ${formatDateOnly(today)}</p>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;">
        <div class="stat-card">
          <h3 style="color:var(--gold-primary);margin-bottom:15px;">💵 เงินสด</h3>
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
      </div>
    `;
    
    window.currentCloseSummary = {
      user: userName,
      date: today.toISOString(),
      cashLAK: cashReceived.LAK,
      cashTHB: cashReceived.THB,
      cashUSD: cashReceived.USD,
      oldGold: JSON.stringify(oldGoldReceived)
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

var _autoRefreshInterval = null;

function startAutoRefresh() {
  stopAutoRefresh();
  _autoRefreshInterval = setInterval(function() {
    checkPendingClose();
    if (typeof loadPendingTransferCount === 'function') loadPendingTransferCount();
  }, 30000);
}

function stopAutoRefresh() {
  if (_autoRefreshInterval) {
    clearInterval(_autoRefreshInterval);
    _autoRefreshInterval = null;
  }
}

async function checkPendingClose() {
  const closeBtn = document.getElementById('closeWorkBtn');
  const reviewBtn = document.getElementById('reviewCloseBtn');
  
  if (!currentUser) {
    if (closeBtn) closeBtn.style.display = 'none';
    if (reviewBtn) reviewBtn.style.display = 'none';
    var tcBtn = document.getElementById('transferCashBtn');
    if (tcBtn) tcBtn.style.display = 'none';
    return;
  }
  
  if (currentUser.role === 'Manager') {
    if (closeBtn) closeBtn.style.display = 'none';
    var tcBtn2 = document.getElementById('transferCashBtn');
    if (tcBtn2) tcBtn2.style.display = 'none';
    
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
    var tcBtn3 = document.getElementById('transferCashBtn');
    if (tcBtn3) tcBtn3.style.display = 'inline-block';
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
            html += '<tr><td style="padding:6px 0;white-space:nowrap;">' + (productNames[pid] || pid) + '</td><td style="text-align:right;font-weight:bold;padding:6px 0;white-space:nowrap;">' + obj[pid] + ' ชิ้น</td></tr>';
          }
        });
      } catch(e) {}
      return hasData ? html : '<tr><td colspan="2" style="text-align:center;color:var(--text-secondary);padding:15px 0;">ไม่มี</td></tr>';
    };
    
    document.getElementById('closeDetailContent').innerHTML = `
      <div style="text-align:center;margin-bottom:20px;">
        <p style="font-size:20px;color:var(--gold-primary);font-weight:bold;">${closeRecord[1]}</p>
        <p style="color:var(--text-secondary);">Close ID: ${closeRecord[0]} — ส่งเมื่อ: ${formatDateTime(closeRecord[7])}</p>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:20px;">
        <div class="stat-card" style="padding:20px;">
          <h3 style="color:var(--gold-primary);margin-bottom:15px;font-size:15px;white-space:nowrap;">💵 เงินสด</h3>
          <table style="width:100%;border-collapse:collapse;">
            <tr><td style="padding:6px 0;">LAK</td><td style="text-align:right;font-weight:bold;font-size:18px;padding:6px 0;">${formatNumber(closeRecord[3])}</td></tr>
            <tr><td style="padding:6px 0;">THB</td><td style="text-align:right;font-weight:bold;font-size:18px;padding:6px 0;">${formatNumber(closeRecord[4])}</td></tr>
            <tr><td style="padding:6px 0;">USD</td><td style="text-align:right;font-weight:bold;font-size:18px;padding:6px 0;">${formatNumber(closeRecord[5])}</td></tr>
          </table>
        </div>
        <div class="stat-card" style="padding:20px;">
          <h3 style="color:#ff9800;margin-bottom:15px;font-size:15px;white-space:nowrap;">🥇 ทองเก่า (IN)</h3>
          <table style="width:100%;border-collapse:collapse;">${buildGoldTable(closeRecord[6])}</table>
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

async function rejectClose() {
  if (!currentCloseId) return;
  if (!confirm('ปฏิเสธการปิดงาน ' + currentCloseId + '?')) return;

  try {
    showLoading();
    const result = await callAppsScript('REJECT_CLOSE', {
      closeId: currentCloseId,
      approvedBy: currentUser.nickname
    });
    if (result.success) {
      alert('✅ ปฏิเสธ Close สำเร็จ');
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

var _transferCashBalances = { LAK: 0, THB: 0, USD: 0 };

async function openTransferCashModal() {
  try {
    showLoading();
    var userName = currentUser.nickname;
    var userSheetData = await fetchSheetData("'" + userName + "'!A:I");
    
    _transferCashBalances = { LAK: 0, THB: 0, USD: 0 };
    if (userSheetData && userSheetData.length > 1) {
      for (var i = 1; i < userSheetData.length; i++) {
        var r = userSheetData[i];
        var method = String(r[4] || '').trim();
        var currency = String(r[3] || '').trim();
        var amount = parseFloat(r[2]) || 0;
        if (method === 'Cash' && _transferCashBalances.hasOwnProperty(currency)) {
          _transferCashBalances[currency] += amount;
        }
      }
    }
    
    document.getElementById('transferCashCurrency').value = 'LAK';
    document.getElementById('transferCashAmount').value = '';
    updateTransferCashBalance();
    
    hideLoading();
    openModal('transferCashModal');
  } catch (e) {
    hideLoading();
    alert('❌ Error: ' + e.message);
  }
}

function updateTransferCashBalance() {
  var currency = document.getElementById('transferCashCurrency').value;
  var bal = _transferCashBalances[currency] || 0;
  document.getElementById('transferCashBalance').innerHTML =
    '<div style="display:flex;justify-content:space-between;align-items:center;">' +
    '<span style="color:var(--text-secondary);">ยอดเงินสด ' + currency + ' ของคุณ</span>' +
    '<span style="font-size:20px;font-weight:bold;color:var(--gold-primary);">' + formatNumber(bal) + ' ' + currency + '</span>' +
    '</div>';
}

async function confirmTransferCash() {
  var currency = document.getElementById('transferCashCurrency').value;
  var amount = parseFloat(document.getElementById('transferCashAmount').value) || 0;
  
  if (amount <= 0) {
    alert('❌ กรุณากรอกจำนวนเงิน');
    return;
  }
  
  var bal = _transferCashBalances[currency] || 0;
  if (amount > bal) {
    alert('❌ ยอดเงินไม่พอ! มี ' + formatNumber(bal) + ' ' + currency + ' แต่ต้องการย้าย ' + formatNumber(amount) + ' ' + currency);
    return;
  }
  
  if (!confirm('ยืนยันย้ายเงิน ' + formatNumber(amount) + ' ' + currency + ' เข้าร้าน?')) return;
  
  try {
    showLoading();
    var result = await callAppsScript('TRANSFER_CASH_TO_SHOP', {
      user: currentUser.nickname,
      currency: currency,
      amount: amount
    });
    
    if (result.success) {
      alert('✅ ย้ายเงินเข้าร้านสำเร็จ!');
      closeModal('transferCashModal');
    } else {
      alert('❌ Error: ' + result.message);
    }
    hideLoading();
  } catch (e) {
    hideLoading();
    alert('❌ Error: ' + e.message);
  }
}