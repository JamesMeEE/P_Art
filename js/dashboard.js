var _dashReportInterval = null;

async function loadDashboard() {
  try {
    showLoading();

    var today = new Date();
    var todayStr = formatDateOnly(today);
    document.getElementById('dashboardDate').textContent = todayStr;

    var dbData = await fetchSheetData('_database!A1:G31');

    var salesMoney = 0, salesTx = 0;
    var bbMoney = 0, bbTx = 0;
    var wdMoney = 0, wdTx = 0;
    var todayPL = 0;

    if (dbData.length >= 27) {
      var summaryRow = dbData[26];
      salesMoney = parseFloat(summaryRow[0]) || 0;
      salesTx = parseInt(summaryRow[1]) || 0;
      bbMoney = parseFloat(summaryRow[2]) || 0;
      bbTx = parseInt(summaryRow[3]) || 0;
      wdMoney = parseFloat(summaryRow[4]) || 0;
      wdTx = parseInt(summaryRow[5]) || 0;
      todayPL = parseFloat(summaryRow[6]) || 0;
    }

    var wacPerG = 0, wacPerBaht = 0;
    if (dbData.length >= 31) {
      var newGoldG = parseFloat(dbData[30][0]) || 0;
      var newValue = parseFloat(dbData[30][1]) || 0;
      var oldGoldG = parseFloat(dbData[30][2]) || 0;
      var oldValue = parseFloat(dbData[30][3]) || 0;
      var totalGoldG_wac = newGoldG + oldGoldG;
      var totalCost_wac = Math.round(oldValue / 1000) * 1000 + Math.round(newValue / 1000) * 1000;
      if (totalGoldG_wac > 0) {
        wacPerG = Math.round(totalCost_wac / totalGoldG_wac / 1000) * 1000;
        wacPerBaht = Math.round(wacPerG * 15 / 1000) * 1000;
      }
    }

    document.getElementById('dashSalesBox').innerHTML =
      '<h3 style="color:var(--gold-primary);margin-bottom:10px;">💰 SALES</h3>' +
      '<p style="font-size:18px;margin:5px 0;font-weight:bold;">' + formatNumber(Math.round(salesMoney)) + ' <span style="font-size:12px;">LAK</span></p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:3px 0;">Transactions: <b>' + salesTx + '</b></p>';

    document.getElementById('dashBuybackBox').innerHTML =
      '<h3 style="color:var(--gold-primary);margin-bottom:10px;">🔄 BUYBACK</h3>' +
      '<p style="font-size:18px;margin:5px 0;font-weight:bold;">' + formatNumber(Math.round(bbMoney)) + ' <span style="font-size:12px;">LAK</span></p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:3px 0;">Transactions: <b>' + bbTx + '</b></p>';

    document.getElementById('dashWithdrawBox').innerHTML =
      '<h3 style="color:var(--gold-primary);margin-bottom:10px;">📤 WITHDRAW</h3>' +
      '<p style="font-size:18px;margin:5px 0;font-weight:bold;">' + formatNumber(Math.round(wdMoney)) + ' <span style="font-size:12px;">LAK</span></p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:3px 0;">Transactions: <b>' + wdTx + '</b></p>';

    document.getElementById('dashPLBox').innerHTML =
      '<h3 style="color:var(--gold-primary);margin-bottom:10px;">📈 P/L</h3>' +
      '<p style="font-size:24px;font-weight:bold;color:' + (todayPL >= 0 ? '#4caf50' : '#f44336') + ';margin:10px 0;">' + formatNumber(Math.round(todayPL)) + ' <span style="font-size:12px;">LAK</span></p>';

    document.getElementById('dashWACBox').innerHTML =
      '<h3 style="color:var(--gold-primary);margin-bottom:10px;">⚖ WAC</h3>' +
      '<p style="font-size:12px;color:var(--text-secondary);margin:2px 0;">ราคา/g</p>' +
      '<p style="font-size:20px;font-weight:bold;margin:2px 0;">' + formatNumber(Math.round(wacPerG)) + ' <span style="font-size:12px;">LAK</span></p>' +
      '<p style="font-size:12px;color:var(--text-secondary);margin:8px 0 2px 0;">ราคา/บาท</p>' +
      '<p style="font-size:20px;font-weight:bold;margin:2px 0;">' + formatNumber(Math.round(wacPerBaht)) + ' <span style="font-size:12px;">LAK</span></p>';

    await refreshDashReport();

    var newStock = { pieces: 0, goldG: 0 };
    var oldStock = { pieces: 0, goldG: 0 };
    var cash = { LAK: 0, THB: 0, USD: 0 };
    var bank = { LAK: 0, THB: 0, USD: 0 };

    if (dbData.length >= 7) {
      var newGoldRow = dbData[6];
      ['G01','G02','G03','G04','G05','G06','G07'].forEach(function(p, i) {
        var qty = parseFloat(newGoldRow[i]) || 0;
        newStock.pieces += qty;
        newStock.goldG += qty * getGoldWeight(p);
      });
    }
    if (dbData.length >= 10) {
      var oldGoldRow = dbData[9];
      ['G01','G02','G03','G04','G05','G06','G07'].forEach(function(p, i) {
        var qty = parseFloat(oldGoldRow[i]) || 0;
        oldStock.pieces += qty;
        oldStock.goldG += qty * getGoldWeight(p);
      });
    }
    if (dbData.length >= 17) {
      cash.LAK = parseFloat(dbData[16][0]) || 0;
      cash.THB = parseFloat(dbData[16][1]) || 0;
      cash.USD = parseFloat(dbData[16][2]) || 0;
    }
    if (dbData.length >= 23) {
      bank.LAK = (parseFloat(dbData[19][0]) || 0) + (parseFloat(dbData[22][0]) || 0);
      bank.THB = (parseFloat(dbData[19][1]) || 0) + (parseFloat(dbData[22][1]) || 0);
      bank.USD = (parseFloat(dbData[19][2]) || 0) + (parseFloat(dbData[22][2]) || 0);
    }

    document.getElementById('dashNewStockBox').innerHTML =
      '<h3 style="color:#4caf50;margin-bottom:10px;">💎 NEW STOCK</h3>' +
      '<p style="font-size:20px;margin:5px 0;font-weight:bold;">' + newStock.pieces + ' <span style="font-size:13px;">ชิ้น</span></p>' +
      '<p style="font-size:16px;color:var(--text-secondary);">' + newStock.goldG.toFixed(2) + ' g</p>';

    document.getElementById('dashOldStockBox').innerHTML =
      '<h3 style="color:#ff9800;margin-bottom:10px;">🥇 OLD STOCK</h3>' +
      '<p style="font-size:20px;margin:5px 0;font-weight:bold;">' + oldStock.pieces + ' <span style="font-size:13px;">ชิ้น</span></p>' +
      '<p style="font-size:16px;color:var(--text-secondary);">' + oldStock.goldG.toFixed(2) + ' g</p>';

    document.getElementById('dashCashBox').innerHTML =
      '<h3 style="color:var(--gold-primary);margin-bottom:10px;">💵 CASH</h3>' +
      '<p style="font-size:16px;margin:3px 0;">' + formatNumber(cash.LAK) + ' <span style="font-size:12px;">LAK</span></p>' +
      '<p style="font-size:16px;margin:3px 0;">' + formatNumber(cash.THB) + ' <span style="font-size:12px;">THB</span></p>' +
      '<p style="font-size:16px;margin:3px 0;">' + formatNumber(cash.USD) + ' <span style="font-size:12px;">USD</span></p>';

    document.getElementById('dashBankBox').innerHTML =
      '<h3 style="color:var(--gold-primary);margin-bottom:10px;">🏦 BANK</h3>' +
      '<p style="font-size:16px;margin:3px 0;">' + formatNumber(bank.LAK) + ' <span style="font-size:12px;">LAK</span></p>' +
      '<p style="font-size:16px;margin:3px 0;">' + formatNumber(bank.THB) + ' <span style="font-size:12px;">THB</span></p>' +
      '<p style="font-size:16px;margin:3px 0;">' + formatNumber(bank.USD) + ' <span style="font-size:12px;">USD</span></p>';

    var totalGoldG = newStock.goldG + oldStock.goldG;
    var totalCashLAK = cash.LAK + bank.LAK;
    var totalCashTHB = cash.THB + bank.THB;
    var totalCashUSD = cash.USD + bank.USD;

    document.getElementById('dashTotalGoldBox').innerHTML =
      '<h3 style="color:var(--gold-primary);margin-bottom:10px;">TOTAL GOLD</h3>' +
      '<p style="font-size:28px;margin:5px 0;font-weight:bold;">' + totalGoldG.toFixed(2) + ' g</p>' +
      '<p style="font-size:13px;color:var(--text-secondary);">NEW: ' + newStock.goldG.toFixed(2) + ' g | OLD: ' + oldStock.goldG.toFixed(2) + ' g</p>';

    document.getElementById('dashTotalCashBox').innerHTML =
      '<h3 style="color:var(--gold-primary);margin-bottom:10px;">TOTAL CASH + BANK</h3>' +
      '<p style="font-size:18px;margin:5px 0;">' + formatNumber(totalCashLAK) + ' <span style="font-size:12px;">LAK</span></p>' +
      '<p style="font-size:18px;margin:3px 0;">' + formatNumber(totalCashTHB) + ' <span style="font-size:12px;">THB</span></p>' +
      '<p style="font-size:18px;margin:3px 0;">' + formatNumber(totalCashUSD) + ' <span style="font-size:12px;">USD</span></p>';

    startDashReportRefresh();
    hideLoading();
  } catch(error) {
    console.error('Error loading dashboard:', error);
    hideLoading();
  }
}

async function refreshDashReport() {
  try {
    var result = await callAppsScript('GET_LIVE_REPORT');
    if (!result.data) return;
    var net = result.data.netTotal || 0;

    document.getElementById('dashReportBox').innerHTML =
      '<h3 style="color:var(--gold-primary);margin-bottom:10px;">📋 Report</h3>' +
      '<p style="font-size:24px;font-weight:bold;margin:15px 0;">' + net.toFixed(2) + ' g</p>';
  } catch(e) {}
}

function startDashReportRefresh() {
  stopDashReportRefresh();
  _dashReportInterval = setInterval(refreshDashReport, 10000);
}

function stopDashReportRefresh() {
  if (_dashReportInterval) {
    clearInterval(_dashReportInterval);
    _dashReportInterval = null;
  }
}

function getGoldWeight(productId) {
  var weights = {
    'G01': 150,
    'G02': 75,
    'G03': 30,
    'G04': 15,
    'G05': 7.5,
    'G06': 3.75,
    'G07': 1
  };
  return weights[productId] || 0;
}