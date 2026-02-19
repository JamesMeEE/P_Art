var _dashReportInterval = null;

function calcItemsGrams(itemsJson) {
  try {
    var items = JSON.parse(itemsJson);
    var total = 0;
    items.forEach(function(item) { total += (getGoldWeight(item.productId) || 0) * item.qty; });
    return total;
  } catch(e) { return 0; }
}

function isTodayRow(dateValue) {
  var today = new Date();
  var todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  var todayEnd = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59, 999);
  var d;
  if (dateValue instanceof Date) {
    d = dateValue;
  } else if (typeof dateValue === 'string') {
    if (dateValue.includes('/')) {
      var parts = dateValue.split(' ');
      var dp = parts[0].split('/');
      d = new Date(parseInt(dp[2]), parseInt(dp[1]) - 1, parseInt(dp[0]));
    } else {
      d = new Date(dateValue);
    }
  }
  if (!d || isNaN(d.getTime())) return false;
  return d >= todayStart && d <= todayEnd;
}

async function loadDashboard() {
  try {
    showLoading();

    var today = new Date();
    var todayStr = formatDateOnly(today);
    document.getElementById('dashboardDate').textContent = todayStr;

    var dbPromise = fetchSheetData('_database!A1:G31');
    var txPromise = Promise.all([
      fetchSheetData('Sells!A:L'),
      fetchSheetData('Tradeins!A:N'),
      fetchSheetData('Exchanges!A:N'),
      fetchSheetData('Switches!A:N'),
      fetchSheetData('FreeExchanges!A:L'),
      fetchSheetData('Buybacks!A:L'),
      fetchSheetData('Withdraws!A:J')
    ]);

    var dbData = await dbPromise;
    var txResults = await txPromise;

    var sellRows = txResults[0].slice(1).filter(function(r) { return isTodayRow(r[9]) && (r[10] === 'COMPLETED' || r[10] === 'PAID'); });
    var tradeinRows = txResults[1].slice(1).filter(function(r) { return isTodayRow(r[11]) && (r[12] === 'COMPLETED' || r[12] === 'PAID'); });
    var exchangeRows = txResults[2].slice(1).filter(function(r) { return isTodayRow(r[11]) && (r[12] === 'COMPLETED' || r[12] === 'PAID'); });
    var switchRows = txResults[3].slice(1).filter(function(r) { return isTodayRow(r[11]) && (r[12] === 'COMPLETED' || r[12] === 'PAID'); });
    var freeExRows = txResults[4].slice(1).filter(function(r) { return isTodayRow(r[7]) && (r[8] === 'COMPLETED' || r[8] === 'PAID'); });
    var buybackRows = txResults[5].slice(1).filter(function(r) { return isTodayRow(r[9]) && (r[10] === 'COMPLETED' || r[10] === 'PAID'); });
    var withdrawRows = txResults[6].slice(1).filter(function(r) { return isTodayRow(r[6]) && (r[7] === 'COMPLETED' || r[7] === 'PAID'); });

    var sellMoney = 0; sellRows.forEach(function(r) { sellMoney += parseFloat(r[3]) || 0; });
    var tradeinMoney = 0; tradeinRows.forEach(function(r) { tradeinMoney += parseFloat(r[6]) || 0; });
    var exchangeMoney = 0; exchangeRows.forEach(function(r) { exchangeMoney += parseFloat(r[6]) || 0; });
    var switchMoney = 0; switchRows.forEach(function(r) { switchMoney += parseFloat(r[6]) || 0; });
    var freeExMoney = 0; freeExRows.forEach(function(r) { freeExMoney += parseFloat(r[5]) || 0; });

    var salesTotal = sellMoney + tradeinMoney + exchangeMoney + switchMoney + freeExMoney;
    var salesTotalTx = sellRows.length + tradeinRows.length + exchangeRows.length + switchRows.length + freeExRows.length;

    var salesOldGIn = 0;
    var salesNewGOut = 0;

    sellRows.forEach(function(r) { salesNewGOut += calcItemsGrams(r[2]); });
    tradeinRows.forEach(function(r) { salesOldGIn += calcItemsGrams(r[2]); salesNewGOut += calcItemsGrams(r[3]); });
    exchangeRows.forEach(function(r) { salesOldGIn += calcItemsGrams(r[2]); salesNewGOut += calcItemsGrams(r[3]); });
    switchRows.forEach(function(r) { salesOldGIn += calcItemsGrams(r[2]); salesNewGOut += calcItemsGrams(r[3]); });
    freeExRows.forEach(function(r) { salesOldGIn += calcItemsGrams(r[2]); salesNewGOut += calcItemsGrams(r[3]); });

    var bbMoney = 0; buybackRows.forEach(function(r) { bbMoney += parseFloat(r[6]) || parseFloat(r[3]) || 0; });
    var bbOldGIn = 0; buybackRows.forEach(function(r) { bbOldGIn += calcItemsGrams(r[2]); });

    var wdMoney = 0; withdrawRows.forEach(function(r) { wdMoney += parseFloat(r[4]) || 0; });
    var wdNewGOut = 0; withdrawRows.forEach(function(r) { wdNewGOut += calcItemsGrams(r[2]); });

    var totalOldGIn = salesOldGIn + bbOldGIn;
    var totalNewGOut = salesNewGOut + wdNewGOut;
    var netSellBaht = (totalNewGOut - totalOldGIn) / 15;

    var todayPL = 0;
    if (dbData.length >= 27) {
      todayPL = parseFloat(dbData[26][6]) || 0;
    }

    var wacPerG = 0, wacPerBaht = 0;
    if (dbData.length >= 31) {
      var newGoldG = parseFloat(dbData[30][0]) || 0;
      var newValue = parseFloat(dbData[30][1]) || 0;
      var oldGoldG = parseFloat(dbData[30][2]) || 0;
      var oldValue = parseFloat(dbData[30][3]) || 0;
      var totalGoldG_wac = newGoldG + oldGoldG;
      var totalCost_wac = oldValue + newValue;
      if (totalGoldG_wac > 0) {
        wacPerG = totalCost_wac / totalGoldG_wac;
        wacPerBaht = wacPerG * 15;
      }
    }

    document.getElementById('dashSalesBox').innerHTML =
      '<h3 style="color:var(--gold-primary);margin-bottom:8px;">💰 SALES</h3>' +
      '<p style="font-size:18px;margin:3px 0;font-weight:bold;">' + formatNumber(Math.round(salesTotal)) + ' <span style="font-size:12px;">LAK</span></p>' +
      '<p style="font-size:11px;color:var(--text-secondary);margin:2px 0;">Tx: <b>' + salesTotalTx + '</b></p>' +
      '<div style="border-top:1px solid var(--border-color);margin:6px 0;padding-top:6px;font-size:11px;color:var(--text-secondary);line-height:1.6;">' +
      'Sell: ' + formatNumber(Math.round(sellMoney)) + ' (' + sellRows.length + ')<br>' +
      'Trade-in: ' + formatNumber(Math.round(tradeinMoney)) + ' (' + tradeinRows.length + ')<br>' +
      'Exchange: ' + formatNumber(Math.round(exchangeMoney)) + ' (' + exchangeRows.length + ')<br>' +
      'Switch: ' + formatNumber(Math.round(switchMoney)) + ' (' + switchRows.length + ')<br>' +
      'Free Ex: ' + formatNumber(Math.round(freeExMoney)) + ' (' + freeExRows.length + ')' +
      '</div>' +
      '<div style="border-top:1px solid var(--border-color);margin:6px 0;padding-top:6px;font-size:12px;">' +
      '<span style="color:#ff9800;">◀ Old In: ' + salesOldGIn.toFixed(2) + ' g</span><br>' +
      '<span style="color:#4caf50;">▶ New Out: ' + salesNewGOut.toFixed(2) + ' g</span>' +
      '</div>';

    document.getElementById('dashBuybackBox').innerHTML =
      '<h3 style="color:var(--gold-primary);margin-bottom:8px;">🔄 BUYBACK</h3>' +
      '<p style="font-size:18px;margin:3px 0;font-weight:bold;">' + formatNumber(Math.round(bbMoney)) + ' <span style="font-size:12px;">LAK</span></p>' +
      '<p style="font-size:11px;color:var(--text-secondary);margin:2px 0;">Tx: <b>' + buybackRows.length + '</b></p>' +
      '<div style="border-top:1px solid var(--border-color);margin:6px 0;padding-top:6px;font-size:12px;">' +
      '<span style="color:#ff9800;">◀ Old In: ' + bbOldGIn.toFixed(2) + ' g</span>' +
      '</div>';

    document.getElementById('dashWithdrawBox').innerHTML =
      '<h3 style="color:var(--gold-primary);margin-bottom:8px;">📤 WITHDRAW</h3>' +
      '<p style="font-size:18px;margin:3px 0;font-weight:bold;">' + formatNumber(Math.round(wdMoney)) + ' <span style="font-size:12px;">LAK</span></p>' +
      '<p style="font-size:11px;color:var(--text-secondary);margin:2px 0;">Tx: <b>' + withdrawRows.length + '</b></p>' +
      '<div style="border-top:1px solid var(--border-color);margin:6px 0;padding-top:6px;font-size:12px;">' +
      '<span style="color:#4caf50;">▶ New Out: ' + wdNewGOut.toFixed(2) + ' g</span>' +
      '</div>';

    var netColor = netSellBaht >= 0 ? '#4caf50' : '#f44336';
    document.getElementById('dashNetSellBox').innerHTML =
      '<h3 style="color:var(--gold-primary);margin-bottom:8px;">⚖ NET SELL</h3>' +
      '<p style="font-size:24px;margin:8px 0;font-weight:bold;color:' + netColor + ';">' + netSellBaht.toFixed(2) + ' <span style="font-size:13px;">บาท</span></p>' +
      '<div style="border-top:1px solid var(--border-color);margin:6px 0;padding-top:6px;font-size:11px;color:var(--text-secondary);line-height:1.6;">' +
      'New Out ทั้งหมด: ' + totalNewGOut.toFixed(2) + ' g<br>' +
      'Old In ทั้งหมด: ' + totalOldGIn.toFixed(2) + ' g<br>' +
      'Net: ' + (totalNewGOut - totalOldGIn).toFixed(2) + ' g ÷ 15' +
      '</div>';

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