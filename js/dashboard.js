var _dashReportInterval = null;

async function loadDashboard() {
  try {
    showLoading();

    var today = new Date();
    var todayStr = formatDateOnly(today);
    document.getElementById('dashboardDate').textContent = todayStr;

    var todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    var todayEnd = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);

    var results = await Promise.all([
      fetchSheetData('Sells!A:L'),
      fetchSheetData('Tradeins!A:N'),
      fetchSheetData('Buybacks!A:L'),
      fetchSheetData('Exchanges!A:N'),
      fetchSheetData('Withdraws!A:J'),
      fetchSheetData('Switches!A:N'),
      fetchSheetData('FreeExchanges!A:J'),
      fetchSheetData('_database!A1:G20'),
      callAppsScript('GET_WAC')
    ]);

    var sellData = results[0], tradeinData = results[1], buybackData = results[2];
    var exchangeData = results[3], withdrawData = results[4], switchData = results[5];
    var freeExData = results[6], dbData = results[7], wacResult = results[8];

    var sales = { moneyLAK: 0, newGoldOutG: 0, oldGoldInG: 0, txCount: 0 };
    var bb = { moneyLAK: 0, oldGoldInG: 0, txCount: 0 };
    var wd = { moneyLAK: 0, newGoldOutG: 0, txCount: 0 };

    sellData.slice(1).forEach(function(row) {
      var date = parseSheetDate(row[9]);
      if (date && date >= todayStart && date <= todayEnd && row[10] === 'COMPLETED') {
        sales.moneyLAK += parseFloat(row[3]) || 0;
        sales.txCount++;
        try {
          JSON.parse(row[2]).forEach(function(item) {
            sales.newGoldOutG += getGoldWeight(item.productId) * item.qty;
          });
        } catch(e) {}
      }
    });

    tradeinData.slice(1).forEach(function(row) {
      var date = parseSheetDate(row[11]);
      if (date && date >= todayStart && date <= todayEnd && row[12] === 'COMPLETED') {
        var diff = parseFloat(row[4]) || 0;
        var fee = parseFloat(row[5]) || 0;
        var premium = parseFloat(row[6]) || 0;
        sales.moneyLAK += (diff + fee + premium);
        sales.txCount++;
        try {
          JSON.parse(row[2]).forEach(function(item) { sales.oldGoldInG += getGoldWeight(item.productId) * item.qty; });
          JSON.parse(row[3]).forEach(function(item) { sales.newGoldOutG += getGoldWeight(item.productId) * item.qty; });
        } catch(e) {}
      }
    });

    exchangeData.slice(1).forEach(function(row) {
      var date = parseSheetDate(row[11]);
      if (date && date >= todayStart && date <= todayEnd && row[12] === 'COMPLETED') {
        sales.moneyLAK += (parseFloat(row[4]) || 0) + (parseFloat(row[5]) || 0);
        sales.txCount++;
        try {
          JSON.parse(row[2]).forEach(function(item) { sales.oldGoldInG += getGoldWeight(item.productId) * item.qty; });
          JSON.parse(row[3]).forEach(function(item) { sales.newGoldOutG += getGoldWeight(item.productId) * item.qty; });
        } catch(e) {}
      }
    });

    switchData.slice(1).forEach(function(row) {
      var date = parseSheetDate(row[11]);
      if (date && date >= todayStart && date <= todayEnd && row[12] === 'COMPLETED') {
        sales.moneyLAK += parseFloat(row[6]) || 0;
        sales.txCount++;
        try {
          JSON.parse(row[2]).forEach(function(item) { sales.oldGoldInG += getGoldWeight(item.productId) * item.qty; });
          JSON.parse(row[3]).forEach(function(item) { sales.newGoldOutG += getGoldWeight(item.productId) * item.qty; });
        } catch(e) {}
      }
    });

    freeExData.slice(1).forEach(function(row) {
      var date = parseSheetDate(row[7]);
      if (date && date >= todayStart && date <= todayEnd && row[8] === 'COMPLETED') {
        sales.moneyLAK += parseFloat(row[4]) || 0;
        sales.txCount++;
        try {
          JSON.parse(row[2]).forEach(function(item) { sales.oldGoldInG += getGoldWeight(item.productId) * item.qty; });
          JSON.parse(row[3]).forEach(function(item) { sales.newGoldOutG += getGoldWeight(item.productId) * item.qty; });
        } catch(e) {}
      }
    });

    buybackData.slice(1).forEach(function(row) {
      var date = parseSheetDate(row[9]);
      if (date && date >= todayStart && date <= todayEnd && (row[10] === 'COMPLETED' || row[10] === 'PAID')) {
        bb.moneyLAK += parseFloat(row[6]) || 0;
        bb.txCount++;
        try {
          JSON.parse(row[2]).forEach(function(item) { bb.oldGoldInG += getGoldWeight(item.productId) * item.qty; });
        } catch(e) {}
      }
    });

    withdrawData.slice(1).forEach(function(row) {
      var date = parseSheetDate(row[6]);
      if (date && date >= todayStart && date <= todayEnd && row[7] === 'COMPLETED') {
        wd.moneyLAK += parseFloat(row[4]) || 0;
        wd.txCount++;
        try {
          JSON.parse(row[2]).forEach(function(item) { wd.newGoldOutG += getGoldWeight(item.productId) * item.qty; });
        } catch(e) {}
      }
    });

    document.getElementById('dashSalesBox').innerHTML =
      '<h3 style="color:var(--gold-primary);margin-bottom:10px;">📊 SALES</h3>' +
      '<p style="font-size:11px;color:var(--text-secondary);margin-bottom:8px;">Sell · Trade-in · Exchange · Switch · Free Ex</p>' +
      '<p style="font-size:18px;margin:5px 0;font-weight:bold;">' + formatNumber(sales.moneyLAK) + ' <span style="font-size:12px;">LAK</span></p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:3px 0;">New Gold Out: <b style="color:#f44336;">' + sales.newGoldOutG.toFixed(2) + ' g</b></p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:3px 0;">Old Gold In: <b style="color:#4caf50;">' + sales.oldGoldInG.toFixed(2) + ' g</b></p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:3px 0;">Transactions: <b>' + sales.txCount + '</b></p>';

    document.getElementById('dashBuybackBox').innerHTML =
      '<h3 style="color:var(--gold-primary);margin-bottom:10px;">💰 BUYBACK</h3>' +
      '<p style="font-size:18px;margin:5px 0;font-weight:bold;">' + formatNumber(bb.moneyLAK) + ' <span style="font-size:12px;">LAK</span></p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:3px 0;">Old Gold In: <b style="color:#4caf50;">' + bb.oldGoldInG.toFixed(2) + ' g</b></p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:3px 0;">Transactions: <b>' + bb.txCount + '</b></p>';

    document.getElementById('dashWithdrawBox').innerHTML =
      '<h3 style="color:var(--gold-primary);margin-bottom:10px;">📤 WITHDRAW</h3>' +
      '<p style="font-size:18px;margin:5px 0;font-weight:bold;">' + formatNumber(wd.moneyLAK) + ' <span style="font-size:12px;">LAK</span></p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:3px 0;">New Gold Out: <b style="color:#f44336;">' + wd.newGoldOutG.toFixed(2) + ' g</b></p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:3px 0;">Transactions: <b>' + wd.txCount + '</b></p>';

    document.getElementById('dashPLBox').innerHTML =
      '<h3 style="color:var(--gold-primary);margin-bottom:10px;">📈 P/L</h3>' +
      '<p style="font-size:14px;color:var(--text-secondary);text-align:center;padding:20px 0;">Coming soon...</p>';

    var wacPerG = wacResult.data ? wacResult.data.wacPerG || 0 : 0;
    var wacPerBaht = wacResult.data ? wacResult.data.wacPerBaht || 0 : 0;
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

    if (dbData.length >= 14) {
      cash.LAK = parseFloat(dbData[13][0]) || 0;
      cash.THB = parseFloat(dbData[13][1]) || 0;
      cash.USD = parseFloat(dbData[13][2]) || 0;
    }

    if (dbData.length >= 20) {
      bank.LAK = (parseFloat(dbData[16][0]) || 0) + (parseFloat(dbData[19][0]) || 0);
      bank.THB = (parseFloat(dbData[16][1]) || 0) + (parseFloat(dbData[19][1]) || 0);
      bank.USD = (parseFloat(dbData[16][2]) || 0) + (parseFloat(dbData[19][2]) || 0);
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
    var carry = result.data.carryForward || 0;
    var net = result.data.netTotal || 0;
    var diff = net - carry;
    var diffColor = diff >= 0 ? '#4caf50' : '#f44336';
    var diffSign = diff >= 0 ? '+' : '';

    document.getElementById('dashReportBox').innerHTML =
      '<h3 style="color:var(--gold-primary);margin-bottom:10px;">📋 REPORT</h3>' +
      '<p style="font-size:12px;color:var(--text-secondary);margin:2px 0;">ยอดยกมา (g)</p>' +
      '<p style="font-size:16px;font-weight:bold;margin:2px 0;">' + carry.toFixed(2) + '</p>' +
      '<p style="font-size:12px;color:var(--text-secondary);margin:6px 0 2px 0;">ยอดสุทธิ (g)</p>' +
      '<p style="font-size:16px;font-weight:bold;margin:2px 0;">' + net.toFixed(2) + '</p>' +
      '<p style="font-size:12px;color:var(--text-secondary);margin:6px 0 2px 0;">ส่วนต่าง</p>' +
      '<p style="font-size:18px;font-weight:bold;color:' + diffColor + ';margin:2px 0;">' + diffSign + diff.toFixed(2) + ' g</p>';
  } catch(e) {}
}

function startDashReportRefresh() {
  stopDashReportRefresh();
  _dashReportInterval = setInterval(refreshDashReport, 5000);
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