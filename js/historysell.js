var historySellDateFrom = '';
var historySellDateTo = '';

async function loadHistorySell() {
  try {
    showLoading();
    var today = getTodayDateString();
    if (!historySellDateFrom && !historySellDateTo) {
      historySellDateFrom = today;
      historySellDateTo = today;
    }
    document.getElementById('historySellDateFrom').value = historySellDateFrom;
    document.getElementById('historySellDateTo').value = historySellDateTo;

    var fromEl = document.getElementById('historySellDateFrom');
    var toEl = document.getElementById('historySellDateTo');
    fromEl.onchange = function() { historySellDateFrom = this.value; loadHistorySell(); };
    toEl.onchange = function() { historySellDateTo = this.value; loadHistorySell(); };

    var results = await Promise.all([
      fetchSheetData('Sells!A:L'),
      fetchSheetData('Tradeins!A:N'),
      fetchSheetData('Exchanges!A:N'),
      fetchSheetData('Switches!A:N'),
      fetchSheetData('FreeExchanges!A:L'),
      fetchSheetData('Withdraws!A:J')
    ]);

    var all = [];

    results[0].slice(1).forEach(function(r) {
      all.push({ type: 'SELL', id: r[0], phone: r[1], items: formatItemsForTable(r[2]), total: parseFloat(r[3]) || 0, status: r[10] || '', sale: r[11] || '', date: r[9], raw: r });
    });
    results[1].slice(1).forEach(function(r) {
      all.push({ type: 'TRADE-IN', id: r[0], phone: r[1], items: formatItemsForTable(r[2]) + ' → ' + formatItemsForTable(r[3]), total: parseFloat(r[6]) || 0, status: r[12] || '', sale: r[13] || '', date: r[11], raw: r });
    });
    results[2].slice(1).forEach(function(r) {
      all.push({ type: 'EXCHANGE', id: r[0], phone: r[1], items: formatItemsForTable(r[2]) + ' → ' + formatItemsForTable(r[3]), total: parseFloat(r[6]) || 0, status: r[12] || '', sale: r[13] || '', date: r[11], raw: r });
    });
    results[3].slice(1).forEach(function(r) {
      all.push({ type: 'SWITCH', id: r[0], phone: r[1], items: formatItemsForTable(r[2]) + ' → ' + formatItemsForTable(r[3]), total: parseFloat(r[6]) || 0, status: r[12] || '', sale: r[13] || '', date: r[11], raw: r });
    });
    results[4].slice(1).forEach(function(r) {
      all.push({ type: 'FREE EX', id: r[0], phone: r[1], items: formatItemsForTable(r[2]) + ' → ' + formatItemsForTable(r[3]), total: parseFloat(r[5]) || 0, status: r[8] || '', sale: r[9] || '', date: r[7], raw: r });
    });
    results[5].slice(1).forEach(function(r) {
      all.push({ type: 'WITHDRAW', id: r[0], phone: r[1], items: formatItemsForTable(r[2]), total: parseFloat(r[4]) || 0, status: r[7] || '', sale: r[8] || '', date: r[6], raw: r });
    });

    all = filterHistoryByDate(all, historySellDateFrom, historySellDateTo);
    all.sort(function(a, b) { return parseHistoryDate(b.date) - parseHistoryDate(a.date); });

    var tbody = document.getElementById('historySellTable');
    if (all.length === 0) {
      tbody.innerHTML = '<tr><td colspan="9" style="text-align:center;padding:40px;">No records</td></tr>';
    } else {
      tbody.innerHTML = all.map(function(r) {
        var typeColors = { 'SELL': '#4caf50', 'TRADE-IN': '#2196f3', 'EXCHANGE': '#9c27b0', 'SWITCH': '#ff9800', 'FREE EX': '#00bcd4', 'WITHDRAW': '#f44336' };
        var color = typeColors[r.type] || '#888';
        var actions = '-';
        if (r.status === 'COMPLETED' || r.status === 'PAID') {
          var detail = encodeURIComponent(JSON.stringify([['Type', r.type], ['Transaction ID', r.id], ['Phone', r.phone], ['Items', r.items], ['Total', formatNumber(r.total) + ' LAK'], ['Date', formatDateTime(r.date)], ['Status', r.status], ['Sale', r.sale]]));
          actions = '<button class="btn-action" onclick="viewTransactionDetail(\'' + r.type + '\',\'' + detail + '\')" style="background:#555;">👁 View</button>';
        } else if (r.status === 'PENDING') {
          actions = '<span style="color:var(--text-secondary);">Pending Review</span>';
        } else if (r.status === 'READY') {
          actions = '<span style="color:var(--text-secondary);">Ready</span>';
        }
        return '<tr>' +
          '<td><span style="background:' + color + ';color:#fff;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:bold;">' + r.type + '</span></td>' +
          '<td>' + r.id + '</td>' +
          '<td>' + r.phone + '</td>' +
          '<td>' + r.items + '</td>' +
          '<td>' + formatNumber(r.total) + '</td>' +
          '<td><span class="status-badge status-' + (r.status || '').toLowerCase() + '">' + r.status + '</span></td>' +
          '<td>' + r.sale + '</td>' +
          '<td style="font-size:12px;">' + formatDateTime(r.date) + '</td>' +
          '<td>' + actions + '</td>' +
          '</tr>';
      }).join('');
    }

    hideLoading();
  } catch(e) {
    console.error('Error loading history sell:', e);
    hideLoading();
  }
}

function filterHistoryByDate(data, from, to) {
  var fromDate = null, toDate = null;
  if (from) { var p = from.split('-'); fromDate = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]), 0, 0, 0, 0); }
  if (to) { var p = to.split('-'); toDate = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]), 23, 59, 59, 999); }
  return data.filter(function(r) {
    var d = parseHistoryDate(r.date);
    if (!d || isNaN(d)) return false;
    if (fromDate && d < fromDate) return false;
    if (toDate && d > toDate) return false;
    return true;
  });
}

function parseHistoryDate(dateValue) {
  if (!dateValue) return 0;
  if (dateValue instanceof Date) return dateValue.getTime();
  if (typeof dateValue === 'string') {
    if (dateValue.includes('/')) {
      var parts = dateValue.split(' ');
      var dp = parts[0].split('/');
      var day = parseInt(dp[0]), month = parseInt(dp[1]) - 1, year = parseInt(dp[2]);
      var h = 0, m = 0;
      if (parts[1]) { var tp = parts[1].split(':'); h = parseInt(tp[0]) || 0; m = parseInt(tp[1]) || 0; }
      return new Date(year, month, day, h, m).getTime();
    }
    return new Date(dateValue).getTime();
  }
  return 0;
}

function resetHistorySellDateFilter() {
  historySellDateFrom = getTodayDateString();
  historySellDateTo = getTodayDateString();
  loadHistorySell();
}
