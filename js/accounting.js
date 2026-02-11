async function loadAccounting() {
  try {
    showLoading();

    var today = new Date();
    today.setHours(0, 0, 0, 0);

    var yesterday = new Date(today);
    yesterday.setDate(yesterday.getDate() - 1);

    var accountingData = await fetchSheetData('Accounting!A:M');

    var existingDates = new Set();
    if (accountingData.length > 1) {
      accountingData.slice(1).forEach(function(row) {
        if (row[0]) {
          var d = parseSheetDate(row[0]);
          if (d) {
            existingDates.add(d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0'));
          }
        }
      });
    }

    var yesterdayKey = yesterday.getFullYear() + '-' + String(yesterday.getMonth()+1).padStart(2,'0') + '-' + String(yesterday.getDate()).padStart(2,'0');

    if (!existingDates.has(yesterdayKey)) {
      if (accountingData.length > 1) {
        var lastRecord = accountingData[accountingData.length - 1];
        var lastDateObj = parseSheetDate(lastRecord[0]);

        if (lastDateObj) {
          lastDateObj.setHours(0, 0, 0, 0);
          var diffTime = yesterday.getTime() - lastDateObj.getTime();
          var diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));

          if (diffDays > 0) {
            for (var i = 1; i <= diffDays; i++) {
              var targetDate = new Date(lastDateObj);
              targetDate.setDate(targetDate.getDate() + i);
              var targetKey = targetDate.getFullYear() + '-' + String(targetDate.getMonth()+1).padStart(2,'0') + '-' + String(targetDate.getDate()).padStart(2,'0');
              if (!existingDates.has(targetKey)) {
                await saveAccountingForDate(targetDate);
                existingDates.add(targetKey);
              }
            }
          }
        }
      } else {
        await saveAccountingForDate(yesterday);
      }
    }

    await loadTodayStats();
    await loadAccountingHistory();

    hideLoading();
  } catch (error) {
    console.error('Error loading accounting:', error);
    hideLoading();
  }
}

async function saveAccountingForDate(targetDate) {
  try {
    var sells = await fetchSheetData('Sells!A:L');
    var tradeins = await fetchSheetData('Tradeins!A:N');
    var exchanges = await fetchSheetData('Exchanges!A:N');
    var buybacks = await fetchSheetData('Buybacks!A:L');
    var withdraws = await fetchSheetData('Withdraws!A:J');

    var dayStart = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate());
    var dayEnd = new Date(targetDate.getFullYear(), targetDate.getMonth(), targetDate.getDate(), 23, 59, 59);

    var stats = {
      sellMoney: 0, sellGold: 0,
      tradeinMoney: 0, tradeinGold: 0,
      exchangeMoney: 0, exchangeGold: 0,
      buybackMoney: 0, buybackGold: 0,
      withdrawMoney: 0, withdrawGold: 0,
      incompleteMoney: 0, incompleteGold: 0
    };

    sells.slice(1).forEach(function(row) {
      var date = parseSheetDate(row[9]);
      if (date && date >= dayStart && date <= dayEnd) {
        if (row[10] === 'COMPLETED') {
          stats.sellMoney += parseFloat(row[3]) || 0;
          stats.sellGold += calcGold(row[2]);
        } else {
          stats.incompleteMoney += parseFloat(row[3]) || 0;
          stats.incompleteGold += calcGold(row[2]);
        }
      }
    });

    tradeins.slice(1).forEach(function(row) {
      var date = parseSheetDate(row[11]);
      if (date && date >= dayStart && date <= dayEnd) {
        var totalLAK = (parseFloat(row[4]) || 0) + (parseFloat(row[5]) || 0) + (parseFloat(row[6]) || 0);
        var goldG = calcGold(row[3]);
        if (row[12] === 'COMPLETED') {
          stats.tradeinMoney += totalLAK;
          stats.tradeinGold += goldG;
        } else {
          stats.incompleteMoney += totalLAK;
          stats.incompleteGold += goldG;
        }
      }
    });

    exchanges.slice(1).forEach(function(row) {
      var date = parseSheetDate(row[11]);
      if (date && date >= dayStart && date <= dayEnd) {
        var totalLAK = parseFloat(row[6]) || 0;
        var goldG = calcGold(row[3]);
        if (row[12] === 'COMPLETED') {
          stats.exchangeMoney += totalLAK;
          stats.exchangeGold += goldG;
        } else {
          stats.incompleteMoney += totalLAK;
          stats.incompleteGold += goldG;
        }
      }
    });

    buybacks.slice(1).forEach(function(row) {
      var date = parseSheetDate(row[9]);
      if (date && date >= dayStart && date <= dayEnd) {
        var totalLAK = parseFloat(row[6]) || 0;
        var goldG = calcGold(row[2]);
        if (row[10] === 'COMPLETED' || row[10] === 'PAID') {
          stats.buybackMoney += totalLAK;
          stats.buybackGold += goldG;
        } else {
          stats.incompleteMoney += totalLAK;
          stats.incompleteGold += goldG;
        }
      }
    });

    withdraws.slice(1).forEach(function(row) {
      var date = parseSheetDate(row[6]);
      if (date && date >= dayStart && date <= dayEnd) {
        var totalLAK = parseFloat(row[4]) || 0;
        var goldG = calcGold(row[2]);
        if (row[7] === 'COMPLETED') {
          stats.withdrawMoney += totalLAK;
          stats.withdrawGold += goldG;
        } else {
          stats.incompleteMoney += totalLAK;
          stats.incompleteGold += goldG;
        }
      }
    });

    var dateStr = targetDate.toISOString().split('T')[0];

    await callAppsScript('SAVE_ACCOUNTING', {
      date: dateStr,
      sellMoney: stats.sellMoney,
      sellGold: stats.sellGold,
      tradeinMoney: stats.tradeinMoney,
      tradeinGold: stats.tradeinGold,
      exchangeMoney: stats.exchangeMoney,
      exchangeGold: stats.exchangeGold,
      buybackMoney: stats.buybackMoney,
      buybackGold: stats.buybackGold,
      withdrawMoney: stats.withdrawMoney,
      withdrawGold: stats.withdrawGold,
      incompleteMoney: stats.incompleteMoney,
      incompleteGold: stats.incompleteGold
    });
  } catch (error) {
    console.error('Error saving accounting for date:', error);
  }
}

async function loadTodayStats() {
  try {
    var results = await Promise.all([
      fetchSheetData('Sells!A:L'),
      fetchSheetData('Tradeins!A:N'),
      fetchSheetData('Exchanges!A:N'),
      fetchSheetData('Buybacks!A:L'),
      fetchSheetData('Withdraws!A:J'),
      fetchSheetData('Switches!A:N'),
      fetchSheetData('FreeExchanges!A:J'),
      fetchSheetData('CashBank!A:I'),
      callAppsScript('GET_WAC'),
      fetchSheetData('Diff!A:I')
    ]);

    var sells = results[0], tradeins = results[1], exchanges = results[2];
    var buybacks = results[3], withdraws = results[4], switchData = results[5];
    var freeExData = results[6], cashbankData = results[7], wacResult = results[8];
    var diffData = results[9];

    var wacPerG = wacResult.data ? wacResult.data.wacPerG || 0 : 0;

    var today = new Date();
    var todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    var todayEnd = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);

    var sell = { newGoldG: 0, txCount: 0 };
    var tradein = { newGoldG: 0, oldGoldG: 0, moneyNoP: 0, txCount: 0 };
    var exchange = { newGoldG: 0, oldGoldG: 0, txCount: 0 };
    var wd = { newGoldG: 0, txCount: 0 };
    var freeEx = { newGoldG: 0, oldGoldG: 0, txCount: 0 };
    var sw = { newGoldG: 0, oldGoldG: 0, txCount: 0 };
    var bb = { oldGoldG: 0, txCount: 0 };
    var otherExpenseLAK = 0;
    var incomplete = { money: 0, gold: 0 };

    sells.slice(1).forEach(function(row) {
      var date = parseSheetDate(row[9]);
      if (date && date >= todayStart && date <= todayEnd) {
        if (row[10] === 'COMPLETED') {
          sell.txCount++;
          try {
            JSON.parse(row[2]).forEach(function(item) {
              sell.newGoldG += getGoldWeight(item.productId) * item.qty;
            });
          } catch(e) {}
        } else if (row[10] !== 'REJECTED') {
          incomplete.money += parseFloat(row[3]) || 0;
          incomplete.gold += calcGold(row[2]);
        }
      }
    });

    tradeins.slice(1).forEach(function(row) {
      var date = parseSheetDate(row[11]);
      if (date && date >= todayStart && date <= todayEnd) {
        if (row[12] === 'COMPLETED') {
          tradein.txCount++;
          var diff = parseFloat(row[4]) || 0;
          var fee = parseFloat(row[5]) || 0;
          tradein.moneyNoP += (diff + fee);
          try {
            JSON.parse(row[2]).forEach(function(item) { tradein.oldGoldG += getGoldWeight(item.productId) * item.qty; });
            JSON.parse(row[3]).forEach(function(item) { tradein.newGoldG += getGoldWeight(item.productId) * item.qty; });
          } catch(e) {}
        } else if (row[12] !== 'REJECTED') {
          var totalLAK = (parseFloat(row[4]) || 0) + (parseFloat(row[5]) || 0) + (parseFloat(row[6]) || 0);
          incomplete.money += totalLAK;
          incomplete.gold += calcGold(row[3]);
        }
      }
    });

    exchanges.slice(1).forEach(function(row) {
      var date = parseSheetDate(row[11]);
      if (date && date >= todayStart && date <= todayEnd) {
        if (row[12] === 'COMPLETED') {
          exchange.txCount++;
          try {
            JSON.parse(row[2]).forEach(function(item) { exchange.oldGoldG += getGoldWeight(item.productId) * item.qty; });
            JSON.parse(row[3]).forEach(function(item) { exchange.newGoldG += getGoldWeight(item.productId) * item.qty; });
          } catch(e) {}
        } else if (row[12] !== 'REJECTED') {
          incomplete.money += parseFloat(row[6]) || 0;
          incomplete.gold += calcGold(row[3]);
        }
      }
    });

    switchData.slice(1).forEach(function(row) {
      var date = parseSheetDate(row[11]);
      if (date && date >= todayStart && date <= todayEnd) {
        if (row[12] === 'COMPLETED') {
          sw.txCount++;
          try {
            JSON.parse(row[2]).forEach(function(item) { sw.oldGoldG += getGoldWeight(item.productId) * item.qty; });
            JSON.parse(row[3]).forEach(function(item) { sw.newGoldG += getGoldWeight(item.productId) * item.qty; });
          } catch(e) {}
        } else if (row[12] !== 'REJECTED') {
          incomplete.money += parseFloat(row[6]) || 0;
          incomplete.gold += calcGold(row[3]);
        }
      }
    });

    freeExData.slice(1).forEach(function(row) {
      var date = parseSheetDate(row[7]);
      if (date && date >= todayStart && date <= todayEnd) {
        if (row[8] === 'COMPLETED') {
          freeEx.txCount++;
          try {
            JSON.parse(row[2]).forEach(function(item) { freeEx.oldGoldG += getGoldWeight(item.productId) * item.qty; });
            JSON.parse(row[3]).forEach(function(item) { freeEx.newGoldG += getGoldWeight(item.productId) * item.qty; });
          } catch(e) {}
        } else if (row[8] !== 'REJECTED') {
          incomplete.money += parseFloat(row[5]) || 0;
          incomplete.gold += calcGold(row[3]);
        }
      }
    });

    buybacks.slice(1).forEach(function(row) {
      var date = parseSheetDate(row[9]);
      if (date && date >= todayStart && date <= todayEnd) {
        if (row[10] === 'COMPLETED' || row[10] === 'PAID') {
          bb.txCount++;
          try {
            JSON.parse(row[2]).forEach(function(item) { bb.oldGoldG += getGoldWeight(item.productId) * item.qty; });
          } catch(e) {}
        } else if (row[10] !== 'REJECTED') {
          incomplete.money += parseFloat(row[6]) || 0;
          incomplete.gold += calcGold(row[2]);
        }
      }
    });

    withdraws.slice(1).forEach(function(row) {
      var date = parseSheetDate(row[6]);
      if (date && date >= todayStart && date <= todayEnd) {
        if (row[7] === 'COMPLETED') {
          wd.txCount++;
          try {
            JSON.parse(row[2]).forEach(function(item) { wd.newGoldG += getGoldWeight(item.productId) * item.qty; });
          } catch(e) {}
        } else if (row[7] !== 'REJECTED') {
          incomplete.money += parseFloat(row[4]) || 0;
          incomplete.gold += calcGold(row[2]);
        }
      }
    });

    cashbankData.slice(1).forEach(function(row) {
      var date = parseSheetDate(row[7]);
      if (date && date >= todayStart && date <= todayEnd) {
        if (row[1] === 'OTHER_EXPENSE') {
          var amt = parseFloat(row[2]) || 0;
          var cur = row[3];
          if (cur === 'THB') amt = amt * (currentExchangeRates?.THB_Sell || 0);
          else if (cur === 'USD') amt = amt * (currentExchangeRates?.USD_Sell || 0);
          otherExpenseLAK += Math.abs(amt);
        }
      }
    });

    var sellCostLAK = wacPerG * sell.newGoldG;
    var tradeinDiffBaht = (tradein.newGoldG - tradein.oldGoldG) / 15;
    var tradeinAvg = tradeinDiffBaht !== 0 ? tradein.moneyNoP / tradeinDiffBaht : 0;
    var bbCostLAK = wacPerG * bb.oldGoldG;

    var gpDiff = 0;
    if (diffData && diffData.length > 1) {
      diffData.slice(1).forEach(function(row) {
        var date = parseSheetDate(row[8]);
        if (date && date >= todayStart && date <= todayEnd) {
          gpDiff += parseFloat(row[7]) || 0;
        }
      });
    }
    var pl = gpDiff - otherExpenseLAK;

    document.getElementById('accountingStats').innerHTML =
      '<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:15px;margin-bottom:15px;">' +

      '<div class="stat-card">' +
      '<h3 style="color:var(--gold-primary);margin-bottom:8px;">SELL</h3>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:2px 0;">ต้นทุน (WAC)</p>' +
      '<p style="font-size:16px;font-weight:bold;margin:2px 0;">' + formatNumber(Math.round(sellCostLAK)) + ' <span style="font-size:11px;">LAK</span></p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:6px 0 2px;">New Gold Out</p>' +
      '<p style="font-size:16px;font-weight:bold;margin:2px 0;">' + sell.newGoldG.toFixed(2) + ' g</p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:6px 0 2px;">Transactions</p>' +
      '<p style="font-size:16px;font-weight:bold;margin:2px 0;">' + sell.txCount + '</p>' +
      '</div>' +

      '<div class="stat-card">' +
      '<h3 style="color:var(--gold-primary);margin-bottom:8px;">TRADE-IN</h3>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:2px 0;">ราคาเฉลี่ย/บาท</p>' +
      '<p style="font-size:16px;font-weight:bold;margin:2px 0;">' + formatNumber(Math.round(tradeinAvg)) + ' <span style="font-size:11px;">LAK</span></p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:6px 0 2px;">New Gold Out</p>' +
      '<p style="font-size:14px;font-weight:bold;margin:2px 0;">' + tradein.newGoldG.toFixed(2) + ' g</p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:4px 0 2px;">Old Gold In</p>' +
      '<p style="font-size:14px;font-weight:bold;margin:2px 0;">' + tradein.oldGoldG.toFixed(2) + ' g</p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:4px 0 2px;">Transactions</p>' +
      '<p style="font-size:14px;font-weight:bold;margin:2px 0;">' + tradein.txCount + '</p>' +
      '</div>' +

      '<div class="stat-card">' +
      '<h3 style="color:var(--gold-primary);margin-bottom:8px;">EXCHANGE</h3>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:2px 0;">New Gold Out</p>' +
      '<p style="font-size:16px;font-weight:bold;margin:2px 0;">' + exchange.newGoldG.toFixed(2) + ' g</p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:6px 0 2px;">Old Gold In</p>' +
      '<p style="font-size:16px;font-weight:bold;margin:2px 0;">' + exchange.oldGoldG.toFixed(2) + ' g</p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:6px 0 2px;">Transactions</p>' +
      '<p style="font-size:16px;font-weight:bold;margin:2px 0;">' + exchange.txCount + '</p>' +
      '</div>' +

      '<div class="stat-card">' +
      '<h3 style="color:var(--gold-primary);margin-bottom:8px;">WITHDRAW</h3>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:2px 0;">New Gold Out</p>' +
      '<p style="font-size:16px;font-weight:bold;margin:2px 0;">' + wd.newGoldG.toFixed(2) + ' g</p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:6px 0 2px;">Transactions</p>' +
      '<p style="font-size:16px;font-weight:bold;margin:2px 0;">' + wd.txCount + '</p>' +
      '</div>' +

      '</div>' +

      '<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:15px;margin-bottom:15px;">' +

      '<div class="stat-card">' +
      '<h3 style="color:var(--gold-primary);margin-bottom:8px;">FREE EXCHANGE</h3>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:2px 0;">New Gold Out</p>' +
      '<p style="font-size:16px;font-weight:bold;margin:2px 0;">' + freeEx.newGoldG.toFixed(2) + ' g</p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:6px 0 2px;">Old Gold In</p>' +
      '<p style="font-size:16px;font-weight:bold;margin:2px 0;">' + freeEx.oldGoldG.toFixed(2) + ' g</p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:6px 0 2px;">Transactions</p>' +
      '<p style="font-size:16px;font-weight:bold;margin:2px 0;">' + freeEx.txCount + '</p>' +
      '</div>' +

      '<div class="stat-card">' +
      '<h3 style="color:var(--gold-primary);margin-bottom:8px;">SWITCH</h3>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:2px 0;">New Gold Out</p>' +
      '<p style="font-size:16px;font-weight:bold;margin:2px 0;">' + sw.newGoldG.toFixed(2) + ' g</p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:6px 0 2px;">Old Gold In</p>' +
      '<p style="font-size:16px;font-weight:bold;margin:2px 0;">' + sw.oldGoldG.toFixed(2) + ' g</p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:6px 0 2px;">Transactions</p>' +
      '<p style="font-size:16px;font-weight:bold;margin:2px 0;">' + sw.txCount + '</p>' +
      '</div>' +

      '<div class="stat-card">' +
      '<h3 style="color:var(--gold-primary);margin-bottom:8px;">BUYBACK</h3>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:2px 0;">ต้นทุน (WAC)</p>' +
      '<p style="font-size:16px;font-weight:bold;margin:2px 0;">' + formatNumber(Math.round(bbCostLAK)) + ' <span style="font-size:11px;">LAK</span></p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:6px 0 2px;">Old Gold In</p>' +
      '<p style="font-size:16px;font-weight:bold;margin:2px 0;">' + bb.oldGoldG.toFixed(2) + ' g</p>' +
      '<p style="font-size:13px;color:var(--text-secondary);margin:6px 0 2px;">Transactions</p>' +
      '<p style="font-size:16px;font-weight:bold;margin:2px 0;">' + bb.txCount + '</p>' +
      '</div>' +

      '<div class="stat-card" style="border:2px solid #c62828;background:linear-gradient(135deg,#1a1a1a 0%,#2d1a1a 100%);">' +
      '<h3 style="color:#ef5350;margin-bottom:8px;">INCOMPLETE</h3>' +
      '<p style="font-size:16px;color:#ef5350;font-weight:bold;margin:5px 0;">' + formatNumber(Math.round(incomplete.money)) + ' <span style="font-size:11px;">LAK</span></p>' +
      '<p style="font-size:14px;color:#ef5350;margin:3px 0;">' + incomplete.gold.toFixed(2) + ' g</p>' +
      '</div>' +

      '</div>' +

      '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:15px;">' +

      '<div class="stat-card">' +
      '<h3 style="color:var(--gold-primary);margin-bottom:8px;">GP / Diff</h3>' +
      '<p style="font-size:20px;font-weight:bold;color:' + (gpDiff >= 0 ? '#4caf50' : '#f44336') + ';margin:10px 0;">' + formatNumber(Math.round(gpDiff)) + ' <span style="font-size:12px;">LAK</span></p>' +
      '</div>' +

      '<div class="stat-card">' +
      '<h3 style="color:var(--gold-primary);margin-bottom:8px;">Other Expense</h3>' +
      '<p style="font-size:20px;font-weight:bold;color:#ff9800;margin:10px 0;">' + formatNumber(Math.round(otherExpenseLAK)) + ' <span style="font-size:12px;">LAK</span></p>' +
      '</div>' +

      '<div class="stat-card" style="border:2px solid var(--gold-primary);">' +
      '<h3 style="color:var(--gold-primary);margin-bottom:8px;">P/L</h3>' +
      '<p style="font-size:24px;font-weight:bold;color:' + (pl >= 0 ? '#4caf50' : '#f44336') + ';margin:10px 0;">' + formatNumber(Math.round(pl)) + ' <span style="font-size:12px;">LAK</span></p>' +
      '</div>' +

      '</div>';

  } catch (error) {
    console.error('Error loading today stats:', error);
  }
}

async function loadAccountingHistory() {
  try {
    var data = await fetchSheetData('Accounting!A:M');
    var tbody = document.getElementById('accountingTable');

    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="13" style="text-align:center;padding:40px;">No records</td></tr>';
      return;
    }

    var history = data.slice(1).reverse();

    tbody.innerHTML = history.map(function(row) {
      return '<tr>' +
        '<td>' + row[0] + '</td>' +
        '<td>' + formatNumber(row[1]) + ' LAK</td>' +
        '<td>' + parseFloat(row[2] || 0).toFixed(2) + ' g</td>' +
        '<td>' + formatNumber(row[3]) + ' LAK</td>' +
        '<td>' + parseFloat(row[4] || 0).toFixed(2) + ' g</td>' +
        '<td>' + formatNumber(row[5]) + ' LAK</td>' +
        '<td>' + parseFloat(row[6] || 0).toFixed(2) + ' g</td>' +
        '<td>' + formatNumber(row[7]) + ' LAK</td>' +
        '<td>' + parseFloat(row[8] || 0).toFixed(2) + ' g</td>' +
        '<td>' + formatNumber(row[9]) + ' LAK</td>' +
        '<td>' + parseFloat(row[10] || 0).toFixed(2) + ' g</td>' +
        '<td>' + formatNumber(row[11]) + ' LAK</td>' +
        '<td>' + parseFloat(row[12] || 0).toFixed(2) + ' g</td>' +
      '</tr>';
    }).join('');
  } catch (error) {
    console.error('Error loading accounting history:', error);
  }
}

async function filterAccountingHistory() {
  try {
    var startDate = document.getElementById('accountingStartDate').value;
    var endDate = document.getElementById('accountingEndDate').value;

    if (!startDate || !endDate) return;

    var data = await fetchSheetData('Accounting!A:M');
    var tbody = document.getElementById('accountingTable');

    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="13" style="text-align:center;padding:40px;">No records</td></tr>';
      return;
    }

    var filtered = data.slice(1).filter(function(row) {
      return row[0] >= startDate && row[0] <= endDate;
    }).reverse();

    if (filtered.length === 0) {
      tbody.innerHTML = '<tr><td colspan="13" style="text-align:center;padding:40px;">No records in this date range</td></tr>';
      return;
    }

    tbody.innerHTML = filtered.map(function(row) {
      return '<tr>' +
        '<td>' + row[0] + '</td>' +
        '<td>' + formatNumber(row[1]) + ' LAK</td>' +
        '<td>' + parseFloat(row[2] || 0).toFixed(2) + ' g</td>' +
        '<td>' + formatNumber(row[3]) + ' LAK</td>' +
        '<td>' + parseFloat(row[4] || 0).toFixed(2) + ' g</td>' +
        '<td>' + formatNumber(row[5]) + ' LAK</td>' +
        '<td>' + parseFloat(row[6] || 0).toFixed(2) + ' g</td>' +
        '<td>' + formatNumber(row[7]) + ' LAK</td>' +
        '<td>' + parseFloat(row[8] || 0).toFixed(2) + ' g</td>' +
        '<td>' + formatNumber(row[9]) + ' LAK</td>' +
        '<td>' + parseFloat(row[10] || 0).toFixed(2) + ' g</td>' +
        '<td>' + formatNumber(row[11]) + ' LAK</td>' +
        '<td>' + parseFloat(row[12] || 0).toFixed(2) + ' g</td>' +
      '</tr>';
    }).join('');
  } catch (error) {
    console.error('Error filtering accounting:', error);
  }
}

function checkAccountingFilter() {
  var startDate = document.getElementById('accountingStartDate').value;
  var endDate = document.getElementById('accountingEndDate').value;
  if (startDate && endDate) filterAccountingHistory();
}

async function showTodayAccounting() {
  var today = new Date().toISOString().split('T')[0];
  document.getElementById('accountingStartDate').value = today;
  document.getElementById('accountingEndDate').value = today;
  filterAccountingHistory();
}

function calcGold(itemsStr) {
  var total = 0;
  try {
    JSON.parse(itemsStr).forEach(function(item) {
      total += getGoldWeight(item.productId) * item.qty;
    });
  } catch(e) {}
  return total;
}