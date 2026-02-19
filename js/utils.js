function isManager() {
  return currentUser && (currentUser.role === 'Manager' || currentUser.role === 'Admin');
}

function formatNumber(num) {
  return new Intl.NumberFormat('en-US').format(Math.round(num));
}

function formatWeight(num) {
  return new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(num);
}

function parseSheetDate(dateValue) {
  if (!dateValue) return null;
  
  try {
    var result = null;

    if (dateValue instanceof Date) {
      result = dateValue;
    } else if (typeof dateValue === 'number') {
      result = new Date((dateValue - 25569) * 86400 * 1000);
    } else if (typeof dateValue === 'string') {
      if (dateValue.includes('/')) {
        var parts = dateValue.split(' ');
        var dateParts = parts[0].split('/');
        var day = parseInt(dateParts[0]);
        var month = parseInt(dateParts[1]) - 1;
        var year = parseInt(dateParts[2]);
        
        if (parts.length > 1 && parts[1] && parts[1].includes(':')) {
          var timeParts = parts[1].split(':');
          var hour = parseInt(timeParts[0]) || 0;
          var minute = parseInt(timeParts[1]) || 0;
          var second = parseInt(timeParts[2]) || 0;
          result = new Date(year, month, day, hour, minute, second);
        } else {
          result = new Date(year, month, day);
        }
      } else {
        var isoMatch = dateValue.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (isoMatch) {
          var tMatch = dateValue.match(/(\d{2}):(\d{2}):?(\d{2})?/);
          if (tMatch) {
            result = new Date(parseInt(isoMatch[1]), parseInt(isoMatch[2]) - 1, parseInt(isoMatch[3]), parseInt(tMatch[1]) || 0, parseInt(tMatch[2]) || 0, parseInt(tMatch[3]) || 0);
          } else {
            result = new Date(parseInt(isoMatch[1]), parseInt(isoMatch[2]) - 1, parseInt(isoMatch[3]));
          }
        } else {
          result = new Date(dateValue);
        }
      }
    } else {
      result = new Date(dateValue);
    }

    if (result && !isNaN(result.getTime())) return result;
    return null;
  } catch (error) {
    return null;
  }
}

function formatDateOnly(dateInput) {
  if (!dateInput) return '-';
  
  try {
    let d;
    
    if (typeof dateInput === 'string') {
      const parts = dateInput.split(' ')[0].split('/');
      if (parts.length === 3) {
        const day = parseInt(parts[0]);
        const month = parseInt(parts[1]);
        const year = parseInt(parts[2]);
        d = new Date(year, month - 1, day);
      } else {
        d = new Date(dateInput);
      }
    } else if (typeof dateInput === 'number') {
      d = new Date((dateInput - 25569) * 86400 * 1000);
    } else if (dateInput instanceof Date) {
      d = dateInput;
    } else {
      return '-';
    }
    
    if (isNaN(d.getTime())) return '-';
    
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    return `${day}/${month}/${year}`;
  } catch (error) {
    return '-';
  }
}

function formatDateTime(dateInput) {
  if (!dateInput) return '-';
  
  try {
    let d;
    
    if (typeof dateInput === 'string') {
      if (dateInput.includes('/') && dateInput.includes(':')) {
        const [datePart, timePart] = dateInput.split(' ');
        const dateParts = datePart.split('/');
        const timeParts = timePart.split(':');
        
        const day = parseInt(dateParts[0]);
        const month = parseInt(dateParts[1]);
        const year = parseInt(dateParts[2]);
        const hour = parseInt(timeParts[0]);
        const minute = parseInt(timeParts[1]);
        const second = timeParts[2] ? parseInt(timeParts[2]) : 0;
        
        d = new Date(year, month - 1, day, hour, minute, second);
      } else {
        d = new Date(dateInput);
      }
    } else if (typeof dateInput === 'number') {
      d = new Date((dateInput - 25569) * 86400 * 1000);
    } else if (dateInput instanceof Date) {
      d = dateInput;
    } else {
      return '-';
    }
    
    if (isNaN(d.getTime())) return '-';
    
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    const hour = String(d.getHours()).padStart(2, '0');
    const minute = String(d.getMinutes()).padStart(2, '0');
    return `${day}/${month}/${year} ${hour}:${minute}`;
  } catch (error) {
    return '-';
  }
}

const formatDate = formatDateTime;

function formatItemsForTable(itemsJson) {
  try {
    const items = JSON.parse(itemsJson);
    return items.map(item => {
      const product = FIXED_PRODUCTS.find(p => p.id === item.productId);
      return `${product.name}: ${item.qty} unit`;
    }).join('<br>');
  } catch (error) {
    return itemsJson;
  }
}

function formatItemsForDisplay(itemsJson) {
  try {
    const items = JSON.parse(itemsJson);
    return items.map(item => {
      const product = FIXED_PRODUCTS.find(p => p.id === item.productId);
      return `• ${product.name} × ${item.qty}`;
    }).join('\n');
  } catch (error) {
    return itemsJson;
  }
}

function calculatePremiumFromItems(itemsJson) {
  try {
    const items = JSON.parse(itemsJson);
    let totalPremium = 0;
    items.forEach(item => {
      if (PREMIUM_PRODUCTS.includes(item.productId)) {
        totalPremium += PREMIUM_PER_PIECE * item.qty;
      }
    });
    return totalPremium;
  } catch {
    return 0;
  }
}

function filterTodayData(data, dateColumnIndex, createdByIndex) {
  const today = new Date();
  const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  
  return data.filter(row => {
    const dateValue = row[dateColumnIndex];
    const createdBy = row[createdByIndex];
    
    let rowDate;
    if (dateValue instanceof Date) {
      rowDate = dateValue;
    } else if (typeof dateValue === 'string') {
      if (dateValue.includes('/')) {
        const parts = dateValue.split(' ');
        const dateParts = parts[0].split('/');
        const day = parseInt(dateParts[0]);
        const month = parseInt(dateParts[1]) - 1;
        const year = parseInt(dateParts[2]);
        rowDate = new Date(year, month, day);
      } else {
        var isoMatch = dateValue.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (isoMatch) {
          rowDate = new Date(parseInt(isoMatch[1]), parseInt(isoMatch[2]) - 1, parseInt(isoMatch[3]));
        } else {
          rowDate = new Date(dateValue);
        }
      }
    } else {
      rowDate = new Date(dateValue);
    }
    
    const rowDateStart = new Date(rowDate.getFullYear(), rowDate.getMonth(), rowDate.getDate());
    const isToday = rowDateStart.getTime() === todayStart.getTime();
    
    if (isManager()) {
      return isToday;
    } else if (currentUser.role === 'User') {
      return isToday && createdBy === currentUser.nickname;
    }
    return isToday;
  });
}

var _isSubmitting = false;
var _submitTimeout = null;

function showLoading() {
  document.getElementById('loading').classList.add('active');
  if (_isSubmitting && !_submitTimeout) {
    _submitTimeout = setTimeout(function() { _isSubmitting = false; hideLoading(); }, 30000);
  }
}

function hideLoading() {
  document.getElementById('loading').classList.remove('active');
  _isSubmitting = false;
  if (_submitTimeout) { clearTimeout(_submitTimeout); _submitTimeout = null; }
}

function startSubmit() {
  if (_isSubmitting) return false;
  _isSubmitting = true;
  showLoading();
  _submitTimeout = setTimeout(function() { _isSubmitting = false; }, 30000);
  return true;
}

function endSubmit() {
  _isSubmitting = false;
  if (_submitTimeout) { clearTimeout(_submitTimeout); _submitTimeout = null; }
  hideLoading();
}

function openModal(modalId) {
  document.getElementById(modalId).classList.add('active');
}

function closeModal(modalId) {
  document.getElementById(modalId).classList.remove('active');
}

function roundTo1000(num) {
  return Math.ceil(num / 1000) * 1000;
}

function calculateSellPrice(productId, sell1Baht) {
  let price = 0;
  switch(productId) {
    case 'G01': price = sell1Baht * 10; break;
    case 'G02': price = sell1Baht * 5; break;
    case 'G03': price = sell1Baht * 2; break;
    case 'G04': price = sell1Baht; break;
    case 'G05': price = (sell1Baht / 2); break;
    case 'G06': price = (sell1Baht / 4); break;
    case 'G07': price = Math.round(((sell1Baht / 15) + 120000) / 1000) * 1000; break;
  }
  return price;
}

function calculateBuybackPrice(productId, sell1Baht) {
  const buyback1B = sell1Baht - 530000;
  
  let price = 0;
  switch(productId) {
    case 'G01': price = buyback1B * 10; break;
    case 'G02': price = buyback1B * 5; break;
    case 'G03': price = buyback1B * 2; break;
    case 'G04': price = buyback1B; break;
    case 'G05': price = buyback1B / 2; break;
    case 'G06': price = buyback1B / 4; break;
    case 'G07': price = buyback1B / 15; break;
  }
  return Math.round(price / 1000) * 1000;
}

function filterByDateRange(data, dateColumnIndex, createdByIndex, dateFrom, dateTo) {
  var from = null;
  var to = null;

  if (dateFrom) {
    var fParts = dateFrom.split('-');
    from = new Date(parseInt(fParts[0]), parseInt(fParts[1]) - 1, parseInt(fParts[2]), 0, 0, 0, 0);
  }
  if (dateTo) {
    var tParts = dateTo.split('-');
    to = new Date(parseInt(tParts[0]), parseInt(tParts[1]) - 1, parseInt(tParts[2]), 23, 59, 59, 999);
  }
  
  return data.filter(row => {
    const dateValue = row[dateColumnIndex];
    const createdBy = row[createdByIndex];
    
    let rowDate;
    if (dateValue instanceof Date) {
      rowDate = dateValue;
    } else if (typeof dateValue === 'string') {
      if (dateValue.includes('/')) {
        const parts = dateValue.split(' ');
        const dateParts = parts[0].split('/');
        const day = parseInt(dateParts[0]);
        const month = parseInt(dateParts[1]) - 1;
        const year = parseInt(dateParts[2]);
        if (parts.length > 1 && parts[1]) {
          const timeParts = parts[1].split(':');
          rowDate = new Date(year, month, day, parseInt(timeParts[0]) || 0, parseInt(timeParts[1]) || 0, parseInt(timeParts[2]) || 0);
        } else {
          rowDate = new Date(year, month, day);
        }
      } else {
        var isoMatch = dateValue.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (isoMatch) {
          var tMatch = dateValue.match(/(\d{2}):(\d{2}):?(\d{2})?/);
          if (tMatch) {
            rowDate = new Date(parseInt(isoMatch[1]), parseInt(isoMatch[2]) - 1, parseInt(isoMatch[3]), parseInt(tMatch[1]) || 0, parseInt(tMatch[2]) || 0, parseInt(tMatch[3]) || 0);
          } else {
            rowDate = new Date(parseInt(isoMatch[1]), parseInt(isoMatch[2]) - 1, parseInt(isoMatch[3]));
          }
        } else {
          rowDate = new Date(dateValue);
        }
      }
    } else {
      rowDate = new Date(dateValue);
    }
    
    let inRange = true;
    if (from) {
      inRange = inRange && rowDate >= from;
    }
    if (to) {
      inRange = inRange && rowDate <= to;
    }
    
    if (isManager()) {
      return inRange;
    } else if (currentUser.role === 'User') {
      return inRange && createdBy === currentUser.nickname;
    }
    return inRange;
  });
}

function getTodayDateString() {
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, '0');
  const day = String(today.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function viewTransactionDetail(type, jsonData) {
  var row = JSON.parse(decodeURIComponent(jsonData));
  var html = '<div style="padding:20px;">';
  html += '<h3 style="color:var(--gold-primary);margin-bottom:15px;">' + type.toUpperCase() + ' Detail</h3>';
  html += '<table style="width:100%;border-collapse:collapse;">';
  row.forEach(function(item) {
    html += '<tr style="border-bottom:1px solid var(--border-color);">';
    html += '<td style="padding:8px 12px;color:var(--text-secondary);white-space:nowrap;">' + item[0] + '</td>';
    html += '<td style="padding:8px 12px;font-weight:600;">' + item[1] + '</td>';
    html += '</tr>';
  });
  html += '</table></div>';

  var modal = document.getElementById('viewDetailModal');
  if (!modal) {
    modal = document.createElement('div');
    modal.id = 'viewDetailModal';
    modal.className = 'modal';
    modal.innerHTML = '<div class="modal-content" style="max-width:500px;"><div id="viewDetailContent"></div><div style="text-align:right;padding:0 20px 20px;"><button class="btn-secondary" onclick="closeModal(\'viewDetailModal\')">Close</button></div></div>';
    document.body.appendChild(modal);
  }
  document.getElementById('viewDetailContent').innerHTML = html;
  openModal('viewDetailModal');
}