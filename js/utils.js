function formatNumber(num) {
  return new Intl.NumberFormat('en-US').format(Math.round(num));
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
    console.error('formatDateOnly error:', error, dateInput);
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
    console.error('formatDateTime error:', error, dateInput);
    return '-';
  }
}

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
      return `${product.name}: ${item.qty} unit`;
    }).join('\n');
  } catch (error) {
    return itemsJson;
  }
}

function calculatePremiumFromItems(itemsJson) {
  try {
    const items = JSON.parse(itemsJson);
    let premium = 0;
    items.forEach(item => {
      if (PREMIUM_PRODUCTS.includes(item.productId)) {
        premium += PREMIUM_PER_PIECE * item.qty;
      }
    });
    return premium;
  } catch {
    return 0;
  }
}

function filterTodayData(data, dateColumnIndex, createdByIndex) {
  const today = new Date();
  const todayStart = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 0, 0, 0);
  const todayEnd = new Date(today.getFullYear(), today.getMonth(), today.getDate(), 23, 59, 59);
  
  console.log('🔍 Filter Debug:', {
    todayStart,
    todayEnd,
    userRole: currentUser?.role,
    userNickname: currentUser?.nickname,
    totalRows: data.length
  });
  
  const filtered = data.filter(row => {
    const dateStr = row[dateColumnIndex];
    const createdBy = row[createdByIndex];
    
    // Parse DD/MM/YYYY HH:mm format
    let rowDate;
    if (typeof dateStr === 'string' && dateStr.includes('/')) {
      const parts = dateStr.split(' ');
      const dateParts = parts[0].split('/');
      const day = parseInt(dateParts[0]);
      const month = parseInt(dateParts[1]) - 1; // Month is 0-indexed
      const year = parseInt(dateParts[2]);
      rowDate = new Date(year, month, day);
    } else {
      rowDate = new Date(dateStr);
    }
    
    const isToday = rowDate >= todayStart && rowDate <= todayEnd;
    
    console.log('Row check:', {
      date: dateStr,
      parsedDate: rowDate,
      isToday,
      createdBy,
      matches: currentUser.role === 'Manager' ? 'Manager sees all' : (createdBy === currentUser.nickname)
    });
    
    if (currentUser.role === 'Manager') {
      return isToday;
    } else if (currentUser.role === 'User') {
      return isToday && createdBy === currentUser.nickname;
    }
    return isToday;
  });
  
  console.log('✅ Filtered result:', filtered.length, 'rows');
  
  return filtered;
}

function showLoading() {
  document.getElementById('loading').classList.add('active');
}

function hideLoading() {
  document.getElementById('loading').classList.remove('active');
}

function openModal(modalId) {
  document.getElementById(modalId).classList.add('active');
}

function closeModal(modalId) {
  document.getElementById(modalId).classList.remove('active');
}
