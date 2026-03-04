var _notifInterval = null;
var _notifData = [];
var _notifDropdownOpen = false;
var _markedReadIds = {};

function startNotificationPolling() {
  if (_notifInterval) clearInterval(_notifInterval);
  pollNotifications();
  _notifInterval = setInterval(pollNotifications, 5000);
}

function stopNotificationPolling() {
  if (_notifInterval) { clearInterval(_notifInterval); _notifInterval = null; }
}

async function pollNotifications() {
  try {
    var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + CONFIG.SHEET_ID + '/values/' + encodeURIComponent('_notifications!A:I') + '?key=' + CONFIG.API_KEY;
    var resp = await fetch(url);
    if (!resp.ok) {
      _notifData = [];
      updateNotifBadge();
      return;
    }
    var json = await resp.json();
    var data = json.values || [];
    if (data.length <= 1) {
      _notifData = [];
      updateNotifBadge();
      return;
    }

    var user = currentUser.name;
    var role = currentUser.role;
    var filtered = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var targetRole = String(row[3] || '');
      var targetUser = String(row[4] || '').trim();
      var createdBy = String(row[6] || '').trim();
      var readBy = String(row[8] || '').trim();

      if (createdBy === user) continue;

      var isTarget = false;
      if (targetUser && targetUser === user) {
        isTarget = true;
      } else if (targetRole && targetRole.indexOf(role) >= 0 && !targetUser) {
        isTarget = true;
      }

      if (isTarget) {
        var readList = readBy.split(',').map(function(s) { return s.trim(); });
        var isRead = readList.indexOf(user) >= 0 || !!_markedReadIds[row[0]];
        filtered.push({
          id: row[0],
          type: row[1],
          message: row[2],
          tab: String(row[5] || ''),
          createdAt: row[7],
          read: isRead
        });
      }
    }

    _notifData = filtered.reverse();
    updateNotifBadge();
  } catch(e) {}
}

function updateNotifBadge() {
  var badge = document.getElementById('notifBadge');
  if (!badge) return;
  var unread = _notifData.filter(function(n) { return !n.read; }).length;
  if (unread > 0) {
    badge.textContent = unread > 99 ? '99+' : unread;
    badge.style.display = 'flex';
  } else {
    badge.style.display = 'none';
  }
}

function toggleNotifDropdown() {
  var dropdown = document.getElementById('notifDropdown');
  if (!dropdown) return;

  _notifDropdownOpen = !_notifDropdownOpen;
  dropdown.style.display = _notifDropdownOpen ? 'block' : 'none';

  if (_notifDropdownOpen) {
    renderNotifList();
    markAllRead();
  }
}

function renderNotifList() {
  var list = document.getElementById('notifList');
  if (!list) return;

  if (_notifData.length === 0) {
    list.innerHTML = '<div style="padding:15px;text-align:center;color:var(--text-secondary);">ไม่มีการแจ้งเตือน</div>';
    return;
  }

  list.innerHTML = _notifData.slice(0, 30).map(function(n) {
    var icon = '📌';
    if (n.type === 'NEW_TX') icon = '🆕';
    else if (n.type === 'TX_APPROVED') icon = '✅';
    else if (n.type === 'TX_REJECTED') icon = '❌';
    else if (n.type === 'PRICE_UPDATE') icon = '💰';
    else if (n.type === 'CASHBANK') icon = '🏦';
    else if (n.type === 'TRANSFER') icon = '🔄';
    else if (n.type === 'STOCK') icon = '📦';

    var time = '';
    try { time = formatDateTime(n.createdAt); } catch(e) {}

    var bg = n.read ? 'transparent' : 'rgba(212,175,55,0.08)';

    return '<div onclick="goToNotifTab(\'' + n.tab + '\')" style="padding:10px 15px;border-bottom:1px solid var(--border-color);cursor:pointer;background:' + bg + ';">' +
      '<div style="font-size:13px;">' + icon + ' ' + n.message + '</div>' +
      '<div style="font-size:11px;color:var(--text-secondary);margin-top:3px;">' + time + '</div>' +
      '</div>';
  }).join('');
}

function goToNotifTab(tab) {
  if (tab) showSection(tab);
  _notifDropdownOpen = false;
  var dropdown = document.getElementById('notifDropdown');
  if (dropdown) dropdown.style.display = 'none';
}

async function markAllRead() {
  var unread = _notifData.filter(function(n) { return !n.read; });
  if (unread.length === 0) return;

  _notifData.forEach(function(n) {
    n.read = true;
    _markedReadIds[n.id] = true;
  });
  updateNotifBadge();

  try {
    await callAppsScript('MARK_NOTIFICATIONS_READ', { user: currentUser.name });
  } catch(e) {}
}

async function refreshPage() {
  var btn = document.getElementById('refreshBtn');
  if (btn) {
    btn.style.animation = 'spin 0.8s linear infinite';
    btn.style.pointerEvents = 'none';
  }

  var activeTab = document.querySelector('.nav-btn.active');
  var tabName = 'dashboard';
  if (activeTab) {
    var onclick = activeTab.getAttribute('onclick');
    var match = onclick.match(/showSection\('(.+?)'\)/);
    if (match) tabName = match[1];
  }

  try {
    invalidateCache();
    await showSection(tabName);
    await pollNotifications();
    if (typeof checkPendingClose === 'function') checkPendingClose();
  } catch(e) {}

  if (btn) {
    btn.style.animation = '';
    btn.style.pointerEvents = '';
  }
}

document.addEventListener('click', function(e) {
  var bell = document.getElementById('notifBell');
  var dropdown = document.getElementById('notifDropdown');
  if (bell && dropdown && !bell.contains(e.target) && !dropdown.contains(e.target)) {
    _notifDropdownOpen = false;
    dropdown.style.display = 'none';
  }
});
