function setupManagerUI() {
  var managerHideButtons = ['addSellBtn', 'addTradeinBtn', 'addBuybackBtn', 'addExchangeBtn', 'addSwitchBtn', 'addFreeExchangeBtn', 'addWithdrawBtn', 'withdrawBtn'];
  managerHideButtons.forEach(function(id) {
    var el = document.getElementById(id);
    if (el) el.style.display = 'none';
  });
  var hideSections = ["'sell'", "'tradein'", "'exchange'", "'switch'", "'freeexchange'", "'withdraw'"];
  document.querySelectorAll('.nav-btn').forEach(function(btn) {
    var oc = (btn.getAttribute('onclick') || '') + '';
    for (var i = 0; i < hideSections.length; i++) {
      if (oc.indexOf(hideSections[i]) !== -1) {
        btn.style.display = 'none';
        break;
      }
    }
    if (oc.indexOf("'buyback'") !== -1) btn.textContent = '◑ History Buyback';
  });
  var hBtn = document.getElementById('navHistorySell');
  if (hBtn) hBtn.style.display = '';
  var bbTitle = document.getElementById('buybackTitle');
  if (bbTitle) bbTitle.textContent = 'History Buyback';
}

var _inactivityTimer = null;
var INACTIVITY_LIMIT = 60 * 60 * 1000;

function resetInactivityTimer() {
  localStorage.setItem('lastActivity', Date.now());
  if (_inactivityTimer) clearTimeout(_inactivityTimer);
  _inactivityTimer = setTimeout(function() {
    if (currentUser) {
      alert('⏰ ไม่มีการใช้งาน 1 ชั่วโมง — ออกจากระบบอัตโนมัติ');
      logout();
    }
  }, INACTIVITY_LIMIT);
}

function startInactivityWatch() {
  ['click', 'keydown', 'mousemove', 'touchstart', 'scroll'].forEach(function(evt) {
    document.addEventListener(evt, resetInactivityTimer, { passive: true });
  });
  resetInactivityTimer();
}

function stopInactivityWatch() {
  ['click', 'keydown', 'mousemove', 'touchstart', 'scroll'].forEach(function(evt) {
    document.removeEventListener(evt, resetInactivityTimer);
  });
  if (_inactivityTimer) { clearTimeout(_inactivityTimer); _inactivityTimer = null; }
}

function isSessionExpired() {
  var last = parseInt(localStorage.getItem('lastActivity')) || 0;
  return (Date.now() - last) > INACTIVITY_LIMIT;
}

function login() {
  const username = document.getElementById('loginUsername').value;
  const password = document.getElementById('loginPassword').value;

  if (USERS[username] && USERS[username].password === password) {
    currentUser = { username, ...USERS[username] };

    localStorage.setItem('currentUser', JSON.stringify(currentUser));

    document.getElementById('loginScreen').classList.remove('active');
    document.getElementById('mainHeader').style.display = 'block';
    document.getElementById('mainContainer').style.display = 'block';
    document.getElementById('userName').textContent = currentUser.nickname || currentUser.role;
    document.getElementById('userRole').textContent = currentUser.role;
    document.getElementById('userAvatar').textContent = (currentUser.nickname || currentUser.role)[0];

    document.body.className = 'role-' + username;

    if (currentUser.role === 'Accountant') {
      document.body.classList.add('accountant-readonly');
      const readonlyBtns = ['quickSales', 'quickTradein', 'quickBuyback', 'addSellBtn', 'addTradeinBtn', 'addBuybackBtn', 'addExchangeBtn', 'addSwitchBtn', 'addFreeExchangeBtn', 'withdrawBtn'];
      readonlyBtns.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.style.display = 'none';
      });
    }

    if (currentUser.role === 'User') {
      document.querySelectorAll('.date-filter').forEach(function(el) { el.style.display = 'none'; });
      const allowedTabs = ['sell', 'trade-in', 'exchange', 'switch', 'free exchange', 'buyback', 'withdraw'];
      document.querySelectorAll('.nav-btn').forEach(btn => {
        const text = btn.textContent.toLowerCase();
        let isAllowed = false;
        allowedTabs.forEach(tab => {
          if (text.includes(tab)) isAllowed = true;
        });
        if (!isAllowed) btn.style.display = 'none';
      });
    }

    if (isManager()) { setupManagerUI(); }

    fetchExchangeRates();
    checkPendingClose();
    callAppsScript('INIT_STOCK').catch(function(){});
    startAutoRefresh();
    startInactivityWatch();

    if (currentUser.role === 'User') {
      showSection('sell');
    } else {
      loadDashboard();
    }
  } else {
    alert('Invalid username or password');
  }
}

function logout() {
  stopInactivityWatch();
  stopAutoRefresh();
  currentUser = null;

  localStorage.removeItem('currentUser');
  localStorage.removeItem('lastActivity');

  document.getElementById('loginScreen').classList.add('active');
  document.getElementById('mainHeader').style.display = 'none';
  document.getElementById('mainContainer').style.display = 'none';
  document.getElementById('loginUsername').value = '';
  document.getElementById('loginPassword').value = '';
  document.body.className = '';

  localStorage.setItem('cacheBuster', Date.now());

  setTimeout(() => {
    window.location.reload();
  }, 100);
}

(function checkSession() {
  const savedUser = localStorage.getItem('currentUser');
  if (!savedUser) return;

  if (isSessionExpired()) {
    localStorage.removeItem('currentUser');
    localStorage.removeItem('lastActivity');
    return;
  }

  try {
    currentUser = JSON.parse(savedUser);
    document.getElementById('loginScreen').classList.remove('active');
    document.getElementById('mainHeader').style.display = 'block';
    document.getElementById('mainContainer').style.display = 'block';
    document.getElementById('userName').textContent = currentUser.nickname || currentUser.role;
    document.getElementById('userRole').textContent = currentUser.role;
    document.getElementById('userAvatar').textContent = (currentUser.nickname || currentUser.role)[0];

    document.body.className = 'role-' + currentUser.username;

    if (currentUser.role === 'Accountant') {
      document.body.classList.add('accountant-readonly');
      const readonlyBtns = ['quickSales', 'quickTradein', 'quickBuyback', 'addSellBtn', 'addTradeinBtn', 'addBuybackBtn', 'addExchangeBtn', 'addSwitchBtn', 'addFreeExchangeBtn', 'withdrawBtn'];
      readonlyBtns.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.style.display = 'none';
      });
    }

    if (currentUser.role === 'User') {
      document.querySelectorAll('.date-filter').forEach(function(el) { el.style.display = 'none'; });
      const allowedTabs = ['sell', 'trade-in', 'exchange', 'switch', 'free exchange', 'buyback', 'withdraw'];
      document.querySelectorAll('.nav-btn').forEach(btn => {
        const text = btn.textContent.toLowerCase();
        let isAllowed = false;
        allowedTabs.forEach(tab => {
          if (text.includes(tab)) isAllowed = true;
        });
        if (!isAllowed) btn.style.display = 'none';
      });
    }

    if (isManager()) { setupManagerUI(); }

    fetchExchangeRates();
    checkPendingClose();
    callAppsScript('INIT_STOCK').catch(function(){});
    startAutoRefresh();
    startInactivityWatch();

    setTimeout(() => {
      if (currentUser.role === 'User') {
        showSection('sell');
      } else {
        loadDashboard();
      }
    }, 100);
  } catch (error) {
    console.error('Session restore error:', error);
    localStorage.removeItem('currentUser');
    localStorage.removeItem('lastActivity');
  }
})();