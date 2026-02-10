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
      const allowedTabs = ['sell', 'trade-in', 'exchange', 'switch', 'free ex', 'buyback', 'withdraw'];
      document.querySelectorAll('.nav-btn').forEach(btn => {
        const text = btn.textContent.toLowerCase();
        let isAllowed = false;
        allowedTabs.forEach(tab => {
          if (text.includes(tab)) isAllowed = true;
        });
        if (!isAllowed) btn.style.display = 'none';
      });
    }
    
    if (currentUser.role === 'Manager') {
      const managerHideButtons = ['addSellBtn', 'addTradeinBtn', 'addBuybackBtn', 'addExchangeBtn', 'addSwitchBtn', 'addFreeExchangeBtn', 'addWithdrawBtn', 'withdrawBtn'];
      managerHideButtons.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.style.display = 'none';
      });
    }
    
    fetchExchangeRates();
    checkPendingClose();
    
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
  currentUser = null;
  
  localStorage.removeItem('currentUser');
  
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
  if (savedUser) {
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
        const allowedTabs = ['sell', 'trade-in', 'exchange', 'switch', 'free ex', 'buyback', 'withdraw'];
        document.querySelectorAll('.nav-btn').forEach(btn => {
          const text = btn.textContent.toLowerCase();
          let isAllowed = false;
          allowedTabs.forEach(tab => {
            if (text.includes(tab)) isAllowed = true;
          });
          if (!isAllowed) btn.style.display = 'none';
        });
      }
      
      if (currentUser.role === 'Manager') {
        const managerHideButtons = ['addSellBtn', 'addTradeinBtn', 'addBuybackBtn', 'addExchangeBtn', 'addSwitchBtn', 'addFreeExchangeBtn', 'addWithdrawBtn', 'withdrawBtn'];
        managerHideButtons.forEach(id => {
          const el = document.getElementById(id);
          if (el) el.style.display = 'none';
        });
      }
      
      fetchExchangeRates();
      checkPendingClose();
      
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
    }
  }
})();