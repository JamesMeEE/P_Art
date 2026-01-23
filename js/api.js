async function fetchSheetData(range) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.SHEET_ID}/values/${range}?key=${CONFIG.API_KEY}`;
  const response = await fetch(url);
  const data = await response.json();
  return data.values || [];
}

async function callAppsScript(action, params = {}) {
  const response = await fetch(CONFIG.SCRIPT_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      action,
      ...params,
      user: currentUser?.nickname || currentUser?.role || 'Unknown'
    })
  });
  
  const result = await response.json();
  return result;
}

async function fetchExchangeRates() {
  return { LAK: 1, THB: 270, USD: 21500 };
}
