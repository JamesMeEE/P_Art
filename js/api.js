async function fetchSheetData(range) {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.SHEET_ID}/values/${range}?key=${CONFIG.API_KEY}`;
  const response = await fetch(url);
  const data = await response.json();
  return data.values || [];
}

async function callAppsScript(action, params = {}) {
  const queryParams = new URLSearchParams({
    action,
    ...params,
    user: currentUser?.nickname || currentUser?.role || 'Unknown'
  });
  
  const url = `${CONFIG.SCRIPT_URL}?${queryParams.toString()}`;
  
  const response = await fetch(url, {
    method: 'GET',
    redirect: 'follow'
  });
  
  const result = await response.json();
  return result;
}

async function fetchExchangeRates() {
  return { LAK: 1, THB: 270, USD: 21500 };
}
