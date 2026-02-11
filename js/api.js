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

const executeGoogleScript = callAppsScript;

async function fetchExchangeRates() {
  try {
    var prData = await fetchSheetData('PriceRate!A:E');
    if (prData.length > 1) {
      var last = prData[prData.length - 1];
      currentExchangeRates = {
        LAK: 1,
        THB_Sell: parseFloat(last[1]) || 0,
        USD_Sell: parseFloat(last[2]) || 0,
        THB_Buy: parseFloat(last[3]) || 0,
        USD_Buy: parseFloat(last[4]) || 0,
        THB: parseFloat(last[1]) || 0,
        USD: parseFloat(last[2]) || 0
      };
    }
  } catch(e) {
    console.error('Error fetching exchange rates:', e);
  }
  return currentExchangeRates;
}