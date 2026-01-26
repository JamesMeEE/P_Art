async function loadReports() {
  try {
    showLoading();
    
    await checkAndCalculateMissingReports();
    
    const data = await fetchSheetData('Reports!A:C');
    const tbody = document.getElementById('reportsTable');
    
    if (data.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="3" style="text-align: center; padding: 40px;">No reports yet</td></tr>';
      hideLoading();
      return;
    }
    
    const reports = data.slice(1).reverse();
    
    tbody.innerHTML = reports.map(row => `
      <tr>
        <td style="text-align: center;">${row[0]}</td>
        <td style="text-align: center;">${parseFloat(row[1] || 0).toFixed(2)}</td>
        <td style="text-align: center;">${parseFloat(row[2] || 0).toFixed(2)}</td>
      </tr>
    `).join('');
    
    hideLoading();
  } catch (error) {
    console.error('Error loading reports:', error);
    document.getElementById('reportsTable').innerHTML = '<tr><td colspan="3" style="text-align: center; padding: 40px; color: #f44336;">Error loading reports</td></tr>';
    hideLoading();
  }
}

async function checkAndCalculateMissingReports() {
  try {
    const result = await callAppsScript('AUTO_CALCULATE_REPORTS', {});
    if (result.calculated > 0) {
      console.log(`Auto-calculated ${result.calculated} missing reports`);
    }
  } catch (error) {
    console.error('Error auto-calculating reports:', error);
  }
}