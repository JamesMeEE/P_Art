async function loadReports() {
  try {
    showLoading();
    
    const data = await fetchSheetData('Reports!A:C');
    
    if (data.length > 1) {
      const today = new Date();
      const yesterday = new Date(today);
      yesterday.setDate(yesterday.getDate() - 1);
      yesterday.setHours(0, 0, 0, 0);
      
      const yesterdayKey = `${yesterday.getFullYear()}-${String(yesterday.getMonth()+1).padStart(2,'0')}-${String(yesterday.getDate()).padStart(2,'0')}`;
      
      let hasYesterday = false;
      data.slice(1).forEach(row => {
        if (row[0]) {
          const d = parseSheetDate(row[0]);
          if (d) {
            const dateKey = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
            if (dateKey === yesterdayKey) {
              hasYesterday = true;
            }
          }
        }
      });
      
      if (!hasYesterday) {
        await checkAndCalculateMissingReports();
      }
    } else {
      await checkAndCalculateMissingReports();
    }
    
    const updatedData = await fetchSheetData('Reports!A:C');
    const tbody = document.getElementById('reportsTable');
    
    if (updatedData.length <= 1) {
      tbody.innerHTML = '<tr><td colspan="3" style="text-align: center; padding: 40px;">No reports yet</td></tr>';
      hideLoading();
      return;
    }
    
    const reports = updatedData.slice(1).reverse();
    
    tbody.innerHTML = reports.map(row => `
      <tr>
        <td style="text-align: center;">${formatDateOnly(row[0])}</td>
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