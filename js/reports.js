async function loadReports() {
  try {
    showLoading();
    
    document.getElementById('reportsContent').innerHTML = `
      <div style="text-align: center; padding: 60px 20px;">
        <div style="font-size: 48px; margin-bottom: 20px;">📊</div>
        <h3 style="color: var(--gold-primary); margin-bottom: 15px;">Reports Coming Soon</h3>
        <p style="color: var(--text-secondary); max-width: 500px; margin: 0 auto;">
          รายงานและการวิเคราะห์ข้อมูลจะเพิ่มในเร็วๆ นี้
        </p>
      </div>
    `;
    
    hideLoading();
  } catch (error) {
    console.error('Error loading reports:', error);
    hideLoading();
  }
}
