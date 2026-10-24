/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * MODULE       : BÁO CÁO HIỆU QUẢ BÁN HÀNG & DOANH THU THEO KÊNH
 * PHIÊN BẢN    : 2.1 (Tải Excel sạch, Tự động dọn rác, Phân tích Kênh / Nhu cầu)
 * NGÀY CẬP NHẬT: 20/09/2026 19:30:00 (GMT+7)
 * TÁC GIẢ      : Trịnh Vũ Thế Vinh & AI Assistant
 * ============================================================================
 */

function generateSalesReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  const ui = SpreadsheetApp.getUi();
  if (!rawSheet) return;

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) return;

  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));
  const idxTenCH   = headers.findIndex(h => h === 'tên cửa hàng');
  const idxNhuCau  = headers.findIndex(h => h === 'nhu cầu hòa mạng');
  const idxDtpk    = headers.findIndex(h => h === 'dtpk');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');

  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, lastRow - CONFIG.DATA_START_ROW_TARGET + 1, lastCol).getValues();
  const storeMap = {}, nhuCauMap = {};
  let totalActiveTB = 0, totalDtpk = 0;

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const status = normalizeHeader(row[idxTinhTrangHD]);
    if (status === 'đã hủy' || status === 'da huy') continue;

    totalActiveTB++;
    const dtpk = parseCurrencyToInteger(row[idxDtpk]) || 0;
    totalDtpk += dtpk;

    const ch = String(row[idxTenCH] || 'Chưa phân loại').trim();
    if (!storeMap[ch]) storeMap[ch] = { count: 0, dtpk: 0 };
    storeMap[ch].count++; storeMap[ch].dtpk += dtpk;

    const nc = String(row[idxNhuCau] || 'Khác').trim();
    if (!nhuCauMap[nc]) nhuCauMap[nc] = { count: 0, dtpk: 0 };
    nhuCauMap[nc].count++; nhuCauMap[nc].dtpk += dtpk;
  }

  let rSheet = ss.getSheetByName(CONFIG.REPORT_SHEET_NAME);
  if (!rSheet) rSheet = ss.insertSheet(CONFIG.REPORT_SHEET_NAME);
  else rSheet.clear();

  const ssId = ss.getId();
  const sheetId = rSheet.getSheetId();
  const directExportUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=xlsx&gid=${sheetId}`;

  rSheet.getRange('A1:J1').merge().setValue('BÁO CÁO TỔNG HỢP HIỆU QUẢ PHÁT TRIỂN & DOANH THU MOBIFIBER')
    .setFontWeight('bold').setFontSize(13).setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  
  rSheet.getRange('A2:H2').merge().setValue(`📊 Tổng thuê bao đang hoạt động: ${totalActiveTB} TB   |   💰 Tổng Doanh thu Phân kỳ (DTPK): ${totalDtpk.toLocaleString('vi-VN')} VNĐ`)
    .setFontWeight('bold').setFontSize(11).setBackground('#D2E3FC').setFontColor('#174EA6').setHorizontalAlignment('center');

  rSheet.getRange('I2:J2').merge()
    .setFormula(`=HYPERLINK("${directExportUrl}"; "📥 TẢI FILE EXCEL")`)
    .setFontWeight('bold').setFontSize(11).setBackground('#1A73E8').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  rSheet.getRange('A4:D4').merge().setValue('1. HIỆU QUẢ THEO CỬA HÀNG / KÊNH BÁN').setFontWeight('bold').setBackground('#E8F0FE').setFontColor('#174EA6');
  rSheet.getRange('A5:D5').setValues([['TÊN CỬA HÀNG', 'SỐ THUÊ BAO', 'DOANH THU (DTPK)', 'TỶ TRỌNG']]).setFontWeight('bold').setBackground('#1B365D').setFontColor('#FFFFFF');

  const storeRows = Object.keys(storeMap).map(k => [k, storeMap[k].count, storeMap[k].dtpk, totalDtpk > 0 ? (storeMap[k].dtpk / totalDtpk) : 0]).sort((a, b) => b[1] - a[1]);
  if (storeRows.length > 0) {
    rSheet.getRange(6, 1, storeRows.length, 4).setValues(storeRows);
    rSheet.getRange(6, 2, storeRows.length, 1).setHorizontalAlignment('center');
    rSheet.getRange(6, 3, storeRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');
    rSheet.getRange(6, 4, storeRows.length, 1).setNumberFormat('0.0%').setHorizontalAlignment('right');
  }

  rSheet.getRange('F4:I4').merge().setValue('2. CƠ CẤU PHÁT TRIỂN (NHU CẦU)').setFontWeight('bold').setBackground('#E8F0FE').setFontColor('#174EA6');
  rSheet.getRange('F5:I5').setValues([['NHU CẦU HÒA MẠNG', 'SỐ THUÊ BAO', 'DOANH THU (DTPK)', 'TỶ TRỌNG']]).setFontWeight('bold').setBackground('#00558F').setFontColor('#FFFFFF');

  const nhuCauRows = Object.keys(nhuCauMap).map(k => [k, nhuCauMap[k].count, nhuCauMap[k].dtpk, totalDtpk > 0 ? (nhuCauMap[k].dtpk / totalDtpk) : 0]).sort((a, b) => b[1] - a[1]);
  if (nhuCauRows.length > 0) {
    rSheet.getRange(6, 6, nhuCauRows.length, 4).setValues(nhuCauRows);
    rSheet.getRange(6, 7, nhuCauRows.length, 1).setHorizontalAlignment('center');
    rSheet.getRange(6, 8, nhuCauRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');
    rSheet.getRange(6, 9, nhuCauRows.length, 1).setNumberFormat('0.0%').setHorizontalAlignment('right');
  }

  rSheet.autoResizeColumns(1, 10);
  ss.setActiveSheet(rSheet);
  ui.alert('✅ Báo cáo Bán hàng & Doanh thu theo Kênh đã được cập nhật thành công!');
}