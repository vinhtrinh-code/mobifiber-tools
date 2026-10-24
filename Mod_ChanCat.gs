/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * MODULE       : QUẢN LÝ CHẶN CẮT, NỢ CƯỚC & THU HỒI ONT
 * PHIÊN BẢN    : 2.1 (Tải Excel sạch, Tự động dọn rác, Phân loại rõ nợ cước / KHYC / Thu hồi ONT)
 * NGÀY CẬP NHẬT: 20/09/2026 19:30:00 (GMT+7)
 * TÁC GIẢ      : Trịnh Vũ Thế Vinh & AI Assistant
 * ============================================================================
 */

function generateDisconnectionReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  const ui = SpreadsheetApp.getUi();
  if (!rawSheet) return;

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) return;

  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));
  const idxMaKH        = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxTenTB       = headers.findIndex(h => h === 'tên thuê bao');
  const idxSDT         = headers.findIndex(h => h === 'sđt liên hệ');
  const idxDiaChi      = headers.findIndex(h => h === 'địa chỉ lắp đặt');
  const idxGoiCuoc     = headers.findIndex(h => h === 'gói cước');
  const idxGiaCuoc     = headers.findIndex(h => h === 'giá cước');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxChanCat     = headers.findIndex(h => h === 'tình trạng chặn cắt');
  const idxThietBi     = headers.findIndex(h => h === 'loại thiết bị ont');
  const idxSerial      = headers.findIndex(h => h === 'serial ont');
  const idxTenCH       = headers.findIndex(h => h === 'tên cửa hàng');
  const idxTenNV       = headers.findIndex(h => h === 'tên nhân viên');
  const idxNgayKetThuc = headers.findIndex(h => h === 'ngày kết thúc gói cước');

  const numRows = lastRow - CONFIG.DATA_START_ROW_TARGET + 1;
  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, numRows, lastCol).getValues();

  const reportRows = [];
  let countDebt = 0, countKHYC = 0, countCancelled = 0;

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const chanCat = String(row[idxChanCat] || '').trim();
    const tinhTrangHD = String(row[idxTinhTrangHD] || '').trim();
    const isChanCat = chanCat && normalizeHeader(chanCat) !== 'bình thường';
    const isCancelled = normalizeHeader(tinhTrangHD) === 'đã hủy' || normalizeHeader(tinhTrangHD) === 'da huy';
    if (!isChanCat && !isCancelled) continue;

    let phanLoai = '', bgColor = '#ffffff';
    if (isCancelled) {
      phanLoai = '🔴 Hợp đồng đã hủy (Cần thu hồi ONT)'; bgColor = '#FCE8E6'; countCancelled++;
    } else if (normalizeHeader(chanCat).includes('nợ cước') || normalizeHeader(chanCat).includes('no cuoc')) {
      phanLoai = '🟠 Chặn do nợ cước (Cần đôn đốc)'; bgColor = '#FEF7E0'; countDebt++;
    } else {
      phanLoai = '🟡 Chặn theo yêu cầu KHYC'; bgColor = '#FEF9E7'; countKHYC++;
    }

    let cleanMaKH = String(row[idxMaKH] || '').replace(/^'+/, '').trim();
    if (/^\d{8}$/.test(cleanMaKH)) cleanMaKH = '00' + cleanMaKH;
    let cleanSDT = String(row[idxSDT] || '').replace(/^'+/, '').trim();
    if (/^\d{9}$/.test(cleanSDT)) cleanSDT = '0' + cleanSDT;

    reportRows.push({
      bgColor,
      values: [
        phanLoai, "'" + cleanMaKH, '  ' + (row[idxTenTB] || ''), "'" + cleanSDT, '  ' + (row[idxDiaChi] || ''),
        row[idxGoiCuoc] || '', parseCurrencyToInteger(row[idxGiaCuoc]), formatDateTimeStr(row[idxNgayKetThuc]),
        chanCat || 'Bình thường', tinhTrangHD, '  ' + (row[idxThietBi] || ''), "'" + String(row[idxSerial] || '').replace(/^'+/, ''),
        '  ' + (row[idxTenCH] || ''), '  ' + (row[idxTenNV] || '')
      ]
    });
  }

  let destSheet = ss.getSheetByName(CONFIG.CHANCAT_SHEET_NAME);
  if (!destSheet) destSheet = ss.insertSheet(CONFIG.CHANCAT_SHEET_NAME);
  else destSheet.clear();

  const ssId = ss.getId();
  const sheetId = destSheet.getSheetId();
  const directExportUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=xlsx&gid=${sheetId}`;

  destSheet.getRange(1, 1, 1, 14).merge().setValue('DANH SÁCH THUÊ BAO CHẶN CẮT, NỢ CƯỚC & THU HỒI THIẾT BỊ ONT')
    .setFontWeight('bold').setFontSize(13).setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  
  destSheet.getRange(2, 1, 1, 12).merge().setValue(`📊 Tổng cộng: ${reportRows.length} TB   |   🟠 Nợ cước: ${countDebt} TB   |   🟡 Chặn KHYC: ${countKHYC} TB   |   🔴 Đã hủy: ${countCancelled} TB`)
    .setFontWeight('bold').setFontSize(11).setBackground('#E8F0FE').setFontColor('#174EA6').setHorizontalAlignment('center');

  // Nút tải Excel trên sheet
  destSheet.getRange(2, 13, 1, 2).merge()
    .setFormula(`=HYPERLINK("${directExportUrl}"; "📥 TẢI FILE EXCEL")`)
    .setFontWeight('bold').setFontSize(11).setBackground('#1A73E8').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  const headersOut = ['Phân loại theo dõi', 'MÃ KHÁCH HÀNG', 'TÊN THUÊ BAO', 'SĐT LIÊN HỆ', 'ĐỊA CHỈ LẮP ĐẶT', 'GÓI CƯỚC', 'GIÁ CƯỚC', 'HẠN GÓI CƯỚC', 'TÌNH TRẠNG CHẶN CẮT', 'TÌNH TRẠNG HĐ', 'LOẠI THIẾT BỊ ONT', 'SERIAL ONT', 'TÊN CỬA HÀNG', 'TÊN NHÂN VIÊN'];
  destSheet.getRange(4, 1, 1, headersOut.length).setValues([headersOut]).setFontWeight('bold').setBackground('#C5221F').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  destSheet.setFrozenRows(4);

  if (reportRows.length > 0) {
    const outData = reportRows.map(r => r.values);
    destSheet.getRange(5, 1, outData.length, headersOut.length).setValues(outData);
    for (let r = 0; r < reportRows.length; r++) destSheet.getRange(5 + r, 1).setBackground(reportRows[r].bgColor).setFontWeight('bold');
    destSheet.getRange(5, 2, outData.length, 1).setNumberFormat('@').setHorizontalAlignment('center');
    destSheet.getRange(5, 4, outData.length, 1).setNumberFormat('@').setHorizontalAlignment('center');
    destSheet.getRange(5, 7, outData.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');
    destSheet.getRange(5, 8, outData.length, 3).setHorizontalAlignment('center');
    destSheet.getRange(5, 12, outData.length, 1).setNumberFormat('@').setHorizontalAlignment('center');

    const exactWidths = [190, 120, 180, 115, 260, 110, 110, 155, 150, 130, 130, 130, 220, 160];
    for (let c = 0; c < exactWidths.length; c++) destSheet.setColumnWidth(c + 1, exactWidths[c]);
  }

  ss.setActiveSheet(destSheet);
  ui.alert(`✅ Đã trích xuất ${reportRows.length} thuê bao cần xử lý vào sheet "${CONFIG.CHANCAT_SHEET_NAME}"!`);
}