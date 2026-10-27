/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * MODULE       : Mod_SapXep.gs
 * PHIÊN BẢN    : 2.0 (Sắp xếp dữ liệu phân tầng 4 cấp trên RAM & Bảo toàn định dạng)
 * TÁC GIẢ      : Trịnh Vũ Thế Vinh & AI Assistant
 * ============================================================================
 */

function sortMobiFiberData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  const ui = SpreadsheetApp.getUi();

  if (!rawSheet) {
    ui.alert(`⚠️ Không tìm thấy sheet dữ liệu "${CONFIG.TARGET_SHEET_NAME}"!`);
    return;
  }

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();

  if (lastRow < CONFIG.DATA_START_ROW_TARGET) {
    ui.alert(`ℹ️ Sheet "${CONFIG.TARGET_SHEET_NAME}" chưa có dữ liệu để sắp xếp.`);
    return;
  }

  // 1. Đọc và chuẩn hóa tiêu đề cột
  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));
  
  const idxMaKH = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxMaCH = headers.findIndex(h => h === 'mã cửa hàng');
  const idxMaNV = headers.findIndex(h => h === 'mã nhân viên');
  const idxNgayDauNoi = headers.findIndex(h => h === 'ngày đấu nối');
  const idxNgayBatDau = headers.findIndex(h => h === 'ngày băt đầu gói cước' || h === 'ngày bắt đầu gói cước');
  const idxSDT = headers.findIndex(h => h === 'sđt liên hệ');
  const idxCCCD = headers.findIndex(h => h === 'số cmnd/cccd');
  const idxGiaCuoc = headers.findIndex(h => h === 'giá cước');
  const idxDtpk = headers.findIndex(h => h === 'dtpk');

  const numRows = lastRow - CONFIG.DATA_START_ROW_TARGET + 1;
  const rawRange = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, numRows, lastCol);
  const rawValues = rawRange.getValues();

  // 2. Thuật toán sắp xếp 4 cấp ưu tiên trên RAM
  rawValues.sort((rowA, rowB) => {
    // CẤP 1: TÌNH TRẠNG HỢP ĐỒNG (Đang hoạt động lên trước, Đã hủy xuống sau)
    const statusA = normalizeHeader(String(rowA[idxTinhTrangHD] || ''));
    const statusB = normalizeHeader(String(rowB[idxTinhTrangHD] || ''));
    
    const isCancelledA = statusA.includes('hủy') || statusA.includes('huy') || statusA.includes('thanh lý') || statusA.includes('chấm dứt');
    const isCancelledB = statusB.includes('hủy') || statusB.includes('huy') || statusB.includes('thanh lý') || statusB.includes('chấm dứt');

    if (isCancelledA !== isCancelledB) {
      return isCancelledA ? 1 : -1; // Đang hoạt động ưu tiên xếp trước
    }

    // CẤP 2: MÃ CỬA HÀNG (A-Z)
    const maCHA = String(rowA[idxMaCH] || '').trim();
    const maCHB = String(rowB[idxMaCH] || '').trim();
    const compCH = maCHA.localeCompare(maCHB, 'vi', { numeric: true });
    if (compCH !== 0) return compCH;

    // CẤP 3: MÃ NHÂN VIÊN (A-Z)
    const maNVA = String(rowA[idxMaNV] || '').trim();
    const maNVB = String(rowB[idxMaNV] || '').trim();
    const compNV = maNVA.localeCompare(maNVB, 'vi', { numeric: true });
    if (compNV !== 0) return compNV;

    // CẤP 4: THỜI GIAN (Ngày đấu nối / Bắt đầu: Từ Mới nhất -> Cũ nhất)
    const dateA = parseToDate(rowA[idxNgayDauNoi]) || parseToDate(rowA[idxNgayBatDau]) || new Date(0);
    const dateB = parseToDate(rowB[idxNgayDauNoi]) || parseToDate(rowB[idxNgayBatDau]) || new Date(0);
    
    return dateB.getTime() - dateA.getTime();
  });

  // 3. Chuẩn hóa bảo toàn định dạng chuỗi cho Mã KH, SĐT, CCCD
  for (let r = 0; r < rawValues.length; r++) {
    if (idxMaKH !== -1) {
      let code = String(rawValues[r][idxMaKH] || '').replace(/[^0-9]/g, '').trim();
      if (code.length === 8) code = '00' + code;
      else if (code.length > 0 && code.length < 10) code = code.padStart(10, '0');
      if (code) rawValues[r][idxMaKH] = "'" + code;
    }
    if (idxSDT !== -1) {
      let sdt = String(rawValues[r][idxSDT] || '').replace(/[^0-9]/g, '').trim();
      if (sdt.length === 9) sdt = '0' + sdt;
      if (sdt) rawValues[r][idxSDT] = "'" + sdt;
    }
    if (idxCCCD !== -1) {
      let cccd = String(rawValues[r][idxCCCD] || '').replace(/[^0-9]/g, '').trim();
      if (cccd.length === 11) cccd = '0' + cccd;
      if (cccd) rawValues[r][idxCCCD] = "'" + cccd;
    }
  }

  // 4. Ghi đè dữ liệu đã sắp xếp chuẩn về bảng tính
  rawRange.setValues(rawValues);

  // 5. Khóa định dạng số nguyên cho các cột tiền tệ
  if (idxGiaCuoc !== -1) rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, idxGiaCuoc + 1, numRows, 1).setNumberFormat('#,##0');
  if (idxDtpk !== -1) rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, idxDtpk + 1, numRows, 1).setNumberFormat('#,##0');
  if (idxMaKH !== -1) rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, idxMaKH + 1, numRows, 1).setNumberFormat('@');
  if (idxSDT !== -1) rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, idxSDT + 1, numRows, 1).setNumberFormat('@');

  ui.alert(`✅ Đã sắp xếp chuẩn hóa thành công ${numRows} thuê bao trong bảng "${CONFIG.TARGET_SHEET_NAME}" theo 4 cấp ưu tiên!`);
}