/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * MODULE       : SẮP XẾP DỮ LIỆU PHÂN TẦNG NGHIỆP VỤ
 * PHIÊN BẢN    : 2.1 (Phân tầng Thuê bao nghiệp vụ, chuẩn hóa dd/MM/yyyy)
 * NGÀY CẬP NHẬT: 20/09/2026 19:30:00 (GMT+7)
 * TÁC GIẢ      : Trịnh Vũ Thế Vinh & AI Assistant
 * ============================================================================
 */

function sortMobiFiberData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Xác nhận', 'Sắp xếp lại toàn bộ bảng raw_mobifiber theo đúng quy tắc phân tầng nghiệp vụ?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) return;

  const headers = sheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));
  const idxLoaiKH     = headers.findIndex(h => h === 'loại khách hàng');
  const idxTinhTrang  = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxNhuCau     = headers.findIndex(h => h === 'nhu cầu hòa mạng');
  const idxMaCH       = headers.findIndex(h => h === 'mã cửa hàng');
  const idxMaNV       = headers.findIndex(h => h === 'mã nhân viên');
  const idxNgayDauNoi = headers.findIndex(h => h === 'ngày đấu nối');
  const idxNgaySinh   = headers.findIndex(h => h === 'ngày sinh');

  const contractStatusPriority = { 'đã hủy': 1, 'da huy': 1, 'chờ thi công': 2, 'cho thi cong': 2, 'đang hoạt động': 3, 'dang hoat dong': 3 };
  const nhuCauPriority = { 'đổi gói': 1, 'doi goi': 1, 'gia hạn': 2, 'gia han': 2, 'đăng ký mới': 3, 'dang ky moi': 3 };

  const numRows = lastRow - CONFIG.DATA_START_ROW_TARGET + 1;
  const dataRange = sheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, numRows, lastCol);
  const data = dataRange.getValues();

  data.sort((rowA, rowB) => {
    const isNghiepVuA = normalizeHeader(rowA[idxLoaiKH]) === 'thuê bao nghiệp vụ';
    const isNghiepVuB = normalizeHeader(rowB[idxLoaiKH]) === 'thuê bao nghiệp vụ';

    if (isNghiepVuA && !isNghiepVuB) return -1;
    if (!isNghiepVuA && isNghiepVuB) return 1;

    const ttA = contractStatusPriority[normalizeHeader(rowA[idxTinhTrang])] || 99;
    const ttB = contractStatusPriority[normalizeHeader(rowB[idxTinhTrang])] || 99;
    if (ttA !== ttB) return ttA - ttB;

    if (isNghiepVuA && isNghiepVuB) {
      const dateA = parseToDate(rowA[idxNgayDauNoi]);
      const dateB = parseToDate(rowB[idxNgayDauNoi]);
      return (dateA ? dateA.getTime() : 9999999999999) - (dateB ? dateB.getTime() : 9999999999999);
    }

    if (idxNhuCau !== -1) {
      const ncA = nhuCauPriority[normalizeHeader(rowA[idxNhuCau])] || 99;
      const ncB = nhuCauPriority[normalizeHeader(rowB[idxNhuCau])] || 99;
      if (ncA !== ncB) return ncA - ncB;
    }

    if (idxMaCH !== -1) {
      const compCH = String(rowA[idxMaCH] || '').trim().localeCompare(String(rowB[idxMaCH] || '').trim(), 'vi', { numeric: true });
      if (compCH !== 0) return compCH;
    }

    if (idxMaNV !== -1) {
      const compNV = String(rowA[idxMaNV] || '').trim().localeCompare(String(rowB[idxMaNV] || '').trim(), 'vi', { numeric: true });
      if (compNV !== 0) return compNV;
    }

    const dateA = parseToDate(rowA[idxNgayDauNoi]);
    const dateB = parseToDate(rowB[idxNgayDauNoi]);
    return (dateA ? dateA.getTime() : 9999999999999) - (dateB ? dateB.getTime() : 9999999999999);
  });

  dataRange.setValues(data);

  const keyIdx = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const cccdIdx = headers.findIndex(h => h === 'số cmnd/cccd');
  const sdtIdx = headers.findIndex(h => h === 'sđt liên hệ');
  [keyIdx, cccdIdx, sdtIdx].forEach(idx => {
    if (idx !== -1) sheet.getRange(CONFIG.DATA_START_ROW_TARGET, idx + 1, numRows, 1).setNumberFormat('@');
  });

  const giaCuocIdx = headers.findIndex(h => h === 'giá cước');
  const dtpkIdx = headers.findIndex(h => h === 'dtpk');
  [giaCuocIdx, dtpkIdx].forEach(idx => {
    if (idx !== -1) sheet.getRange(CONFIG.DATA_START_ROW_TARGET, idx + 1, numRows, 1).setNumberFormat('#,##0');
  });

  if (idxNgaySinh !== -1) {
    sheet.getRange(CONFIG.DATA_START_ROW_TARGET, idxNgaySinh + 1, numRows, 1).setNumberFormat('dd/MM/yyyy');
  }

  CONFIG.DATETIME_FIELDS.forEach(fName => {
    const dIdx = headers.findIndex(h => h === normalizeHeader(fName));
    if (dIdx !== -1) {
      sheet.getRange(CONFIG.DATA_START_ROW_TARGET, dIdx + 1, numRows, 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
    }
  });

  ui.alert(`✅ Đã sắp xếp thành công toàn bộ ${numRows} dòng dữ liệu!`);
}