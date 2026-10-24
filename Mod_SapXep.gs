/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * FILE        : Mod_SapXep.gs
 * PHIÊN BẢN   : 3.0 (High-Performance Engine: In-Memory Multi-Tier Sorting)
 * NGÀY TẠO    : 20/09/2026 11:30:00 (GMT+7)
 * NGÀY CẬP NHẬT: 24/10/2026 20:05:00 (GMT+7)
 * TÁC GIẢ     : Trịnh Vũ Thế Vinh & AI Assistant
 * MÔ TẢ       : Sắp xếp bảng dữ liệu thô raw_mobifiber theo đa tiêu chí thời gian & mã KH trên RAM
 * ============================================================================
 */

function sortMobiFiberData() {
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  const ui = SpreadsheetApp.getUi();

  if (!rawSheet) {
    ui.alert(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);
    return;
  }

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) {
    ui.alert('Chưa có dữ liệu thuê bao để sắp xếp!');
    return;
  }

  // 1. XÁC ĐỊNH VỊ TRÍ CỘT TIÊU CHÍ TRÊN RAM
  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));
  const idxMaKH = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxNgayThiCong = headers.findIndex(h => h === 'ngày hoàn tất thi công');
  const idxNgayDauNoi = headers.findIndex(h => h === 'ngày đấu nối');
  const idxNgayBatDau = headers.findIndex(h => h === 'ngày băt đầu gói cước' || h === 'ngày bắt đầu gói cước');

  const numRows = lastRow - CONFIG.DATA_START_ROW_TARGET + 1;
  const dataRange = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, numRows, lastCol);
  const data = dataRange.getValues();

  // 2. THUẬT TOÁN SẮP XẾP ĐA TẦNG TRÊN RAM (IN-MEMORY SORT)
  data.sort((rowA, rowB) => {
    // Tiêu chí 1: Ngày hoàn tất thi công (hoặc ngày đấu nối / ngày bắt đầu)
    const dateA = (idxNgayThiCong !== -1 ? parseToDate(rowA[idxNgayThiCong]) : null) ||
                  (idxNgayDauNoi !== -1 ? parseToDate(rowA[idxNgayDauNoi]) : null) ||
                  (idxNgayBatDau !== -1 ? parseToDate(rowA[idxNgayBatDau]) : null);

    const dateB = (idxNgayThiCong !== -1 ? parseToDate(rowB[idxNgayThiCong]) : null) ||
                  (idxNgayDauNoi !== -1 ? parseToDate(rowB[idxNgayDauNoi]) : null) ||
                  (idxNgayBatDau !== -1 ? parseToDate(rowB[idxNgayBatDau]) : null);

    const timeA = dateA ? dateA.getTime() : 0;
    const timeB = dateB ? dateB.getTime() : 0;

    // Sắp xếp ngày giảm dần (Mới nhất lên trên)
    if (timeA !== timeB) {
      return timeB - timeA;
    }

    // Tiêu chí 2: Nếu cùng ngày thì sắp xếp theo Mã khách hàng tăng dần
    if (idxMaKH !== -1) {
      const codeA = String(rowA[idxMaKH] || '').replace(/[^0-9]/g, '');
      const codeB = String(rowB[idxMaKH] || '').replace(/[^0-9]/g, '');
      return codeA.localeCompare(codeB, undefined, { numeric: true });
    }

    return 0;
  });

  // 3. GHI ĐÈ LẠI DỮ LIỆU ĐÃ SẮP XẾP (1 LẦN DUY NHẤT)
  dataRange.setValues(data);

  const duration = ((new Date().getTime() - startTime) / 1000).toFixed(2);
  Logger.log(`⚡ Sắp xếp ${numRows} dòng dữ liệu raw_mobifiber hoàn tất trong ${duration}s!`);
  ui.alert(`✅ Đã sắp xếp thành công ${numRows} thuê bao theo thời gian phát triển mới nhất (Thời gian xử lý: ${duration}s)!`);
}