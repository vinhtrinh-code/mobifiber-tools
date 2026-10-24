/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * FILE         : Code.gs
 * PHIÊN BẢN    : 2.9 (Tích hợp Module Báo cáo Doanh thu Phân bổ DTPK)
 * NGÀY CẬP NHẬT: 21/09/2026 14:00:00 (GMT+7)
 * TÁC GIẢ      : Trịnh Vũ Thế Vinh & AI Assistant
 * ============================================================================
 */

const CONFIG = {
  TARGET_SHEET_NAME: 'raw_mobifiber',
  LOG_SHEET_NAME: 'log_capnhat',
  CHANCAT_SHEET_NAME: 'theo_doi_chancat',
  REPORT_SHEET_NAME: 'baocao_banhang',

  HEADER_ROW_TARGET: 2,
  DATA_START_ROW_TARGET: 3,
  KEY_COLUMN_NAME: 'MÃ KHÁCH HÀNG',
  NAME_COLUMN_NAME: 'TÊN THUÊ BAO',

  CRITICAL_LOG_FIELDS: [
    'NHU CẦU HÒA MẠNG',
    'GÓI CƯỚC',
    'GIÁ CƯỚC',
    'TÌNH TRẠNG HỢP ĐỒNG',
    'NGÀY HOÀN TẤT THI CÔNG',
    'NGÀY BĂT ĐẦU GÓI CƯỚC',
    'NGÀY KẾT THÚC GÓI CƯỚC',
    'TÌNH TRẠNG CHẶN CẮT'
  ],

  CURRENCY_INTEGER_FIELDS: ['GIÁ CƯỚC', 'DTPK'],
  DATE_ONLY_FIELDS: ['NGÀY SINH'],
  DATETIME_FIELDS: [
    'NGÀY ĐẤU NỐI',
    'NGÀY HOÀN TẤT THI CÔNG',
    'NGÀY BĂT ĐẦU GÓI CƯỚC',
    'NGÀY KẾT THÚC GÓI CƯỚC'
  ]
};

// Khởi tạo Menu chính trên Google Sheets
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ MobiFiber Tools')
    // NHÓM 1: QUẢN TRỊ DỮ LIỆU THÔ
    .addItem('📥 Cập nhật dữ liệu', 'showUploadDialog')
    .addItem('📊 Sắp xếp dữ liệu', 'sortMobiFiberData')
    .addItem('⚙️ Cài đặt & Khởi tạo Tham số chuẩn hóa', 'initThamSoCaiDat')
    .addSeparator()

    // NHÓM 2: CÁC MODULE BÁO CÁO & TRÍCH XUẤT LIỀN KỀ NHAU
    .addItem('🎯 Báo cáo PTM', 'showPtmReportDialog')
    .addItem('📈 Báo cáo Doanh thu Phân bổ (DTPK)', 'showRevenueAllocationDialog')
    .addItem('⏰ Kết xuất Danh sách Gia hạn', 'showRenewalPromptDialog')
    .addItem('🛡️ Báo cáo TB Nghiệp vụ/BCA', 'showSpecialReportDialog')
    .addItem('🎯 Trích xuất Danh sách Đa điều kiện', 'showQueryDialog')
    .addItem('📊 Báo cáo Bán hàng & Doanh thu theo Kênh', 'generateSalesReport')
    .addItem('🚫 Quản lý Chặn cắt & Thu hồi ONT', 'generateDisconnectionReport')
    .addSeparator()

    // NHÓM 3: TIỆN ÍCH TẢI VỀ, DỌN RÁC & TRA CỨU
    .addItem('📥 Tải sheet hiện tại về Excel (.xlsx)', 'downloadCurrentActiveSheetClean')
    .addItem('🧹 Dọn sạch file rác tạm trên Drive', 'cleanAllOrphanedExportFiles')
    .addItem('🔍 Tra cứu thông tin thuê bao (360°)', 'showCustomerLookupDialog')
    .addToUi();
}

/**
 * TẢI SHEET HIỆN TẠI VỀ EXCEL 1-CLICK
 */
function downloadCurrentActiveSheetClean() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  if (sheetName === CONFIG.TARGET_SHEET_NAME || sheetName === CONFIG.LOG_SHEET_NAME) {
    SpreadsheetApp.getUi().alert('ℹ️ Chức năng này chỉ áp dụng cho các sheet báo cáo (PTM, Gia hạn, Đặc thù, DTPB)!');
    return;
  }

  const ssId = ss.getId();
  const sheetId = sheet.getSheetId();
  const exportUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=xlsx&gid=${sheetId}`;
  const fileName = `${sheetName}.xlsx`;

  const html = `
  <!DOCTYPE html>
  <html>
  <head>
    <base target="_top">
    <style>
      body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; padding: 20px; text-align: center; font-size: 13px; color: #202124; margin: 0; background: #fff; }
      .spinner { font-size: 28px; margin-bottom: 6px; }
      .title { font-weight: bold; color: #1a73e8; margin-bottom: 4px; }
      .sub { color: #5f6368; font-size: 11px; }
    </style>
  </head>
  <body>
    <div class="spinner">⏳</div>
    <div class="title">Đang tải: ${fileName}</div>
    <div class="sub">Tệp đang được lưu về máy, cửa sổ sẽ tự đóng...</div>
    <iframe id="dlIframe" src="${exportUrl}" style="display:none;"></iframe>

    <script>
      setTimeout(function() {
        google.script.host.close();
      }, 1200);
    </script>
  </body>
  </html>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(html).setWidth(340).setHeight(130);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Đang tải file Excel...');
}