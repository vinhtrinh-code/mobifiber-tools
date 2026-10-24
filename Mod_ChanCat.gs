/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * FILE        : Mod_ChanCat.gs
 * PHIÊN BẢN   : 3.0 (High-Performance Engine: In-Memory Aggregation & Batch Highlighting)
 * NGÀY TẠO    : 20/09/2026 19:30:00 (GMT+7)
 * NGÀY CẬP NHẬT: 24/10/2026 19:35:00 (GMT+7)
 * TÁC GIẢ     : Trịnh Vũ Thế Vinh & AI Assistant
 * MÔ TẢ       : Quản lý thuê bao chặn cắt, nợ cước và điều phối thu hồi thiết bị ONT
 * ============================================================================
 */

function generateDisconnectionReport() {
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) {
    SpreadsheetApp.getUi().alert(`Không tìm thấy sheet dữ liệu gốc "${CONFIG.TARGET_SHEET_NAME}"!`);
    return;
  }

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) {
    SpreadsheetApp.getUi().alert('Sheet dữ liệu chưa có bản ghi thuê bao!');
    return;
  }

  // 1. ĐỌC DỮ LIỆU ĐÍCH VÀO RAM (1 LẦN DUY NHẤT)
  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));
  const idxMaKH = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxTenTB = headers.findIndex(h => h === 'tên thuê bao');
  const idxSDT = headers.findIndex(h => h === 'sđt liên hệ');
  const idxDiaChi = headers.findIndex(h => h === 'địa chỉ lắp đặt');
  const idxGoiCuoc = headers.findIndex(h => h === 'gói cước');
  const idxGiaCuoc = headers.findIndex(h => h === 'giá cước');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxChanCat = headers.findIndex(h => h === 'tình trạng chặn cắt');
  const idxThietBi = headers.findIndex(h => h === 'loại thiết bị ont');
  const idxSerial = headers.findIndex(h => h === 'serial ont');
  const idxTenCH = headers.findIndex(h => h === 'tên cửa hàng');
  const idxTenNV = headers.findIndex(h => h === 'tên nhân viên');
  const idxNgayKetThuc = headers.findIndex(h => h === 'ngày kết thúc gói cước');

  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, lastRow - CONFIG.DATA_START_ROW_TARGET + 1, lastCol).getValues();

  // 2. LỌC & PHÂN LOẠI TRONG BỘ NHỚ
  const reportRows = [];
  const statusColors = []; // Mảng màu Matrix Formatting
  let countDebt = 0;        // Nợ cước
  let countKHYC = 0;        // KHYC tạm ngưng
  let countCancelled = 0;   // Hủy / Cần thu hồi ONT

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const statusHD = normalizeHeader(String(row[idxTinhTrangHD] || ''));
    const statusChan = normalizeHeader(String(row[idxChanCat] || ''));

    const isCancelled = statusHD.includes('hủy') || statusHD.includes('huy') || statusHD.includes('thanh lý') || statusHD.includes('chấm dứt');
    const isDebt = statusChan.includes('nợ') || statusChan.includes('no cuoc') || statusChan.includes('khóa cước');
    const isKHYC = statusChan.includes('khyc') || statusChan.includes('yêu cầu') || statusChan.includes('yeu cau');

    // Chỉ lấy các thuê bao thuộc 1 trong 3 nhóm biến động
    if (!isCancelled && !isDebt && !isKHYC) continue;

    let groupName = '';
    let rowColor = '#ffffff';

    if (isCancelled) {
      groupName = 'Hủy - Thu hồi ONT';
      rowColor = '#fce8e6'; // Đỏ nhạt
      countCancelled++;
    } else if (isDebt) {
      groupName = 'Chặn nợ cước';
      rowColor = '#fef7e0'; // Vàng nhạt
      countDebt++;
    } else if (isKHYC) {
      groupName = 'Tạm ngưng KHYC';
      rowColor = '#e8f0fe'; // Xanh dương nhạt
      countKHYC++;
    }

    let cleanMaKH = String(row[idxMaKH] || '').replace(/^'+/, '').trim();
    if (/^\d{8}$/.test(cleanMaKH)) cleanMaKH = '00' + cleanMaKH;

    const giaCuoc = parseCurrency(row[idxGiaCuoc]);
    const dKetThuc = parseToDate(row[idxNgayKetThuc]);
    const ketThucStr = dKetThuc ? Utilities.formatDate(dKetThuc, 'Asia/Saigon', 'dd/MM/yyyy') : '';

    reportRows.push([
      reportRows.length + 1,
      groupName,
      cleanMaKH,
      row[idxTenTB],
      row[idxSDT],
      row[idxDiaChi],
      row[idxGoiCuoc],
      giaCuoc,
      ketThucStr,
      row[idxThietBi],
      row[idxSerial],
      row[idxTenCH],
      row[idxTenNV]
    ]);

    statusColors.push(new Array(13).fill(rowColor));
  }

  // 3. TẠO HOẶC LÀM SẠCH SHEET THEO DÕI
  const targetSheetName = CONFIG.CHANCAT_SHEET_NAME || 'theo_doi_chancat';
  let reportSheet = ss.getSheetByName(targetSheetName);
  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet(targetSheetName);
  }

  const nowFormatted = Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm');

  // Khối Card KPI & Header
  const headerMatrix = [
    ['BÁO CÁO THEO DÕI CHẶN CẮT & THU HỒI THIẾT BỊ ONT', '', '', '', '', '', '', '', '', '', '', '', ''],
    [
      `📊 TỔNG CỘNG: ${reportRows.length} TB | 🔴 Nợ cước: ${countDebt} | 🟠 Thu hồi ONT: ${countCancelled} | 🔵 KHYC: ${countKHYC}`,
      '', '', '', '', '', '', '', '', '', '', `🕒 Lập lúc: ${nowFormatted}`, ''
    ],
    [
      'STT', 'PHÂN LOẠI', 'MÃ KHÁCH HÀNG', 'TÊN THUÊ BAO', 'SĐT LIÊN HỆ', 'ĐỊA CHỈ LẮP ĐẶT',
      'GÓI CƯỚC', 'GIÁ CƯỚC', 'HẾT HẠN GÓI', 'LOẠI ONT', 'SERIAL ONT', 'TÊN CỬA HÀNG', 'NHÂN VIÊN QUẢN LÝ'
    ]
  ];

  // Ghi khối Header 1 lần
  reportSheet.getRange(1, 1, 3, 13).setValues(headerMatrix);

  // Định dạng Header
  reportSheet.getRange('A1:M1').merge().setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  reportSheet.getRange('A2:K2').merge().setBackground('#f8f9fa').setFontColor('#202124').setFontWeight('bold');
  reportSheet.getRange('L2:M2').merge().setBackground('#f8f9fa').setFontColor('#5f6368').setHorizontalAlignment('right');
  reportSheet.getRange('A3:M3').setBackground('#34a853').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  // 4. BATCH WRITE & MATRIX FORMATTING CHI TIẾT
  if (reportRows.length > 0) {
    const dataRange = reportSheet.getRange(4, 1, reportRows.length, 13);
    dataRange.setValues(reportRows);
    dataRange.setBackgrounds(statusColors);

    // Format số & căn lề đồng loạt
    reportSheet.getRange(4, 1, reportRows.length, 1).setHorizontalAlignment('center');
    reportSheet.getRange(4, 2, reportRows.length, 1).setHorizontalAlignment('center').setFontWeight('bold');
    reportSheet.getRange(4, 3, reportRows.length, 1).setNumberFormat('@').setHorizontalAlignment('center');
    reportSheet.getRange(4, 8, reportRows.length, 1).setNumberFormat('#,##0" đ"');
    reportSheet.getRange(4, 9, reportRows.length, 1).setHorizontalAlignment('center');
    reportSheet.getRange(4, 11, reportRows.length, 1).setFontFamily('Roboto Mono').setHorizontalAlignment('center');
  }

  // Tối ưu độ rộng cột
  reportSheet.setColumnWidth(1, 50);
  reportSheet.setColumnWidth(2, 140);
  reportSheet.setColumnWidth(3, 110);
  reportSheet.setColumnWidth(4, 180);
  reportSheet.setColumnWidth(5, 100);
  reportSheet.setColumnWidth(6, 230);
  reportSheet.setColumnWidth(7, 100);
  reportSheet.setColumnWidth(8, 110);
  reportSheet.setColumnWidth(9, 100);
  reportSheet.setColumnWidth(10, 100);
  reportSheet.setColumnWidth(11, 130);
  reportSheet.setColumnWidth(12, 150);
  reportSheet.setColumnWidth(13, 140);

  const duration = ((new Date().getTime() - startTime) / 1000).toFixed(2);
  Logger.log(`⚡ Tạo báo cáo Chặn cắt & Thu hồi ONT hoàn tất trong ${duration}s!`);
  ss.setActiveSheet(reportSheet);
}