/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * FILE        : Mod_BaoCaoKenh.gs
 * PHIÊN BẢN   : 3.0 (High-Performance Engine: Dual In-Memory Aggregation & Matrix Formatting)
 * NGÀY TẠO    : 19/09/2026 14:00:00 (GMT+7)
 * NGÀY CẬP NHẬT: 24/10/2026 19:40:00 (GMT+7)
 * TÁC GIẢ     : Trịnh Vũ Thế Vinh & AI Assistant
 * MÔ TẢ       : Tổng hợp báo cáo hiệu quả bán hàng, doanh thu và cơ cấu nhu cầu theo Kênh/Cửa hàng
 * ============================================================================
 */

function generateSalesReport() {
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
    ui.alert('Sheet raw_mobifiber chưa có dữ liệu thuê bao!');
    return;
  }

  // 1. ĐỌC DỮ LIỆU VÀ HEADERS VÀO RAM (1 LẦN DUY NHẤT)
  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));
  const idxGiaCuoc = headers.findIndex(h => h === 'giá cước');
  const idxChuKy = headers.findIndex(h => h === 'chu kỳ sử dụng');
  const idxDtpk = headers.findIndex(h => h === 'dtpk');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxTenCH = headers.findIndex(h => h === 'tên cửa hàng');
  const idxMaCH = headers.findIndex(h => h === 'mã cửa hàng');
  const idxNhuCau = headers.findIndex(h => h === 'nhu cầu hòa mạng');

  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, lastRow - CONFIG.DATA_START_ROW_TARGET + 1, lastCol).getValues();

  // 2. DUAL IN-MEMORY AGGREGATION
  const channelMap = new Map(); // Key: Tên Cửa Hàng -> { maCH, count, revenue, dtpk }
  const demandMap = new Map();  // Key: Nhu Cầu Hòa Mạng -> { count, dtpk }

  let grandTotalSub = 0;
  let grandTotalRev = 0;
  let grandTotalDtpk = 0;

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const statusHD = normalizeHeader(String(row[idxTinhTrangHD] || ''));

    // Bỏ qua hợp đồng đã hủy hoặc thanh lý
    if (statusHD.includes('hủy') || statusHD.includes('huy') || statusHD.includes('thanh lý') || statusHD.includes('chấm dứt')) {
      continue;
    }

    const giaCuoc = parseCurrency(row[idxGiaCuoc]);
    const chuKy = parseInt(row[idxChuKy], 10) || 1;
    let dtpk = parseCurrency(row[idxDtpk]);
    if (!dtpk && giaCuoc > 0) dtpk = Math.round(giaCuoc / chuKy);

    grandTotalSub++;
    grandTotalRev += giaCuoc;
    grandTotalDtpk += dtpk;

    // Gom nhóm theo Kênh / Cửa hàng
    const tenCH = String(row[idxTenCH] || 'Chưa xác định kênh').trim();
    const maCH = String(row[idxMaCH] || '').trim();

    if (!channelMap.has(tenCH)) {
      channelMap.set(tenCH, { maCH: maCH, count: 0, revenue: 0, dtpk: 0 });
    }
    const cStat = channelMap.get(tenCH);
    cStat.count++;
    cStat.revenue += giaCuoc;
    cStat.dtpk += dtpk;

    // Gom nhóm theo Nhu cầu hòa mạng
    let nhuCau = String(row[idxNhuCau] || 'Khác').trim();
    if (!nhuCau) nhuCau = 'Khác';

    if (!demandMap.has(nhuCau)) {
      demandMap.set(nhuCau, { count: 0, dtpk: 0 });
    }
    const dStat = demandMap.get(nhuCau);
    dStat.count++;
    dStat.dtpk += dtpk;
  }

  // 3. TẠO HOẶC XÓA SHEET BÁO CÁO
  const targetReportName = CONFIG.REPORT_SHEET_NAME || 'baocao_banhang';
  let rSheet = ss.getSheetByName(targetReportName);
  if (rSheet) {
    rSheet.clear();
  } else {
    rSheet = ss.insertSheet(targetReportName);
  }

  const nowStr = Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm');

  // Khối Tiêu đề chung & Card KPI
  const banner = [
    ['BÁO CÁO HIỆU QUẢ BÁN HÀNG & DOANH THU THEO KÊNH PHÂN PHỐI', '', '', '', '', '', '', '', ''],
    [
      `🎯 TỔNG THUÊ BAO: ${grandTotalSub.toLocaleString('vi-VN')} TB | DOANH THU: ${grandTotalRev.toLocaleString('vi-VN')} đ | DTPK: ${grandTotalDtpk.toLocaleString('vi-VN')} đ`,
      '', '', '', '', '', '', `🕒 Lập lúc: ${nowStr}`, ''
    ]
  ];
  rSheet.getRange(1, 1, 2, 9).setValues(banner);
  rSheet.getRange('A1:I1').merge().setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  rSheet.getRange('A2:G2').merge().setBackground('#e8f0fe').setFontColor('#174ea6').setFontWeight('bold');
  rSheet.getRange('H2:I2').merge().setBackground('#e8f0fe').setFontColor('#5f6368').setHorizontalAlignment('right');

  // 4. XÂY DỰNG BẢNG 1: HIỆU QUẢ THEO KÊNH (Cột A -> D)
  const b1Header = ['KÊNH / CỬA HÀNG', 'SỐ THUÊ BAO', 'DOANH THU (DTPK)', 'TỶ TRỌNG (%)'];
  const b1Data = [];

  // Sắp xếp Kênh giảm dần theo DTPK
  const sortedChannels = Array.from(channelMap.entries()).sort((a, b) => b[1].dtpk - a[1].dtpk);
  sortedChannels.forEach(([name, val]) => {
    const ratio = grandTotalDtpk > 0 ? (val.dtpk / grandTotalDtpk) : 0;
    b1Data.push([name, val.count, val.dtpk, ratio]);
  });
  b1Data.push(['TỔNG CỘNG', grandTotalSub, grandTotalDtpk, grandTotalDtpk > 0 ? 1 : 0]);

  rSheet.getRange(4, 1, 1, 4).setValues([b1Header]).setBackground('#34a853').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  rSheet.getRange(5, 1, b1Data.length, 4).setValues(b1Data);

  // Format Bảng 1
  rSheet.getRange(5, 2, b1Data.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('center');
  rSheet.getRange(5, 3, b1Data.length, 1).setNumberFormat('#,##0" đ"').setHorizontalAlignment('right');
  rSheet.getRange(5, 4, b1Data.length, 1).setNumberFormat('0.00%').setHorizontalAlignment('right');
  rSheet.getRange(4 + b1Data.length, 1, 1, 4).setBackground('#ea4335').setFontColor('#ffffff').setFontWeight('bold');

  // 5. XÂY DỰNG BẢNG 2: CƠ CẤU NHU CẦU HÒA MẠNG (Cột F -> I)
  const b2Header = ['NHU CẦU HÒA MẠNG', 'SỐ THUÊ BAO', 'DOANH THU (DTPK)', 'TỶ TRỌNG (%)'];
  const b2Data = [];

  // Sắp xếp Nhu cầu giảm dần theo Số lượng
  const sortedDemands = Array.from(demandMap.entries()).sort((a, b) => b[1].count - a[1].count);
  sortedDemands.forEach(([demand, val]) => {
    const ratio = grandTotalDtpk > 0 ? (val.dtpk / grandTotalDtpk) : 0;
    b2Data.push([demand, val.count, val.dtpk, ratio]);
  });
  b2Data.push(['TỔNG CỘNG', grandTotalSub, grandTotalDtpk, grandTotalDtpk > 0 ? 1 : 0]);

  rSheet.getRange(4, 6, 1, 4).setValues([b2Header]).setBackground('#00558f').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  rSheet.getRange(5, 6, b2Data.length, 4).setValues(b2Data);

  // Format Bảng 2
  rSheet.getRange(5, 7, b2Data.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('center');
  rSheet.getRange(5, 8, b2Data.length, 1).setNumberFormat('#,##0" đ"').setHorizontalAlignment('right');
  rSheet.getRange(5, 9, b2Data.length, 1).setNumberFormat('0.00%').setHorizontalAlignment('right');
  rSheet.getRange(4 + b2Data.length, 6, 1, 4).setBackground('#ea4335').setFontColor('#ffffff').setFontWeight('bold');

  // Căn chỉnh độ rộng các cột
  rSheet.setColumnWidth(1, 190);
  rSheet.setColumnWidth(2, 110);
  rSheet.setColumnWidth(3, 140);
  rSheet.setColumnWidth(4, 110);
  rSheet.setColumnWidth(5, 40); // Cột đệm giữa 2 bảng
  rSheet.setColumnWidth(6, 180);
  rSheet.setColumnWidth(7, 110);
  rSheet.setColumnWidth(8, 140);
  rSheet.setColumnWidth(9, 110);

  const duration = ((new Date().getTime() - startTime) / 1000).toFixed(2);
  Logger.log(`⚡ Tạo Báo cáo Kênh hoàn tất trong ${duration}s!`);
  ss.setActiveSheet(rSheet);
}