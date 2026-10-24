/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * MODULE       : BÁO CÁO THUÊ BAO PHÁT TRIỂN MỚI (PTM)
 * PHIÊN BẢN    : 2.6 (Thuật toán sắp xếp phân tầng 4 cấp: Khu vực -> Mã CH -> Mã NV -> Ngày cũ đến mới)
 * NGÀY CẬP NHẬT: 21/09/2026 10:30:00 (GMT+7)
 * TÁC GIẢ      : Trịnh Vũ Thế Vinh & AI Assistant
 * MÔ TẢ        : 
 *   - Bảng 1: Cơ cấu PTM theo KHU VỰC (Cột TỶ TRỌNG DTPK vừa khít tiêu đề + 2px padding = 124px)
 *   - Bảng 2: Danh sách chi tiết 18 cột chuẩn (có cột KHU VỰC, Tên CH rút gọn, Ngày bắt đầu rút gọn)
 *   - Thuật toán sắp xếp chi tiết 4 cấp:
 *       1. KHU VỰC (A-Z)
 *       2. MÃ CỬA HÀNG (tăng dần)
 *       3. MÃ NHÂN VIÊN (tăng dần)
 *       4. DÒNG THỜI GIAN (từ cũ nhất đến mới nhất)
 *   - Đánh số lại cột STT tự động từ 1 đến hết
 *   - Tự động hủy file tạm trên Drive sau 5 giây chống rác bộ nhớ
 * ============================================================================
 */

function showPtmReportDialog() {
  const html = HtmlService.createHtmlOutput(getPtmDialogHtml())
    .setWidth(490)
    .setHeight(430)
    .setTitle('Báo cáo PTM');
  SpreadsheetApp.getUi().showModalDialog(html, 'Báo cáo Thuê bao Phát triển mới (PTM)');
}

/**
 * Xử lý quét dữ liệu PTM và tạo bảng báo cáo chuẩn nghiệp vụ
 */
function processPtmReport(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) throw new Error('Bảng raw_mobifiber chưa có dữ liệu.');

  // 1. Nạp bản đồ tham số chuẩn hóa từ thamso_caidat
  const mapping = loadThamSoMapping();

  // 2. Phân tích khoảng thời gian lọc
  let startDate, endDate, sheetTitleSuffix, sheetName;
  if (params.type === 'month') {
    const month = parseInt(params.month, 10);
    const year = parseInt(params.year, 10);
    const mStr = month < 10 ? '0' + month : '' + month;
    startDate = new Date(year, month - 1, 1, 0, 0, 0);
    endDate = new Date(year, month, 0, 23, 59, 59);
    sheetName = `PTM_T${mStr}_${year}`;
    sheetTitleSuffix = `THÁNG ${mStr}/${year}`;
  } else {
    const startParts = params.from.split('/');
    const endParts = params.to.split('/');
    startDate = new Date(parseInt(startParts[2], 10), parseInt(startParts[1], 10) - 1, parseInt(startParts[0], 10), 0, 0, 0);
    endDate = new Date(parseInt(endParts[2], 10), parseInt(endParts[1], 10) - 1, parseInt(endParts[0], 10), 23, 59, 59);
    const d1 = startParts[0] + startParts[1];
    const d2 = endParts[0] + endParts[1];
    sheetName = `PTM_${d1}_${d2}_${endParts[2]}`;
    sheetTitleSuffix = `TỪ ${params.from} ĐẾN ${params.to}`;
  }

  // 3. Đọc Headers
  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol)
    .getValues()[0].map(h => normalizeHeader(h));

  const idxMaKH        = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxTenTB       = headers.findIndex(h => h === 'tên thuê bao');
  const idxSDT         = headers.findIndex(h => h === 'sđt liên hệ');
  const idxDiaChi      = headers.findIndex(h => h === 'địa chỉ lắp đặt');
  const idxGoiCuoc     = headers.findIndex(h => h === 'gói cước');
  const idxGiaCuoc     = headers.findIndex(h => h === 'giá cước');
  const idxChuKy       = headers.findIndex(h => h === 'chu kỳ sử dụng');
  const idxDtpk        = headers.findIndex(h => h === 'dtpk');
  const idxNgayDauNoi  = headers.findIndex(h => h === 'ngày đấu nối');
  const idxNgayThiCong = headers.findIndex(h => h === 'ngày hoàn tất thi công');
  const idxNgayBatDau  = headers.findIndex(h => h === 'ngày băt đầu gói cước' || h === 'ngày bắt đầu gói cước');
  const idxThietBi     = headers.findIndex(h => h === 'loại thiết bị ont');
  const idxSerial      = headers.findIndex(h => h === 'serial ont');
  const idxLoaiKH      = headers.findIndex(h => h === 'loại khách hàng');
  const idxNhuCau      = headers.findIndex(h => h === 'nhu cầu hòa mạng');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxMaCH        = headers.findIndex(h => h === 'mã cửa hàng');
  const idxTenCH       = headers.findIndex(h => h === 'tên cửa hàng');
  const idxMaNV        = headers.findIndex(h => h === 'mã nhân viên');
  const idxTenNV       = headers.findIndex(h => h === 'tên nhân viên');

  const numRows = lastRow - CONFIG.DATA_START_ROW_TARGET + 1;
  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, numRows, lastCol).getValues();

  const ptmRowsData = [];
  const regionSummaryMap = {}; // Thống kê gom nhóm theo KHU VỰC
  let totalRevenue = 0;
  let totalDtpk = 0;

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];

    // Tiêu chí 1: NHU CẦU HÒA MẠNG = "Đăng ký mới"
    const nhuCau = normalizeHeader(row[idxNhuCau]);
    if (nhuCau !== 'đăng ký mới' && nhuCau !== 'dang ky moi') continue;

    // Tiêu chí 3: Loại trừ "Thuê bao nghiệp vụ"
    const loaiKH = normalizeHeader(row[idxLoaiKH]);
    if (loaiKH === 'thuê bao nghiệp vụ' || loaiKH === 'thue bao nghiep vu') continue;

    // Tiêu chí 4: TÌNH TRẠNG HỢP ĐỒNG = "Đang hoạt động"
    const tinhTrang = normalizeHeader(row[idxTinhTrangHD]);
    if (tinhTrang !== 'đang hoạt động' && tinhTrang !== 'dang hoat dong') continue;

    const chuKyRaw = row[idxChuKy];
    const chuKy = chuKyRaw ? parseInt(chuKyRaw, 10) : 0;

    // Tiêu chí 2: XÁC ĐỊNH NGÀY GHI NHẬN PTM
    let ptmDate = null;
    const dDauNoi = parseToDate(row[idxNgayDauNoi]);
    const dBatDau = parseToDate(row[idxNgayBatDau]);
    const dThiCong = parseToDate(row[idxNgayThiCong]);

    if (chuKy === 1) {
      if (dDauNoi && dThiCong) {
        ptmDate = (dThiCong.getTime() > dDauNoi.getTime()) ? dThiCong : dDauNoi;
      } else {
        ptmDate = dThiCong || dDauNoi || dBatDau;
      }
    } else {
      ptmDate = dBatDau;
    }

    if (!ptmDate) continue;
    if (ptmDate < startDate || ptmDate > endDate) continue;

    const giaCuoc = parseCurrencyToInteger(row[idxGiaCuoc]) || 0;
    const dtpk = parseCurrencyToInteger(row[idxDtpk]) || (chuKy > 0 ? Math.round(giaCuoc / chuKy) : 0);
    totalRevenue += giaCuoc;
    totalDtpk += dtpk;

    // Xác định Khu vực từ thamso_caidat
    const rawMaNV = String(row[idxMaNV] || '').trim();
    let khuVuc = 'Chưa phân bổ KV';
    if (rawMaNV && mapping.staffRegion && mapping.staffRegion.has(normalizeHeader(rawMaNV))) {
      const staffInfo = mapping.staffRegion.get(normalizeHeader(rawMaNV));
      if (staffInfo.khuVuc && !staffInfo.khuVuc.includes('Chưa phân bổ')) {
        khuVuc = staffInfo.khuVuc;
      }
    }

    // Thống kê theo KHU VỰC cho Bảng 1
    if (!regionSummaryMap[khuVuc]) regionSummaryMap[khuVuc] = { count: 0, revenue: 0, dtpk: 0 };
    regionSummaryMap[khuVuc].count++;
    regionSummaryMap[khuVuc].revenue += giaCuoc;
    regionSummaryMap[khuVuc].dtpk += dtpk;

    let cleanMaKH = String(row[idxMaKH] || '').replace(/^'+/, '').trim();
    if (/^\d{8}$/.test(cleanMaKH)) cleanMaKH = '00' + cleanMaKH;

    let cleanSDT = String(row[idxSDT] || '').replace(/^'+/, '').trim();
    if (/^\d{9}$/.test(cleanSDT)) cleanSDT = '0' + cleanSDT;

    const paddedTenTB   = row[idxTenTB] ? '  ' + String(row[idxTenTB]).trim() : '';
    const paddedDiaChi  = row[idxDiaChi] ? '  ' + String(row[idxDiaChi]).trim() : '';
    const paddedThietBi = row[idxThietBi] ? '  ' + String(row[idxThietBi]).trim() : '';
    const cleanSerial   = String(row[idxSerial] || '').replace(/^'+/, '').trim();
    const cleanMaCH     = String(row[idxMaCH] || '').replace(/^'+/, '').trim();
    const cleanMaNV     = String(row[idxMaNV] || '').replace(/^'+/, '').trim();

    // Tra cứu tên Cửa hàng rút gọn
    const rawTenCH = String(row[idxTenCH] || '').trim();
    const shortTenCH = getShortValue(mapping.stores, rawTenCH);
    const paddedTenCH = shortTenCH ? '  ' + shortTenCH : '';

    const paddedTenNV = row[idxTenNV] ? '  ' + String(row[idxTenNV]).trim() : '';

    // Mốc thời gian phục vụ sắp xếp cũ đến mới
    const sortTime = ptmDate ? ptmDate.getTime() : (dDauNoi ? dDauNoi.getTime() : 9999999999999);

    ptmRowsData.push({
      khuVuc: khuVuc,
      maCH: cleanMaCH,
      maNV: cleanMaNV,
      sortTime: sortTime,
      rowValues: [
        0,                                              // 1. STT (sẽ đánh số sau khi sắp xếp)
        "'" + cleanMaKH,                                // 2. MÃ KHÁCH HÀNG
        paddedTenTB,                                    // 3. TÊN THUÊ BAO
        "'" + cleanSDT,                                 // 4. SĐT LIÊN HỆ
        paddedDiaChi,                                   // 5. ĐỊA CHỈ LẮP ĐẶT
        row[idxGoiCuoc] || '',                          // 6. GÓI CƯỚC
        giaCuoc,                                        // 7. GIÁ CƯỚC
        chuKy > 0 ? chuKy : '',                         // 8. CHU KỲ
        dtpk,                                           // 9. DTPK
        formatDateTimeStr(row[idxNgayDauNoi]),          // 10. NGÀY ĐẤU NỐI
        formatDateTimeStr(row[idxNgayBatDau]),          // 11. NGÀY BẮT ĐẦU
        paddedThietBi,                                  // 12. LOẠI THIẾT BỊ ONT
        "'" + cleanSerial,                              // 13. SERIAL ONT
        khuVuc,                                         // 14. KHU VỰC
        "'" + cleanMaCH,                                // 15. MÃ CỬA HÀNG
        "'" + cleanMaNV,                                // 16. MÃ NHÂN VIÊN
        paddedTenCH,                                    // 17. TÊN CỬA HÀNG (Rút gọn)
        paddedTenNV                                     // 18. TÊN NHÂN VIÊN
      ]
    });
  }

  // =========================================================================
  // 4. THUẬT TOÁN SẮP XẾP PHÂN TẦNG 4 CẤP:
  //    Cấp 1: KHU VỰC (A-Z)
  //    Cấp 2: MÃ CỬA HÀNG (tăng dần)
  //    Cấp 3: MÃ NHÂN VIÊN (tăng dần)
  //    Cấp 4: DÒNG THỜI GIAN (từ cũ nhất đến mới nhất: sortTime tăng dần)
  // =========================================================================
  ptmRowsData.sort((a, b) => {
    // 1. So sánh Khu vực
    const compKV = a.khuVuc.localeCompare(b.khuVuc, 'vi', { sensitivity: 'base' });
    if (compKV !== 0) return compKV;

    // 2. So sánh Mã Cửa hàng
    const compCH = a.maCH.localeCompare(b.maCH, 'vi', { numeric: true });
    if (compCH !== 0) return compCH;

    // 3. So sánh Mã Nhân viên
    const compNV = a.maNV.localeCompare(b.maNV, 'vi', { numeric: true });
    if (compNV !== 0) return compNV;

    // 4. So sánh Dòng thời gian: từ cũ nhất đến mới nhất (tăng dần)
    return a.sortTime - b.sortTime;
  });

  // Đánh lại số thứ tự (STT) liên tục từ 1 đến hết
  const ptmFinalRows = ptmRowsData.map((item, index) => {
    item.rowValues[0] = index + 1; // Gán STT chuẩn
    return item.rowValues;
  });

  // 5. Tạo hoặc làm mới Sheet
  let destSheet = ss.getSheetByName(sheetName);
  if (!destSheet) {
    destSheet = ss.insertSheet(sheetName);
  } else {
    destSheet.clear();
  }

  const ssId = ss.getId();
  const sheetId = destSheet.getSheetId();
  const directExportUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=xlsx&gid=${sheetId}`;

  // DÒNG 1: Tiêu đề lớn (Trải rộng 18 cột)
  destSheet.getRange(1, 1, 1, 18).merge()
    .setValue(`BÁO CÁO THUÊ BAO MOBIFIBER PHÁT TRIỂN MỚI (PTM) - ${sheetTitleSuffix}`)
    .setFontWeight('bold').setFontSize(13).setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');

  // DÒNG 2: TỔNG THUÊ BAO PTM & NGÀY LẬP
  destSheet.getRange(2, 1, 1, 7).merge()
    .setValue(`🎯 TỔNG THUÊ BAO PTM: ${ptmFinalRows.length} TB`)
    .setFontWeight('bold').setFontSize(11).setBackground('#D2E3FC').setFontColor('#174EA6').setHorizontalAlignment('left');

  destSheet.getRange(2, 8, 1, 5).merge()
    .setValue(`🕒 Ngày lập: ${Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm')}`)
    .setFontSize(10).setBackground('#D2E3FC').setFontColor('#5F6368').setHorizontalAlignment('right');

  // DÒNG 3: DOANH THU GÓI CƯỚC
  destSheet.getRange(3, 1, 1, 12).merge()
    .setValue(`💰 DOANH THU GÓI CƯỚC: ${totalRevenue.toLocaleString('vi-VN')} VNĐ`)
    .setFontWeight('bold').setFontSize(11).setBackground('#E8F0FE').setFontColor('#174EA6').setHorizontalAlignment('left');

  // DÒNG 4: DOANH THU PHÂN KỲ (DTPK)
  destSheet.getRange(4, 1, 1, 12).merge()
    .setValue(`📈 DOANH THU PHÂN KỲ (DTPK): ${totalDtpk.toLocaleString('vi-VN')} VNĐ`)
    .setFontWeight('bold').setFontSize(11).setBackground('#FEF7E0').setFontColor('#B06000').setHorizontalAlignment('left');

  // NÚT BẤM TẢI FILE EXCEL TRÊN GOOGLE SHEETS (HỢP NHẤT M2:N4)
  destSheet.getRange(2, 13, 3, 2).merge()
    .setFormula(`=HYPERLINK("${directExportUrl}"; "📥 TẢI FILE EXCEL")`)
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground('#1A73E8')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // ==========================================
  // KHỐI 2: TỔNG HỢP THEO KHU VỰC (Dòng 6 -> ...)
  // ==========================================
  destSheet.getRange('A6:G6').merge()
    .setValue('📊 1. CƠ CẤU PTM THEO KHU VỰC / CỤM ĐỊA BÀN')
    .setFontWeight('bold').setBackground('#D2E3FC').setFontColor('#174EA6');

  destSheet.getRange('A7:C7').merge().setValue('KHU VỰC').setFontWeight('bold').setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  destSheet.getRange('D7').setValue('SỐ TB PTM').setFontWeight('bold').setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  destSheet.getRange('E7').setValue('DOANH THU THU').setFontWeight('bold').setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  destSheet.getRange('F7').setValue('DTPK').setFontWeight('bold').setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  destSheet.getRange('G7').setValue('TỶ TRỌNG DTPK').setFontWeight('bold').setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');

  const regionRows = Object.keys(regionSummaryMap).map(k => [
    '  ' + k,
    regionSummaryMap[k].count,
    regionSummaryMap[k].revenue,
    regionSummaryMap[k].dtpk,
    totalDtpk > 0 ? (regionSummaryMap[k].dtpk / totalDtpk) : 0
  ]).sort((a, b) => b[1] - a[1]);

  let curRow = 8;
  if (regionRows.length > 0) {
    for (let s = 0; s < regionRows.length; s++) {
      const rowIdx = curRow + s;
      destSheet.getRange(rowIdx, 1, 1, 3).merge().setValue(regionRows[s][0]).setHorizontalAlignment('left');
      destSheet.getRange(rowIdx, 4).setValue(regionRows[s][1]).setHorizontalAlignment('center');
      destSheet.getRange(rowIdx, 5).setValue(regionRows[s][2]).setNumberFormat('#,##0').setHorizontalAlignment('right');
      destSheet.getRange(rowIdx, 6).setValue(regionRows[s][3]).setNumberFormat('#,##0').setHorizontalAlignment('right');
      destSheet.getRange(rowIdx, 7).setValue(regionRows[s][4]).setNumberFormat('0.0%').setHorizontalAlignment('right');
    }
    curRow += regionRows.length;
  }

  // ==========================================
  // KHỐI 3: DANH SÁCH CHI TIẾT 18 CỘT CHUẨN
  // (Khoảng cách đúng 1 dòng trống: curRow + 1)
  // ==========================================
  const listStartRow = curRow + 1;
  destSheet.getRange(listStartRow, 1, 1, 18).merge()
    .setValue(`📋 DANH SÁCH CHI TIẾT THUÊ BAO PTM (${ptmFinalRows.length} THUÊ BAO)`)
    .setFontWeight('bold').setFontSize(12).setBackground('#E8F0FE').setFontColor('#174EA6');

  const detailHeaders = [
    'STT', 'MÃ KHÁCH HÀNG', 'TÊN THUÊ BAO', 'SĐT LIÊN HỆ', 'ĐỊA CHỈ LẮP ĐẶT',
    'GÓI CƯỚC', 'GIÁ CƯỚC', 'CHU KỲ', 'DTPK', 'NGÀY ĐẤU NỐI', 'NGÀY BẮT ĐẦU',
    'LOẠI THIẾT BỊ ONT', 'SERIAL ONT', 'KHU VỰC', 'MÃ CỬA HÀNG', 'MÃ NHÂN VIÊN',
    'TÊN CỬA HÀNG', 'TÊN NHÂN VIÊN'
  ];

  const headerRowIdx = listStartRow + 1;
  destSheet.getRange(headerRowIdx, 1, 1, detailHeaders.length)
    .setValues([detailHeaders])
    .setFontWeight('bold').setBackground('#00558F').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  destSheet.setFrozenRows(headerRowIdx);

  if (ptmFinalRows.length > 0) {
    const dataStartRow = headerRowIdx + 1;
    destSheet.getRange(dataStartRow, 1, ptmFinalRows.length, detailHeaders.length).setValues(ptmFinalRows);

    destSheet.getRange(dataStartRow, 1, ptmFinalRows.length, 1).setHorizontalAlignment('center'); // 1. STT
    destSheet.getRange(dataStartRow, 2, ptmFinalRows.length, 1).setNumberFormat('@').setHorizontalAlignment('center'); // 2. Mã KH
    destSheet.getRange(dataStartRow, 3, ptmFinalRows.length, 1).setHorizontalAlignment('left'); // 3. Tên TB
    destSheet.getRange(dataStartRow, 4, ptmFinalRows.length, 1).setNumberFormat('@').setHorizontalAlignment('center'); // 4. SĐT
    destSheet.getRange(dataStartRow, 5, ptmFinalRows.length, 1).setHorizontalAlignment('left'); // 5. Địa chỉ
    destSheet.getRange(dataStartRow, 6, ptmFinalRows.length, 1).setHorizontalAlignment('center'); // 6. Gói cước
    destSheet.getRange(dataStartRow, 7, ptmFinalRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right'); // 7. Giá cước
    destSheet.getRange(dataStartRow, 8, ptmFinalRows.length, 1).setHorizontalAlignment('center'); // 8. Chu kỳ
    destSheet.getRange(dataStartRow, 9, ptmFinalRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right'); // 9. DTPK
    destSheet.getRange(dataStartRow, 10, ptmFinalRows.length, 2).setHorizontalAlignment('center'); // 10, 11. Ngày đấu nối & Ngày bắt đầu
    destSheet.getRange(dataStartRow, 12, ptmFinalRows.length, 1).setHorizontalAlignment('left'); // 12. Loại ONT
    destSheet.getRange(dataStartRow, 13, ptmFinalRows.length, 1).setNumberFormat('@').setHorizontalAlignment('center'); // 13. Serial ONT
    destSheet.getRange(dataStartRow, 14, ptmFinalRows.length, 1).setHorizontalAlignment('center'); // 14. KHU VỰC
    destSheet.getRange(dataStartRow, 15, ptmFinalRows.length, 2).setNumberFormat('@').setHorizontalAlignment('center'); // 15, 16. Mã CH & Mã NV
    destSheet.getRange(dataStartRow, 17, ptmFinalRows.length, 2).setHorizontalAlignment('left'); // 17, 18. Tên CH (Rút gọn) & Tên NV

    // CẤU HÌNH ĐỘ RỘNG 18 CỘT CHUẨN
    const exactColumnWidths = [
      50,  // 1. STT
      120, // 2. Mã KH
      180, // 3. Tên thuê bao
      115, // 4. SĐT
      260, // 5. Địa chỉ lắp đặt
      110, // 6. Gói cước
      110, // 7. Giá cước (Cột G)
      75,  // 8. Chu kỳ
      110, // 9. DTPK
      155, // 10. Ngày đấu nối
      155, // 11. Ngày bắt đầu
      130, // 12. Loại ONT
      130, // 13. Serial ONT
      110, // 14. KHU VỰC
      120, // 15. Mã cửa hàng
      120, // 16. Mã nhân viên
      180, // 17. Tên cửa hàng (Đã rút gọn)
      160  // 18. Tên nhân viên
    ];

    for (let c = 0; c < exactColumnWidths.length; c++) {
      destSheet.setColumnWidth(c + 1, exactColumnWidths[c]);
    }

    // YÊU CẦU ĐẶC BIỆT: CỘT G (Cột 7) - Cột "TỶ TRỌNG DTPK" của Bảng 1
    // Đo độ dài tiêu đề "TỶ TRỌNG DTPK" (13 ký tự) + 2px padding 2 bên -> 124px
    destSheet.setColumnWidth(7, 124);
  }

  ss.setActiveSheet(destSheet);

  return {
    success: true,
    sheetName: sheetName,
    count: ptmFinalRows.length,
    revenue: totalRevenue.toLocaleString('vi-VN'),
    dtpk: totalDtpk.toLocaleString('vi-VN')
  };
}

/**
 * Tạo file Excel sạch độc lập (đã xóa ô M2:N4) để tải về máy
 */
function prepareCleanPtmDownloadUrl(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(sheetName);
  if (!sourceSheet) throw new Error(`Không tìm thấy sheet "${sheetName}"!`);

  cleanupOldTempFiles(sheetName);

  const tempSs = SpreadsheetApp.create(sheetName);
  const tempSsId = tempSs.getId();

  try {
    const copiedSheet = sourceSheet.copyTo(tempSs);
    copiedSheet.setName(sheetName);

    const defaultSheet = tempSs.getSheets()[0];
    if (defaultSheet.getSheetId() !== copiedSheet.getSheetId()) {
      tempSs.deleteSheet(defaultSheet);
    }

    // Xóa bỏ nút tải ô M2:N4 trên file tải về
    copiedSheet.getRange('M2:N4').breakApart().clear().setBackground('#D2E3FC');
    copiedSheet.getRange('H2:N2').merge().setValue(sourceSheet.getRange('H2').getValue()).setHorizontalAlignment('right');
    copiedSheet.getRange('A3:N3').merge().setValue(sourceSheet.getRange('A3').getValue()).setHorizontalAlignment('left');
    copiedSheet.getRange('A4:N4').merge().setValue(sourceSheet.getRange('A4').getValue()).setHorizontalAlignment('left');

    SpreadsheetApp.flush();

    const file = DriveApp.getFileById(tempSsId);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const downloadUrl = `https://docs.google.com/spreadsheets/d/${tempSsId}/export?format=xlsx`;

    return {
      success: true,
      downloadUrl: downloadUrl,
      fileName: `${sheetName}.xlsx`,
      fileId: tempSsId
    };
  } catch (err) {
    DriveApp.getFileById(tempSsId).setTrashed(true);
    throw err;
  }
}

function getPtmDialogHtml() {
  const now = new Date();
  const curMonth = now.getMonth() + 1;
  const curYear = now.getFullYear();

  return `
  <!DOCTYPE html>
  <html>
  <head>
    <base target="_top">
    <style>
      body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; padding: 14px; margin: 0; background: #f8f9fa; font-size: 13px; color: #202124; }
      .opt-card { background: white; border: 1px solid #dadce0; border-radius: 8px; padding: 12px 14px; margin-bottom: 10px; }
      .radio-label { font-weight: 600; display: flex; align-items: center; gap: 8px; cursor: pointer; color: #1a73e8; }
      .inputs { margin-top: 8px; padding-left: 24px; }
      .hint { font-size: 11px; color: #5f6368; margin-top: 4px; line-height: 1.4; font-style: italic; }
      input[type="text"] { padding: 7px 10px; border: 1px solid #dadce0; border-radius: 4px; font-size: 13px; box-sizing: border-box; }
      input[type="text"]:focus { border-color: #1a73e8; outline: none; }
      .row { display: flex; gap: 8px; align-items: center; margin-top: 6px; }
      button { width: 100%; border: none; padding: 10px; font-size: 13px; font-weight: bold; border-radius: 6px; cursor: pointer; margin-top: 6px; display: block; box-sizing: border-box; text-align: center; }
      .btn-primary { background: #1a73e8; color: white; }
      .btn-primary:hover { background: #1557b0; }
      .btn-primary:disabled { background: #dadce0; cursor: not-allowed; }
      .btn-download { background: #0b8043; color: white; }
      .btn-download:hover { background: #086334; }
      .btn-close { background: #f1f3f4; color: #202124; border: 1px solid #dadce0; }
      .btn-close:hover { background: #e8eaed; }
      .alert-banner { display: none; background: #fce8e6; border: 1px solid #f5c2c7; color: #c5221f; padding: 8px 12px; border-radius: 6px; font-size: 12px; margin-bottom: 10px; }
      .status { font-size: 12px; margin-top: 8px; padding: 8px; border-radius: 4px; display: none; text-align: center; }
      .status.info { background: #e8f0fe; color: #174ea6; display: block; }
      #successView { display: none; text-align: center; padding: 6px 4px; }
      .success-title { font-size: 16px; font-weight: bold; color: #137333; margin-bottom: 8px; }
      .kpi-card { background: white; border: 1px solid #dadce0; border-radius: 6px; padding: 10px; margin-bottom: 8px; text-align: left; }
    </style>
  </head>
  <body>
    <div id="formView">
      <div id="errorBanner" class="alert-banner"></div>
      <div class="opt-card">
        <label class="radio-label">
          <input type="radio" name="reportType" value="month" checked onchange="toggleType()">
          <strong>1. Lọc theo Tháng kế hoạch</strong>
        </label>
        <div class="inputs" id="monthBox">
          <span>Tháng:</span>
          <input type="text" id="monthVal" value="${curMonth < 10 ? '0' + curMonth : curMonth}" style="width: 45px; text-align: center;"> /
          <span>Năm:</span>
          <input type="text" id="yearVal" value="${curYear}" style="width: 65px; text-align: center;">
          <div class="hint">💡 Lọc thuê bao PTM có ngày bắt đầu cước (hoặc ngày thi công với CK 1) trong tháng.</div>
        </div>
      </div>

      <div class="opt-card">
        <label class="radio-label">
          <input type="radio" name="reportType" value="range" onchange="toggleType()">
          <strong>2. Lọc theo Khoảng ngày tùy chọn</strong>
        </label>
        <div class="inputs" id="rangeBox" style="opacity: 0.5; pointer-events: none;">
          <div class="row"><span style="width: 65px;">Từ ngày:</span><input type="text" id="fromDate" placeholder="01/09/2026" style="width: 110px;"></div>
          <div class="row"><span style="width: 65px;">Đến ngày:</span><input type="text" id="toDate" placeholder="15/09/2026" style="width: 110px;"></div>
          <div class="hint">💡 Thích hợp xuất báo cáo thi đua tuần hoặc chiến dịch bán hàng ngắn ngày.</div>
        </div>
      </div>

      <button id="btnRun" class="btn-primary" onclick="runReport()">🚀 Xuất Báo cáo PTM</button>
      <div id="statusBox" class="status"></div>
    </div>

    <div id="successView">
      <div class="success-title">🎉 XUẤT BÁO CÁO PTM THÀNH CÔNG!</div>
      <div class="kpi-card">
        <div style="font-weight: 600; color: #1a73e8;" id="resSheetName">Trang tính: ...</div>
        <div style="margin-top: 6px; font-size: 13px; line-height: 1.6;">
          🎯 <b>Số lượng thuê bao PTM:</b> <span id="resCount" style="color: #c5221f; font-weight: bold;">0</span> TB<br>
          💰 <b>Doanh thu gói cước:</b> <span id="resRevenue" style="color: #137333; font-weight: bold;">0</span> VNĐ<br>
          📈 <b>Doanh thu phân kỳ (DTPK):</b> <span id="resDtpk" style="color: #b06000; font-weight: bold;">0</span> VNĐ
        </div>
      </div>
      <button id="btnDlPtm" class="btn-download" onclick="startCleanPtmDownload()">📥 Tải file Excel (.xlsx) về máy</button>
      <button class="btn-close" onclick="closeDialogCleanly()">Đóng cửa sổ</button>
    </div>

    <iframe id="downloadPtmFrame" style="display:none;"></iframe>

    <script>
      let currentCreatedPtmSheet = '';
      let currentTempPtmFileId = null;

      window.onload = function() {
        const mInput = document.getElementById('monthVal');
        if (mInput) {
          mInput.focus();
          mInput.select();
        }
      };

      function showError(msg) {
        const b = document.getElementById('errorBanner');
        b.innerText = '⚠️ ' + msg;
        b.style.display = 'block';
      }

      function hideError() {
        document.getElementById('errorBanner').style.display = 'none';
      }

      function toggleType() {
        hideError();
        const isMonth = document.querySelector('input[name="reportType"]:checked').value === 'month';
        document.getElementById('monthBox').style.opacity = isMonth ? '1' : '0.5';
        document.getElementById('monthBox').style.pointerEvents = isMonth ? 'auto' : 'none';
        document.getElementById('rangeBox').style.opacity = isMonth ? '0.5' : '1';
        document.getElementById('rangeBox').style.pointerEvents = isMonth ? 'none' : 'auto';
        if (isMonth) {
          document.getElementById('monthVal').focus();
        } else {
          document.getElementById('fromDate').focus();
        }
      }

      function isValidDateStr(str) {
        if (!/^\\d{1,2}\\/\\d{1,2}\\/\\d{4}$/.test(str)) return false;
        const p = str.split('/');
        const d = parseInt(p[0], 10), m = parseInt(p[1], 10), y = parseInt(p[2], 10);

        if (m < 1 || m > 12 || d < 1 || d > 31) return false;
        const dt = new Date(y, m - 1, d);
        return dt.getFullYear() === y && (dt.getMonth() + 1) === m && dt.getDate() === d;
      }

      function parseDateObj(str) {
        const p = str.split('/');
        return new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10));
      }

      function runReport() {
        hideError();
        const type = document.querySelector('input[name="reportType"]:checked').value;
        const btn = document.getElementById('btnRun');
        const status = document.getElementById('statusBox');

        let params = { type: type };
        if (type === 'month') {
          const m = document.getElementById('monthVal').value.trim();
          const y = document.getElementById('yearVal').value.trim();
          if (!m || !y) { showError('Vui lòng nhập đầy đủ Tháng và Năm!'); return; }
          const mNum = parseInt(m, 10), yNum = parseInt(y, 10);
          if (isNaN(mNum) || mNum < 1 || mNum > 12) { showError('Tháng phải là số từ 1 đến 12!'); return; }
          if (isNaN(yNum) || yNum < 2000 || yNum > 2100) { showError('Năm không hợp lệ!'); return; }
          params.month = m; params.year = y;
        } else {
          const f = document.getElementById('fromDate').value.trim();
          const t = document.getElementById('toDate').value.trim();
          if (!f || !t) { showError('Vui lòng nhập đầy đủ Từ ngày và Đến ngày!'); return; }
          if (!isValidDateStr(f)) { showError('Định dạng "Từ ngày" không đúng (phải là dd/MM/yyyy)!'); return; }
          if (!isValidDateStr(t)) { showError('Định dạng "Đến ngày" không đúng (phải là dd/MM/yyyy)!'); return; }
          if (parseDateObj(f) > parseDateObj(t)) { showError('"Từ ngày" không được lớn hơn "Đến ngày"!'); return; }
          params.from = f; params.to = t;
        }

        btn.disabled = true;
        btn.innerText = '⏳ Đang quét dữ liệu PTM...';
        status.className = 'status info';
        status.innerText = 'Hệ thống đang trích xuất dữ liệu và tối ưu bảng...';

        google.script.run
          .withSuccessHandler(function(res) {
            currentCreatedPtmSheet = res.sheetName;
            document.getElementById('formView').style.display = 'none';
            document.getElementById('resSheetName').innerText = '📁 Trang tính: ' + res.sheetName;
            document.getElementById('resCount').innerText = res.count;
            document.getElementById('resRevenue').innerText = res.revenue;
            document.getElementById('resDtpk').innerText = res.dtpk;
            document.getElementById('successView').style.display = 'block';
          })
          .withFailureHandler(function(err) {
            btn.disabled = false;
            btn.innerText = '🚀 Xuất Báo cáo PTM';
            status.style.display = 'none';
            showError(err.message);
          })
          .processPtmReport(params);
      }

      function startCleanPtmDownload() {
        const btn = document.getElementById('btnDlPtm');
        btn.disabled = true;
        btn.innerText = '⏳ Đang chuẩn bị file Excel sạch...';

        google.script.run
          .withSuccessHandler(function(res) {
            btn.disabled = false;
            btn.innerText = '📥 Tải file Excel (.xlsx) về máy';
            currentTempPtmFileId = res.fileId;
            document.getElementById('downloadPtmFrame').src = res.downloadUrl;

            // Tự động xóa file tạm sau 5 giây
            setTimeout(function() {
              if (currentTempPtmFileId) {
                google.script.run.deleteTempFile(currentTempPtmFileId);
                currentTempPtmFileId = null;
              }
            }, 5000);
          })
          .withFailureHandler(function(err) {
            btn.disabled = false;
            btn.innerText = '📥 Tải file Excel (.xlsx) về máy';
            showError('Lỗi tải file: ' + err.message);
          })
          .prepareCleanPtmDownloadUrl(currentCreatedPtmSheet);
      }

      function closeDialogCleanly() {
        if (currentTempPtmFileId) {
          google.script.run.deleteTempFile(currentTempPtmFileId);
          currentTempPtmFileId = null;
        }
        google.script.host.close();
      }
    </script>
  </body>
  </html>
  `;
}