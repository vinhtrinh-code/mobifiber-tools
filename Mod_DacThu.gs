/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * MODULE       : BÁO CÁO THUÊ BAO ĐẶC THÙ (NGHIỆP VỤ / BCA)
 * PHIÊN BẢN    : 4.1 (Bổ sung Tỷ trọng Thuê bao vào Bảng Tổng hợp Khu vực)
 * NGÀY CẬP NHẬT: 29/10/2026
 * TÁC GIẢ      : Trịnh Vũ Thế Vinh & AI Assistant
 * ============================================================================
 */

function showSpecialReportDialog() {
  const html = HtmlService.createHtmlOutput(getSpecialDialogHtml())
    .setWidth(500)
    .setHeight(480)
    .setTitle('Báo cáo TB Nghiệp vụ/BCA');
  SpreadsheetApp.getUi().showModalDialog(html, 'Báo cáo Thuê bao Nghiệp vụ / BCA');
}

/**
 * Xử lý lọc dữ liệu và tạo trang báo cáo đặc thù trên Google Sheets
 */
function processSpecialReport(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);
  
  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) throw new Error('Bảng raw_mobifiber chưa có dữ liệu.');
  
  // 1. Nạp bản đồ tham số chuẩn hóa từ thamso_caidat
  const mapping = loadThamSoMapping();
  
  // 2. Phân tích đối tượng lọc
  const targetGroup = params.targetGroup || 'all';
  let groupTitle = '', groupCode = '';
  if (targetGroup === 'nv') {
    groupTitle = 'THUÊ BAO NGHIỆP VỤ';
    groupCode = 'NV';
  } else if (targetGroup === 'bca') {
    groupTitle = 'THUÊ BAO BCA / HÀNH CHÍNH & LLVT';
    groupCode = 'BCA';
  } else {
    groupTitle = 'THUÊ BAO NGHIỆP VỤ & BCA / LLVT';
    groupCode = 'ALL';
  }
  
  // 3. Phân tích thời gian lọc
  const timeType = params.timeType || 'month';
  let startDate = null, endDate = null, timeTitle = '', timeCode = '';
  
  if (timeType === 'month') {
    const month = parseInt(params.month, 10);
    const year = parseInt(params.year, 10);
    const mStr = month < 10 ? '0' + month : '' + month;
    startDate = new Date(year, month - 1, 1, 0, 0, 0);
    endDate = new Date(year, month, 0, 23, 59, 59);
    timeTitle = `THÁNG ${mStr}/${year}`;
    timeCode = `T${mStr}_${year}`;
  } else if (timeType === 'range') {
    const startParts = params.from.split('/');
    const endParts = params.to.split('/');
    startDate = new Date(parseInt(startParts[2], 10), parseInt(startParts[1], 10) - 1, parseInt(startParts[0], 10), 0, 0, 0);
    endDate = new Date(parseInt(endParts[2], 10), parseInt(endParts[1], 10) - 1, parseInt(endParts[0], 10), 23, 59, 59);
    const d1 = startParts[0] + startParts[1];
    const d2 = endParts[0] + endParts[1];
    timeTitle = `TỪ ${params.from} ĐẾN ${params.to}`;
    timeCode = `${d1}_${d2}_${endParts[2]}`;
  } else {
    timeTitle = 'TOÀN BỘ DANH SÁCH LŨY KẾ';
    timeCode = 'LuyKe';
  }
  
  const sheetName = `DacThu_${groupCode}_${timeCode}`;
  
  // 4. Đọc Headers
  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol)
    .getValues()[0].map(h => normalizeHeader(h));
  
  const idxMaKH        = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxLoaiKH      = headers.findIndex(h => h === 'loại khách hàng');
  const idxTenTB       = headers.findIndex(h => h === 'tên thuê bao');
  const idxSDT         = headers.findIndex(h => h === 'sđt liên hệ');
  const idxDiaChi      = headers.findIndex(h => h === 'địa chỉ lắp đặt');
  const idxGoiCuoc     = headers.findIndex(h => h === 'gói cước');
  const idxGiaCuoc     = headers.findIndex(h => h === 'giá cước');
  const idxChuKy       = headers.findIndex(h => h === 'chu kỳ sử dụng');
  const idxDtpk        = headers.findIndex(h => h === 'dtpk');
  const idxThietBi     = headers.findIndex(h => h === 'loại thiết bị ont');
  const idxSerial      = headers.findIndex(h => h === 'serial ont');
  const idxNgayDauNoi  = headers.findIndex(h => h === 'ngày đấu nối');
  const idxNgayBatDau  = headers.findIndex(h => h === 'ngày băt đầu gói cước' || h === 'ngày bắt đầu gói cước');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxChanCat     = headers.findIndex(h => h === 'tình trạng chặn cắt');
  const idxMaCH        = headers.findIndex(h => h === 'mã cửa hàng');
  const idxTenCH       = headers.findIndex(h => h === 'tên cửa hàng');
  const idxMaNV        = headers.findIndex(h => h === 'mã nhân viên');
  const idxTenNV       = headers.findIndex(h => h === 'tên nhân viên');
  
  const numRows = lastRow - CONFIG.DATA_START_ROW_TARGET + 1;
  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, numRows, lastCol).getValues();
  
  const specRowsData = [];
  const regionSummaryMap = {};
  let totalRevenue = 0;
  
  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const rawStatusHD = String(row[idxTinhTrangHD] || '').trim();
    const normStatusHD = normalizeHeader(rawStatusHD);
  
    // Lọc triệt để hợp đồng hủy
    const isCancelled = normStatusHD.includes('hủy') || normStatusHD.includes('huy') ||
                        normStatusHD.includes('thanh lý') || normStatusHD.includes('thanh ly') ||
                        normStatusHD.includes('chấm dứt') || normStatusHD.includes('cham dut');
    if (isCancelled) continue;
  
    const rawLoaiKH = String(row[idxLoaiKH] || '').trim();
    const normLoaiKH = normalizeHeader(rawLoaiKH);
  
    const isNghiepVu = normLoaiKH.includes('thuê bao nghiệp vụ') || normLoaiKH.includes('thue bao nghiep vu');
    const isBCA = normLoaiKH.includes('bca') || normLoaiKH.includes('vũ trang') ||
                  normLoaiKH.includes('vu trang') || normLoaiKH.includes('hành chính') ||
                  normLoaiKH.includes('công an');
  
    if (targetGroup === 'nv' && !isNghiepVu) continue;
    if (targetGroup === 'bca' && !isBCA) continue;
    if (targetGroup === 'all' && !isNghiepVu && !isBCA) continue;
  
    const dDauNoi = parseToDate(row[idxNgayDauNoi]);
    const dBatDau = parseToDate(row[idxNgayBatDau]);
    const checkDate = dDauNoi || dBatDau;
  
    if (timeType !== 'all') {
      if (!checkDate) continue;
      if (checkDate < startDate || checkDate > endDate) continue;
    }
  
    const giaCuoc = parseCurrencyToInteger(row[idxGiaCuoc]) || 0;
    const chuKy = row[idxChuKy] ? parseInt(row[idxChuKy], 10) : 0;
    const dtpk = parseCurrencyToInteger(row[idxDtpk]) || (chuKy > 0 ? Math.round(giaCuoc / chuKy) : 0);
    totalRevenue += giaCuoc;
  
    // Xác định Khu vực từ thamso_caidat
    const rawMaNV = String(row[idxMaNV] || '').trim();
    let khuVuc = 'Chưa phân bổ KV';
    if (rawMaNV && mapping.staffRegion && mapping.staffRegion.has(normalizeHeader(rawMaNV))) {
      const staffInfo = mapping.staffRegion.get(normalizeHeader(rawMaNV));
      if (staffInfo.khuVuc && !staffInfo.khuVuc.includes('Chưa phân bổ')) {
        khuVuc = staffInfo.khuVuc;
      }
    }
  
    if (!regionSummaryMap[khuVuc]) regionSummaryMap[khuVuc] = { count: 0, revenue: 0 };
    regionSummaryMap[khuVuc].count++;
    regionSummaryMap[khuVuc].revenue += giaCuoc;
  
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
  
    // Tra cứu tên Cửa hàng rút gọn từ thamso_caidat
    const rawTenCH = String(row[idxTenCH] || '').trim();
    const shortTenCH = getShortValue(mapping.stores, rawTenCH);
    const paddedTenCH = shortTenCH ? '  ' + shortTenCH : '';
  
    const paddedTenNV = row[idxTenNV] ? '  ' + String(row[idxTenNV]).trim() : '';
    
    // Rút gọn LOẠI KHÁCH HÀNG
    const shortLoaiKH = getShortValue(mapping.custTypes, rawLoaiKH);
    const paddedLoaiKH = shortLoaiKH ? '  ' + shortLoaiKH : '';

    const sortTime = dDauNoi ? dDauNoi.getTime() : (dBatDau ? dBatDau.getTime() : 9999999999999);
  
    specRowsData.push({
      isNghiepVu: isNghiepVu ? 1 : 2,
      khuVuc: khuVuc,
      maCH: cleanMaCH,
      maNV: cleanMaNV,
      sortTime: sortTime,
      rowValues: [
        0,                                              // 1. STT
        khuVuc,                                         // 2. KHU VỰC
        "'" + cleanMaKH,                                // 3. MÃ KHÁCH HÀNG
        paddedLoaiKH,                                   // 4. LOẠI KHÁCH HÀNG (Rút gọn)
        paddedTenTB,                                    // 5. TÊN THUÊ BAO
        "'" + cleanSDT,                                 // 6. SĐT LIÊN HỆ
        paddedDiaChi,                                   // 7. ĐỊA CHỈ LẮP ĐẶT
        row[idxGoiCuoc] || '',                          // 8. GÓI CƯỚC
        giaCuoc,                                        // 9. GIÁ CƯỚC
        chuKy > 0 ? chuKy : '',                         // 10. CHU KỲ
        dtpk,                                           // 11. DTPK
        formatDateTimeStr(row[idxNgayDauNoi]),          // 12. NGÀY ĐẤU NỐI
        formatDateTimeStr(row[idxNgayBatDau]),          // 13. NGÀY BẮT ĐẦU
        paddedThietBi,                                  // 14. LOẠI THIẾT BỊ ONT
        "'" + cleanSerial,                              // 15. SERIAL ONT
        rawStatusHD || 'Đang hoạt động',                // 16. TÌNH TRẠNG HỢP ĐỒNG
        row[idxChanCat] || 'Bình thường',               // 17. TÌNH TRẠNG CHẶN CẮT
        "'" + cleanMaCH,                                // 18. MÃ CỬA HÀNG
        "'" + cleanMaNV,                                // 19. MÃ NHÂN VIÊN
        paddedTenCH,                                    // 20. TÊN CỬA HÀNG (Rút gọn)
        paddedTenNV                                     // 21. TÊN NHÂN VIÊN
      ]
    });
  }
  
  // Sắp xếp phân tầng 4 cấp
  specRowsData.sort((a, b) => {
    if (a.isNghiepVu !== b.isNghiepVu) return a.isNghiepVu - b.isNghiepVu;
    const compKV = a.khuVuc.localeCompare(b.khuVuc, 'vi', { sensitivity: 'base' });
    if (compKV !== 0) return compKV;
    const compCH = a.maCH.localeCompare(b.maCH, 'vi', { numeric: true });
    if (compCH !== 0) return compCH;
    const compNV = a.maNV.localeCompare(b.maNV, 'vi', { numeric: true });
    if (compNV !== 0) return compNV;
    return a.sortTime - b.sortTime;
  });
  
  const specFinalRows = specRowsData.map((item, index) => {
    item.rowValues[0] = index + 1;
    return item.rowValues;
  });
  
  // Tạo hoặc làm mới Sheet
  let destSheet = ss.getSheetByName(sheetName);
  if (!destSheet) {
    destSheet = ss.insertSheet(sheetName);
  } else {
    destSheet.clear();
  }
  
  const ssId = ss.getId();
  const sheetId = destSheet.getSheetId();
  const directExportUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=xlsx&gid=${sheetId}`;
  
  // DÒNG 1: Tiêu đề lớn
  destSheet.getRange(1, 1, 1, 21).merge()
    .setValue(`BÁO CÁO THUÊ BAO MOBIFIBER ĐẶC THÙ (${groupTitle}) - ${timeTitle}`)
    .setFontWeight('bold').setFontSize(13).setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  
  // DÒNG 2: TỔNG THUÊ BAO & NGÀY LẬP
  destSheet.getRange(2, 1, 1, 8).merge()
    .setValue(`🎯 TỔNG SỐ THUÊ BAO: ${specFinalRows.length} TB`)
    .setFontWeight('bold').setFontSize(11).setBackground('#D2E3FC').setFontColor('#174EA6').setHorizontalAlignment('left');
  
  destSheet.getRange(2, 9, 1, 7).merge()
    .setValue(`🕒 Ngày lập: ${Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm')}`)
    .setFontSize(10).setBackground('#D2E3FC').setFontColor('#5F6368').setHorizontalAlignment('right');
  
  // DÒNG 3: DOANH THU GÓI CƯỚC
  destSheet.getRange(3, 1, 1, 15).merge()
    .setValue(`💰 TỔNG DOANH THU GÓI CƯỚC: ${totalRevenue.toLocaleString('vi-VN')} VNĐ`)
    .setFontWeight('bold').setFontSize(11).setBackground('#E8F0FE').setFontColor('#174EA6').setHorizontalAlignment('left');
  
  // NÚT BẤM TẢI FILE EXCEL TRÊN GOOGLE SHEETS
  destSheet.getRange(2, 16, 2, 6).merge()
    .setFormula(`=HYPERLINK("${directExportUrl}"; "📥 TẢI FILE EXCEL")`)
    .setFontWeight('bold').setFontSize(11).setBackground('#1A73E8').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  
  // KHỐI 2: TỔNG HỢP THEO KHU VỰC (Mở rộng thành 5 Cột)
  destSheet.getRange('A5:G5').merge()
    .setValue('📊 1. CƠ CẤU THEO KHU VỰC / CỤM ĐỊA BÀN')
    .setFontWeight('bold').setBackground('#D2E3FC').setFontColor('#174EA6');
  
  destSheet.getRange('A6:C6').merge().setValue('KHU VỰC').setFontWeight('bold').setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  destSheet.getRange('D6').setValue('SỐ THUÊ BAO').setFontWeight('bold').setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  destSheet.getRange('E6').setValue('TỶ TRỌNG TB (%)').setFontWeight('bold').setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  destSheet.getRange('F6').setValue('TỔNG DTHU').setFontWeight('bold').setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  destSheet.getRange('G6').setValue('TỶ TRỌNG DTHU (%)').setFontWeight('bold').setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  
  const regionRows = Object.keys(regionSummaryMap).map(k => {
    const rCount = regionSummaryMap[k].count;
    const rRev = regionSummaryMap[k].revenue;
    const pctTB = specFinalRows.length > 0 ? (rCount / specFinalRows.length) : 0;
    const pctRev = totalRevenue > 0 ? (rRev / totalRevenue) : 0;
    return [
      '  ' + k,
      rCount,
      pctTB,
      rRev,
      pctRev
    ];
  }).sort((a, b) => b[3] - a[3]); // Sắp xếp theo Doanh thu giảm dần
  
  let curRow = 7;
  if (regionRows.length > 0) {
    for (let s = 0; s < regionRows.length; s++) {
      const rowIdx = curRow + s;
      destSheet.getRange(rowIdx, 1, 1, 3).merge().setValue(regionRows[s][0]).setHorizontalAlignment('left');
      destSheet.getRange(rowIdx, 4).setValue(regionRows[s][1]).setHorizontalAlignment('center');
      destSheet.getRange(rowIdx, 5).setValue(regionRows[s][2]).setNumberFormat('0.0%').setHorizontalAlignment('right');
      destSheet.getRange(rowIdx, 6).setValue(regionRows[s][3]).setNumberFormat('#,##0').setHorizontalAlignment('right');
      destSheet.getRange(rowIdx, 7).setValue(regionRows[s][4]).setNumberFormat('0.0%').setHorizontalAlignment('right');
    }
    curRow += regionRows.length;
  }
  
  // KHỐI 3: DANH SÁCH CHI TIẾT 21 CỘT CHUẨN
  const listStartRow = curRow + 1;
  destSheet.getRange(listStartRow, 1, 1, 21).merge()
    .setValue(`📋 DANH SÁCH CHI TIẾT (${specFinalRows.length} THUÊ BAO) - SẮP XẾP PHÂN TẦNG`)
    .setFontWeight('bold').setFontSize(12).setBackground('#E8F0FE').setFontColor('#174EA6');
  
  const detailHeaders = [
    'STT', 'KHU VỰC', 'MÃ KHÁCH HÀNG', 'LOẠI KHÁCH HÀNG', 'TÊN THUÊ BAO', 'SĐT LIÊN HỆ', 'ĐỊA CHỈ LẮP ĐẶT',
    'GÓI CƯỚC', 'GIÁ CƯỚC', 'CHU KỲ', 'DTPK', 'NGÀY ĐẤU NỐI', 'NGÀY BẮT ĐẦU',
    'LOẠI THIẾT BỊ ONT', 'SERIAL ONT', 'TÌNH TRẠNG HỢP ĐỒNG', 'TÌNH TRẠNG CHẶN CẮT',
    'MÃ CỬA HÀNG', 'MÃ NHÂN VIÊN', 'TÊN CỬA HÀNG', 'TÊN NHÂN VIÊN'
  ];
  
  const headerRowIdx = listStartRow + 1;
  destSheet.getRange(headerRowIdx, 1, 1, detailHeaders.length)
    .setValues([detailHeaders])
    .setFontWeight('bold').setBackground('#00558F').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  destSheet.setFrozenRows(headerRowIdx);
  
  if (specFinalRows.length > 0) {
    const dataStartRow = headerRowIdx + 1;
    destSheet.getRange(dataStartRow, 1, specFinalRows.length, detailHeaders.length).setValues(specFinalRows);
  
    destSheet.getRange(dataStartRow, 1, specFinalRows.length, 2).setHorizontalAlignment('center'); // STT & KHU VỰC
    destSheet.getRange(dataStartRow, 3, specFinalRows.length, 1).setNumberFormat('@').setHorizontalAlignment('center'); // Mã KH
    destSheet.getRange(dataStartRow, 4, specFinalRows.length, 2).setHorizontalAlignment('left'); // Loại KH & Tên TB
    destSheet.getRange(dataStartRow, 6, specFinalRows.length, 1).setNumberFormat('@').setHorizontalAlignment('center'); // SĐT
    destSheet.getRange(dataStartRow, 7, specFinalRows.length, 1).setHorizontalAlignment('left'); // Địa chỉ
    destSheet.getRange(dataStartRow, 8, specFinalRows.length, 1).setHorizontalAlignment('center'); // Gói cước
    destSheet.getRange(dataStartRow, 9, specFinalRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right'); // Giá cước
    destSheet.getRange(dataStartRow, 10, specFinalRows.length, 1).setHorizontalAlignment('center'); // Chu kỳ
    destSheet.getRange(dataStartRow, 11, specFinalRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right'); // DTPK
    destSheet.getRange(dataStartRow, 12, specFinalRows.length, 2).setHorizontalAlignment('center'); // Ngày đấu nối & Ngày bắt đầu
    destSheet.getRange(dataStartRow, 14, specFinalRows.length, 1).setHorizontalAlignment('left'); // Loại ONT
    destSheet.getRange(dataStartRow, 15, specFinalRows.length, 1).setNumberFormat('@').setHorizontalAlignment('center'); // Serial ONT
    destSheet.getRange(dataStartRow, 16, specFinalRows.length, 2).setHorizontalAlignment('center'); // Tình trạng HĐ & Chặn cắt
    destSheet.getRange(dataStartRow, 18, specFinalRows.length, 2).setNumberFormat('@').setHorizontalAlignment('center'); // Mã CH & Mã NV
    destSheet.getRange(dataStartRow, 20, specFinalRows.length, 2).setHorizontalAlignment('left'); // Tên CH & Tên NV
  
    // Tự động co giãn độ rộng cột thông minh (50px -> 250px)
    destSheet.autoResizeColumns(1, detailHeaders.length);
    for (let c = 1; c <= detailHeaders.length; c++) {
      let curWidth = destSheet.getColumnWidth(c);
      let newWidth = curWidth + 15;
      if (newWidth > 250) newWidth = 250;
      if (newWidth < 50) newWidth = 50;
      destSheet.setColumnWidth(c, newWidth);
    }
  }
  
  ss.setActiveSheet(destSheet);
  
  return {
    success: true,
    sheetName: sheetName,
    count: specFinalRows.length,
    revenue: totalRevenue.toLocaleString('vi-VN')
  };
}
  
/**
 * TẠO FILE EXCEL ĐẶC THÙ SẠCH (ĐÃ XÓA Ô NÚT BẤM TRÊN EXCEL XUẤT RA)
 */
function prepareCleanSpecialDownloadUrl(sheetName) {
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
  
    // Xóa nút bấm khỏi bản tải về
    copiedSheet.getRange('P2:V4').breakApart().clear().setBackground('#FFFFFF');
    copiedSheet.getRange('A2:O2').merge().setValue(sourceSheet.getRange('A2').getValue()).setHorizontalAlignment('left');
    copiedSheet.getRange('A3:O3').merge().setValue(sourceSheet.getRange('A3').getValue()).setHorizontalAlignment('left');
  
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
  
/**
 * Giao diện HTML Hộp thoại Lọc Thuê bao Đặc thù
 */
function getSpecialDialogHtml() {
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
      .section-title { font-weight: bold; color: #1a73e8; margin-bottom: 8px; font-size: 13px; }
      .radio-label { font-weight: 500; display: flex; align-items: center; gap: 8px; cursor: pointer; margin-bottom: 6px; }
      .inputs { margin-top: 6px; padding-left: 24px; }
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
      .alert-banner { display: none; background: #fce8e6; border: 1px solid #f5c2c7; color: #c5221f; padding: 8px 12px; border-radius: 6px; font-size: 12px; margin-bottom: 10px; line-height: 1.4; }
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
        <div class="section-title">1. Chọn Đối tượng Thuê bao:</div>
        <label class="radio-label">
          <input type="radio" name="targetGroup" value="all" checked>
          <span><b>Cả hai nhóm</b> (Nghiệp vụ + BCA/LLVT)</span>
        </label>
        <label class="radio-label">
          <input type="radio" name="targetGroup" value="nv">
          <span>Chỉ Thuê bao <b>Nghiệp vụ</b> (Nội bộ MobiFone)</span>
        </label>
        <label class="radio-label">
          <input type="radio" name="targetGroup" value="bca">
          <span>Chỉ Thuê bao <b>BCA / Lực lượng vũ trang & Hành chính</b></span>
        </label>
      </div>
  
      <div class="opt-card">
        <div class="section-title">2. Chọn Kỳ Báo cáo:</div>
        <label class="radio-label">
          <input type="radio" name="timeType" value="month" checked onchange="toggleType()">
          <span>Theo Tháng kế hoạch</span>
        </label>
        <div class="inputs" id="monthBox">
          <span>Tháng:</span>
          <input type="text" id="monthVal" value="${curMonth < 10 ? '0' + curMonth : curMonth}" style="width: 45px; text-align: center;"> /
          <span>Năm:</span>
          <input type="text" id="yearVal" value="${curYear}" style="width: 65px; text-align: center;">
          <div class="hint">💡 Lọc thuê bao có ngày đấu nối hoặc bắt đầu trong tháng.</div>
        </div>
  
        <label class="radio-label" style="margin-top: 8px;">
          <input type="radio" name="timeType" value="range" onchange="toggleType()">
          <span>Theo Khoảng ngày tùy chọn</span>
        </label>
        <div class="inputs" id="rangeBox" style="opacity: 0.5; pointer-events: none;">
          <div class="row">
            <span style="width: 65px;">Từ ngày:</span>
            <input type="text" id="fromDate" placeholder="01/09/2026" style="width: 110px;">
          </div>
          <div class="row">
            <span style="width: 65px;">Đến ngày:</span>
            <input type="text" id="toDate" placeholder="15/09/2026" style="width: 110px;">
          </div>
        </div>
  
        <label class="radio-label" style="margin-top: 8px;">
          <input type="radio" name="timeType" value="all" onchange="toggleType()">
          <span><b>Toàn bộ danh sách lũy kế</b> (Tổng kiểm kê tài sản/ONT)</span>
        </label>
      </div>
  
      <button id="btnRun" class="btn-primary" onclick="runReport()">🚀 Kết xuất Danh sách Đặc thù</button>
      <div id="statusBox" class="status"></div>
    </div>
  
    <div id="successView">
      <div class="success-title">🎉 KẾT XUẤT THÀNH CÔNG!</div>
      <div class="kpi-card">
        <div style="font-weight: 600; color: #1a73e8;" id="resSheetName">Trang tính: ...</div>
        <div style="margin-top: 6px; font-size: 13px; line-height: 1.6;">
          🎯 <b>Số lượng thuê bao:</b> <span id="resCount" style="color: #c5221f; font-weight: bold;">0</span> TB<br>
          💰 <b>Tổng doanh thu gói cước:</b> <span id="resRevenue" style="color: #137333; font-weight: bold;">0</span> VNĐ
        </div>
      </div>
  
      <button id="btnDlSpec" class="btn-download" onclick="startCleanSpecialDownload()">📥 Tải file Excel (.xlsx) về máy</button>
      <button class="btn-close" onclick="closeDialogCleanly()">Đóng cửa sổ</button>
    </div>
  
    <iframe id="downloadSpecFrame" style="display:none;"></iframe>
  
    <script>
      let currentCreatedSpecSheet = '';
      let currentTempSpecFileId = null;
  
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
        const isMonth = document.querySelector('input[name="timeType"]:checked').value === 'month';
        const isRange = document.querySelector('input[name="timeType"]:checked').value === 'range';
       
        document.getElementById('monthBox').style.opacity = isMonth ? '1' : '0.5';
        document.getElementById('monthBox').style.pointerEvents = isMonth ? 'auto' : 'none';
        document.getElementById('rangeBox').style.opacity = isRange ? '1' : '0.5';
        document.getElementById('rangeBox').style.pointerEvents = isRange ? 'auto' : 'none';
  
        if (isMonth) document.getElementById('monthVal').focus();
        else if (isRange) document.getElementById('fromDate').focus();
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
        const targetGroup = document.querySelector('input[name="targetGroup"]:checked').value;
        const timeType = document.querySelector('input[name="timeType"]:checked').value;
        const btn = document.getElementById('btnRun');
        const status = document.getElementById('statusBox');
  
        let params = { targetGroup: targetGroup, timeType: timeType };
  
        if (timeType === 'month') {
          const m = document.getElementById('monthVal').value.trim();
          const y = document.getElementById('yearVal').value.trim();
          if (!m || !y) { showError('Vui lòng nhập đầy đủ Tháng và Năm!'); return; }
          const mNum = parseInt(m, 10), yNum = parseInt(y, 10);
          if (isNaN(mNum) || mNum < 1 || mNum > 12) { showError('Tháng phải là số từ 1 đến 12!'); return; }
          if (isNaN(yNum) || yNum < 2000 || yNum > 2100) { showError('Năm không hợp lệ!'); return; }
          params.month = m; params.year = y;
        } else if (timeType === 'range') {
          const f = document.getElementById('fromDate').value.trim();
          const t = document.getElementById('toDate').value.trim();
          if (!f || !t) { showError('Vui lòng nhập đầy đủ Từ ngày và Đến ngày!'); return; }
          if (!isValidDateStr(f)) { showError('Định dạng "Từ ngày" không đúng (phải là dd/MM/yyyy)!'); return; }
          if (!isValidDateStr(t)) { showError('Định dạng "Đến ngày" không đúng (phải là dd/MM/yyyy)!'); return; }
          if (parseDateObj(f) > parseDateObj(t)) { showError('"Từ ngày" không được lớn hơn "Đến ngày"!'); return; }
          params.from = f; params.to = t;
        }
  
        btn.disabled = true;
        btn.innerText = '⏳ Đang quét dữ liệu đặc thù...';
        status.className = 'status info';
        status.innerText = 'Hệ thống đang trích xuất dữ liệu và sắp xếp phân tầng...';
  
        google.script.run
          .withSuccessHandler(function(res) {
            currentCreatedSpecSheet = res.sheetName;
            document.getElementById('formView').style.display = 'none';
            document.getElementById('resSheetName').innerText = '📁 Trang tính: ' + res.sheetName;
            document.getElementById('resCount').innerText = res.count;
            document.getElementById('resRevenue').innerText = res.revenue;
            document.getElementById('successView').style.display = 'block';
          })
          .withFailureHandler(function(err) {
            btn.disabled = false;
            btn.innerText = '🚀 Kết xuất Danh sách Đặc thù';
            status.style.display = 'none';
            showError(err.message);
          })
          .processSpecialReport(params);
      }
  
      function startCleanSpecialDownload() {
        const btn = document.getElementById('btnDlSpec');
        btn.disabled = true;
        btn.innerText = '⏳ Đang chuẩn bị file Excel sạch...';
  
        google.script.run
          .withSuccessHandler(function(res) {
            btn.disabled = false;
            btn.innerText = '📥 Tải file Excel (.xlsx) về máy';
            currentTempSpecFileId = res.fileId;
            document.getElementById('downloadSpecFrame').src = res.downloadUrl;
  
            // Tự động xóa file tạm sau 5 giây
            setTimeout(function() {
              if (currentTempSpecFileId) {
                google.script.run.deleteTempFile(currentTempSpecFileId);
                currentTempSpecFileId = null;
              }
            }, 5000);
          })
          .withFailureHandler(function(err) {
            btn.disabled = false;
            btn.innerText = '📥 Tải file Excel (.xlsx) về máy';
            showError('Lỗi tải file: ' + err.message);
          })
          .prepareCleanSpecialDownloadUrl(currentCreatedSpecSheet);
      }
  
      function closeDialogCleanly() {
        if (currentTempSpecFileId) {
          google.script.run.deleteTempFile(currentTempSpecFileId);
          currentTempSpecFileId = null;
        }
        google.script.host.close();
      }
    </script>
  </body>
  </html>
  `;
}