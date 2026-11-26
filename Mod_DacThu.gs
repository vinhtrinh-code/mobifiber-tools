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
  // Bảng 5 (ngoại lệ) trong thamso_caidat + phân loại thống nhất với tab Thuê bao BCA / Nghiệp vụ
  const dtTs = ss.getSheetByName(DT5.TS_SHEET);
  dt5EnsureExceptionTable_(dtTs);
  const dtEx = dt5ReadExceptions_(dtTs);
  const dtNewXM = [];
  
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
    if (isCancelled || rawStatusHD !== 'Đang hoạt động') continue;
  
    const rawLoaiKH = String(row[idxLoaiKH] || '').trim();
    const normLoaiKH = normalizeHeader(rawLoaiKH);
  
    const dtC = dt5Classify_(row[idxMaKH], rawLoaiKH, row[idxTenTB], row[idxGoiCuoc], dtEx);
    if (dtC.add) dtNewXM.push([String(row[idxMaKH] || '').replace(/^'+/, '').trim(), String(row[idxTenTB] || '')]);
    if (dtC.skip) continue;
    const isNghiepVu = dtC.grp === 'NV';
    const isBCA = dtC.grp === 'BCA';
  
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
  
  // Tự bắt thuê bao hành chính/LLVT chưa rõ vào Bảng 5 (Cần xác minh)
  const dtSeen = {};
  dtNewXM.forEach(function (x) {
    const nid = normalizeCustomerId(x[0]);
    if (dtEx[nid] || dtSeen[nid]) return;
    dtSeen[nid] = true;
    dt5AppendException_(dtTs, x[0], x[1], DT5.XL_XAC_MINH, 'Tự bắt: loại KH hành chính/LLVT, không có dấu hiệu công an');
  });

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
/* ===== MODULE THUE BAO DAC THU v5 (BCA / NGHIEP VU) - tab rieng, tu cap nhat, bang ngoai le trong thamso_caidat ===== */

var DT5 = {
  SHEET_BCA: 'Thuê bao BCA',
  SHEET_NV: 'Thuê bao Nghiệp vụ',
  TS_SHEET: 'thamso_caidat',
  EX_COL: 15,
  EX_TITLE_ROW: 2,
  EX_HEAD_ROW: 3,
  EX_FIRST_ROW: 4,
  EX_HEADERS: ['MÃ KHÁCH HÀNG', 'TÊN THUÊ BAO', 'XỬ LÝ', 'LÝ DO', 'NGÀY CẬP NHẬT'],
  XL_LOAI_TRU: 'Loại trừ',
  XL_XAC_MINH: 'Cần xác minh',
  XL_BCA: 'Tính BCA',
  XL_NV: 'Tính Nghiệp vụ',
  AREA_ORDER: ['TT.KDVT', 'KV Đà Lạt', 'KV Bảo Lộc', 'KV Gia Nghĩa', 'KV Phan Thiết']
};

function capNhatThuBaoDacThu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var res = buildThuBaoDacThu_(ss);
  SpreadsheetApp.getActive().toast('BCA: ' + res.bca + ' | Nghiệp vụ: ' + res.nv + ' | Loại trừ: ' + res.loaiTru + ' | Cần xác minh: ' + res.xacMinh, 'Thuê bao đặc thù', 8);
  return res;
}

function dt5ReadExceptions_(ts) {
  var map = {};
  var last = ts.getLastRow();
  if (last < DT5.EX_FIRST_ROW) return map;
  var vals = ts.getRange(DT5.EX_FIRST_ROW, DT5.EX_COL, last - DT5.EX_FIRST_ROW + 1, 5).getValues();
  for (var i = 0; i < vals.length; i++) {
    var id = normalizeCustomerId(String(vals[i][0] || '').replace(/^'+/, ''));
    if (id === '' || id === '0') continue;
    map[id] = { xuLy: String(vals[i][2] || '').trim(), ten: String(vals[i][1] || ''), row: DT5.EX_FIRST_ROW + i };
  }
  return map;
}

function dt5EnsureExceptionTable_(ts) {
  var c = DT5.EX_COL;
  var t = ts.getRange(DT5.EX_TITLE_ROW, c);
  if (!t.getValue()) {
    t.setValue('BẢNG 5: THUÊ BAO NGOẠI LỆ (LOẠI TRỪ / CẦN XÁC MINH / ÉP NHÓM)').setFontWeight('bold');
    ts.getRange(DT5.EX_HEAD_ROW, c, 1, 5).setValues([DT5.EX_HEADERS]).setFontWeight('bold').setBackground('#1B365D').setFontColor('#FFFFFF');
    ts.getRange(DT5.EX_FIRST_ROW, c, 500, 1).setNumberFormat('@');
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList([DT5.XL_LOAI_TRU, DT5.XL_XAC_MINH, DT5.XL_BCA, DT5.XL_NV], true).setAllowInvalid(false).build();
    ts.getRange(DT5.EX_FIRST_ROW, c + 2, 500, 1).setDataValidation(rule);
    ts.setColumnWidth(c, 120); ts.setColumnWidth(c + 1, 260); ts.setColumnWidth(c + 2, 120); ts.setColumnWidth(c + 3, 260); ts.setColumnWidth(c + 4, 110);
  }
}

function dt5AppendException_(ts, id, ten, xuLy, lyDo) {
  var last = ts.getLastRow();
  var colVals = last >= DT5.EX_FIRST_ROW ? ts.getRange(DT5.EX_FIRST_ROW, DT5.EX_COL, last - DT5.EX_FIRST_ROW + 1, 1).getValues() : [];
  var r = DT5.EX_FIRST_ROW;
  for (var i = 0; i < colVals.length; i++) { if (String(colVals[i][0] || '').trim() !== '') r = DT5.EX_FIRST_ROW + i + 1; }
  ts.getRange(r, DT5.EX_COL).setNumberFormat('@');
  ts.getRange(r, DT5.EX_COL, 1, 5).setValues([[id, ten, xuLy, lyDo, Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy')]]);
}

function dt5SeedExceptions_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ts = ss.getSheetByName(DT5.TS_SHEET);
  dt5EnsureExceptionTable_(ts);
  var ex = dt5ReadExceptions_(ts);
  var raw = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  var data = raw.getDataRange().getValues();
  var h = data[CONFIG.HEADER_ROW_TARGET - 1].map(normalizeHeader);
  var iId = h.indexOf(normalizeHeader('MÃ KHÁCH HÀNG')), iTen = h.indexOf(normalizeHeader('TÊN THUÊ BAO'));
  var n = 0;
  for (var i = CONFIG.DATA_START_ROW_TARGET - 1; i < data.length; i++) {
    var ten = String(data[i][iTen] || '');
    if (ten.toUpperCase().indexOf('VĂN PHÒNG HĐND VÀ UBND XÃ ĐẠ TẺH 2') === -1) continue;
    var id = normalizeCustomerId(data[i][iId]);
    if (ex[id]) continue;
    dt5AppendException_(ts, String(data[i][iId]).trim(), ten, DT5.XL_LOAI_TRU, 'UBND xã, không phải công an');
    n++;
  }
  return n;
}

function buildThuBaoDacThu_(ss) {
  var raw = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!raw) throw new Error('Không tìm thấy sheet raw_mobifiber');
  var ts = ss.getSheetByName(DT5.TS_SHEET);
  dt5EnsureExceptionTable_(ts);
  var ex = dt5ReadExceptions_(ts);
  dt6Ensure_(ts);
  dt7Ensure_(ts);
  [['Mũi Né', 'KV Phan Thiết'], ['Cát Tiên', 'KV Bảo Lộc'], ['Đạ Tẻh', 'KV Bảo Lộc'], ['Đơn Dương', 'KV Đà Lạt'], ['Đức Trọng', 'KV Đà Lạt'], ['Lạc Dương', 'KV Đà Lạt'],
   ['Gia Nghĩa', 'KV Gia Nghĩa'], ['Đắk Nông', 'KV Gia Nghĩa'], ['Đắk Mil', 'KV Gia Nghĩa'], ['Đắk Song', 'KV Gia Nghĩa'], ['Krông Nô', 'KV Gia Nghĩa'], ["Đắk R'Lấp", 'KV Gia Nghĩa'], ['Đắk Glong', 'KV Gia Nghĩa'], ['Cư Jút', 'KV Gia Nghĩa'], ['Tuy Đức', 'KV Gia Nghĩa']].forEach(function (p) { dt8AddIfMissing_(ts, p[0], p[1], ''); });
  var abbr = dt7ReadAbbr_(ts), kvRules = dt8Read_(ts);
  var ex6 = dt6Read_(ts), kvOld = {};
  var mapping = loadThamSoMapping();

  var data = raw.getDataRange().getValues();
  var h = data[CONFIG.HEADER_ROW_TARGET - 1].map(normalizeHeader);
  function col(name) { var k = h.indexOf(normalizeHeader(name)); if (k === -1) throw new Error('Thiếu cột ' + name); return k; }
  var iId = col('MÃ KHÁCH HÀNG'), iLoai = col('LOẠI KHÁCH HÀNG'), iTen = col('TÊN THUÊ BAO'), iGoi = col('GÓI CƯỚC'),
      iTt = col('TÌNH TRẠNG HỢP ĐỒNG'), iNv = col('MÃ NHÂN VIÊN'), iDt = col('DTPK'), iCh = col('TÊN CỬA HÀNG');
  var iDv = col('DỊCH VỤ'), iDnn = col('TÊN DOANH NGHIỆP'), iAddr = col('ĐỊA CHỈ LẮP ĐẶT'), iUser = col('USER'), iGia = col('GIÁ CƯỚC'), iDn = col('NGÀY ĐẤU NỐI'), iKt = col('NGÀY KẾT THÚC GÓI CƯỚC'), iBd;
  try { iBd = col('NGÀY BĂT ĐẦU GÓI CƯỚC'); } catch (e1) { iBd = col('NGÀY BẮT ĐẦU GÓI CƯỚC'); }

  var bca = [], nv = [], nBCAcho = 0, nNVhuy = 0, nLoaiTru = 0, nXM = 0, newXM = [];
  for (var i = CONFIG.DATA_START_ROW_TARGET - 1; i < data.length; i++) {
    var r = data[i];
    var rawId = String(r[iId] || '').trim();
    if (!rawId) continue;
    var id = normalizeCustomerId(rawId);
    var loai = String(r[iLoai] || '').toLowerCase();
    var ten = String(r[iTen] || '');
    var tenU = ten.toUpperCase();
    var goi = String(r[iGoi] || '').toUpperCase();
    var tt = String(r[iTt] || '').trim();

    var cls = dt5Classify_(rawId, r[iLoai], ten, r[iGoi], ex);
    var grp = cls.grp, pl = '';
    if (cls.skip) {
      if (cls.why === 'loaitru') nLoaiTru++;
      else if (cls.why === 'xm') nXM++;
      else if (cls.why === 'auto') { newXM.push([rawId, ten]); nXM++; }
      continue;
    }
    if (!grp) continue;
    if (grp === 'BCA') {
      if (/^(PHÒNG|BAN|CỤC|ĐỘI|TRẠI|ĐỒN|TRUNG TÂM|VĂN PHÒNG)\s/.test(tenU.trim())) pl = 'Phòng/Ban khác';
    else if (tenU.indexOf('CÔNG AN TỈNH') !== -1) pl = 'Công an tỉnh';
      else if (/CÔNG AN (PHƯỜNG|XÃ)/.test(tenU)) pl = 'Công an phường/xã';
      else if (tenU.indexOf('CÔNG AN') !== -1) pl = 'Phòng/Ban khác';
      else pl = 'CBCS / cá nhân';
    }
    if (tt !== 'Đang hoạt động') { if (grp === 'BCA') nBCAcho++; else nNVhuy++; continue; }

    var maNv = String(r[iNv] || '').replace(/^'+/, '').trim();
    var info = mapping.staffRegion.get(normalizeHeader(maNv));
    var area = info && info.khuVuc ? info.khuVuc : 'Chưa phân bổ KV';
    var id10 = dt6NormId_(rawId);
    kvOld[id10] = area;
    if (ex6[id10]) area = ex6[id10].kv;
    else if (grp !== 'BCA') { var kv8 = dt8Match_(kvRules, String(r[iDnn] || '').trim() || ten, String(r[iAddr] || '')); if (kv8) area = kv8; }
    var unitRaw = (String(r[iDnn] || '').trim() || ten.trim()), unitU = unitRaw.toUpperCase();
    var nhom = /test/i.test(ten) ? 'Test' : (unitU.indexOf('TRUNG TÂM MẠNG LƯỚI') !== -1 || /(^|[^A-Z])TTML([^A-Z]|$)/.test(unitU) ? 'TTML' : (unitU.indexOf('CỬA HÀNG') !== -1 || /^CH\s/.test(unitU) ? 'Cửa hàng' : (unitU.indexOf('TRUNG TÂM GIAO DỊCH') !== -1 || /(^|[^A-Z])TTGD([^A-Z]|$)/.test(unitU) ? 'Trung tâm GD' : (unitU.indexOf('VP MFT') !== -1 || /^MOBIFONE LÂM ĐỒNG\s*$/.test(unitU.trim()) ? 'VPCN' : 'Khác'))));
    var rec = { id: String(rawId).replace(/\D/g, '').padStart(10, '0'), ten: dt7Abbr_(abbr, ten.trim()), dnn: dt7Abbr_(abbr, unitRaw), addr: String(r[iAddr] || '').trim(), nhom: nhom, goi: String(r[iGoi] || '').trim(), loai: getShortValue(mapping.custTypes, r[iLoai]), user: String(r[iUser] || '').trim(), dv: String(r[iDv] || '').trim(), gia: parseCurrencyToInteger(r[iGia]) || 0, dn: parseToDate(r[iDn]), bd: parseToDate(r[iBd]), kt: parseToDate(r[iKt]), area: area, pl: pl, dt: parseCurrencyToInteger(r[iDt]) || 0, ch: getShortValue(mapping.stores, r[iCh]), test: /test/i.test(ten) };
    if (grp === 'BCA') bca.push(rec); else nv.push(rec);
  }

  dt6WriteOld_(ts, ex6, kvOld);
  var seen = {};
  for (var k = 0; k < newXM.length; k++) {
    var nid = normalizeCustomerId(newXM[k][0]);
    if (ex[nid] || seen[nid]) continue;
    seen[nid] = true;
    dt5AppendException_(ts, newXM[k][0], newXM[k][1], DT5.XL_XAC_MINH, 'Tự bắt: loại KH hành chính/LLVT, không có dấu hiệu công an');
  }

  var stamp = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm');
  dt5WriteBca_(ss, bca, stamp, nBCAcho, nLoaiTru, nXM);
  dt5WriteNv_(ss, nv, stamp, nNVhuy, nLoaiTru);
  return { bca: bca.length, nv: nv.length, loaiTru: nLoaiTru, xacMinh: nXM };
}

function dt5AreaList_(recs) {
  var set = {};
  recs.forEach(function (x) { set[x.area] = true; });
  var out = DT5.AREA_ORDER.filter(function (a) { return set[a]; });
  Object.keys(set).forEach(function (a) { if (out.indexOf(a) === -1) out.push(a); });
  return out;
}

function dt5Sheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear(); sh.clearFormats();
  sh.setHiddenGridlines(true);
  return sh;
}

function dt5Head_(sh, row, vals) {
  sh.getRange(row, 1, 1, vals.length).setValues([vals]).setFontWeight('bold').setBackground('#1B365D')
    .setFontColor('#FFFFFF').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
}

function chayThuNghiemDacThu() {
  var n = dt5SeedExceptions_();
  var res = capNhatThuBaoDacThu();
  return 'Bảng 5 thêm ' + n + ' dòng; ' + JSON.stringify(res);
}

/* Phan loai thong nhat (giu dong bo voi buildThuBaoDacThu_): tra ve {grp: 'NV'|'BCA'|null, skip, add} */
function dt5Classify_(rawId, loaiKH, tenTB, goi, ex) {
  var id = normalizeCustomerId(String(rawId || '').replace(/^'+/, ''));
  var loai = String(loaiKH || '').toLowerCase();
  var tenU = String(tenTB || '').toUpperCase();
  var g = String(goi || '').toUpperCase();
  var e = ex[id];
  if (e) {
    if (e.xuLy === DT5.XL_LOAI_TRU) return { grp: null, skip: true, add: false, why: 'loaitru' };
    if (e.xuLy === DT5.XL_XAC_MINH) return { grp: null, skip: true, add: false, why: 'xm' };
    if (e.xuLy === DT5.XL_NV) return { grp: 'NV', skip: false, add: false };
    if (e.xuLy === DT5.XL_BCA) return { grp: 'BCA', skip: false, add: false };
  }
  if (loai.indexOf('nghiệp vụ') !== -1) return { grp: 'NV', skip: false, add: false };
  if (g.indexOf('CBCS') !== -1 || loai === 'bca' || loai.indexOf('cán bộ, chiến sĩ') === 0 || tenU.indexOf('CÔNG AN') !== -1) return { grp: 'BCA', skip: false, add: false };
  if (loai.indexOf('vũ trang') !== -1 || loai.indexOf('hành chính') !== -1) return { grp: null, skip: true, add: true, why: 'auto' };
  return { grp: null, skip: true, add: false, why: 'none' };
}

function kiemThuDacThuLuyKe() {
  var r = processSpecialReport({ targetGroup: 'bca', timeType: 'all' });
  return JSON.stringify(r);
}


function taoLichCapNhatDacThu() {
  ScriptApp.getProjectTriggers().forEach(function(t) { if (t.getHandlerFunction() === 'capNhatThuBaoDacThu') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('capNhatThuBaoDacThu').timeBased().everyDays(1).atHour(7).nearMinute(30).create();
}


/* Danh sach chi tiet (de bao cao lanh dao) - ghi duoi cung tab Thue bao BCA / Nghiep vu */
function datLoaiTruCacMaUbnd() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(DT5.TS_SHEET);
  var last = sh.getLastRow();
  if (last < DT5.EX_FIRST_ROW) return 0;
  var rg = sh.getRange(DT5.EX_FIRST_ROW, DT5.EX_COL, last - DT5.EX_FIRST_ROW + 1, 5);
  var v = rg.getValues();
  var n = 0;
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][2]).trim() === DT5.XL_XAC_MINH && String(v[i][1]).toUpperCase().indexOf('UBND') !== -1) {
      v[i][2] = DT5.XL_LOAI_TRU;
      v[i][3] = 'Đã xác minh: UBND xã, không thuộc BCA';
      n++;
    }
  }
  if (n) rg.setValues(v);
  return n;
}

function dt5PlOrder_() { return ['Công an tỉnh', 'Công an phường/xã', 'CBCS / cá nhân', 'Phòng/Ban khác']; }

function dt5DetailSpec_(isBca) {
  function C(h, f, o) { o = o || {}; o.h = h; o.f = f; return o; }
  var cols = [C('KHU VỰC', function (x) { return x.area; }, { w: 130 })];
  if (isBca) cols.push(C('PHÂN LOẠI', function (x) { return x.pl; }, { w: 135 }));
  cols.push(C('MÃ KHÁCH HÀNG', function (x) { return x.id; }, { w: 115, text: true, center: true }));
  cols.push(C('USER', function (x) { return x.user; }, { w: 135, text: true }));
  cols.push(C('TÊN THUÊ BAO', function (x) { return x.ten; }, { w: 230 }));
  if (!isBca) {
    cols.push(C('ĐƠN VỊ SỬ DỤNG', function (x) { return x.dnn; }, { w: 200 }));
    cols.push(C('ĐỊA CHỈ TRIỂN KHAI', function (x) { return x.addr; }, { w: 320 }));
  }
  if (isBca) cols.push(C('LOẠI KHÁCH HÀNG', function (x) { return x.loai; }, { w: 160 }));
  cols.push(C('GÓI CƯỚC', function (x) { return x.goi; }, { w: 130 }));
  cols.push(C('DỊCH VỤ', function (x) { return x.dv; }, { w: 80, center: true }));
  if (isBca) {
    cols.push(C('GIÁ GÓI (đ)', function (x) { return x.gia || 0; }, { w: 110, fmt: '#,##0', right: true }));
    cols.push(C('DTPK (đ/tháng)', function (x) { return x.dt || 0; }, { w: 110, fmt: '#,##0', right: true }));
  }
  cols.push(C('NGÀY ĐẤU NỐI', function (x) { return x.dn || ''; }, { w: 105, fmt: 'dd/MM/yyyy', center: true }));
  cols.push(C('NGÀY BẮT ĐẦU', function (x) { return x.bd || ''; }, { w: 105, fmt: 'dd/MM/yyyy', center: true }));
  cols.push(C('NGÀY KẾT THÚC', function (x) { return x.kt || ''; }, { w: 105, fmt: 'dd/MM/yyyy', center: true, isEnd: true }));
  cols.push(C('CÒN LẠI (ngày)', null, { w: 95, fmt: '0', center: true, days: true }));
  if (isBca) cols.push(C('CỬA HÀNG', function (x) { return x.ch || ''; }, { w: 160 }));
  else cols.push(C('GHI CHÚ', function (x) { return x.test ? 'Thuê bao test' : (!x.addr ? 'Thiếu địa chỉ' : ''); }, { w: 130 }));
  return cols;
}

/* Danh sach chi tiet: sap xep Khu vuc -> Phan loai; co dinh dong tieu de; tra ve dong trong tiep theo */
function dt5WriteDetail_(sh, row0, recs, isBca, areas) {
  var cols = dt5DetailSpec_(isBca);
  var plo = dt5PlOrder_();
  var sorted = recs.slice().sort(function (a, b) {
    if (!isBca) { var tt = (a.test ? 1 : 0) - (b.test ? 1 : 0); if (tt) return tt; }
    var d = areas.indexOf(a.area) - areas.indexOf(b.area);
    if (d) return d;
    if (isBca) { var p = plo.indexOf(a.pl) - plo.indexOf(b.pl); if (p) return p; }
    if (!isBca) { var da = a.dn ? new Date(a.dn).getTime() : 0, db = b.dn ? new Date(b.dn).getTime() : 0; if (da !== db) return da - db; }
    var c = String(a.ten || '').localeCompare(String(b.ten || ''), 'vi');
    if (c) return c;
    return a.id < b.id ? -1 : (a.id > b.id ? 1 : 0);
  });
  var hdr = row0 + 1, first = row0 + 2, n = sorted.length, nc = cols.length;
  sh.getRange(row0, 1).setValue('DANH SÁCH CHI TIẾT (' + n + ' thuê bao)').setFontSize(12).setFontWeight('bold').setFontColor('#174EA6');
  dt5Head_(sh, hdr, cols.map(function (c) { return c.h; }));
  sh.setRowHeight(hdr, 42);
  var today0 = new Date(); today0.setHours(0, 0, 0, 0); today0 = today0.getTime();
  var endIdx = -1, dayIdx = -1;
  cols.forEach(function (c, i) { if (c.isEnd) endIdx = i; if (c.days) dayIdx = i; });
  var L = String.fromCharCode(65 + endIdx);
  var f0 = sh.getFilter(); if (f0) f0.remove();
  if (n) {
    var rows = sorted.map(function (x, k) {
      return cols.map(function (c) {
        if (c.days) return x.kt ? Math.floor((new Date(x.kt).getTime() - today0) / 86400000) : '';
        return c.f(x);
      });
    });
    cols.forEach(function (c, i) {
      var cr = sh.getRange(first, i + 1, n, 1);
      if (c.text) cr.setNumberFormat('@'); else if (c.fmt) cr.setNumberFormat(c.fmt);
    });
    var rg = sh.getRange(first, 1, n, nc);
    rg.setValues(rows).setVerticalAlignment('middle').setWrap(true);
    rg.setBorder(true, true, true, true, true, true, '#DADCE0', SpreadsheetApp.BorderStyle.SOLID);
    cols.forEach(function (c, i) {
      if (c.center) sh.getRange(first, i + 1, n, 1).setHorizontalAlignment('center');
      if (c.right) sh.getRange(first, i + 1, n, 1).setHorizontalAlignment('right');
    });
    var dr = sh.getRange(first, dayIdx + 1, n, 1);
    var r1 = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThanOrEqualTo(0).setBackground('#F4C7C3').setFontColor('#B31412').setRanges([dr]).build();
    var r2 = SpreadsheetApp.newConditionalFormatRule().whenNumberBetween(1, 30).setBackground('#FCE8B2').setFontColor('#B06000').setRanges([dr]).build();
    sh.setConditionalFormatRules([r1, r2]);
    sh.getRange(hdr, 1, n + 1, nc).createFilter();
  }
  cols.forEach(function (c, i) { sh.setColumnWidth(i + 1, c.w); });
  sh.setFrozenRows(hdr);
  return first + n;
}

function dt5WriteBca_(ss, recs, stamp, nCho, nLoaiTru, nXM) {
  var sh = dt5Sheet_(ss, DT5.SHEET_BCA);
  var cats = dt5PlOrder_();
  sh.getRange('A1').setValue('THUÊ BAO BCA (CÔNG AN, CBCS) - THEO KHU VỰC').setFontSize(14).setFontWeight('bold').setFontColor('#174EA6');
  sh.getRange('A2').setValue('Cập nhật: ' + stamp + '  |  Chỉ tính thuê bao Đang hoạt động  |  Nguồn: raw_mobifiber + Bảng 5 thamso_caidat').setFontColor('#5F6368').setFontStyle('italic');
  var head = ['KHU VỰC'].concat(cats).concat(['TỔNG TB', 'TỶ TRỌNG TB (%)', 'DTPK (đ/tháng)', 'TỶ TRỌNG DTPK (%)']);
  dt5Head_(sh, 4, head); sh.setRowHeight(4, 42);
  var areas = dt5AreaList_(recs);
  var totTb = recs.length, totDt = recs.reduce(function (s, x) { return s + (x.dt || 0); }, 0);
  var rows = areas.map(function (a) {
    var sub = recs.filter(function (x) { return x.area === a; });
    var row = [a];
    cats.forEach(function (c) { row.push(sub.filter(function (x) { return x.pl === c; }).length); });
    var dt = sub.reduce(function (s, x) { return s + (x.dt || 0); }, 0);
    row.push(sub.length, totTb ? sub.length / totTb : 0, dt, totDt ? dt / totDt : 0);
    return row;
  });
  var tot = ['TỔNG CỘNG'];
  cats.forEach(function (c) { tot.push(recs.filter(function (x) { return x.pl === c; }).length); });
  tot.push(totTb, 1, totDt, 1);
  rows.push(tot);
  var r0 = 5;
  var rg = sh.getRange(r0, 1, rows.length, head.length);
  rg.setValues(rows).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(r0, 1, rows.length, 1).setHorizontalAlignment('left');
  sh.getRange(r0, 7, rows.length, 1).setNumberFormat('0.0%');
  sh.getRange(r0, 8, rows.length, 1).setNumberFormat('#,##0');
  sh.getRange(r0, 9, rows.length, 1).setNumberFormat('0.0%');
  sh.getRange(r0 + rows.length - 1, 1, 1, head.length).setFontWeight('bold').setBackground('#E8F0FE');
  rg.setBorder(true, true, true, true, true, true, '#DADCE0', SpreadsheetApp.BorderStyle.SOLID);
  var next = dt5WriteDetail_(sh, r0 + rows.length + 1, recs, true, areas);
  var n = next + 1;
  sh.getRange(n, 1).setValue('Ghi chú').setFontWeight('bold');
  sh.getRange(n + 1, 1).setValue('• Thuê bao BCA đang chờ thi công (chưa tính): ' + nCho);
  sh.getRange(n + 2, 1).setValue('• Thuê bao loại trừ theo Bảng 5 (thamso_caidat): ' + nLoaiTru);
  sh.getRange(n + 3, 1).setValue('• Thuê bao cần xác minh (chưa tính vào BCA hay Nghiệp vụ): ' + nXM + '  →  xem Bảng 5, cột XỬ LÝ');
  sh.getRange(n + 4, 1).setValue('• Nhận diện BCA: gói CBCS, hoặc loại KH BCA / Cán bộ chiến sĩ, hoặc tên chứa "Công an".');
  sh.getRange(n + 5, 1).setValue('• Phòng/Ban khác: tên bắt đầu bằng Phòng, Ban, Cục, Đội, Trại, Đồn, Trung tâm, Văn phòng; hoặc có chữ "Công an" nhưng không phải Công an tỉnh / phường / xã.');
  sh.getRange(n + 6, 1).setValue('• CÒN LẠI (ngày) = Ngày kết thúc − hôm nay: tô cam khi còn ≤ 30 ngày, tô đỏ khi đã hết hạn. Thuê bao trả sau (chu kỳ 1) không có ngày kết thúc.');
}


/* ===== BANG 6 (thamso_caidat U:Z): DIEU CHINH KHU VUC - ngoai le, khong sua du lieu goc ===== */
var DT6 = { COL: 21, NCOL: 6, TITLE_ROW: 2, HEAD_ROW: 3, FIRST_ROW: 4 };

function dt6NormId_(v) { var s = String(v || '').replace(/\D/g, ''); return s ? s.padStart(10, '0') : ''; }

function dt6Ensure_(ts) {
  if (ts.getMaxColumns() < DT6.COL + DT6.NCOL - 1) ts.insertColumnsAfter(ts.getMaxColumns(), DT6.COL + DT6.NCOL - 1 - ts.getMaxColumns());
  var head = ['MÃ KHÁCH HÀNG', 'TÊN THUÊ BAO', 'KV CŨ (tự điền)', 'KV MỚI (chọn)', 'GHI CHÚ / LÝ DO', 'NGÀY CẬP NHẬT'];
  var cur = String(ts.getRange(DT6.HEAD_ROW, DT6.COL).getValue() || '');
  if (cur === head[0]) return;
  ts.getRange(DT6.TITLE_ROW, DT6.COL).setValue('BẢNG 6: ĐIỀU CHỈNH KHU VỰC THEO THUÊ BAO (NGOẠI LỆ - KHÔNG SỬA DỮ LIỆU GỐC)').setFontWeight('bold').setFontColor('#174EA6');
  var hr = ts.getRange(DT6.HEAD_ROW, DT6.COL, 1, DT6.NCOL);
  hr.setValues([head]).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1B365D').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  ts.getRange(DT6.FIRST_ROW, DT6.COL, 200, 1).setNumberFormat('@');
  ts.getRange(DT6.FIRST_ROW, DT6.COL + 5, 200, 1).setNumberFormat('dd/MM/yyyy');
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(DT5.AREA_ORDER, true).setAllowInvalid(false).build();
  ts.getRange(DT6.FIRST_ROW, DT6.COL + 3, 200, 1).setDataValidation(rule);
  [130, 260, 120, 130, 380, 105].forEach(function (w, i) { ts.setColumnWidth(DT6.COL + i, w); });
  ts.getRange(DT6.FIRST_ROW, DT6.COL, 1, DT6.NCOL).setValues([['0046836447', 'CÔNG AN PHƯỜNG ĐÔNG GIA NGHĨA', '', 'KV Gia Nghĩa', 'Gán về KV Gia Nghĩa theo yêu cầu quản lý (không sửa dữ liệu gốc raw_mobifiber)', new Date()]]);
}

function dt6Read_(ts) {
  var last = ts.getLastRow(), map = {};
  if (last < DT6.FIRST_ROW) return map;
  var v = ts.getRange(DT6.FIRST_ROW, DT6.COL, last - DT6.FIRST_ROW + 1, DT6.NCOL).getValues();
  v.forEach(function (r, i) {
    var id = dt6NormId_(r[0]), kv = String(r[3] || '').trim();
    if (id && kv) map[id] = { kv: kv, row: DT6.FIRST_ROW + i };
  });
  return map;
}

function dt6WriteOld_(ts, ex6, old) {
  Object.keys(ex6).forEach(function (id) {
    var o = old[id];
    ts.getRange(ex6[id].row, DT6.COL + 2).setValue(o !== undefined ? o : '(không có trong danh sách BCA/Nghiệp vụ đang hoạt động)');
  });
}


/* ===== BANG 7 (AB:AC) viet tat ten don vi; BANG 8 (AE:AG) tu khoa dia ban -> khu vuc ===== */
var DT7 = { COL: 28, NCOL: 2 }, DT8 = { COL: 31, NCOL: 3 };

function dt7Table_(ts, spec, title, head, seeds, widths, kvCol) {
  if (ts.getMaxColumns() < spec.COL + spec.NCOL - 1) ts.insertColumnsAfter(ts.getMaxColumns(), spec.COL + spec.NCOL - 1 - ts.getMaxColumns());
  if (String(ts.getRange(3, spec.COL).getValue() || '') === head[0]) return;
  ts.getRange(2, spec.COL).setValue(title).setFontWeight('bold').setFontColor('#174EA6');
  ts.getRange(3, spec.COL, 1, spec.NCOL).setValues([head]).setFontWeight('bold').setFontColor('#FFFFFF').setBackground('#1B365D').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  if (seeds.length) ts.getRange(4, spec.COL, seeds.length, spec.NCOL).setValues(seeds);
  widths.forEach(function (w, i) { ts.setColumnWidth(spec.COL + i, w); });
  if (kvCol) ts.getRange(4, spec.COL + kvCol - 1, 200, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(DT5.AREA_ORDER, true).setAllowInvalid(false).build());
}

function dt7Ensure_(ts) {
  dt7Table_(ts, DT7, 'BẢNG 7: VIẾT TẮT TÊN ĐƠN VỊ (tab Thuê bao BCA / Nghiệp vụ)', ['CỤM TỪ GỐC', 'VIẾT TẮT'], [
    ['TRUNG TÂM MẠNG LƯỚI', 'TTML'],
    ['CHI NHÁNH TỔNG CÔNG TY VIỄN THÔNG MOBIFONE', 'CN TCT MOBIFONE'],
    ['TRUNG TÂM GIAO DỊCH', 'TTGD'],
    ['TRUNG TÂM KINH DOANH VIỄN THÔNG', 'TT.KDVT'],
    ['CỬA HÀNG', 'CH']], [300, 150], 0);
  dt7Table_(ts, DT8, 'BẢNG 8: TỪ KHÓA ĐỊA BÀN → KHU VỰC (thuê bao nghiệp vụ thuộc TT.KDVT)', ['TỪ KHÓA (tên đơn vị / địa chỉ)', 'KHU VỰC', 'GHI CHÚ'], [
    ['Phan Thiết', 'KV Phan Thiết', 'Ưu tiên tìm trong tên đơn vị trước, địa chỉ sau; khớp dòng đầu tiên'],
    ['Bắc Bình', 'KV Phan Thiết', ''], ['Hàm Thuận', 'KV Phan Thiết', 'gồm Hàm Thuận Nam'], ['Tánh Linh', 'KV Phan Thiết', ''],
    ['La Gi', 'KV Phan Thiết', ''], ['Lagi', 'KV Phan Thiết', ''], ['Hàm Tân', 'KV Phan Thiết', ''], ['Đức Linh', 'KV Phan Thiết', ''], ['Liên Hương', 'KV Phan Thiết', ''],
    ['Bảo Lộc', 'KV Bảo Lộc', ''], ['Bảo Lâm', 'KV Bảo Lộc', ''], ['Di Linh', 'KV Bảo Lộc', ''], ['Đạ Huoai', 'KV Bảo Lộc', ''],
    ['Xuân Hương', 'KV Đà Lạt', ''], ['Đà Lạt', 'KV Đà Lạt', '']], [230, 120, 330], 2);
}

function dt8AddIfMissing_(ts, k, kv, note) {
  var rows = dt7Read_(ts, DT8);
  for (var i = 0; i < rows.length; i++) if (String(rows[i][0]).trim().toLowerCase() === k.toLowerCase()) return;
  var r = 4; for (var j = rows.length - 1; j >= 0; j--) if (String(rows[j][0]).trim()) { r = 4 + j + 1; break; }
  ts.getRange(r, DT8.COL, 1, DT8.NCOL).setValues([[k, kv, note || '']]);
}

function dt7Read_(ts, spec) {
  var last = ts.getLastRow(); if (last < 4) return [];
  return ts.getRange(4, spec.COL, last - 3, spec.NCOL).getValues();
}

function dt7ReadAbbr_(ts) {
  var rows = dt7Read_(ts, DT7).filter(function (r) { return String(r[0] || '').trim() && String(r[1] || '').trim(); });
  rows.sort(function (a, b) { return String(b[0]).length - String(a[0]).length; });
  return rows.map(function (r) {
    return { re: new RegExp(String(r[0]).trim().replace(/[.*+?^$()|[\]\\{}]/g, '\\$&'), 'gi'), to: String(r[1]).trim() };
  });
}

function dt7Abbr_(list, s) {
  s = String(s || '');
  list.forEach(function (x) { s = s.replace(x.re, x.to); });
  return s.replace(/\s+/g, ' ').trim();
}

function dt8Read_(ts) {
  return dt7Read_(ts, DT8).filter(function (r) { return String(r[0] || '').trim() && String(r[1] || '').trim(); })
    .map(function (r) { return { k: String(r[0]).trim().toLowerCase(), kv: String(r[1]).trim() }; });
}

function dt8Match_(rules, unit, addr) {
  var src = [String(unit || '').toLowerCase(), String(addr || '').toLowerCase()];
  for (var p = 0; p < src.length; p++) {
    if (!src[p]) continue;
    for (var i = 0; i < rules.length; i++) if (src[p].indexOf(rules[i].k) !== -1) return rules[i].kv;
  }
  return '';
}

function dt5WriteNv_(ss, recs, stamp, nHuy, nLoaiTru) {
  var sh = dt5Sheet_(ss, DT5.SHEET_NV);
  sh.getRange('A1').setValue('THUÊ BAO NGHIỆP VỤ (PHỤC VỤ NỘI BỘ - KHÔNG GHI NHẬN DOANH THU)').setFontSize(14).setFontWeight('bold').setFontColor('#174EA6');
  sh.getRange('A2').setValue('Cập nhật: ' + stamp + '  |  Chỉ tính thuê bao Đang hoạt động  |  Nguồn: raw_mobifiber + Bảng 5, 6, 7, 8 thamso_caidat').setFontColor('#5F6368').setFontStyle('italic');
  var groups = ['Cửa hàng', 'Trung tâm GD', 'TTML', 'VPCN', 'Khác', 'Test'];
  var head = ['KHU VỰC'].concat(groups).concat(['TỔNG TB', 'TỶ TRỌNG (%)']);
  dt5Head_(sh, 4, head); sh.setRowHeight(4, 42);
  var areas = dt5AreaList_(recs), tot = recs.length;
  var rows = areas.map(function (a) {
    var sub = recs.filter(function (x) { return x.area === a; });
    var row = [a];
    groups.forEach(function (g) { row.push(sub.filter(function (x) { return x.nhom === g; }).length); });
    row.push(sub.length, tot ? sub.length / tot : 0);
    return row;
  });
  var tr = ['TỔNG CỘNG'];
  groups.forEach(function (g) { tr.push(recs.filter(function (x) { return x.nhom === g; }).length); });
  tr.push(tot, 1); rows.push(tr);
  var r0 = 5;
  var rg = sh.getRange(r0, 1, rows.length, head.length);
  rg.setNumberFormat('0');
  rg.setValues(rows).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(r0, 1, rows.length, 1).setHorizontalAlignment('left');
  sh.getRange(r0, 9, rows.length, 1).setNumberFormat('0.0%');
  sh.getRange(r0 + rows.length - 1, 1, 1, head.length).setFontWeight('bold').setBackground('#E8F0FE');
  rg.setBorder(true, true, true, true, true, true, '#DADCE0', SpreadsheetApp.BorderStyle.SOLID);
  var next = dt5WriteDetail_(sh, r0 + rows.length + 1, recs, false, areas);
  var n = next + 1;
  sh.getRange(n, 1).setValue('Ghi chú').setFontWeight('bold');
  sh.getRange(n + 1, 1).setValue('• Thuê bao nghiệp vụ đã hủy (không tính): ' + nHuy);
  sh.getRange(n + 2, 1).setValue('• Khu vực: theo mã nhân viên tạo thuê bao (Bảng 4); khu vực nghiệp vụ được xác định theo tên đơn vị / địa chỉ lắp đặt (Bảng 8); không khớp từ khóa thì giữ khu vực của nhân viên tạo. Gán lại từng mã: Bảng 6.');
  sh.getRange(n + 3, 1).setValue('• Nhóm theo tên đơn vị: Cửa hàng, Trung tâm GD, TTML, VPCN (VP MFT và MobiFone Lâm Đồng chung), Khác; Test = tên thuê bao có chữ "test". Tên đơn vị viết tắt theo Bảng 7.');
  sh.getRange(n + 4, 1).setValue('• Địa chỉ lấy từ ĐỊA CHỈ LẮP ĐẶT trong dữ liệu gốc; dòng ghi "Thiếu địa chỉ" cần tra cứu và bổ sung.');
  sh.getRange(n + 5, 1).setValue('• CÒN LẠI (ngày) = Ngày kết thúc − hôm nay: tô cam khi còn ≤ 30 ngày, tô đỏ khi đã hết hạn.');
}
