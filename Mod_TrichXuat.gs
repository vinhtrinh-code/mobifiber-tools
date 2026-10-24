/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * MODULE       : TRÍCH XUẤT DANH SÁCH THUÊ BAO ĐA ĐIỀU KIỆN
 * PHIÊN BẢN    : 1.0 (Lọc Đa Tiêu Chí Động, Bảng Chi Tiết 20 Cột Chuẩn, Tải Excel Sạch)
 * NGÀY CẬP NHẬT: 20/09/2026 21:45:00 (GMT+7)
 * TÁC GIẢ      : Trịnh Vũ Thế Vinh & AI Assistant
 * MÔ TẢ        : 
 *   - Lọc đa chiều: Cửa hàng, Nhân viên, Phường xã, Loại KH, Tình trạng HĐ, Chặn cắt, Gói cước
 *   - Danh sách chi tiết 20 cột chuẩn: Đầy đủ hợp đồng, chặn cắt, thiết bị ONT và vòng đời ngày tháng
 *   - Dashboard KPI ở đầu trang: Tổng TB, Doanh thu gói cước, DTPK
 *   - Tự động co giãn cột (max 300px), lề trái thoáng mắt, chống lỗi ### trên Microsoft Excel
 *   - Tải file Excel (.xlsx) sạch qua iframe ngầm, tự động xóa ô nút tải trên file tải về
 * ============================================================================
 */

function showQueryDialog() {
  const html = HtmlService.createHtmlOutput(getQueryDialogHtml())
    .setWidth(540)
    .setHeight(560)
    .setTitle('Trích xuất Danh sách Đa điều kiện');
  SpreadsheetApp.getUi().showModalDialog(html, 'Trích xuất Danh sách Thuê bao MobiFiber');
}

/**
 * Lấy danh sách các giá trị duy nhất (Unique values) từ sheet raw_mobifiber để nạp vào Dropdown
 */
function getFilterDropdownData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) return { stores: [], custTypes: [], contractStatuses: [], blockStatuses: [] };

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) {
    return { stores: [], custTypes: [], contractStatuses: [], blockStatuses: [] };
  }

  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol)
    .getValues()[0].map(h => normalizeHeader(h));

  const idxTenCH       = headers.findIndex(h => h === 'tên cửa hàng');
  const idxLoaiKH      = headers.findIndex(h => h === 'loại khách hàng');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxChanCat     = headers.findIndex(h => h === 'tình trạng chặn cắt');

  const numRows = lastRow - CONFIG.DATA_START_ROW_TARGET + 1;
  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, numRows, lastCol).getValues();

  const storeSet = new Set();
  const custTypeSet = new Set();
  const contractStatusSet = new Set();
  const blockStatusSet = new Set();

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    if (idxTenCH !== -1 && row[idxTenCH]) storeSet.add(String(row[idxTenCH]).trim());
    if (idxLoaiKH !== -1 && row[idxLoaiKH]) custTypeSet.add(String(row[idxLoaiKH]).trim());
    if (idxTinhTrangHD !== -1 && row[idxTinhTrangHD]) contractStatusSet.add(String(row[idxTinhTrangHD]).trim());
    if (idxChanCat !== -1 && row[idxChanCat]) blockStatusSet.add(String(row[idxChanCat]).trim());
  }

  return {
    stores: Array.from(storeSet).sort(),
    custTypes: Array.from(custTypeSet).sort(),
    contractStatuses: Array.from(contractStatusSet).sort(),
    blockStatuses: Array.from(blockStatusSet).sort()
  };
}

/**
 * Xử lý lọc dữ liệu theo tổ hợp điều kiện và tạo trang báo cáo
 */
function processAdvancedQuery(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) throw new Error('Bảng raw_mobifiber chưa có dữ liệu.');

  // Đọc Headers
  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol)
    .getValues()[0].map(h => normalizeHeader(h));

  const idxMaKH        = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxLoaiKH      = headers.findIndex(h => h === 'loại khách hàng');
  const idxTenTB       = headers.findIndex(h => h === 'tên thuê bao');
  const idxSDT         = headers.findIndex(h => h === 'sđt liên hệ');
  const idxDiaChi      = headers.findIndex(h => h === 'địa chỉ lắp đặt');
  const idxPhuongXa    = headers.findIndex(h => h === 'phường xã' || h === 'mã phường xã');
  const idxGoiCuoc     = headers.findIndex(h => h === 'gói cước');
  const idxGiaCuoc     = headers.findIndex(h => h === 'giá cước');
  const idxChuKy       = headers.findIndex(h => h === 'chu kỳ sử dụng');
  const idxDtpk        = headers.findIndex(h => h === 'dtpk');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxChanCat     = headers.findIndex(h => h === 'tình trạng chặn cắt');
  const idxThietBi     = headers.findIndex(h => h === 'loại thiết bị ont');
  const idxSerial      = headers.findIndex(h => h === 'serial ont');
  const idxNgayDauNoi  = headers.findIndex(h => h === 'ngày đấu nối');
  const idxNgayThiCong = headers.findIndex(h => h === 'ngày hoàn tất thi công');
  const idxNgayBatDau  = headers.findIndex(h => h === 'ngày băt đầu gói cước' || h === 'ngày bắt đầu gói cước');
  const idxNgayKetThuc = headers.findIndex(h => h === 'ngày kết thúc gói cước');
  const idxTenCH       = headers.findIndex(h => h === 'tên cửa hàng');
  const idxMaNV        = headers.findIndex(h => h === 'mã nhân viên');
  const idxTenNV       = headers.findIndex(h => h === 'tên nhân viên');

  const numRows = lastRow - CONFIG.DATA_START_ROW_TARGET + 1;
  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, numRows, lastCol).getValues();

  // Chuẩn hóa điều kiện lọc
  const filterStore    = params.store ? normalizeHeader(params.store) : '';
  const filterStaff    = params.staff ? normalizeHeader(params.staff) : '';
  const filterWard     = params.ward ? normalizeHeader(params.ward) : '';
  const filterCustType = params.custType ? normalizeHeader(params.custType) : '';
  const filterStatusHD = params.statusHD ? normalizeHeader(params.statusHD) : '';
  const filterStatusBlock = params.statusBlock ? normalizeHeader(params.statusBlock) : '';
  const filterPackage  = params.pkg ? normalizeHeader(params.pkg) : '';

  // Phân tích ngày (nếu có)
  let startDate = null, endDate = null;
  if (params.from && params.to) {
    const sParts = params.from.split('/');
    const eParts = params.to.split('/');
    startDate = new Date(parseInt(sParts[2], 10), parseInt(sParts[1], 10) - 1, parseInt(sParts[0], 10), 0, 0, 0);
    endDate = new Date(parseInt(eParts[2], 10), parseInt(eParts[1], 10) - 1, parseInt(eParts[0], 10), 23, 59, 59);
  }

  const matchedRows = [];
  let totalRevenue = 0;
  let totalDtpk = 0;

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];

    // 1. Lọc Cửa hàng
    if (filterStore) {
      const storeVal = normalizeHeader(row[idxTenCH]);
      if (storeVal !== filterStore) continue;
    }

    // 2. Lọc Nhân viên (Mã hoặc Tên)
    if (filterStaff) {
      const staffCode = normalizeHeader(row[idxMaNV]);
      const staffName = normalizeHeader(row[idxTenNV]);
      if (!staffCode.includes(filterStaff) && !staffName.includes(filterStaff)) continue;
    }

    // 3. Lọc Phường xã
    if (filterWard) {
      const wardVal = normalizeHeader(row[idxPhuongXa]);
      const addrVal = normalizeHeader(row[idxDiaChi]);
      if (!wardVal.includes(filterWard) && !addrVal.includes(filterWard)) continue;
    }

    // 4. Lọc Loại khách hàng
    if (filterCustType) {
      const cTypeVal = normalizeHeader(row[idxLoaiKH]);
      if (cTypeVal !== filterCustType) continue;
    }

    // 5. Lọc Tình trạng hợp đồng
    if (filterStatusHD) {
      const sHdVal = normalizeHeader(row[idxTinhTrangHD]);
      if (sHdVal !== filterStatusHD) continue;
    }

    // 6. Lọc Tình trạng chặn cắt
    if (filterStatusBlock) {
      const sBlockVal = normalizeHeader(row[idxChanCat]);
      if (sBlockVal !== filterStatusBlock) continue;
    }

    // 7. Lọc Gói cước
    if (filterPackage) {
      const pkgVal = normalizeHeader(row[idxGoiCuoc]);
      if (!pkgVal.includes(filterPackage)) continue;
    }

    // 8. Lọc Thời gian (theo Ngày đấu nối hoặc Ngày bắt đầu)
    if (startDate && endDate) {
      const dDauNoi = parseToDate(row[idxNgayDauNoi]);
      const dBatDau = parseToDate(row[idxNgayBatDau]);
      const checkDate = dDauNoi || dBatDau;
      if (!checkDate || checkDate < startDate || checkDate > endDate) continue;
    }

    const giaCuoc = parseCurrencyToInteger(row[idxGiaCuoc]) || 0;
    const chuKy = row[idxChuKy] ? parseInt(row[idxChuKy], 10) : 0;
    const dtpk = parseCurrencyToInteger(row[idxDtpk]) || (chuKy > 0 ? Math.round(giaCuoc / chuKy) : 0);
    totalRevenue += giaCuoc;
    totalDtpk += dtpk;

    let cleanMaKH = String(row[idxMaKH] || '').replace(/^'+/, '').trim();
    if (/^\d{8}$/.test(cleanMaKH)) cleanMaKH = '00' + cleanMaKH;

    let cleanSDT = String(row[idxSDT] || '').replace(/^'+/, '').trim();
    if (/^\d{9}$/.test(cleanSDT)) cleanSDT = '0' + cleanSDT;

    const paddedLoaiKH  = row[idxLoaiKH] ? '  ' + String(row[idxLoaiKH]).trim() : '';
    const paddedTenTB   = row[idxTenTB] ? '  ' + String(row[idxTenTB]).trim() : '';
    const paddedDiaChi  = row[idxDiaChi] ? '  ' + String(row[idxDiaChi]).trim() : '';
    const paddedThietBi = row[idxThietBi] ? '  ' + String(row[idxThietBi]).trim() : '';
    const paddedTenCH   = row[idxTenCH] ? '  ' + String(row[idxTenCH]).trim() : '';
    const paddedTenNV   = row[idxTenNV] ? '  ' + String(row[idxTenNV]).trim() : '';
    const cleanSerial   = String(row[idxSerial] || '').replace(/^'+/, '').trim();

    matchedRows.push([
      matchedRows.length + 1,                             // 1. STT
      "'" + cleanMaKH,                                    // 2. MÃ KHÁCH HÀNG
      paddedLoaiKH,                                       // 3. LOẠI KHÁCH HÀNG
      paddedTenTB,                                        // 4. TÊN THUÊ BAO
      "'" + cleanSDT,                                     // 5. SĐT LIÊN HỆ
      paddedDiaChi,                                       // 6. ĐỊA CHỈ LẮP ĐẶT
      row[idxGoiCuoc] || '',                              // 7. GÓI CƯỚC
      giaCuoc,                                            // 8. GIÁ CƯỚC
      chuKy > 0 ? chuKy : '',                             // 9. CHU KỲ
      dtpk,                                               // 10. DTPK
      row[idxTinhTrangHD] || 'Đang hoạt động',             // 11. TÌNH TRẠNG HỢP ĐỒNG
      row[idxChanCat] || 'Bình thường',                   // 12. TÌNH TRẠNG CHẶN CẮT
      paddedThietBi,                                      // 13. LOẠI THIẾT BỊ ONT
      "'" + cleanSerial,                                  // 14. SERIAL ONT
      formatDateTimeStr(row[idxNgayDauNoi]),              // 15. NGÀY ĐẤU NỐI
      formatDateTimeStr(row[idxNgayThiCong]),             // 16. NGÀY HOÀN TẤT THI CÔNG
      formatDateTimeStr(row[idxNgayBatDau]),              // 17. NGÀY BẮT ĐẦU GÓI CƯỚC
      formatDateTimeStr(row[idxNgayKetThuc]),             // 18. NGÀY KẾT THÚC GÓI CƯỚC
      paddedTenCH,                                        // 19. TÊN CỬA HÀNG
      paddedTenNV                                         // 20. TÊN NHÂN VIÊN
    ]);
  }

  // Tên sheet động theo tiêu chí
  const now = new Date();
  const timeStamp = Utilities.formatDate(now, 'Asia/Saigon', 'ddMM_HHmm');
  let slug = 'TuyChon';
  if (params.store) slug = 'CH_' + cleanSlug(params.store);
  else if (params.staff) slug = 'NV_' + cleanSlug(params.staff);
  else if (params.custType) slug = 'LoaiKH_' + cleanSlug(params.custType);
  else if (params.ward) slug = 'PX_' + cleanSlug(params.ward);

  const sheetName = `Loc_${slug}_${timeStamp}`.substring(0, 30);

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

  // DÒNG 1: Tiêu đề lớn (Trải rộng 20 cột)
  destSheet.getRange(1, 1, 1, 20).merge()
    .setValue(`BÁO CÁO TRÍCH XUẤT THUÊ BAO MOBIFIBER THEO ĐIỀU KIỆN LỌC`)
    .setFontWeight('bold').setFontSize(13).setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');

  // DÒNG 2: TỔNG THUÊ BAO & NGÀY LẬP
  destSheet.getRange(2, 1, 1, 9).merge()
    .setValue(`🎯 TỔNG SỐ THUÊ BAO: ${matchedRows.length} TB`)
    .setFontWeight('bold').setFontSize(11).setBackground('#D2E3FC').setFontColor('#174EA6').setHorizontalAlignment('left');

  destSheet.getRange(2, 10, 1, 6).merge()
    .setValue(`🕒 Ngày lập: ${Utilities.formatDate(now, 'Asia/Saigon', 'dd/MM/yyyy HH:mm')}`)
    .setFontSize(10).setBackground('#D2E3FC').setFontColor('#5F6368').setHorizontalAlignment('right');

  // DÒNG 3: TỔNG DOANH THU GÓI CƯỚC
  destSheet.getRange(3, 1, 1, 15).merge()
    .setValue(`💰 TỔNG DOANH THU GÓI CƯỚC: ${totalRevenue.toLocaleString('vi-VN')} VNĐ`)
    .setFontWeight('bold').setFontSize(11).setBackground('#E8F0FE').setFontColor('#174EA6').setHorizontalAlignment('left');

  // DÒNG 4: TỔNG DTPK
  destSheet.getRange(4, 1, 1, 15).merge()
    .setValue(`📈 TỔNG DOANH THU PHÂN KỲ (DTPK): ${totalDtpk.toLocaleString('vi-VN')} VNĐ`)
    .setFontWeight('bold').setFontSize(11).setBackground('#FEF7E0').setFontColor('#B06000').setHorizontalAlignment('left');

  // NÚT BẤM TẢI FILE EXCEL TRÊN SHEET (HỢP NHẤT Ô P2:T4)
  destSheet.getRange(2, 16, 3, 5).merge()
    .setFormula(`=HYPERLINK("${directExportUrl}"; "📥 TẢI FILE EXCEL")`)
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground('#1A73E8')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // DÒNG 6: DANH SÁCH CHI TIẾT 20 CỘT CHUẨN
  const detailHeaders = [
    'STT', 'MÃ KHÁCH HÀNG', 'LOẠI KHÁCH HÀNG', 'TÊN THUÊ BAO', 'SĐT LIÊN HỆ', 'ĐỊA CHỈ LẮP ĐẶT',
    'GÓI CƯỚC', 'GIÁ CƯỚC', 'CHU KỲ', 'DTPK', 'TÌNH TRẠNG HỢP ĐỒNG', 'TÌNH TRẠNG CHẶN CẮT',
    'LOẠI THIẾT BỊ ONT', 'SERIAL ONT', 'NGÀY ĐẤU NỐI', 'NGÀY HOÀN TẤT THI CÔNG',
    'NGÀY BẮT ĐẦU GÓI CƯỚC', 'NGÀY KẾT THÚC GÓI CƯỚC', 'TÊN CỬA HÀNG', 'TÊN NHÂN VIÊN'
  ];

  destSheet.getRange(6, 1, 1, detailHeaders.length)
    .setValues([detailHeaders])
    .setFontWeight('bold')
    .setBackground('#00558F')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
  destSheet.setFrozenRows(6);

  if (matchedRows.length > 0) {
    destSheet.getRange(7, 1, matchedRows.length, detailHeaders.length).setValues(matchedRows);

    destSheet.getRange(7, 1, matchedRows.length, 1).setHorizontalAlignment('center'); // STT
    destSheet.getRange(7, 2, matchedRows.length, 1).setNumberFormat('@').setHorizontalAlignment('center'); // Mã KH
    destSheet.getRange(7, 3, matchedRows.length, 2).setHorizontalAlignment('left'); // Loại KH & Tên TB
    destSheet.getRange(7, 5, matchedRows.length, 1).setNumberFormat('@').setHorizontalAlignment('center'); // SĐT
    destSheet.getRange(7, 6, matchedRows.length, 1).setHorizontalAlignment('left'); // Địa chỉ
    destSheet.getRange(7, 7, matchedRows.length, 1).setHorizontalAlignment('center'); // Gói cước
    destSheet.getRange(7, 8, matchedRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right'); // Giá cước
    destSheet.getRange(7, 9, matchedRows.length, 1).setHorizontalAlignment('center'); // Chu kỳ
    destSheet.getRange(7, 10, matchedRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right'); // DTPK
    destSheet.getRange(7, 11, matchedRows.length, 2).setHorizontalAlignment('center'); // Tình trạng HĐ & Chặn cắt
    destSheet.getRange(7, 13, matchedRows.length, 1).setHorizontalAlignment('left'); // Thiết bị ONT
    destSheet.getRange(7, 14, matchedRows.length, 1).setNumberFormat('@').setHorizontalAlignment('center'); // Serial ONT
    destSheet.getRange(7, 15, matchedRows.length, 4).setHorizontalAlignment('center'); // 4 cột Ngày tháng
    destSheet.getRange(7, 19, matchedRows.length, 2).setHorizontalAlignment('left'); // Cửa hàng & Nhân viên

    // Cấu hình kích thước 20 cột chuẩn
    const exactColumnWidths = [
      50,  // 1. STT
      120, // 2. Mã KH
      160, // 3. Loại KH
      180, // 4. Tên thuê bao
      115, // 5. SĐT
      260, // 6. Địa chỉ
      110, // 7. Gói cước
      110, // 8. Giá cước
      75,  // 9. Chu kỳ
      110, // 10. DTPK
      140, // 11. Tình trạng HĐ
      140, // 12. Tình trạng chặn cắt
      130, // 13. Loại ONT
      130, // 14. Serial ONT
      155, // 15. Ngày đấu nối
      155, // 16. Ngày hoàn tất thi công
      155, // 17. Ngày bắt đầu
      155, // 18. Ngày kết thúc
      220, // 19. Tên cửa hàng
      160  // 20. Tên nhân viên
    ];

    for (let c = 0; c < exactColumnWidths.length; c++) {
      destSheet.setColumnWidth(c + 1, exactColumnWidths[c]);
    }
  }

  ss.setActiveSheet(destSheet);

  return {
    success: true,
    sheetName: sheetName,
    count: matchedRows.length,
    revenue: totalRevenue.toLocaleString('vi-VN'),
    dtpk: totalDtpk.toLocaleString('vi-VN')
  };
}

function cleanSlug(text) {
  return String(text || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]/g, '')
    .substring(0, 12);
}

/**
 * Tạo file Excel sạch độc lập (đã xóa ô P2:T4) để tải về máy
 */
function prepareCleanQueryDownloadUrl(sheetName) {
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

    // Xóa bỏ ô nút bấm tải P2:T4 trên file xuất
    copiedSheet.getRange('P2:T4').breakApart().clear().setBackground('#D2E3FC');
    copiedSheet.getRange('J2:T2').merge().setValue(sourceSheet.getRange('J2').getValue()).setHorizontalAlignment('right');
    copiedSheet.getRange('A3:T3').merge().setValue(sourceSheet.getRange('A3').getValue()).setHorizontalAlignment('left');
    copiedSheet.getRange('A4:T4').merge().setValue(sourceSheet.getRange('A4').getValue()).setHorizontalAlignment('left');

    SpreadsheetApp.flush();

    const file = DriveApp.getFileById(tempSsId);
    file.setName(`${sheetName}.xlsx`);
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
 * Giao diện HTML Hộp thoại Truy vấn Thông minh
 */
function getQueryDialogHtml() {
  return `
  <!DOCTYPE html>
  <html>
  <head>
    <base target="_top">
    <style>
      body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; padding: 14px; margin: 0; background: #f8f9fa; font-size: 13px; color: #202124; }
      .filter-card { background: white; border: 1px solid #dadce0; border-radius: 8px; padding: 12px 14px; margin-bottom: 10px; }
      .section-title { font-weight: bold; color: #1a73e8; margin-bottom: 8px; font-size: 13px; }
      .form-row { display: flex; gap: 10px; margin-bottom: 8px; align-items: center; }
      .form-col { flex: 1; }
      label { display: block; font-size: 11px; font-weight: 600; color: #5f6368; margin-bottom: 3px; }
      select, input[type="text"] { width: 100%; padding: 6px 8px; border: 1px solid #dadce0; border-radius: 4px; font-size: 12px; box-sizing: border-box; background: #fff; }
      select:focus, input[type="text"]:focus { border-color: #1a73e8; outline: none; }
      button { width: 100%; border: none; padding: 11px; font-size: 13px; font-weight: bold; border-radius: 6px; cursor: pointer; margin-top: 6px; display: block; box-sizing: border-box; text-align: center; }
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

      <!-- KHỐI 1: KÊNH BÁN & ĐỊA BÀN -->
      <div class="filter-card">
        <div class="section-title">1. Đơn vị Quản lý & Địa bàn:</div>
        <div class="form-row">
          <div class="form-col">
            <label>CỬA HÀNG:</label>
            <select id="selStore"><option value="">-- Tất cả cửa hàng --</option></select>
          </div>
          <div class="form-col">
            <label>NHÂN VIÊN (Mã/Tên):</label>
            <input type="text" id="inpStaff" placeholder="Ví dụ: Lê Hoàng Lộc...">
          </div>
        </div>
        <div class="form-row" style="margin-bottom:0;">
          <div class="form-col">
            <label>PHƯỜNG XÃ / ĐỊA BÀN:</label>
            <input type="text" id="inpWard" placeholder="Ví dụ: Phường 08, La Gi...">
          </div>
        </div>
      </div>

      <!-- KHỐI 2: KHÁCH HÀNG & HỢP ĐỒNG -->
      <div class="filter-card">
        <div class="section-title">2. Phân khúc, Hợp đồng & Gói cước:</div>
        <div class="form-row">
          <div class="form-col">
            <label>LOẠI KHÁCH HÀNG:</label>
            <select id="selCustType"><option value="">-- Tất cả loại KH --</option></select>
          </div>
          <div class="form-col">
            <label>TÌNH TRẠNG HỢP ĐỒNG:</label>
            <select id="selStatusHD">
              <option value="Đang hoạt động" selected>Đang hoạt động</option>
              <option value="Chờ thi công">Chờ thi công</option>
              <option value="Đã hủy">Đã hủy</option>
              <option value="">-- Tất cả trạng thái --</option>
            </select>
          </div>
        </div>
        <div class="form-row">
          <div class="form-col">
            <label>TÌNH TRẠNG CHẶN CẮT:</label>
            <select id="selStatusBlock">
              <option value="">-- Tất cả tình trạng --</option>
              <option value="Bình thường">Bình thường</option>
              <option value="Chặn 1 chiều KHYC">Chặn 1 chiều KHYC</option>
              <option value="1C Nợ cước">1C Nợ cước</option>
            </select>
          </div>
          <div class="form-col">
            <label>TÊN GÓI CƯỚC:</label>
            <input type="text" id="inpPkg" placeholder="Ví dụ: 12FDN1000, 6WIFI...">
          </div>
        </div>
      </div>

      <!-- KHỐI 3: KHOẢNG NGÀY (TÙY CHỌN) -->
      <div class="filter-card">
        <div class="section-title">3. Khoảng Thời gian (Tùy chọn):</div>
        <div class="form-row" style="margin-bottom:0;">
          <div class="form-col">
            <label>TỪ NGÀY (dd/MM/yyyy):</label>
            <input type="text" id="fromDate" placeholder="01/01/2026">
          </div>
          <div class="form-col">
            <label>ĐẾN NGÀY (dd/MM/yyyy):</label>
            <input type="text" id="toDate" placeholder="31/12/2026">
          </div>
        </div>
      </div>

      <button id="btnRun" class="btn-primary" type="button" onclick="executeAdvancedQuery()">🚀 Trích xuất Danh sách</button>
      <div id="statusBox" class="status"></div>
    </div>

    <!-- MÀN HÌNH HOÀN TẤT & NÚT TẢI SẠCH -->
    <div id="successView">
      <div class="success-title">🎉 TRÍCH XUẤT THÀNH CÔNG!</div>
      <div class="kpi-card">
        <div style="font-weight: 600; color: #1a73e8;" id="resSheetName">Trang tính: ...</div>
        <div style="margin-top: 6px; font-size: 13px; line-height: 1.6;">
          🎯 <b>Số lượng thuê bao:</b> <span id="resCount" style="color: #c5221f; font-weight: bold;">0</span> TB<br>
          💰 <b>Tổng doanh thu gói cước:</b> <span id="resRevenue" style="color: #137333; font-weight: bold;">0</span> VNĐ<br>
          📈 <b>Tổng DTPK:</b> <span id="resDtpk" style="color: #b06000; font-weight: bold;">0</span> VNĐ
        </div>
      </div>

      <button id="btnDlQuery" class="btn-download" type="button" onclick="startCleanQueryDownload()">📥 Tải file Excel (.xlsx) về máy</button>
      <button class="btn-close" type="button" onclick="closeDialogCleanly()">Đóng cửa sổ</button>
    </div>

    <iframe id="downloadQueryFrame" style="display:none;"></iframe>

    <script>
      let currentCreatedQuerySheet = '';
      let currentTempQueryFileId = null;

      window.onload = function() {
        // Tự động nạp dữ liệu Dropdown thực tế từ raw_mobifiber
        google.script.run
          .withSuccessHandler(function(data) {
            const selStore = document.getElementById('selStore');
            data.stores.forEach(function(s) {
              const opt = document.createElement('option');
              opt.value = s; opt.text = s;
              selStore.appendChild(opt);
            });

            const selCustType = document.getElementById('selCustType');
            data.custTypes.forEach(function(c) {
              const opt = document.createElement('option');
              opt.value = c; opt.text = c;
              selCustType.appendChild(opt);
            });
          })
          .getFilterDropdownData();
      };

      function showError(msg) {
        const b = document.getElementById('errorBanner');
        b.innerText = '⚠️ ' + msg;
        b.style.display = 'block';
      }

      function hideError() {
        document.getElementById('errorBanner').style.display = 'none';
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

      function executeAdvancedQuery() {
        try {
          hideError();
          const f = document.getElementById('fromDate').value.trim();
          const t = document.getElementById('toDate').value.trim();

          if ((f && !t) || (!f && t)) {
            showError('Nếu lọc ngày, vui lòng nhập đầy đủ cả Từ ngày và Đến ngày!'); return;
          }
          if (f && t) {
            if (!isValidDateStr(f) || !isValidDateStr(t)) {
              showError('Định dạng ngày không hợp lệ (phải là dd/MM/yyyy)!'); return;
            }
            if (parseDateObj(f) > parseDateObj(t)) {
              showError('"Từ ngày" không được lớn hơn "Đến ngày"!'); return;
            }
          }

          const params = {
            store: document.getElementById('selStore').value,
            staff: document.getElementById('inpStaff').value.trim(),
            ward: document.getElementById('inpWard').value.trim(),
            custType: document.getElementById('selCustType').value,
            statusHD: document.getElementById('selStatusHD').value,
            statusBlock: document.getElementById('selStatusBlock').value,
            pkg: document.getElementById('inpPkg').value.trim(),
            from: f,
            to: t
          };

          const btn = document.getElementById('btnRun');
          const status = document.getElementById('statusBox');

          btn.disabled = true;
          btn.innerText = '⏳ Đang quét cơ sở dữ liệu...';
          status.className = 'status info';
          status.innerText = 'Hệ thống đang trích xuất dữ liệu theo điều kiện...';

          google.script.run
            .withSuccessHandler(function(res) {
              currentCreatedQuerySheet = res.sheetName;
              document.getElementById('formView').style.display = 'none';
              document.getElementById('resSheetName').innerText = '📁 Trang tính: ' + res.sheetName;
              document.getElementById('resCount').innerText = res.count;
              document.getElementById('resRevenue').innerText = res.revenue;
              document.getElementById('resDtpk').innerText = res.dtpk;
              document.getElementById('successView').style.display = 'block';
            })
            .withFailureHandler(function(err) {
              btn.disabled = false;
              btn.innerText = '🚀 Trích xuất Danh sách';
              status.style.display = 'none';
              showError(err.message);
            })
            .processAdvancedQuery(params);
        } catch (e) {
          showError('Có lỗi xảy ra: ' + e.message);
        }
      }

      function startCleanQueryDownload() {
        const btn = document.getElementById('btnDlQuery');
        btn.disabled = true;
        btn.innerText = '⏳ Đang chuẩn bị file Excel sạch...';

        google.script.run
          .withSuccessHandler(function(res) {
            btn.disabled = false;
            btn.innerText = '📥 Tải file Excel (.xlsx) về máy';
            currentTempQueryFileId = res.fileId;
            document.getElementById('downloadQueryFrame').src = res.downloadUrl;

            // Tự động xóa file tạm sau 5 giây
            setTimeout(function() {
              if (currentTempQueryFileId) {
                google.script.run.deleteTempFile(currentTempQueryFileId);
                currentTempQueryFileId = null;
              }
            }, 5000);
          })
          .withFailureHandler(function(err) {
            btn.disabled = false;
            btn.innerText = '📥 Tải file Excel (.xlsx) về máy';
            showError('Lỗi tải file: ' + err.message);
          })
          .prepareCleanQueryDownloadUrl(currentCreatedQuerySheet);
      }

      function closeDialogCleanly() {
        if (currentTempQueryFileId) {
          google.script.run.deleteTempFile(currentTempQueryFileId);
          currentTempQueryFileId = null;
        }
        google.script.host.close();
      }
    </script>
  </body>
  </html>
  `;
}