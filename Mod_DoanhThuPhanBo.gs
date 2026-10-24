/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * MODULE       : BÁO CÁO DOANH THU PHÂN BỔ (DTPK)
 * PHIÊN BẢN    : 1.0 (Phân rã Doanh thu PTM vs DTPK Hiện hữu, Tách Sheet Chi tiết theo nhu cầu)
 * NGÀY CẬP NHẬT: 21/09/2026 14:00:00 (GMT+7)
 * TÁC GIẢ      : Trịnh Vũ Thế Vinh & AI Assistant
 * MÔ TẢ        : 
 *   - Lọc 3 lớp: Loại trừ TB nghiệp vụ, Hợp đồng Đang hoạt động, Chặn cắt Bình thường
 *   - Bóc tách độc lập:
 *       + DOANH THU PTM: Thuê bao phát sinh mới trong kỳ
 *       + DTPK HIỆN HỮU: Thuê bao đơn kỳ trả sau & Thuê bao dài kỳ còn hạn
 *   - Sheet Tổng hợp: Cột A Khu vực không gộp, 8 cột chuẩn, nút Tải Excel & Nút Kết xuất Chi tiết
 *   - Sheet Chi tiết: Kết xuất theo nhu cầu sang tab riêng (19 cột chuẩn), sắp xếp 4 cấp
 *   - Cơ chế tự hủy file tạm trên Drive sau 5 giây chống rác bộ nhớ
 * ============================================================================
 */

function showRevenueAllocationDialog() {
  const html = HtmlService.createHtmlOutput(getRevenueAllocationDialogHtml())
    .setWidth(500)
    .setHeight(450)
    .setTitle('Báo cáo Doanh thu Phân bổ (DTPK)');
  SpreadsheetApp.getUi().showModalDialog(html, 'Báo cáo Doanh thu Phân bổ MobiFiber');
}

/**
 * Xử lý kết xuất Bảng Tổng hợp Doanh thu Phân bổ
 */
function processRevenueAllocationSummary(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) throw new Error('Bảng raw_mobifiber chưa có dữ liệu.');

  // 1. Nạp bản đồ tham số
  const mapping = loadThamSoMapping();

  // 2. Xác định mốc thời gian kỳ báo cáo
  let startDate, endDate, timeTitle, timeCode;
  if (params.type === 'month') {
    const month = parseInt(params.month, 10);
    const year = parseInt(params.year, 10);
    const mStr = month < 10 ? '0' + month : '' + month;
    startDate = new Date(year, month - 1, 1, 0, 0, 0);
    endDate = new Date(year, month, 0, 23, 59, 59);
    timeTitle = `THÁNG ${mStr}/${year}`;
    timeCode = `T${mStr}_${year}`;
  } else {
    const startParts = params.from.split('/');
    const endParts = params.to.split('/');
    startDate = new Date(parseInt(startParts[2], 10), parseInt(startParts[1], 10) - 1, parseInt(startParts[0], 10), 0, 0, 0);
    endDate = new Date(parseInt(endParts[2], 10), parseInt(endParts[1], 10) - 1, parseInt(endParts[0], 10), 23, 59, 59);
    const d1 = startParts[0] + startParts[1];
    const d2 = endParts[0] + endParts[1];
    timeTitle = `TỪ ${params.from} ĐẾN ${params.to}`;
    timeCode = `${d1}_${d2}_${endParts[2]}`;
  }

  const summarySheetName = `DTPB_${timeCode}`;

  // 3. Đọc Headers
  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol)
    .getValues()[0].map(h => normalizeHeader(h));

  const idxMaKH        = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxLoaiKH      = headers.findIndex(h => h === 'loại khách hàng');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxChanCat     = headers.findIndex(h => h === 'tình trạng chặn cắt');
  const idxNhuCau      = headers.findIndex(h => h === 'nhu cầu hòa mạng');
  const idxChuKy       = headers.findIndex(h => h === 'chu kỳ sử dụng');
  const idxGiaCuoc     = headers.findIndex(h => h === 'giá cước');
  const idxDtpk        = headers.findIndex(h => h === 'dtpk');
  const idxNgayDauNoi  = headers.findIndex(h => h === 'ngày đấu nối');
  const idxNgayThiCong = headers.findIndex(h => h === 'ngày hoàn tất thi công');
  const idxNgayBatDau  = headers.findIndex(h => h === 'ngày băt đầu gói cước' || h === 'ngày bắt đầu gói cước');
  const idxNgayKetThuc = headers.findIndex(h => h === 'ngày kết thúc gói cước');
  const idxMaNV        = headers.findIndex(h => h === 'mã nhân viên');

  const numRows = lastRow - CONFIG.DATA_START_ROW_TARGET + 1;
  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, numRows, lastCol).getValues();

  // Khởi tạo các nhóm thống kê theo Khu vực
  const regionMap = {};
  let grandTotalOldTB = 0, grandTotalOldDtpk = 0;
  let grandTotalNewTB = 0, grandTotalNewRevenue = 0;

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];

    // BỘ LỌC 3 LỚP NGHIÊM NGẶT
    const loaiKH = normalizeHeader(row[idxLoaiKH]);
    if (loaiKH === 'thuê bao nghiệp vụ' || loaiKH === 'thue bao nghiep vu') continue;

    const tinhTrangHD = normalizeHeader(row[idxTinhTrangHD]);
    if (tinhTrangHD !== 'đang hoạt động' && tinhTrangHD !== 'dang hoat dong') continue;

    const chanCat = normalizeHeader(row[idxChanCat]);
    if (chanCat !== 'bình thường' && chanCat !== 'binh thuong') continue;

    const chuKyRaw = row[idxChuKy];
    const chuKy = chuKyRaw ? parseInt(chuKyRaw, 10) : 0;
    const giaCuoc = parseCurrencyToInteger(row[idxGiaCuoc]) || 0;
    const dtpk = parseCurrencyToInteger(row[idxDtpk]) || (chuKy > 0 ? Math.round(giaCuoc / chuKy) : 0);

    const dDauNoi = parseToDate(row[idxNgayDauNoi]);
    const dThiCong = parseToDate(row[idxNgayThiCong]);
    const dBatDau = parseToDate(row[idxNgayBatDau]);
    const dKetThuc = parseToDate(row[idxNgayKetThuc]);

    // XÁC ĐỊNH XEM THUÊ BAO CÓ PHẢI LÀ PTM TRONG KỲ KHÔNG
    const nhuCau = normalizeHeader(row[idxNhuCau]);
    let isPtmInPeriod = false;

    if (nhuCau === 'đăng ký mới' || nhuCau === 'dang ky moi') {
      let ptmDate = null;
      if (chuKy === 1) {
        if (dDauNoi && dThiCong) {
          ptmDate = (dThiCong.getTime() > dDauNoi.getTime()) ? dThiCong : dDauNoi;
        } else {
          ptmDate = dThiCong || dDauNoi || dBatDau;
        }
      } else {
        ptmDate = dBatDau;
      }
      if (ptmDate && ptmDate >= startDate && ptmDate <= endDate) {
        isPtmInPeriod = true;
      }
    }

    // XÁC ĐỊNH XEM THUÊ BAO CÓ THỎA MÃN DOANH THU PHÂN BỔ TRONG KỲ
    let isRetainedInPeriod = false;
    if (!isPtmInPeriod) {
      if (chuKy === 1) {
        // Thuê bao đơn kỳ hiện hữu: trả sau hạch toán phân kỳ
        isRetainedInPeriod = true;
      } else {
        // Thuê bao dài kỳ: chưa bị quá hạn so với ngày bắt đầu kỳ báo cáo
        if (dKetThuc && dKetThuc >= startDate) {
          isRetainedInPeriod = true;
        }
      }
    }

    if (!isPtmInPeriod && !isRetainedInPeriod) continue;

    // Xác định Khu vực từ thamso_caidat
    const rawMaNV = String(row[idxMaNV] || '').trim();
    let khuVuc = 'Chưa phân bổ KV';
    if (rawMaNV && mapping.staffRegion && mapping.staffRegion.has(normalizeHeader(rawMaNV))) {
      const staffInfo = mapping.staffRegion.get(normalizeHeader(rawMaNV));
      if (staffInfo.khuVuc && !staffInfo.khuVuc.includes('Chưa phân bổ')) {
        khuVuc = staffInfo.khuVuc;
      }
    }

    if (!regionMap[khuVuc]) {
      regionMap[khuVuc] = { oldTB: 0, oldDtpk: 0, newTB: 0, newRevenue: 0 };
    }

    if (isPtmInPeriod) {
      regionMap[khuVuc].newTB += 1;
      regionMap[khuVuc].newRevenue += dtpk; // Doanh thu phân kỳ của thuê bao PTM
      grandTotalNewTB += 1;
      grandTotalNewRevenue += dtpk;
    } else if (isRetainedInPeriod) {
      regionMap[khuVuc].oldTB += 1;
      regionMap[khuVuc].oldDtpk += dtpk; // Doanh thu phân kỳ của thuê bao cũ
      grandTotalOldTB += 1;
      grandTotalOldDtpk += dtpk;
    }
  }

  const grandTotalDtpk = grandTotalOldDtpk + grandTotalNewRevenue;
  const grandTotalAllTB = grandTotalOldTB + grandTotalNewTB;

  // 4. Tạo hoặc làm mới Sheet Tổng hợp
  let destSheet = ss.getSheetByName(summarySheetName);
  if (!destSheet) {
    destSheet = ss.insertSheet(summarySheetName);
  } else {
    destSheet.clear();
  }

  const ssId = ss.getId();
  const sheetId = destSheet.getSheetId();
  const directExportUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=xlsx&gid=${sheetId}`;

  // DÒNG 1: Tiêu đề lớn
  destSheet.getRange('A1:H1').merge()
    .setValue(`BÁO CÁO DOANH THU PHÂN BỔ MOBIFIBER - ${timeTitle}`)
    .setFontWeight('bold').setFontSize(13).setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');

  // DÒNG 2: TỔNG DOANH THU PHÂN BỔ TOÀN TỈNH
  destSheet.getRange('A2:E2').merge()
    .setValue(`🎯 TỔNG DTPK TOÀN TỈNH: ${grandTotalDtpk.toLocaleString('vi-VN')} VNĐ`)
    .setFontWeight('bold').setFontSize(11).setBackground('#D2E3FC').setFontColor('#174EA6').setHorizontalAlignment('left');

  destSheet.getRange('F2:H2').merge()
    .setValue(`🕒 Ngày lập: ${Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm')}`)
    .setFontSize(10).setBackground('#D2E3FC').setFontColor('#5F6368').setHorizontalAlignment('right');

  // DÒNG 3: TỔNG THUÊ BAO
  destSheet.getRange('A3:H3').merge()
    .setValue(`📊 TỔNG THUÊ BAO: ${grandTotalAllTB} TB   (Hiện hữu: ${grandTotalOldTB} TB  |  PTM mới: ${grandTotalNewTB} TB)`)
    .setFontWeight('bold').setFontSize(11).setBackground('#E8F0FE').setFontColor('#174EA6').setHorizontalAlignment('left');

  // DÒNG 4: HAI NÚT CHỨC NĂNG THAO TÁC TRỰC TIẾP TRÊN SHEET
  destSheet.getRange('A4:D4').merge()
    .setFormula(`=HYPERLINK("#gid=${sheetId}&action=detail"; "📋 KẾT XUẤT DANH SÁCH CHI TIẾT")`)
    .setFontWeight('bold').setFontSize(11).setBackground('#00558F').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  destSheet.getRange('E4:H4').merge()
    .setFormula(`=HYPERLINK("${directExportUrl}"; "📥 TẢI FILE EXCEL TỔNG HỢP")`)
    .setFontWeight('bold').setFontSize(11).setBackground('#1A73E8').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // DÒNG 6: TIÊU ĐỀ BẢNG TỔNG HỢP (8 CỘT CHUẨN, CỘT A KHÔNG GỘP)
  const tableHeaders = [
    'KHU VỰC', 'TB HIỆN HỮU', 'DTPK HIỆN HỮU', 'TB PTM', 'DOANH THU PTM', 'TỔNG SỐ TB', 'TỔNG DTPK (VNĐ)', 'TỶ TRỌNG (%)'
  ];

  destSheet.getRange(6, 1, 1, tableHeaders.length).setValues([tableHeaders])
    .setFontWeight('bold').setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  destSheet.setFrozenRows(6);

  // Điền các dòng dữ liệu Khu vực
  const regionRows = Object.keys(regionMap).map(kv => {
    const d = regionMap[kv];
    const totalTbKv = d.oldTB + d.newTB;
    const totalDtpkKv = d.oldDtpk + d.newRevenue;
    const share = grandTotalDtpk > 0 ? (totalDtpkKv / grandTotalDtpk) : 0;
    return [
      '  ' + kv,
      d.oldTB,
      d.oldDtpk,
      d.newTB,
      d.newRevenue,
      totalTbKv,
      totalDtpkKv,
      share
    ];
  }).sort((a, b) => b[6] - a[6]); // Sắp xếp Khu vực có DTPK lớn nhất lên trước

  let curRow = 7;
  if (regionRows.length > 0) {
    destSheet.getRange(curRow, 1, regionRows.length, tableHeaders.length).setValues(regionRows);
    destSheet.getRange(curRow, 1, regionRows.length, 1).setHorizontalAlignment('left'); // Khu vực
    destSheet.getRange(curRow, 2, regionRows.length, 1).setHorizontalAlignment('center'); // TB Cũ
    destSheet.getRange(curRow, 3, regionRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right'); // DTPK Cũ
    destSheet.getRange(curRow, 4, regionRows.length, 1).setHorizontalAlignment('center'); // TB PTM
    destSheet.getRange(curRow, 5, regionRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right'); // DT PTM
    destSheet.getRange(curRow, 6, regionRows.length, 1).setHorizontalAlignment('center'); // Tổng TB
    destSheet.getRange(curRow, 7, regionRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right'); // Tổng DTPK
    destSheet.getRange(curRow, 8, regionRows.length, 1).setNumberFormat('0.0%').setHorizontalAlignment('right'); // Tỷ trọng
    curRow += regionRows.length;
  }

  // DÒNG TỔNG CỘNG TOÀN TỈNH
  const summaryTotalRow = [
    '  TỔNG CỘNG TOÀN TỈNH',
    grandTotalOldTB,
    grandTotalOldDtpk,
    grandTotalNewTB,
    grandTotalNewRevenue,
    grandTotalAllTB,
    grandTotalDtpk,
    1.0
  ];
  destSheet.getRange(curRow, 1, 1, tableHeaders.length).setValues([summaryTotalRow])
    .setFontWeight('bold').setBackground('#E8F0FE').setFontColor('#174EA6');
  destSheet.getRange(curRow, 1).setHorizontalAlignment('left');
  destSheet.getRange(curRow, 2).setHorizontalAlignment('center');
  destSheet.getRange(curRow, 3).setNumberFormat('#,##0').setHorizontalAlignment('right');
  destSheet.getRange(curRow, 4).setHorizontalAlignment('center');
  destSheet.getRange(curRow, 5).setNumberFormat('#,##0').setHorizontalAlignment('right');
  destSheet.getRange(curRow, 6).setHorizontalAlignment('center');
  destSheet.getRange(curRow, 7).setNumberFormat('#,##0').setHorizontalAlignment('right');
  destSheet.getRange(curRow, 8).setNumberFormat('0.0%').setHorizontalAlignment('right');

  // CẤU HÌNH ĐỘ RỘNG 8 CỘT TỔNG HỢP
  const exactWidths = [160, 110, 130, 100, 130, 110, 140, 124];
  for (let c = 0; c < exactWidths.length; c++) {
    destSheet.setColumnWidth(c + 1, exactWidths[c]);
  }

  ss.setActiveSheet(destSheet);

  return {
    success: true,
    sheetName: summarySheetName,
    timeCode: timeCode,
    from: params.from || '',
    to: params.to || '',
    month: params.month || '',
    year: params.year || '',
    type: params.type,
    totalDtpk: grandTotalDtpk.toLocaleString('vi-VN'),
    totalTB: grandTotalAllTB,
    oldTB: grandTotalOldTB,
    newTB: grandTotalNewTB
  };
}

/**
 * XỬ LÝ KẾT XUẤT SHEET CHI TIẾT RIÊNG BIỆT (19 CỘT CHUẨN)
 * Chỉ kích hoạt khi người dùng yêu cầu
 */
function processRevenueAllocationDetail(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) throw new Error('Bảng raw_mobifiber chưa có dữ liệu.');

  const mapping = loadThamSoMapping();

  let startDate, endDate, timeTitle, timeCode;
  if (params.type === 'month') {
    const month = parseInt(params.month, 10);
    const year = parseInt(params.year, 10);
    const mStr = month < 10 ? '0' + month : '' + month;
    startDate = new Date(year, month - 1, 1, 0, 0, 0);
    endDate = new Date(year, month, 0, 23, 59, 59);
    timeTitle = `THÁNG ${mStr}/${year}`;
    timeCode = `T${mStr}_${year}`;
  } else {
    const startParts = params.from.split('/');
    const endParts = params.to.split('/');
    startDate = new Date(parseInt(startParts[2], 10), parseInt(startParts[1], 10) - 1, parseInt(startParts[0], 10), 0, 0, 0);
    endDate = new Date(parseInt(endParts[2], 10), parseInt(endParts[1], 10) - 1, parseInt(endParts[0], 10), 23, 59, 59);
    const d1 = startParts[0] + startParts[1];
    const d2 = endParts[0] + endParts[1];
    timeTitle = `TỪ ${params.from} ĐẾN ${params.to}`;
    timeCode = `${d1}_${d2}_${endParts[2]}`;
  }

  const detailSheetName = `DTPB_CT_${timeCode}`;

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
  const idxNgayDauNoi  = headers.findIndex(h => h === 'ngày đấu nối');
  const idxNgayThiCong = headers.findIndex(h => h === 'ngày hoàn tất thi công');
  const idxNgayBatDau  = headers.findIndex(h => h === 'ngày băt đầu gói cước' || h === 'ngày bắt đầu gói cước');
  const idxNgayKetThuc = headers.findIndex(h => h === 'ngày kết thúc gói cước');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxChanCat     = headers.findIndex(h => h === 'tình trạng chặn cắt');
  const idxNhuCau      = headers.findIndex(h => h === 'nhu cầu hòa mạng');
  const idxMaCH        = headers.findIndex(h => h === 'mã cửa hàng');
  const idxTenCH       = headers.findIndex(h => h === 'tên cửa hàng');
  const idxMaNV        = headers.findIndex(h => h === 'mã nhân viên');
  const idxTenNV       = headers.findIndex(h => h === 'tên nhân viên');

  const numRows = lastRow - CONFIG.DATA_START_ROW_TARGET + 1;
  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, numRows, lastCol).getValues();

  const detailList = [];

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];

    const loaiKH = normalizeHeader(row[idxLoaiKH]);
    if (loaiKH === 'thuê bao nghiệp vụ' || loaiKH === 'thue bao nghiep vu') continue;

    const tinhTrangHD = normalizeHeader(row[idxTinhTrangHD]);
    if (tinhTrangHD !== 'đang hoạt động' && tinhTrangHD !== 'dang hoat dong') continue;

    const chanCat = normalizeHeader(row[idxChanCat]);
    if (chanCat !== 'bình thường' && chanCat !== 'binh thuong') continue;

    const chuKyRaw = row[idxChuKy];
    const chuKy = chuKyRaw ? parseInt(chuKyRaw, 10) : 0;
    const giaCuoc = parseCurrencyToInteger(row[idxGiaCuoc]) || 0;
    const dtpk = parseCurrencyToInteger(row[idxDtpk]) || (chuKy > 0 ? Math.round(giaCuoc / chuKy) : 0);

    const dDauNoi = parseToDate(row[idxNgayDauNoi]);
    const dThiCong = parseToDate(row[idxNgayThiCong]);
    const dBatDau = parseToDate(row[idxNgayBatDau]);
    const dKetThuc = parseToDate(row[idxNgayKetThuc]);

    const nhuCau = normalizeHeader(row[idxNhuCau]);
    let isPtmInPeriod = false;

    if (nhuCau === 'đăng ký mới' || nhuCau === 'dang ky moi') {
      let ptmDate = null;
      if (chuKy === 1) {
        if (dDauNoi && dThiCong) {
          ptmDate = (dThiCong.getTime() > dDauNoi.getTime()) ? dThiCong : dDauNoi;
        } else {
          ptmDate = dThiCong || dDauNoi || dBatDau;
        }
      } else {
        ptmDate = dBatDau;
      }
      if (ptmDate && ptmDate >= startDate && ptmDate <= endDate) {
        isPtmInPeriod = true;
      }
    }

    let isRetainedInPeriod = false;
    if (!isPtmInPeriod) {
      if (chuKy === 1) {
        isRetainedInPeriod = true;
      } else {
        if (dKetThuc && dKetThuc >= startDate) {
          isRetainedInPeriod = true;
        }
      }
    }

    if (!isPtmInPeriod && !isRetainedInPeriod) continue;

    const rawMaNV = String(row[idxMaNV] || '').trim();
    let khuVuc = 'Chưa phân bổ KV';
    if (rawMaNV && mapping.staffRegion && mapping.staffRegion.has(normalizeHeader(rawMaNV))) {
      const staffInfo = mapping.staffRegion.get(normalizeHeader(rawMaNV));
      if (staffInfo.khuVuc && !staffInfo.khuVuc.includes('Chưa phân bổ')) {
        khuVuc = staffInfo.khuVuc;
      }
    }

    let cleanMaKH = String(row[idxMaKH] || '').replace(/^'+/, '').trim();
    if (/^\d{8}$/.test(cleanMaKH)) cleanMaKH = '00' + cleanMaKH;
    let cleanSDT = String(row[idxSDT] || '').replace(/^'+/, '').trim();
    if (/^\d{9}$/.test(cleanSDT)) cleanSDT = '0' + cleanSDT;

    const paddedTenTB  = row[idxTenTB] ? '  ' + String(row[idxTenTB]).trim() : '';
    const paddedDiaChi = row[idxDiaChi] ? '  ' + String(row[idxDiaChi]).trim() : '';
    const cleanMaCH    = String(row[idxMaCH] || '').replace(/^'+/, '').trim();
    const cleanMaNV    = String(row[idxMaNV] || '').replace(/^'+/, '').trim();

    const rawTenCH = String(row[idxTenCH] || '').trim();
    const shortTenCH = getShortValue(mapping.stores, rawTenCH);
    const paddedTenCH = shortTenCH ? '  ' + shortTenCH : '';
    const paddedTenNV = row[idxTenNV] ? '  ' + String(row[idxTenNV]).trim() : '';

    const phanLoaiText = isPtmInPeriod ? 'PTM trong kỳ' : 'Hiện hữu duy trì';
    const dtpkHienHuu = isPtmInPeriod ? '' : dtpk;
    const doanhThuPtm = isPtmInPeriod ? dtpk : '';

    detailList.push({
      khuVuc: khuVuc,
      orderPhanLoai: isPtmInPeriod ? 1 : 2, // PTM lên trước
      maCH: cleanMaCH,
      maNV: cleanMaNV,
      dtpkVal: dtpk,
      rowValues: [
        0,                                                  // 1. STT
        "'" + cleanMaKH,                                    // 2. MÃ KHÁCH HÀNG
        paddedTenTB,                                        // 3. TÊN THUÊ BAO
        "'" + cleanSDT,                                     // 4. SĐT LIÊN HỆ
        paddedDiaChi,                                       // 5. ĐỊA CHỈ LẮP ĐẶT
        phanLoaiText,                                       // 6. PHÂN LOẠI DOANH THU
        row[idxGoiCuoc] || '',                              // 7. GÓI CƯỚC
        giaCuoc,                                            // 8. GIÁ CƯỚC
        chuKy > 0 ? chuKy : '',                             // 9. CHU KỲ
        dtpkHienHuu,                                        // 10. DTPK HIỆN HỮU
        doanhThuPtm,                                        // 11. DOANH THU PTM
        formatDateTimeStr(row[idxNgayDauNoi]),              // 12. NGÀY ĐẤU NỐI
        formatDateTimeStr(row[idxNgayBatDau]),              // 13. NGÀY BẮT ĐẦU
        formatDateTimeStr(row[idxNgayKetThuc]),             // 14. NGÀY KẾT THÚC
        khuVuc,                                             // 15. KHU VỰC
        "'" + cleanMaCH,                                    // 16. MÃ CỬA HÀNG
        "'" + cleanMaNV,                                    // 17. MÃ NHÂN VIÊN
        paddedTenCH,                                        // 18. TÊN CỬA HÀNG
        paddedTenNV                                         // 19. TÊN NHÂN VIÊN
      ]
    });
  }

  // Sắp xếp phân tầng 4 cấp
  detailList.sort((a, b) => {
    const compKV = a.khuVuc.localeCompare(b.khuVuc, 'vi', { sensitivity: 'base' });
    if (compKV !== 0) return compKV;
    if (a.orderPhanLoai !== b.orderPhanLoai) return a.orderPhanLoai - b.orderPhanLoai;
    const compCH = a.maCH.localeCompare(b.maCH, 'vi', { numeric: true });
    if (compCH !== 0) return compCH;
    const compNV = a.maNV.localeCompare(b.maNV, 'vi', { numeric: true });
    if (compNV !== 0) return compNV;
    return b.dtpkVal - a.dtpkVal; // DTPK giảm dần
  });

  const finalRows = detailList.map((item, index) => {
    item.rowValues[0] = index + 1;
    return item.rowValues;
  });

  // Tạo hoặc làm mới Sheet Chi tiết
  let ctSheet = ss.getSheetByName(detailSheetName);
  if (!ctSheet) {
    ctSheet = ss.insertSheet(detailSheetName);
  } else {
    ctSheet.clear();
  }

  const ssId = ss.getId();
  const ctSheetId = ctSheet.getSheetId();
  const directExportUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=xlsx&gid=${ctSheetId}`;

  // Dòng 1: Tiêu đề lớn
  ctSheet.getRange(1, 1, 1, 19).merge()
    .setValue(`DANH SÁCH CHI TIẾT THUÊ BAO ĐÓNG GÓP DOANH THU PHÂN BỔ - ${timeTitle}`)
    .setFontWeight('bold').setFontSize(13).setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');

  // Dòng 2 & 3: KPI
  ctSheet.getRange(2, 1, 1, 13).merge()
    .setValue(`🎯 TỔNG SỐ THUÊ BAO ĐÓNG GÓP DOANH THU: ${finalRows.length} TB`)
    .setFontWeight('bold').setFontSize(11).setBackground('#D2E3FC').setFontColor('#174EA6').setHorizontalAlignment('left');

  ctSheet.getRange(2, 14, 1, 6).merge()
    .setValue(`🕒 Ngày lập: ${Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm')}`)
    .setFontSize(10).setBackground('#D2E3FC').setFontColor('#5F6368').setHorizontalAlignment('right');

  // Nút Tải Excel Chi tiết
  ctSheet.getRange(3, 14, 1, 6).merge()
    .setFormula(`=HYPERLINK("${directExportUrl}"; "📥 TẢI FILE EXCEL CHI TIẾT")`)
    .setFontWeight('bold').setFontSize(11).setBackground('#1A73E8').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Dòng 5: Headers 19 cột chuẩn
  const detailHeaders = [
    'STT', 'MÃ KHÁCH HÀNG', 'TÊN THUÊ BAO', 'SĐT LIÊN HỆ', 'ĐỊA CHỈ LẮP ĐẶT',
    'PHÂN LOẠI DOANH THU', 'GÓI CƯỚC', 'GIÁ CƯỚC', 'CHU KỲ', 'DTPK HIỆN HỮU', 'DOANH THU PTM',
    'NGÀY ĐẤU NỐI', 'NGÀY BẮT ĐẦU', 'NGÀY KẾT THÚC', 'KHU VỰC', 'MÃ CỬA HÀNG', 'MÃ NHÂN VIÊN',
    'TÊN CỬA HÀNG', 'TÊN NHÂN VIÊN'
  ];

  ctSheet.getRange(5, 1, 1, detailHeaders.length).setValues([detailHeaders])
    .setFontWeight('bold').setBackground('#00558F').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  ctSheet.setFrozenRows(5);

  if (finalRows.length > 0) {
    ctSheet.getRange(6, 1, finalRows.length, detailHeaders.length).setValues(finalRows);
    ctSheet.getRange(6, 1, finalRows.length, 1).setHorizontalAlignment('center'); // STT
    ctSheet.getRange(6, 2, finalRows.length, 1).setNumberFormat('@').setHorizontalAlignment('center'); // Mã KH
    ctSheet.getRange(6, 3, finalRows.length, 1).setHorizontalAlignment('left'); // Tên TB
    ctSheet.getRange(6, 4, finalRows.length, 1).setNumberFormat('@').setHorizontalAlignment('center'); // SĐT
    ctSheet.getRange(6, 5, finalRows.length, 1).setHorizontalAlignment('left'); // Địa chỉ
    ctSheet.getRange(6, 6, finalRows.length, 1).setHorizontalAlignment('center'); // Phân loại
    ctSheet.getRange(6, 7, finalRows.length, 1).setHorizontalAlignment('center'); // Gói cước
    ctSheet.getRange(6, 8, finalRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right'); // Giá cước
    ctSheet.getRange(6, 9, finalRows.length, 1).setHorizontalAlignment('center'); // Chu kỳ
    ctSheet.getRange(6, 10, finalRows.length, 2).setNumberFormat('#,##0').setHorizontalAlignment('right'); // DTPK Cũ & PTM
    ctSheet.getRange(6, 12, finalRows.length, 3).setHorizontalAlignment('center'); // Ngày tháng
    ctSheet.getRange(6, 15, finalRows.length, 1).setHorizontalAlignment('center'); // Khu vực
    ctSheet.getRange(6, 16, finalRows.length, 2).setNumberFormat('@').setHorizontalAlignment('center'); // Mã CH & NV
    ctSheet.getRange(6, 18, finalRows.length, 2).setHorizontalAlignment('left'); // Tên CH & NV

    const exactColWidths = [
      50, 120, 180, 115, 260, 130, 110, 110, 75, 120, 120, 155, 155, 155, 110, 120, 120, 180, 160
    ];
    for (let c = 0; c < exactColWidths.length; c++) {
      ctSheet.setColumnWidth(c + 1, exactColWidths[c]);
    }
  }

  ss.setActiveSheet(ctSheet);

  return {
    success: true,
    sheetName: detailSheetName,
    count: finalRows.length
  };
}

/**
 * TẠO FILE EXCEL SẠCH (TỔNG HỢP HOẶC CHI TIẾT)
 */
function prepareCleanRevenueDownloadUrl(sheetName) {
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

    // Xóa nút bấm trên file xuất
    if (sheetName.includes('_CT_')) {
      copiedSheet.getRange('N3:S3').breakApart().clear().setBackground('#D2E3FC');
    } else {
      copiedSheet.getRange('A4:H4').breakApart().clear().setBackground('#FFFFFF');
    }

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
 * Giao diện Dialog Báo cáo Doanh thu Phân bổ
 */
function getRevenueAllocationDialogHtml() {
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
      .btn-secondary { background: #00558f; color: white; }
      .btn-secondary:hover { background: #003e68; }
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

    <!-- MÀN HÌNH 1: NHẬP KỲ LỌC -->
    <div id="formView">
      <div id="errorBanner" class="alert-banner"></div>

      <div class="opt-card">
        <label class="radio-label">
          <input type="radio" name="reportType" value="month" checked onchange="toggleType()">
          <strong>1. Báo cáo Doanh thu theo Tháng</strong>
        </label>
        <div class="inputs" id="monthBox">
          <span>Tháng:</span>
          <input type="text" id="monthVal" value="${curMonth < 10 ? '0' + curMonth : curMonth}" style="width: 45px; text-align: center;"> /
          <span>Năm:</span>
          <input type="text" id="yearVal" value="${curYear}" style="width: 65px; text-align: center;">
          <div class="hint">💡 Tự động tính DTPK của TB PTM trong tháng và TB hiện hữu còn hạn.</div>
        </div>
      </div>

      <div class="opt-card">
        <label class="radio-label">
          <input type="radio" name="reportType" value="range" onchange="toggleType()">
          <strong>2. Báo cáo theo Khoảng ngày tùy chọn</strong>
        </label>
        <div class="inputs" id="rangeBox" style="opacity: 0.5; pointer-events: none;">
          <div class="row">
            <span style="width: 65px;">Từ ngày:</span>
            <input type="text" id="fromDate" placeholder="01/10/2026" style="width: 110px;">
          </div>
          <div class="row">
            <span style="width: 65px;">Đến ngày:</span>
            <input type="text" id="toDate" placeholder="31/10/2026" style="width: 110px;">
          </div>
        </div>
      </div>

      <button id="btnRun" class="btn-primary" onclick="runSummaryReport()">🚀 Kết xuất Báo cáo Doanh thu</button>
      <div id="statusBox" class="status"></div>
    </div>

    <!-- MÀN HÌNH 2: KẾT QUẢ VÀ CÁC NÚT ĐIỀU KHIỂN -->
    <div id="successView">
      <div class="success-title">🎉 KẾT XUẤT DOANH THU THÀNH CÔNG!</div>
      <div class="kpi-card">
        <div style="font-weight: 600; color: #1a73e8;" id="resSheetName">Bảng tổng hợp: ...</div>
        <div style="margin-top: 6px; font-size: 13px; line-height: 1.6;">
          🎯 <b>Tổng DTPK toàn tỉnh:</b> <span id="resDtpk" style="color: #c5221f; font-weight: bold;">0</span> VNĐ<br>
          📊 <b>Tổng số thuê bao:</b> <span id="resTotalTB" style="font-weight: bold;">0</span> TB<br>
          <i>(Hiện hữu: <span id="resOldTB">0</span> TB | PTM mới: <span id="resNewTB">0</span> TB)</i>
        </div>
      </div>

      <button id="btnDlSummary" class="btn-download" onclick="startDownloadSummary()">📥 Tải Excel Bảng Tổng hợp (.xlsx)</button>
      <button id="btnCreateDetail" class="btn-secondary" onclick="triggerCreateDetailSheet()">📋 Kết xuất Danh sách Chi tiết sang tab riêng</button>
      <button class="btn-close" onclick="closeDialogCleanly()">Đóng cửa sổ</button>
    </div>

    <iframe id="downloadRevenueFrame" style="display:none;"></iframe>

    <script>
      let currentSummarySheetName = '';
      let currentCreatedParams = null;
      let currentTempRevenueFileId = null;

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
        if (isMonth) document.getElementById('monthVal').focus();
        else document.getElementById('fromDate').focus();
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

      function runSummaryReport() {
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
        btn.innerText = '⏳ Đang tính toán doanh thu...';
        status.className = 'status info';
        status.innerText = 'Hệ thống đang tổng hợp DTPK theo từng Khu vực...';

        google.script.run
          .withSuccessHandler(function(res) {
            currentSummarySheetName = res.sheetName;
            currentCreatedParams = params;

            document.getElementById('formView').style.display = 'none';
            document.getElementById('resSheetName').innerText = '📁 Bảng tổng hợp: ' + res.sheetName;
            document.getElementById('resDtpk').innerText = res.totalDtpk;
            document.getElementById('resTotalTB').innerText = res.totalTB;
            document.getElementById('resOldTB').innerText = res.oldTB;
            document.getElementById('resNewTB').innerText = res.newTB;
            document.getElementById('successView').style.display = 'block';
          })
          .withFailureHandler(function(err) {
            btn.disabled = false;
            btn.innerText = '🚀 Kết xuất Báo cáo Doanh thu';
            status.style.display = 'none';
            showError(err.message);
          })
          .processRevenueAllocationSummary(params);
      }

      function startDownloadSummary() {
        const btn = document.getElementById('btnDlSummary');
        btn.disabled = true;
        btn.innerText = '⏳ Đang tạo file Excel...';

        google.script.run
          .withSuccessHandler(function(res) {
            btn.disabled = false;
            btn.innerText = '📥 Tải Excel Bảng Tổng hợp (.xlsx)';
            currentTempRevenueFileId = res.fileId;
            document.getElementById('downloadRevenueFrame').src = res.downloadUrl;

            setTimeout(function() {
              if (currentTempRevenueFileId) {
                google.script.run.deleteTempFile(currentTempRevenueFileId);
                currentTempRevenueFileId = null;
              }
            }, 5000);
          })
          .withFailureHandler(function(err) {
            btn.disabled = false;
            btn.innerText = '📥 Tải Excel Bảng Tổng hợp (.xlsx)';
            alert('Lỗi tải file: ' + err.message);
          })
          .prepareCleanRevenueDownloadUrl(currentSummarySheetName);
      }

      function triggerCreateDetailSheet() {
        if (!currentCreatedParams) return;
        const btn = document.getElementById('btnCreateDetail');
        btn.disabled = true;
        btn.innerText = '⏳ Đang xuất Danh sách Chi tiết sang tab riêng...';

        google.script.run
          .withSuccessHandler(function(res) {
            btn.disabled = false;
            btn.innerText = '✅ Đã tạo sheet: ' + res.sheetName;
            alert('🎉 ĐÃ KẾT XUẤT CHI TIẾT THÀNH CÔNG!\\n\\n' +
                  '📁 Tên trang tính: ' + res.sheetName + '\\n' +
                  '📋 Số lượng: ' + res.count + ' thuê bao đóng góp doanh thu.\\n\\n' +
                  '👉 Trang tính chi tiết đã được mở sẵn cho bạn.');
          })
          .withFailureHandler(function(err) {
            btn.disabled = false;
            btn.innerText = '📋 Kết xuất Danh sách Chi tiết sang tab riêng';
            alert('Lỗi xuất chi tiết: ' + err.message);
          })
          .processRevenueAllocationDetail(currentCreatedParams);
      }

      function closeDialogCleanly() {
        if (currentTempRevenueFileId) {
          google.script.run.deleteTempFile(currentTempRevenueFileId);
          currentTempRevenueFileId = null;
        }
        google.script.host.close();
      }
    </script>
  </body>
  </html>
  `;
}