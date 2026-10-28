/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * FILE        : WebApp_API.gs
 * PHIÊN BẢN   : 13.0 (Executive Dashboard - Tích hợp Tăng trưởng, ARPU, Thu hồi ONT & Cơ cấu KH)
 * NGÀY CẬP NHẬT: 28/10/2026
 * TÁC GIẢ     : Trịnh Vũ Thế Vinh & AI Assistant
 * ============================================================================
 */

const STANDARD_REGIONS = [
  'TT.KDVT',
  'KV Đà Lạt',
  'KV Bảo Lộc',
  'KV Gia Nghĩa',
  'KV Phan Thiết',
  'KV Hàm Tân'
];

function safeParseCurrency(val) {
  if (typeof val === 'number') return Math.round(val);
  if (!val) return 0;
  const clean = String(val).replace(/[^0-9]/g, '');
  return parseInt(clean, 10) || 0;
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setTitle('MobiFiber Executive Dashboard - MobiFone Lâm Đồng')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDateRangeByPeriod(periodType, selectedYear, selectedMonth) {
  const now = new Date();
  const y = parseInt(selectedYear, 10) || now.getFullYear();
  const m = selectedMonth !== undefined ? parseInt(selectedMonth, 10) : now.getMonth();

  let start = null, end = null;
  switch (periodType) {
    case 'week': {
      const currentDay = now.getDay();
      const diffToMonday = (currentDay === 0 ? -6 : 1) - currentDay;
      start = new Date(now.getFullYear(), now.getMonth(), now.getDate() + diffToMonday, 0, 0, 0);
      end = new Date(now.getFullYear(), now.getMonth(), now.getDate() + diffToMonday + 6, 23, 59, 59);
      break;
    }
    case 'month':
      start = new Date(y, m, 1, 0, 0, 0); end = new Date(y, m + 1, 0, 23, 59, 59); break;
    case 'q1': start = new Date(y, 0, 1, 0, 0, 0); end = new Date(y, 2, 31, 23, 59, 59); break;
    case 'q2': start = new Date(y, 3, 1, 0, 0, 0); end = new Date(y, 5, 30, 23, 59, 59); break;
    case 'q3': start = new Date(y, 6, 1, 0, 0, 0); end = new Date(y, 8, 30, 23, 59, 59); break;
    case 'q4': start = new Date(y, 9, 1, 0, 0, 0); end = new Date(y, 11, 31, 23, 59, 59); break;
    case '6tdn': start = new Date(y, 0, 1, 0, 0, 0); end = new Date(y, 5, 30, 23, 59, 59); break;
    case '9tdn': start = new Date(y, 0, 1, 0, 0, 0); end = new Date(y, 8, 30, 23, 59, 59); break;
    case 'year': start = new Date(y, 0, 1, 0, 0, 0); end = new Date(y, 11, 31, 23, 59, 59); break;
    default: start = new Date(y, m, 1, 0, 0, 0); end = new Date(y, m + 1, 0, 23, 59, 59);
  }
  return { start, end };
}

// Hàm lấy kỳ trước để tính Tăng trưởng MoM/YoY
function getPrevDateRange(periodType, y, m) {
  if (periodType === 'month') {
    let pM = m - 1, pY = y;
    if (pM < 0) { pM = 11; pY = y - 1; }
    return { start: new Date(pY, pM, 1, 0, 0, 0), end: new Date(pY, pM + 1, 0, 23, 59, 59) };
  } else if (periodType === 'year') {
    return { start: new Date(y - 1, 0, 1, 0, 0, 0), end: new Date(y - 1, 11, 31, 23, 59, 59) };
  }
  return null; // Các kỳ khác bỏ qua MoM cho nhẹ hệ thống
}

function calculateDtpkForDateRange(rawData, indices, staffRegionMap, startDate, endDate) {
  let grandTotalOldTB = 0, grandTotalOldDtpk = 0;
  let grandTotalNewTB = 0, grandTotalNewRevenue = 0;
  
  let subKhyc = 0, subNeedRenew = 0, subRealDebt = 0, subCancelled = 0;
  let unrecoveredOnt = 0; // Đếm ONT chưa thu hồi
  let segGov = 0, segBiz = 0, segInd = 0; // Cơ cấu Khách hàng

  const regionMap = new Map();
  STANDARD_REGIONS.forEach(reg => {
    regionMap.set(reg, { currentSub: 0, currentDtpk: 0, ptmSub: 0, ptmDtpk: 0, totalSub: 0, totalDtpk: 0 });
  });

  const ptmPackageMap = new Map();
  const startTime = startDate.getTime();
  const endTime = endDate.getTime();

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const loaiKH = normalizeHeader(String(row[indices.idxLoaiKH] || ''));
    if (loaiKH === 'thuê bao nghiệp vụ' || loaiKH === 'thue bao nghiep vu') continue;

    const tinhTrangHD = normalizeHeader(String(row[indices.idxTinhTrangHD] || ''));
    const isCancelled = tinhTrangHD.includes('hủy') || tinhTrangHD.includes('huy') ||
                        tinhTrangHD.includes('thanh lý') || tinhTrangHD.includes('thanh ly') ||
                        tinhTrangHD.includes('chấm dứt') || tinhTrangHD.includes('cham dut');
                        
    // Cảnh báo ONT: Nếu Hủy HĐ mà vẫn còn Serial
    if (isCancelled) {
      subCancelled++;
      const serial = String(row[indices.idxSerial] || '').trim();
      if (serial && serial.toLowerCase() !== 'khong' && serial !== '-' && serial !== '---') {
        unrecoveredOnt++;
      }
      continue;
    }

    const chuKy = parseInt(row[indices.idxChuKy], 10) || 0;
    const giaCuoc = parseCurrencyToInteger(row[indices.idxGiaCuoc]) || 0;
    const dtpk = parseCurrencyToInteger(row[indices.idxDtpk]) || (chuKy > 0 ? Math.round(giaCuoc / chuKy) : 0);

    const dDauNoi = parseToDate(row[indices.idxNgayDauNoi]);
    const dThiCong = parseToDate(row[indices.idxNgayThiCong]);
    const dBatDau = parseToDate(row[indices.idxNgayBatDau]);
    const dKetThuc = parseToDate(row[indices.idxNgayKetThuc]);
    const nhuCau = normalizeHeader(String(row[indices.idxNhuCau] || ''));
    const chanCat = normalizeHeader(String(row[indices.idxChanCat] || ''));

    const isDebt = chanCat.includes('1c') || chanCat.includes('nợ cước') || chanCat.includes('no cuoc');
    const isKhyc = chanCat.includes('khyc') || chanCat.includes('yêu cầu') || chanCat.includes('yeu cau');

    if (isKhyc) {
      subKhyc++;
    } else if (isDebt && (!dKetThuc || dKetThuc.getTime() < startTime)) {
      if (chuKy > 1) subNeedRenew++; else subRealDebt++;
    }

    if (isDebt && dKetThuc && dKetThuc.getTime() < startTime) continue;

    let isPtmInPeriod = false;
    if (nhuCau === 'đăng ký mới' || nhuCau === 'dang ky moi') {
      let ptmDate = null;
      if (chuKy === 1) {
        if (dDauNoi && dThiCong) ptmDate = (dThiCong.getTime() > dDauNoi.getTime()) ? dThiCong : dDauNoi;
        else ptmDate = dThiCong || dDauNoi || dBatDau;
      } else {
        ptmDate = dBatDau;
      }
      if (ptmDate && ptmDate.getTime() >= startTime && ptmDate.getTime() <= endTime) isPtmInPeriod = true;
    }

    let isRetainedInPeriod = false;
    if (!isPtmInPeriod) {
      if (chuKy === 1) {
        const activeDate = dBatDau || dDauNoi;
        if (activeDate && activeDate.getTime() <= endTime) isRetainedInPeriod = true;
      } else {
        if (dKetThuc && dKetThuc.getTime() >= startTime) isRetainedInPeriod = true;
      }
    }

    if (!isPtmInPeriod && !isRetainedInPeriod) continue;

    // Phân loại Cơ cấu KH
    if (loaiKH.includes('hành chính') || loaiKH.includes('cơ quan') || loaiKH.includes('ủy ban') || loaiKH.includes('công an') || loaiKH.includes('bca') || loaiKH.includes('vũ trang') || loaiKH.includes('trường') || loaiKH.includes('y tế') || loaiKH.includes('bệnh viện')) {
      segGov++;
    } else if (loaiKH.includes('doanh nghiệp') || loaiKH.includes('công ty') || loaiKH.includes('tnhh') || loaiKH.includes('cp') || loaiKH.includes('dntn') || loaiKH.includes('hợp tác xã')) {
      segBiz++;
    } else {
      segInd++;
    }

    const rawMaNV = normalizeHeader(String(row[indices.idxMaNV] || ''));
    let khuVuc = 'TT.KDVT';
    if (rawMaNV && staffRegionMap.has(rawMaNV)) khuVuc = staffRegionMap.get(rawMaNV);
    let matchedRegion = STANDARD_REGIONS.find(reg => normalizeHeader(reg) === normalizeHeader(khuVuc)) || 'TT.KDVT';
    const regStat = regionMap.get(matchedRegion);

    if (isPtmInPeriod) {
      regStat.ptmSub += 1; regStat.ptmDtpk += dtpk; regStat.totalSub += 1; regStat.totalDtpk += dtpk;
      grandTotalNewTB += 1; grandTotalNewRevenue += dtpk;
      const goiCuoc = String(row[indices.idxGoiCuoc] || 'Khác').trim().toUpperCase();
      ptmPackageMap.set(goiCuoc, (ptmPackageMap.get(goiCuoc) || 0) + 1);
    } else if (isRetainedInPeriod) {
      regStat.currentSub += 1; regStat.currentDtpk += dtpk; regStat.totalSub += 1; regStat.totalDtpk += dtpk;
      grandTotalOldTB += 1; grandTotalOldDtpk += dtpk;
    }
  }

  return {
    revenueCurrent: grandTotalOldDtpk, revenuePtm: grandTotalNewRevenue, totalRevenue: grandTotalOldDtpk + grandTotalNewRevenue,
    subCurrent: grandTotalOldTB, subPtm: grandTotalNewTB, totalActiveSub: grandTotalOldTB + grandTotalNewTB,
    subNeedRenew, subRealDebt, subKhyc, subCancelled, unrecoveredOnt, 
    segGov, segBiz, segInd,
    regionMap, ptmPackageMap
  };
}

function getDashboardMetrics(periodType = 'month', selectedYear, selectedMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  const yTarget = parseInt(selectedYear, 10) || new Date().getFullYear();
  const mTarget = selectedMonth !== undefined ? parseInt(selectedMonth, 10) : new Date().getMonth();

  if (lastRow < CONFIG.DATA_START_ROW_TARGET) return { success: false };

  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));
  const indices = {
    idxMaKH: headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME)),
    idxLoaiKH: headers.findIndex(h => h === 'loại khách hàng'),
    idxNhuCau: headers.findIndex(h => h === 'nhu cầu hòa mạng'),
    idxGoiCuoc: headers.findIndex(h => h === 'gói cước'),
    idxGiaCuoc: headers.findIndex(h => h === 'giá cước'),
    idxChuKy: headers.findIndex(h => h === 'chu kỳ sử dụng'),
    idxDtpk: headers.findIndex(h => h === 'dtpk'),
    idxTinhTrangHD: headers.findIndex(h => h === 'tình trạng hợp đồng'),
    idxChanCat: headers.findIndex(h => h === 'tình trạng chặn cắt'),
    idxMaNV: headers.findIndex(h => h === 'mã nhân viên'),
    idxNgayDauNoi: headers.findIndex(h => h === 'ngày đấu nối'),
    idxNgayThiCong: headers.findIndex(h => h === 'ngày hoàn tất thi công'),
    idxNgayBatDau: headers.findIndex(h => h === 'ngày băt đầu gói cước' || h === 'ngày bắt đầu gói cước'),
    idxNgayKetThuc: headers.findIndex(h => h === 'ngày kết thúc gói cước'),
    idxSerial: headers.findIndex(h => h === 'serial ont')
  };

  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, lastRow - CONFIG.DATA_START_ROW_TARGET + 1, lastCol).getValues();
  const staffRegionMap = new Map();
  const thamSoSheet = ss.getSheetByName('thamso_caidat');
  if (thamSoSheet && thamSoSheet.getLastRow() >= 4) {
    const tsData = thamSoSheet.getRange(4, 10, thamSoSheet.getLastRow() - 3, 4).getValues();
    tsData.forEach(r => {
      const kv = String(r[0] || '').trim();
      const maNV = normalizeHeader(String(r[2] || ''));
      if (maNV && kv) staffRegionMap.set(maNV, kv);
    });
  }

  // 1. TÍNH DỮ LIỆU HIỆN TẠI
  const { start: dateStart, end: dateEnd } = getDateRangeByPeriod(periodType, yTarget, mTarget);
  const currentResult = calculateDtpkForDateRange(rawData, indices, staffRegionMap, dateStart, dateEnd);

  // 1.5. TÍNH DỮ LIỆU KỲ TRƯỚC (ĐỂ SO SÁNH TĂNG TRƯỞNG)
  let prevTotalRev = 0, prevTotalSub = 0, prevPtmSub = 0;
  const prevDateRange = getPrevDateRange(periodType, yTarget, mTarget);
  if (prevDateRange) {
    const prevResult = calculateDtpkForDateRange(rawData, indices, staffRegionMap, prevDateRange.start, prevDateRange.end);
    prevTotalRev = prevResult.totalRevenue;
    prevTotalSub = prevResult.totalActiveSub;
    prevPtmSub = prevResult.subPtm;
  }

  // 2. TÍNH 12 THÁNG XU HƯỚNG
  const monthlyTrend = [];
  for (let m = 0; m < 12; m++) {
    const mStart = new Date(yTarget, m, 1, 0, 0, 0);
    const mEnd = new Date(yTarget, m + 1, 0, 23, 59, 59);

    if (m <= mTarget) {
      const mRes = calculateDtpkForDateRange(rawData, indices, staffRegionMap, mStart, mEnd);
      monthlyTrend.push({
        monthLabel: 'T' + (m + 1), monthIndex: m, hasData: true,
        currentDtpk: mRes.revenueCurrent, ptmDtpk: mRes.revenuePtm, totalDtpk: mRes.totalRevenue,
        currentSub: mRes.subCurrent, ptmSub: mRes.subPtm
      });
    } else {
      monthlyTrend.push({ monthLabel: 'T' + (m + 1), monthIndex: m, hasData: false, currentDtpk: 0, ptmDtpk: 0, totalDtpk: 0, currentSub: 0, ptmSub: 0 });
    }
  }

  const regionStats = STANDARD_REGIONS.map(regName => {
    const s = currentResult.regionMap.get(regName);
    return { name: regName, currentSub: s.currentSub, ptmSub: s.ptmSub, totalSub: s.totalSub, currentDtpk: s.currentDtpk, ptmDtpk: s.ptmDtpk, totalDtpk: s.totalDtpk };
  });

  const packagePtmDistribution = Array.from(currentResult.ptmPackageMap.entries())
    .map(([pkg, count]) => ({ package: pkg, count: count, ratio: currentResult.subPtm > 0 ? ((count / currentResult.subPtm) * 100).toFixed(1) : 0 }))
    .sort((a, b) => b.count - a.count);

  return {
    success: true, period: periodType, selectedYear: yTarget, selectedMonth: mTarget,
    summary: {
      totalRevenue: currentResult.totalRevenue, revenueCurrent: currentResult.revenueCurrent, revenuePtm: currentResult.revenuePtm,
      totalActiveSub: currentResult.totalActiveSub, subCurrent: currentResult.subCurrent, subPtm: currentResult.subPtm,
      subNeedRenew: currentResult.subNeedRenew, subRealDebt: currentResult.subRealDebt, subKhyc: currentResult.subKhyc, 
      unrecoveredOnt: currentResult.unrecoveredOnt, 
      segGov: currentResult.segGov, segBiz: currentResult.segBiz, segInd: currentResult.segInd,
      prevTotalRev, prevTotalSub, prevPtmSub
    },
    regionStats: regionStats, monthlyTrend: monthlyTrend, packagePtmDistribution: packagePtmDistribution,
    lastUpdated: Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm')
  };
}

function apiSearchSubscriber(keyword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) return [];

  const qClean = String(keyword || '').replace(/[\s\uFEFF\xA0]+/g, '').toLowerCase();
  const qDigits = qClean.replace(/[^0-9]/g, '');
  if (!qClean) return [];

  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));
  const idxMaKH = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxUser = headers.findIndex(h => h === 'user');
  const idxTenTB = headers.findIndex(h => h === 'tên thuê bao');
  const idxSDT = headers.findIndex(h => h === 'sđt liên hệ');
  const idxDiaChi = headers.findIndex(h => h === 'địa chỉ lắp đặt');
  const idxGoiCuoc = headers.findIndex(h => h === 'gói cước');
  const idxGiaCuoc = headers.findIndex(h => h === 'giá cước');
  const idxChuKy = headers.findIndex(h => h === 'chu kỳ sử dụng');
  const idxDtpk = headers.findIndex(h => h === 'dtpk');
  const idxNgayBatDau = headers.findIndex(h => h === 'ngày băt đầu gói cước' || h === 'ngày bắt đầu gói cước');
  const idxNgayKetThuc = headers.findIndex(h => h === 'ngày kết thúc gói cước');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxChanCat = headers.findIndex(h => h === 'tình trạng chặn cắt');
  const idxLoaiONT = headers.findIndex(h => h === 'loại thiết bị ont');
  const idxSerial = headers.findIndex(h => h === 'serial ont');
  const idxTenCH = headers.findIndex(h => h === 'tên cửa hàng');
  const idxTenNV = headers.findIndex(h => h === 'tên nhân viên');
  const idxMaNV = headers.findIndex(h => h === 'mã nhân viên');

  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, lastRow - CONFIG.DATA_START_ROW_TARGET + 1, lastCol).getValues();
  const staffRegionMap = new Map();
  const thamSoSheet = ss.getSheetByName('thamso_caidat');
  if (thamSoSheet && thamSoSheet.getLastRow() >= 4) {
    const tsData = thamSoSheet.getRange(4, 10, thamSoSheet.getLastRow() - 3, 4).getValues();
    tsData.forEach(r => {
      const kv = String(r[0] || '').trim();
      const maNV = normalizeHeader(String(r[2] || ''));
      if (maNV && kv) staffRegionMap.set(maNV, kv);
    });
  }

  const matched = [];
  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    let maKH = String(row[idxMaKH] || '').replace(/[\s\uFEFF\xA0]+/g, '').replace(/^'+/, '');
    if (/^\d{8}$/.test(maKH)) maKH = '00' + maKH;

    const userStr = idxUser !== -1 ? String(row[idxUser] || '').toLowerCase().trim() : '';
    const sdt = String(row[idxSDT] || '').replace(/[^0-9]/g, '');
    const serial = String(row[idxSerial] || '').replace(/[\s\uFEFF\xA0]+/g, '').toLowerCase();

    const matchMaKH = qDigits && maKH.includes(qDigits);
    const matchUser = userStr && userStr.includes(qClean);
    const matchSDT = qDigits && sdt.includes(qDigits);
    const matchSerial = qClean && serial.includes(qClean);

    if (matchMaKH || matchUser || matchSDT || matchSerial) {
      const giaCuoc = parseCurrencyToInteger(row[idxGiaCuoc]) || 0;
      const chuKy = parseInt(row[idxChuKy], 10) || 1;
      let dtpk = parseCurrencyToInteger(row[idxDtpk]) || (chuKy > 0 ? Math.round(giaCuoc / chuKy) : 0);

      const dBatDau = parseToDate(row[idxNgayBatDau]);
      const dKetThuc = parseToDate(row[idxNgayKetThuc]);
      const rawMaNV = normalizeHeader(String(row[idxMaNV] || ''));
      let khuVuc = 'TT.KDVT';
      if (rawMaNV && staffRegionMap.has(rawMaNV)) khuVuc = staffRegionMap.get(rawMaNV);

      matched.push({
        maKH: maKH, user: idxUser !== -1 ? String(row[idxUser] || '').trim() : '',
        tenTB: String(row[idxTenTB] || 'Chưa cập nhật').trim(), sdt: String(row[idxSDT] || '').replace(/^'+/, ''),
        diaChi: String(row[idxDiaChi] || '---').trim(), goiCuoc: String(row[idxGoiCuoc] || '---').trim(),
        giaCuoc: giaCuoc, chuKy: chuKy, dtpk: dtpk,
        ngayBatDau: dBatDau ? Utilities.formatDate(dBatDau, 'Asia/Saigon', 'dd/MM/yyyy') : '---',
        ngayKetThuc: dKetThuc ? Utilities.formatDate(dKetThuc, 'Asia/Saigon', 'dd/MM/yyyy') : '---',
        tinhTrangHD: String(row[idxTinhTrangHD] || 'Đang hoạt động').trim(),
        tinhTrangChan: String(row[idxChanCat] || 'Bình thường').trim(),
        loaiONT: String(row[idxLoaiONT] || '---').trim(),
        serial: String(row[idxSerial] || '---').replace(/^'+/, '').trim(),
        khuVuc: khuVuc, tenCH: String(row[idxTenCH] || '---').trim(), tenNV: String(row[idxTenNV] || '---').trim()
      });
      if (matched.length >= 25) break;
    }
  }
  return matched;
}