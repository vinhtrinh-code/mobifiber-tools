/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * FILE        : WebApp_API.gs
 * PHIÊN BẢN   : 14.0 (Thuật toán Lũy Kế Đa Kỳ, Tổng hợp Doanh thu chính xác)
 * NGÀY CẬP NHẬT: 29/10/2026
 * ============================================================================
 */

const STANDARD_REGIONS = [
  'TT.KDVT', 'KV Đà Lạt', 'KV Bảo Lộc', 'KV Gia Nghĩa', 'KV Phan Thiết', 'KV Hàm Tân'
];

function safeParseCurrency(val) {
  if (typeof val === 'number') return Math.round(val);
  if (!val) return 0;
  return parseInt(String(val).replace(/[^0-9]/g, ''), 10) || 0;
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setTitle('MobiFiber Executive Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Tính toán lát cắt dữ liệu cho 1 khoảng thời gian cố định
function calculateDtpkForDateRange(rawData, indices, staffRegionMap, startDate, endDate) {
  let grandTotalOldTB = 0, grandTotalOldDtpk = 0;
  let grandTotalNewTB = 0, grandTotalNewRevenue = 0;
  let subKhyc = 0, subNeedRenew = 0, subRealDebt = 0, subCancelled = 0;
  let unrecoveredOnt = 0, segGov = 0, segBiz = 0, segInd = 0;

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
    const isCancelled = tinhTrangHD.includes('hủy') || tinhTrangHD.includes('thanh lý') || tinhTrangHD.includes('chấm dứt');
                        
    if (isCancelled) {
      subCancelled++;
      const serial = String(row[indices.idxSerial] || '').trim();
      if (serial && serial.toLowerCase() !== 'khong' && serial !== '-' && serial !== '---') unrecoveredOnt++;
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

    const isDebt = chanCat.includes('1c') || chanCat.includes('nợ cước');
    const isKhyc = chanCat.includes('khyc') || chanCat.includes('yêu cầu');

    if (isKhyc) subKhyc++;
    else if (isDebt && (!dKetThuc || dKetThuc.getTime() < startTime)) {
      if (chuKy > 1) subNeedRenew++; else subRealDebt++;
    }

    if (isDebt && dKetThuc && dKetThuc.getTime() < startTime) continue;

    let isPtmInPeriod = false;
    if (nhuCau.includes('mới')) {
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
    subNeedRenew, subRealDebt, subKhyc, subCancelled, unrecoveredOnt, segGov, segBiz, segInd,
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
  
  // Xác định các tháng thuộc kỳ báo cáo
  let targetMonths = [];
  let endMonthIndex = 0;
  const sm = parseInt(selectedMonth, 10);
  
  if (periodType === 'month') {
    targetMonths = [sm]; endMonthIndex = sm;
  } else if (periodType === 'quarter') {
    if (sm === 1) { targetMonths = [0,1,2]; endMonthIndex = 2; }
    else if (sm === 2) { targetMonths = [3,4,5]; endMonthIndex = 5; }
    else if (sm === 3) { targetMonths = [6,7,8]; endMonthIndex = 8; }
    else { targetMonths = [9,10,11]; endMonthIndex = 11; }
  } else if (periodType === 'ytd') {
    for (let i = 0; i < sm; i++) targetMonths.push(i);
    endMonthIndex = sm - 1;
  } else if (periodType === 'year') {
    for (let i = 0; i < 12; i++) targetMonths.push(i);
    endMonthIndex = 11;
  }

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

  // Khởi tạo biến cộng dồn cho Kỳ Báo cáo
  const monthlyTrend = [];
  const aggregatedSummary = {
    totalRevenue: 0, revenueCurrent: 0, revenuePtm: 0,
    totalActiveSub: 0, subCurrent: 0, subPtm: 0,
    subNeedRenew: 0, subRealDebt: 0, subKhyc: 0, subCancelled: 0, unrecoveredOnt: 0,
    segGov: 0, segBiz: 0, segInd: 0, prevTotalRev: 0, prevTotalSub: 0, prevPtmSub: 0
  };
  
  const aggregatedRegionMap = new Map();
  STANDARD_REGIONS.forEach(reg => {
    aggregatedRegionMap.set(reg, { currentSub: 0, currentDtpk: 0, ptmSub: 0, ptmDtpk: 0, totalSub: 0, totalDtpk: 0 });
  });
  const aggregatedPtmPackageMap = new Map();

  // Chạy vòng lặp tính từng tháng để vẽ Trend 12 tháng, đồng thời cộng dồn vào aggregatedSummary
  for (let m = 0; m < 12; m++) {
    const mStart = new Date(yTarget, m, 1, 0, 0, 0);
    const mEnd = new Date(yTarget, m + 1, 0, 23, 59, 59);
    
    // Tính số liệu riêng cho tháng m
    const mRes = calculateDtpkForDateRange(rawData, indices, staffRegionMap, mStart, mEnd);
    
    monthlyTrend.push({
      monthLabel: 'T' + (m + 1), monthIndex: m, hasData: true,
      currentDtpk: mRes.revenueCurrent, ptmDtpk: mRes.revenuePtm, totalDtpk: mRes.totalRevenue,
      currentSub: mRes.subCurrent, ptmSub: mRes.subPtm
    });

    // Nếu tháng m nằm trong kỳ báo cáo do người dùng chọn (Ví dụ Q1 -> tháng 1,2,3)
    if (targetMonths.includes(m)) {
      // 1. DOANH THU & PTM LÀ TỔNG CỘNG CỦA CÁC THÁNG TRONG KỲ
      aggregatedSummary.revenueCurrent += mRes.revenueCurrent;
      aggregatedSummary.revenuePtm += mRes.revenuePtm;
      aggregatedSummary.totalRevenue += mRes.totalRevenue;
      aggregatedSummary.subPtm += mRes.subPtm;
      
      for (const [pkg, count] of mRes.ptmPackageMap.entries()) {
        aggregatedPtmPackageMap.set(pkg, (aggregatedPtmPackageMap.get(pkg) || 0) + count);
      }
      
      for (const reg of STANDARD_REGIONS) {
        const s1 = aggregatedRegionMap.get(reg);
        const s2 = mRes.regionMap.get(reg);
        s1.ptmSub += s2.ptmSub;
        s1.ptmDtpk += s2.ptmDtpk;
        s1.currentDtpk += s2.currentDtpk;
        s1.totalDtpk += s2.totalDtpk;
      }

      // 2. THUÊ BAO HIỆN HỮU & CÁC TRẠNG THÁI LÀ LÁT CẮT (SNAPSHOT) CỦA THÁNG CUỐI KỲ
      if (m === endMonthIndex) {
        aggregatedSummary.subCurrent = mRes.subCurrent;
        aggregatedSummary.totalActiveSub = mRes.totalActiveSub;
        aggregatedSummary.subNeedRenew = mRes.subNeedRenew;
        aggregatedSummary.subRealDebt = mRes.subRealDebt;
        aggregatedSummary.subKhyc = mRes.subKhyc;
        aggregatedSummary.subCancelled = mRes.subCancelled;
        aggregatedSummary.unrecoveredOnt = mRes.unrecoveredOnt;
        aggregatedSummary.segGov = mRes.segGov;
        aggregatedSummary.segBiz = mRes.segBiz;
        aggregatedSummary.segInd = mRes.segInd;

        for (const reg of STANDARD_REGIONS) {
          const s1 = aggregatedRegionMap.get(reg);
          const s2 = mRes.regionMap.get(reg);
          s1.currentSub = s2.currentSub;
          s1.totalSub = s2.totalSub;
        }
      }
    }
  }

  // Đóng gói mảng RegionStats
  const regionStats = STANDARD_REGIONS.map(regName => {
    const s = aggregatedRegionMap.get(regName);
    return { name: regName, currentSub: s.currentSub, ptmSub: s.ptmSub, totalSub: s.totalSub, currentDtpk: s.currentDtpk, ptmDtpk: s.ptmDtpk, totalDtpk: s.totalDtpk };
  });

  const packagePtmDistribution = Array.from(aggregatedPtmPackageMap.entries())
    .map(([pkg, count]) => ({ package: pkg, count: count, ratio: aggregatedSummary.subPtm > 0 ? ((count / aggregatedSummary.subPtm) * 100).toFixed(1) : 0 }))
    .sort((a, b) => b.count - a.count);

  return {
    success: true, period: periodType, selectedYear: yTarget, selectedMonth: selectedMonth,
    summary: aggregatedSummary,
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

    if (qDigits && maKH.includes(qDigits) || userStr.includes(qClean) || qDigits && sdt.includes(qDigits) || qClean && serial.includes(qClean)) {
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