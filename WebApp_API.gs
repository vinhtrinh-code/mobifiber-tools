/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * FILE        : WebApp_API.gs
 * PHIÊN BẢN   : 5.0 (Hỗ trợ Tra cứu USER, Tách Nợ cước/KHYC & Dual-View Analytics)
 * NGÀY TẠO    : 24/10/2026 20:15:00 (GMT+7)
 * NGÀY CẬP NHẬT: 25/10/2026 10:00:00 (GMT+7)
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
    .setTitle('MobiFiber Portal - MobiFone Lâm Đồng')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDateRangeByPeriod(periodType) {
  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth();
  const d = now.getDate();

  let start = null;
  let end = new Date(y, m, d, 23, 59, 59);

  switch (periodType) {
    case 'week': {
      const currentDay = now.getDay();
      const diffToMonday = (currentDay === 0 ? -6 : 1) - currentDay;
      start = new Date(y, m, d + diffToMonday, 0, 0, 0);
      break;
    }
    case 'month':
      start = new Date(y, m, 1, 0, 0, 0);
      end = new Date(y, m + 1, 0, 23, 59, 59);
      break;
    case 'q1':
      start = new Date(y, 0, 1, 0, 0, 0);
      end = new Date(y, 2, 31, 23, 59, 59);
      break;
    case 'q2':
      start = new Date(y, 3, 1, 0, 0, 0);
      end = new Date(y, 5, 30, 23, 59, 59);
      break;
    case 'q3':
      start = new Date(y, 6, 1, 0, 0, 0);
      end = new Date(y, 8, 30, 23, 59, 59);
      break;
    case 'q4':
      start = new Date(y, 9, 1, 0, 0, 0);
      end = new Date(y, 11, 31, 23, 59, 59);
      break;
    case '6tdn':
      start = new Date(y, 0, 1, 0, 0, 0);
      end = new Date(y, 5, 30, 23, 59, 59);
      break;
    case '9tdn':
      start = new Date(y, 0, 1, 0, 0, 0);
      end = new Date(y, 8, 30, 23, 59, 59);
      break;
    case 'year':
      start = new Date(y, 0, 1, 0, 0, 0);
      end = new Date(y, 11, 31, 23, 59, 59);
      break;
    default:
      start = new Date(y, m, 1, 0, 0, 0);
      end = new Date(y, m + 1, 0, 23, 59, 59);
  }

  return { start, end };
}

/**
 * API BÓC TÁCH DỮ LIỆU DASHBOARD CHO CẢ DOANH THU & THUÊ BAO
 */
function getDashboardMetrics(periodType = 'month') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();

  const regionMap = new Map();
  STANDARD_REGIONS.forEach(reg => {
    regionMap.set(reg, {
      currentSub: 0,
      currentDtpk: 0,
      ptmSub: 0,
      ptmDtpk: 0,
      totalSub: 0,
      totalDtpk: 0
    });
  });

  if (lastRow < CONFIG.DATA_START_ROW_TARGET) {
    return {
      period: periodType,
      summary: {
        totalRevenue: 0,
        revenueCurrent: 0,
        revenuePtm: 0,
        totalActiveSub: 0,
        subCurrent: 0,
        subPtm: 0,
        subDebt: 0,
        subKhyc: 0,
        subCancelled: 0
      },
      regionStats: STANDARD_REGIONS.map(name => ({
        name, currentSub: 0, ptmSub: 0, totalSub: 0, currentDtpk: 0, ptmDtpk: 0, totalDtpk: 0
      })),
      packagePtmDistribution: [],
      lastUpdated: Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm')
    };
  }

  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));
  const idxMaKH = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxUser = headers.findIndex(h => h === 'user');
  const idxLoaiKH = headers.findIndex(h => h === 'loại khách hàng');
  const idxNhuCau = headers.findIndex(h => h === 'nhu cầu hòa mạng');
  const idxGoiCuoc = headers.findIndex(h => h === 'gói cước');
  const idxGiaCuoc = headers.findIndex(h => h === 'giá cước');
  const idxChuKy = headers.findIndex(h => h === 'chu kỳ sử dụng');
  const idxDtpk = headers.findIndex(h => h === 'dtpk');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxChanCat = headers.findIndex(h => h === 'tình trạng chặn cắt');
  const idxMaNV = headers.findIndex(h => h === 'mã nhân viên');
  const idxNgayDauNoi = headers.findIndex(h => h === 'ngày đấu nối');
  const idxNgayThiCong = headers.findIndex(h => h === 'ngày hoàn tất thi công');
  const idxNgayBatDau = headers.findIndex(h => h === 'ngày băt đầu gói cước' || h === 'ngày bắt đầu gói cước');
  const idxNgayKetThuc = headers.findIndex(h => h === 'ngày kết thúc gói cước');

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

  const { start: dateStart, end: dateEnd } = getDateRangeByPeriod(periodType);

  let subDebt = 0;
  let subKhyc = 0;
  let subCancelled = 0;
  let revenueCurrent = 0;
  let revenuePtm = 0;
  let subCurrent = 0;
  let subPtm = 0;

  const ptmPackageMap = new Map();

  for (let i = 0; i < rawData.length; i++) {
    const r = rawData[i];
    const statusHD = normalizeHeader(String(r[idxTinhTrangHD] || ''));
    const statusChan = normalizeHeader(String(r[idxChanCat] || ''));
    const loaiKH = normalizeHeader(String(r[idxLoaiKH] || ''));
    const nhuCau = normalizeHeader(String(r[idxNhuCau] || ''));

    const isCancelled = statusHD.includes('hủy') || statusHD.includes('huy') || statusHD.includes('thanh lý') || statusHD.includes('chấm dứt');
    const isDebt = statusChan.includes('nợ') || statusChan.includes('no cuoc') || statusChan.includes('khóa cước');
    const isKhyc = statusChan.includes('khyc') || statusChan.includes('yêu cầu') || statusChan.includes('yeu cau');
    const isNghiepVu = loaiKH.includes('nghiệp vụ') || loaiKH.includes('nghiep vu');

    if (isCancelled) {
      subCancelled++;
      continue;
    }
    if (isDebt) subDebt++;
    if (isKhyc) subKhyc++;

    const chuKy = parseInt(r[idxChuKy], 10) || 1;
    const giaCuoc = safeParseCurrency(r[idxGiaCuoc]);
    let dtpk = safeParseCurrency(r[idxDtpk]);
    if (!dtpk && giaCuoc > 0) dtpk = Math.round(giaCuoc / chuKy);

    // Xác định Khu vực chuẩn
    const maNV = normalizeHeader(String(r[idxMaNV] || ''));
    let khuVuc = 'TT.KDVT';
    if (maNV && staffRegionMap.has(maNV)) {
      khuVuc = staffRegionMap.get(maNV);
    }
    let matchedRegion = STANDARD_REGIONS.find(reg => normalizeHeader(reg) === normalizeHeader(khuVuc)) || 'TT.KDVT';
    const regStat = regionMap.get(matchedRegion);

    // 1. RULE PTM CHUẨN
    const isPtmDemand = (nhuCau === 'đăng ký mới' || nhuCau === 'dang ky moi');
    const isPtmActive = (statusHD === 'đang hoạt động' || statusHD === 'dang hoat dong');
    let ptmDate = null;

    if (isPtmDemand && !isNghiepVu && isPtmActive) {
      const dDauNoi = parseToDate(r[idxNgayDauNoi]);
      const dThiCong = parseToDate(r[idxNgayThiCong]);
      const dBatDau = parseToDate(r[idxNgayBatDau]);

      if (chuKy === 1) {
        if (dDauNoi && dThiCong) {
          ptmDate = (dThiCong.getTime() > dDauNoi.getTime()) ? dThiCong : dDauNoi;
        } else {
          ptmDate = dThiCong || dDauNoi || dBatDau;
        }
      } else {
        ptmDate = dBatDau;
      }
    }

    const isPtmInPeriod = ptmDate && (ptmDate >= dateStart && ptmDate <= dateEnd);

    // 2. RULE DOANH THU & HIỆN HỮU CHUẨN
    const isNormalActive = !isNghiepVu && 
                           (statusHD === 'đang hoạt động' || statusHD === 'dang hoat dong') && 
                           (statusChan === 'bình thường' || statusChan === 'binh thuong' || !statusChan);

    if (isPtmInPeriod) {
      subPtm++;
      revenuePtm += dtpk;

      regStat.ptmSub++;
      regStat.ptmDtpk += dtpk;
      regStat.totalSub++;
      regStat.totalDtpk += dtpk;

      const goiCuoc = String(r[idxGoiCuoc] || 'Khác').trim().toUpperCase();
      ptmPackageMap.set(goiCuoc, (ptmPackageMap.get(goiCuoc) || 0) + 1);

    } else if (isNormalActive) {
      const dKetThuc = parseToDate(r[idxNgayKetThuc]);
      const isValidCurrent = !dKetThuc || dKetThuc >= dateStart;

      if (isValidCurrent) {
        subCurrent++;
        revenueCurrent += dtpk;

        regStat.currentSub++;
        regStat.currentDtpk += dtpk;
        regStat.totalSub++;
        regStat.totalDtpk += dtpk;
      }
    }
  }

  const regionStats = STANDARD_REGIONS.map(regName => {
    const s = regionMap.get(regName);
    return {
      name: regName,
      currentSub: s.currentSub,
      ptmSub: s.ptmSub,
      totalSub: s.totalSub,
      currentDtpk: s.currentDtpk,
      ptmDtpk: s.ptmDtpk,
      totalDtpk: s.totalDtpk
    };
  });

  const packagePtmDistribution = Array.from(ptmPackageMap.entries())
    .map(([pkg, count]) => ({
      package: pkg,
      count: count,
      ratio: subPtm > 0 ? ((count / subPtm) * 100).toFixed(1) : 0
    }))
    .sort((a, b) => b.count - a.count);

  return {
    period: periodType,
    summary: {
      totalRevenue: revenueCurrent + revenuePtm,
      revenueCurrent: revenueCurrent,
      revenuePtm: revenuePtm,
      totalActiveSub: subCurrent + subPtm,
      subCurrent: subCurrent,
      subPtm: subPtm,
      subDebt: subDebt,
      subKhyc: subKhyc,
      subCancelled: subCancelled
    },
    regionStats: regionStats,
    packagePtmDistribution: packagePtmDistribution,
    lastUpdated: Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm')
  };
}

/**
 * 3. API TRA CỨU HỒ SƠ 360° (ĐÃ BỔ SUNG TÌM KIẾM THEO USER)
 */
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

    // Bổ sung tìm kiếm theo USER
    const matchMaKH = qDigits && maKH.includes(qDigits);
    const matchUser = userStr && userStr.includes(qClean);
    const matchSDT = qDigits && sdt.includes(qDigits);
    const matchSerial = qClean && serial.includes(qClean);

    if (matchMaKH || matchUser || matchSDT || matchSerial) {
      const giaCuoc = safeParseCurrency(row[idxGiaCuoc]);
      const chuKy = parseInt(row[idxChuKy], 10) || 1;
      let dtpk = safeParseCurrency(row[idxDtpk]);
      if (!dtpk && giaCuoc > 0) dtpk = Math.round(giaCuoc / chuKy);

      const dBatDau = parseToDate(row[idxNgayBatDau]);
      const dKetThuc = parseToDate(row[idxNgayKetThuc]);

      const maNV = normalizeHeader(String(row[idxMaNV] || ''));
      let khuVuc = 'TT.KDVT';
      if (maNV && staffRegionMap.has(maNV)) {
        khuVuc = staffRegionMap.get(maNV);
      }

      matched.push({
        maKH: maKH,
        user: idxUser !== -1 ? String(row[idxUser] || '').trim() : '',
        tenTB: String(row[idxTenTB] || 'Chưa cập nhật').trim(),
        sdt: String(row[idxSDT] || '').replace(/^'+/, ''),
        diaChi: String(row[idxDiaChi] || '---').trim(),
        goiCuoc: String(row[idxGoiCuoc] || '---').trim(),
        giaCuoc: giaCuoc,
        chuKy: chuKy,
        dtpk: dtpk,
        ngayBatDau: dBatDau ? Utilities.formatDate(dBatDau, 'Asia/Saigon', 'dd/MM/yyyy') : '---',
        ngayKetThuc: dKetThuc ? Utilities.formatDate(dKetThuc, 'Asia/Saigon', 'dd/MM/yyyy') : '---',
        tinhTrangHD: String(row[idxTinhTrangHD] || 'Đang hoạt động').trim(),
        tinhTrangChan: String(row[idxChanCat] || 'Bình thường').trim(),
        loaiONT: String(row[idxLoaiONT] || '---').trim(),
        serial: String(row[idxSerial] || '---').replace(/^'+/, '').trim(),
        khuVuc: khuVuc,
        tenCH: String(row[idxTenCH] || '---').trim(),
        tenNV: String(row[idxTenNV] || '---').trim()
      });

      if (matched.length >= 25) break;
    }
  }

  return matched;
}