/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * FILE        : WebApp_API.gs
 * PHIÊN BẢN   : 3.0 (PTM-Strict Rule Engine & 6 Standard Regions Integration)
 * NGÀY TẠO    : 24/10/2026 20:15:00 (GMT+7)
 * NGÀY CẬP NHẬT: 24/10/2026 21:50:00 (GMT+7)
 * TÁC GIẢ     : Trịnh Vũ Thế Vinh & AI Assistant
 * MÔ TẢ       : Backend API phục vụ Web Dashboard đa kỳ chuẩn quy tắc PTM & DTPB
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
    .setTitle('MobiFiber Dashboard - MobiFone Lâm Đồng')
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
    case 'quarter': {
      const qMonth = Math.floor(m / 3) * 3;
      start = new Date(y, qMonth, 1, 0, 0, 0);
      end = new Date(y, qMonth + 3, 0, 23, 59, 59);
      break;
    }
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

function getDashboardMetrics(periodType = 'month') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();

  // Khởi tạo khung 6 khu vực
  const regionMap = new Map();
  STANDARD_REGIONS.forEach(reg => {
    regionMap.set(reg, { currentSub: 0, currentDtpk: 0, ptmSub: 0, ptmRevenue: 0, totalDtpk: 0 });
  });

  if (lastRow < CONFIG.DATA_START_ROW_TARGET) {
    return {
      period: periodType,
      totalSub: 0,
      activeSub: 0,
      blockedSub: 0,
      cancelledSub: 0,
      ptmCount: 0,
      ptmRevenue: 0,
      totalRevenueDtpk: 0,
      regionStats: STANDARD_REGIONS.map(name => ({ name, currentSub: 0, ptmSub: 0, totalSub: 0, totalDtpk: 0 })),
      packageDistribution: [],
      lastUpdated: Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm')
    };
  }

  // 1. ĐỌC DỮ LIỆU ĐÍCH VÀ BẢN ĐỒ THAM SỐ
  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));
  const idxMaKH = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
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

  // Nạp map nhân viên -> khu vực
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

  // 2. TÍNH TOÁN THEO CHUẨN RULE PTM & DTPB
  const { start: dateStart, end: dateEnd } = getDateRangeByPeriod(periodType);

  let activeSub = 0;
  let blockedSub = 0;
  let cancelledSub = 0;
  let ptmCount = 0;
  let ptmRevenue = 0;
  let totalRevenueDtpk = 0;

  const ptmPackageMap = new Map(); // Gom tất cả gói cước PTM trong kỳ

  for (let i = 0; i < rawData.length; i++) {
    const r = rawData[i];
    const statusHD = normalizeHeader(String(r[idxTinhTrangHD] || ''));
    const statusChan = normalizeHeader(String(r[idxChanCat] || ''));
    const loaiKH = normalizeHeader(String(r[idxLoaiKH] || ''));
    const nhuCau = normalizeHeader(String(r[idxNhuCau] || ''));

    const isCancelled = statusHD.includes('hủy') || statusHD.includes('huy') || statusHD.includes('thanh lý') || statusHD.includes('chấm dứt');
    const isBlocked = statusChan.includes('nợ') || statusChan.includes('khóa') || statusChan.includes('chặn') || statusChan.includes('khyc');
    const isNghiepVu = loaiKH.includes('nghiệp vụ') || loaiKH.includes('nghiep vu');

    // Thống kê 3 trạng thái
    if (isCancelled) {
      cancelledSub++;
    } else if (isBlocked) {
      blockedSub++;
    } else {
      activeSub++;
    }

    // Loại trừ hoàn toàn thuê bao đã hủy khỏi doanh thu
    if (isCancelled) continue;

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

    // XÉT TIÊU CHÍ PTM TRONG KỲ (CHUẨN RULE Mod_PTM.gs)
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

    if (isPtmInPeriod) {
      ptmCount++;
      ptmRevenue += giaCuoc;
      regStat.ptmSub++;
      regStat.ptmRevenue += giaCuoc;

      // Gom toàn bộ gói cước PTM
      const goiCuoc = String(r[idxGoiCuoc] || 'Khác').trim().toUpperCase();
      if (!ptmPackageMap.has(goiCuoc)) {
        ptmPackageMap.set(goiCuoc, { count: 0, revenue: 0, dtpk: 0 });
      }
      const pInfo = ptmPackageMap.get(goiCuoc);
      pInfo.count++;
      pInfo.revenue += giaCuoc;
      pInfo.dtpk += dtpk;
    }

    // XÉT TIÊU CHÍ DOANH THU TOÀN TỈNH & DTPK HIỆN HỮU (CHUẨN RULE Mod_DoanhThuPhanBo.gs)
    const isQualifyRevenue = !isNghiepVu && 
                            (statusHD === 'đang hoạt động' || statusHD === 'dang hoat dong') && 
                            (statusChan === 'bình thường' || statusChan === 'binh thuong' || !statusChan);

    if (isQualifyRevenue) {
      const dKetThuc = parseToDate(r[idxNgayKetThuc]);
      const isValidCurrent = !dKetThuc || dKetThuc >= dateStart;

      if (isPtmInPeriod) {
        totalRevenueDtpk += dtpk;
        regStat.totalDtpk += dtpk;
      } else if (isValidCurrent) {
        totalRevenueDtpk += dtpk;
        regStat.currentSub++;
        regStat.currentDtpk += dtpk;
        regStat.totalDtpk += dtpk;
      }
    }
  }

  // 3. ĐÓNG GÓI DỮ LIỆU JSON
  const regionStats = STANDARD_REGIONS.map(regName => {
    const s = regionMap.get(regName);
    return {
      name: regName,
      currentSub: s.currentSub,
      ptmSub: s.ptmSub,
      totalSub: s.currentSub + s.ptmSub,
      totalDtpk: s.totalDtpk,
      ptmRevenue: s.ptmRevenue
    };
  });

  // Toàn bộ gói cước PTM sắp xếp từ lớn đến nhỏ theo số lượng đăng ký
  const packageDistribution = Array.from(ptmPackageMap.entries())
    .map(([pkg, info]) => ({
      package: pkg,
      count: info.count,
      revenue: info.revenue,
      dtpk: info.dtpk,
      ratio: ptmCount > 0 ? ((info.count / ptmCount) * 100).toFixed(1) : 0
    }))
    .sort((a, b) => b.count - a.count);

  return {
    period: periodType,
    totalSub: activeSub + blockedSub + cancelledSub,
    activeSub: activeSub,
    blockedSub: blockedSub,
    cancelledSub: cancelledSub,
    ptmCount: ptmCount,
    ptmRevenue: ptmRevenue,
    totalRevenueDtpk: totalRevenueDtpk,
    regionStats: regionStats,
    packageDistribution: packageDistribution,
    lastUpdated: Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm')
  };
}