/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * FILE        : WebApp_API.gs
 * PHIÊN BẢN   : 2.0 (Multi-Period Filter, Standardized 6 Regions & Strict Revenue Rules)
 * NGÀY TẠO    : 24/10/2026 20:15:00 (GMT+7)
 * NGÀY CẬP NHẬT: 24/10/2026 21:00:00 (GMT+7)
 * TÁC GIẢ     : Trịnh Vũ Thế Vinh & AI Assistant
 * MÔ TẢ       : Backend API phục vụ Web Dashboard đa kỳ (Tuần, Tháng, Quý, 6T, 9T, Năm)
 * ============================================================================
 */

// Danh sách 6 Đơn vị chuẩn hóa áp dụng xuyên suốt toàn dự án
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

/**
 * Xác định khoảng ngày (Start - End) dựa theo kỳ báo cáo
 */
function getDateRangeByPeriod(periodType) {
  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth(); // 0-indexed (0 = T1)
  const d = now.getDate();

  let start = null;
  let end = new Date(y, m, d, 23, 59, 59);

  switch (periodType) {
    case 'week': {
      // Tuần hiện tại (tính từ Thứ 2 đến Chủ nhật)
      const currentDay = now.getDay();
      const diffToMonday = (currentDay === 0 ? -6 : 1) - currentDay;
      start = new Date(y, m, d + diffToMonday, 0, 0, 0);
      break;
    }
    case 'month': // Mặc định: Tháng hiện tại
      start = new Date(y, m, 1, 0, 0, 0);
      end = new Date(y, m + 1, 0, 23, 59, 59);
      break;
    case 'quarter': {
      // Quý hiện tại
      const qMonth = Math.floor(m / 3) * 3;
      start = new Date(y, qMonth, 1, 0, 0, 0);
      end = new Date(y, qMonth + 3, 0, 23, 59, 59);
      break;
    }
    case '6tdn': // 6 Tháng đầu năm (01/01 - 30/06)
      start = new Date(y, 0, 1, 0, 0, 0);
      end = new Date(y, 5, 30, 23, 59, 59);
      break;
    case '9tdn': // 9 Tháng đầu năm (01/01 - 30/09)
      start = new Date(y, 0, 1, 0, 0, 0);
      end = new Date(y, 8, 30, 23, 59, 59);
      break;
    case 'year': // Cả năm hiện tại
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
 * API Bóc tách dữ liệu Dashboard theo kỳ (Mặc định: 'month')
 */
function getDashboardMetrics(periodType = 'month') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  
  // Khởi tạo khung số liệu 6 Khu vực chuẩn hóa (Mặc định bằng 0)
  const regionMap = new Map();
  STANDARD_REGIONS.forEach(reg => {
    regionMap.set(reg, { count: 0, dtpk: 0 });
  });

  if (lastRow < CONFIG.DATA_START_ROW_TARGET) {
    return {
      period: periodType,
      totalSub: 0,
      activeSub: 0,
      blockedSub: 0,
      cancelledSub: 0,
      totalRevenueDtpk: 0,
      regionStats: STANDARD_REGIONS.map(name => ({ name, count: 0, dtpk: 0 })),
      packageDistribution: [],
      lastUpdated: Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm')
    };
  }

  // 1. ĐỌC DỮ LIỆU VÀ HEADERS TRONG RAM
  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));
  const idxMaKH = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxLoaiKH = headers.findIndex(h => h === 'loại khách hàng');
  const idxGoiCuoc = headers.findIndex(h => h === 'gói cước');
  const idxGiaCuoc = headers.findIndex(h => h === 'giá cước');
  const idxChuKy = headers.findIndex(h => h === 'chu kỳ sử dụng');
  const idxDtpk = headers.findIndex(h => h === 'dtpk');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxChanCat = headers.findIndex(h => h === 'tình trạng chặn cắt');
  const idxMaNV = headers.findIndex(h => h === 'mã nhân viên');
  const idxNgayDauNoi = headers.findIndex(h => h === 'ngày đấu nối');
  const idxNgayBatDau = headers.findIndex(h => h === 'ngày băt đầu gói cước' || h === 'ngày bắt đầu gói cước');

  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, lastRow - CONFIG.DATA_START_ROW_TARGET + 1, lastCol).getValues();

  // Đọc bảng tham số ánh xạ nhân sự -> Khu vực
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

  // 2. PHÂN LOẠI & TÍNH TOÁN THEO KỲ
  const { start: dateStart, end: dateEnd } = getDateRangeByPeriod(periodType);

  let totalSub = 0;
  let activeSub = 0;
  let blockedSub = 0;
  let cancelledSub = 0;
  let totalRevenueDtpk = 0;

  const packageMap = new Map(); // Key: Tên gói -> { count, dtpk }

  for (let i = 0; i < rawData.length; i++) {
    const r = rawData[i];

    // Lọc thời gian phát triển theo kỳ
    const dDauNoi = parseToDate(r[idxNgayDauNoi]);
    const dBatDau = parseToDate(r[idxNgayBatDau]);
    const checkDate = dDauNoi || dBatDau;

    // Nếu lọc theo tuần/quý/6T/9T/năm, kiểm tra mốc ngày
    if (periodType !== 'all') {
      if (!checkDate || checkDate < dateStart || checkDate > dateEnd) {
        continue;
      }
    }

    totalSub++;

    const statusHD = normalizeHeader(String(r[idxTinhTrangHD] || ''));
    const statusChan = normalizeHeader(String(r[idxChanCat] || ''));
    const loaiKH = normalizeHeader(String(r[idxLoaiKH] || ''));

    const isCancelled = statusHD.includes('hủy') || statusHD.includes('huy') || statusHD.includes('thanh lý') || statusHD.includes('chấm dứt');
    const isBlocked = statusChan.includes('nợ') || statusChan.includes('khóa') || statusChan.includes('chặn') || statusChan.includes('khyc');
    const isNghiepVu = loaiKH.includes('nghiệp vụ') || loaiKH.includes('nghiep vu');

    // Thống kê số lượng 3 nhóm thuê bao
    if (isCancelled) {
      cancelledSub++;
    } else if (isBlocked) {
      blockedSub++;
    } else {
      activeSub++;
    }

    // TIÊU CHÍ TÍNH DOANH THU TOÀN TỈNH & PHÂN BỔ:
    // 1. Không phải Thuê bao nghiệp vụ
    // 2. Tình trạng hợp đồng: Đang hoạt động
    // 3. Tình trạng chặn cắt: Bình thường
    const isQualifyForRevenue = !isNghiepVu && 
                               (statusHD === 'đang hoạt động' || statusHD === 'dang hoat dong') && 
                               (statusChan === 'bình thường' || statusChan === 'binh thuong' || !statusChan);

    if (isQualifyForRevenue) {
      const giaCuoc = safeParseCurrency(r[idxGiaCuoc]);
      const chuKy = parseInt(r[idxChuKy], 10) || 1;
      let dtpk = safeParseCurrency(r[idxDtpk]);
      if (!dtpk && giaCuoc > 0) dtpk = Math.round(giaCuoc / chuKy);

      totalRevenueDtpk += dtpk;

      // Xác định Khu vực
      const maNV = normalizeHeader(String(r[idxMaNV] || ''));
      let khuVuc = 'Chưa phân bổ KV';
      if (maNV && staffRegionMap.has(maNV)) {
        khuVuc = staffRegionMap.get(maNV);
      }

      // Khớp vào 6 Khu vực chuẩn
      let matchedRegion = STANDARD_REGIONS.find(reg => normalizeHeader(reg) === normalizeHeader(khuVuc));
      if (!matchedRegion) {
        matchedRegion = khuVuc.includes('TT') ? 'TT.KDVT' : 'TT.KDVT';
      }

      if (regionMap.has(matchedRegion)) {
        const regStat = regionMap.get(matchedRegion);
        regStat.count++;
        regStat.dtpk += dtpk;
      }

      // Thống kê Gói cước
      const goiCuoc = String(r[idxGoiCuoc] || 'Khác').trim().toUpperCase();
      if (!packageMap.has(goiCuoc)) {
        packageMap.set(goiCuoc, { count: 0, dtpk: 0 });
      }
      const pkgStat = packageMap.get(goiCuoc);
      pkgStat.count++;
      pkgStat.dtpk += dtpk;
    }
  }

  // 3. ĐÓNG GÓI JSON CHO DASHBOARD
  const regionStats = STANDARD_REGIONS.map(regName => {
    const data = regionMap.get(regName);
    return {
      name: regName,
      count: data.count,
      dtpk: data.dtpk
    };
  });

  // Top 6 gói cước có doanh thu DTPK hoặc số lượng lớn nhất
  const packageDistribution = Array.from(packageMap.entries())
    .map(([pkg, info]) => ({
      package: pkg,
      count: info.count,
      dtpk: info.dtpk,
      ratio: totalRevenueDtpk > 0 ? ((info.dtpk / totalRevenueDtpk) * 100).toFixed(1) : 0
    }))
    .sort((a, b) => b.dtpk - a.dtpk)
    .slice(0, 6);

  return {
    period: periodType,
    totalSub: totalSub,
    activeSub: activeSub,
    blockedSub: blockedSub,
    cancelledSub: cancelledSub,
    totalRevenueDtpk: totalRevenueDtpk,
    regionStats: regionStats,
    packageDistribution: packageDistribution,
    lastUpdated: Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm')
  };
}