/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * FILE        : WebApp_API.gs
 * PHIÊN BẢN   : 1.1 (Tương thích chuẩn hàm hệ thống & Tự động parse tiền tệ)
 * NGÀY TẠO    : 24/10/2026 20:15:00 (GMT+7)
 * NGÀY CẬP NHẬT: 24/10/2026 20:30:00 (GMT+7)
 * TÁC GIẢ     : Trịnh Vũ Thế Vinh & AI Assistant
 * MÔ TẢ       : Khởi chạy Web App (doGet) và cung cấp API bóc tách dữ liệu cho Dashboard
 * ============================================================================
 */

/**
 * Hàm phân tích tiền tệ an toàn độc lập (Tránh lỗi ReferenceError)
 */
function safeParseCurrency(val) {
  if (typeof val === 'number') return Math.round(val);
  if (!val) return 0;
  const clean = String(val).replace(/[^0-9]/g, '');
  return parseInt(clean, 10) || 0;
}

/**
 * Endpoint điều phối Web Application
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Dashboard')
    .evaluate()
    .setTitle('MobiFiber Dashboard - MobiFone Lâm Đồng')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * API trích xuất dữ liệu tổng quan cho Dashboard (Xử lý In-Memory)
 */
function getDashboardMetrics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) {
    return {
      totalSub: 0,
      activeSub: 0,
      blockedSub: 0,
      cancelledSub: 0,
      totalRevenue: 0,
      totalDtpk: 0,
      regionStats: [],
      packageDistribution: [],
      lastUpdated: Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm')
    };
  }

  // 1. ĐỌC DỮ LIỆU VÀ HEADERS VÀO RAM
  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));
  const idxMaKH = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxGoiCuoc = headers.findIndex(h => h === 'gói cước');
  const idxGiaCuoc = headers.findIndex(h => h === 'giá cước');
  const idxChuKy = headers.findIndex(h => h === 'chu kỳ sử dụng');
  const idxDtpk = headers.findIndex(h => h === 'dtpk');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxChanCat = headers.findIndex(h => h === 'tình trạng chặn cắt');
  const idxMaNV = headers.findIndex(h => h === 'mã nhân viên');

  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, lastRow - CONFIG.DATA_START_ROW_TARGET + 1, lastCol).getValues();

  // Đọc danh mục Khu vực từ thamso_caidat
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

  // 2. KHỞI TẠO BIẾN THỐNG KÊ
  let totalSub = rawData.length;
  let activeSub = 0;
  let blockedSub = 0;
  let cancelledSub = 0;
  let totalRevenue = 0;
  let totalDtpk = 0;

  const regionMap = new Map();
  const packageMap = new Map();

  for (let i = 0; i < rawData.length; i++) {
    const r = rawData[i];
    const statusHD = normalizeHeader(String(r[idxTinhTrangHD] || ''));
    const statusChan = normalizeHeader(String(r[idxChanCat] || ''));

    const isCancelled = statusHD.includes('hủy') || statusHD.includes('huy') || statusHD.includes('thanh lý') || statusHD.includes('chấm dứt');
    const isBlocked = statusChan.includes('nợ') || statusChan.includes('khóa') || statusChan.includes('chặn') || statusChan.includes('khyc');

    if (isCancelled) {
      cancelledSub++;
    } else if (isBlocked) {
      blockedSub++;
    } else {
      activeSub++;
    }

    const giaCuoc = safeParseCurrency(r[idxGiaCuoc]);
    const chuKy = parseInt(r[idxChuKy], 10) || 1;
    let dtpk = safeParseCurrency(r[idxDtpk]);
    if (!dtpk && giaCuoc > 0) dtpk = Math.round(giaCuoc / chuKy);

    // Chỉ tính doanh thu cho các hợp đồng chưa bị thanh lý/hủy
    if (!isCancelled) {
      totalRevenue += giaCuoc;
      totalDtpk += dtpk;

      // Gom theo Khu vực
      const maNV = normalizeHeader(String(r[idxMaNV] || ''));
      let khuVuc = 'Chưa phân bổ KV';
      if (maNV && staffRegionMap.has(maNV)) {
        khuVuc = staffRegionMap.get(maNV);
      }

      if (!regionMap.has(khuVuc)) {
        regionMap.set(khuVuc, { count: 0, dtpk: 0 });
      }
      const reg = regionMap.get(khuVuc);
      reg.count++;
      reg.dtpk += dtpk;

      // Gom theo Gói cước (Top 6 gói phổ biến nhất)
      const goiCuoc = String(r[idxGoiCuoc] || 'Khác').trim().toUpperCase();
      packageMap.set(goiCuoc, (packageMap.get(goiCuoc) || 0) + 1);
    }
  }

  // 3. ĐÓNG GÓI DỮ LIỆU JSON
  const regionStats = Array.from(regionMap.entries()).map(([name, data]) => ({
    name: name,
    count: data.count,
    dtpk: data.dtpk
  })).sort((a, b) => b.dtpk - a.dtpk);

  const packageDistribution = Array.from(packageMap.entries())
    .map(([pkg, count]) => ({ package: pkg, count: count }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 6);

  return {
    totalSub: totalSub,
    activeSub: activeSub,
    blockedSub: blockedSub,
    cancelledSub: cancelledSub,
    totalRevenue: totalRevenue,
    totalDtpk: totalDtpk,
    regionStats: regionStats,
    packageDistribution: packageDistribution,
    lastUpdated: Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm')
  };
}