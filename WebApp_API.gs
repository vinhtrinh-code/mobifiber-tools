/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * FILE        : WebApp_API.gs
 * PHIÊN BẢN   : 3.1 (Hỗ trợ Dashboard & Cổng Web Tra Cứu Thuê Bao 360°)
 * NGÀY TẠO    : 24/10/2026 20:15:00 (GMT+7)
 * NGÀY CẬP NHẬT: 24/10/2026 22:15:00 (GMT+7)
 * TÁC GIẢ     : Trịnh Vũ Thế Vinh & AI Assistant
 * MÔ TẢ       : Điều hướng trang & Cung cấp API Tra cứu 360° siêu tốc cho Web App
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

/**
 * ĐIỀU HƯỚNG WEB APP:
 * - Mặc định: Dashboard
 * - Link kèm "?page=lookup": Trang tra cứu 360°
 */
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'lookup'; // Mặc định mở Tra cứu nếu muốn
  
  if (page === 'dashboard') {
    return HtmlService.createTemplateFromFile('Dashboard')
      .evaluate()
      .setTitle('MobiFiber Dashboard - MobiFone Lâm Đồng')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Mặc định nạp giao diện Tra cứu
  return HtmlService.createTemplateFromFile('Lookup')
    .evaluate()
    .setTitle('Tra Cứu MobiFiber 360° - MobiFone Lâm Đồng')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * API TRA CỨU HỒ SƠ THUÊ BAO 360° CHO WEB APP (XỬ LÝ IN-MEMORY)
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

  // Đọc Headers
  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));
  const idxMaKH = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
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

  // Đọc danh mục Khu vực
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

    const sdt = String(row[idxSDT] || '').replace(/[^0-9]/g, '');
    const serial = String(row[idxSerial] || '').replace(/[\s\uFEFF\xA0]+/g, '').toLowerCase();

    // Đối khớp từ khóa
    const matchMaKH = qDigits && maKH.includes(qDigits);
    const matchSDT = qDigits && sdt.includes(qDigits);
    const matchSerial = qClean && serial.includes(qClean);

    if (matchMaKH || matchSDT || matchSerial) {
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

      if (matched.length >= 25) break; // Giới hạn tối đa 25 kết quả trả về
    }
  }

  return matched;
}