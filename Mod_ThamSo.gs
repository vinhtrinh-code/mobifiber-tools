/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * FILE        : Mod_ThamSo.gs
 * PHIÊN BẢN   : 3.0 (Master Data Engine: Fast Init, Matrix Layout & O(1) Cache)
 * NGÀY TẠO    : 18/09/2026 14:00:00 (GMT+7)
 * NGÀY CẬP NHẬT: 24/10/2026 20:10:00 (GMT+7)
 * TÁC GIẢ     : Trịnh Vũ Thế Vinh & AI Assistant
 * MÔ TẢ       : Quản trị, khởi tạo và đồng bộ bảng danh mục tham số chuẩn hóa (thamso_caidat)
 * ============================================================================
 */

const THAMSO_SHEET_NAME = 'thamso_caidat';

/**
 * KHỞI TẠO HOẶC CẤU TRÚC LẠI SHEET THAM SỐ CHUẨN HÓA SIÊU TỐC
 */
function initThamSoCaiDat() {
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let tsSheet = ss.getSheetByName(THAMSO_SHEET_NAME);

  if (!tsSheet) {
    tsSheet = ss.insertSheet(THAMSO_SHEET_NAME);
  }

  // 1. THIẾT LẬP KHỐI BANNER CHÍNH
  const bannerMatrix = [
    ['BẢNG THAM SỐ CHUẨN HÓA HỆ THỐNG MOBIFIBER', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['Cập nhật tự động khi phát sinh tham số mới từ file nguồn. Phân bổ khu vực theo Mã nhân viên.', '', '', '', '', '', '', '', '', '', '', '', '']
  ];
  tsSheet.getRange(1, 1, 2, 13).setValues(bannerMatrix);
  tsSheet.getRange('A1:M1').merge().setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  tsSheet.getRange('A2:M2').merge().setBackground('#e8f0fe').setFontColor('#174ea6').setFontStyle('italic').setHorizontalAlignment('center');

  // 2. KHỞI TẠO TIÊU ĐỀ 4 BẢNG DANH MỤC (HÀNG 3)
  const headerRow3 = [
    'LOẠI KHÁCH HÀNG', 'MÃ PHÂN LOẠI', '', 
    'MÔ HÌNH TRIỂN KHAI', 'MÃ MÔ HÌNH', '', 
    'MÃ CỬA HÀNG', 'TÊN CỬA HÀNG', '', 
    'KHU VỰC PHÂN BỔ', 'MÃ CỬA HÀNG', 'MÃ NHÂN VIÊN', 'TÊN NHÂN VIÊN'
  ];
  tsSheet.getRange(3, 1, 1, 13).setValues([headerRow3]);

  // Định dạng tiêu đề từng bảng màu sắc trực quan
  tsSheet.getRange('A3:B3').setBackground('#34a853').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  tsSheet.getRange('D3:E3').setBackground('#fbbc04').setFontColor('#202124').setFontWeight('bold').setHorizontalAlignment('center');
  tsSheet.getRange('G3:H3').setBackground('#ea4335').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  tsSheet.getRange('J3:M3').setBackground('#00558f').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  // Cố định hàng tiêu đề
  tsSheet.setFrozenRows(3);

  // 3. TỰ ĐỘNG QUÉT VÀ NẠP DANH MỤC TỪ RAW_MOBIFIBER NẾU CHƯA CÓ DỮ LIỆU
  const lastRow = tsSheet.getLastRow();
  let autoSyncCount = 0;

  if (lastRow <= 3) {
    const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
    if (rawSheet && rawSheet.getLastRow() >= CONFIG.DATA_START_ROW_TARGET) {
      const rawLastRow = rawSheet.getLastRow();
      const rawLastCol = rawSheet.getLastColumn();
      const rHeaders = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, rawLastCol).getValues()[0].map(h => normalizeHeader(h));
      
      const idxLoaiKH = rHeaders.findIndex(h => h === 'loại khách hàng');
      const idxMoHinh = rHeaders.findIndex(h => h === 'loại mô hình triển khai');
      const idxMaCH = rHeaders.findIndex(h => h === 'mã cửa hàng');
      const idxTenCH = rHeaders.findIndex(h => h === 'tên cửa hàng');
      const idxMaNV = rHeaders.findIndex(h => h === 'mã nhân viên');
      const idxTenNV = rHeaders.findIndex(h => h === 'tên nhân viên');

      const rawRows = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, rawLastRow - CONFIG.DATA_START_ROW_TARGET + 1, rawLastCol).getValues();

      const loaiKhSet = new Set();
      const moHinhSet = new Set();
      const storeMap = new Map();
      const staffMap = new Map();

      rawRows.forEach(r => {
        if (idxLoaiKH !== -1 && r[idxLoaiKH]) loaiKhSet.add(String(r[idxLoaiKH]).trim());
        if (idxMoHinh !== -1 && r[idxMoHinh]) moHinhSet.add(String(r[idxMoHinh]).trim());
        if (idxMaCH !== -1 && idxTenCH !== -1 && r[idxMaCH]) {
          storeMap.set(String(r[idxMaCH]).trim(), String(r[idxTenCH] || '').trim());
        }
        if (idxMaNV !== -1 && r[idxMaNV]) {
          const mNv = String(r[idxMaNV]).trim();
          const tNv = idxTenNV !== -1 ? String(r[idxTenNV] || '').trim() : '';
          const mCh = idxMaCH !== -1 ? String(r[idxMaCH] || '').trim() : '';
          staffMap.set(mNv, { tenNV: tNv, maCH: mCh });
        }
      });

      // Ghi Loại KH ban đầu
      const arrLoaiKH = Array.from(loaiKhSet).map(k => [k, '']);
      if (arrLoaiKH.length > 0) tsSheet.getRange(4, 1, arrLoaiKH.length, 2).setValues(arrLoaiKH);

      // Ghi Mô hình ban đầu
      const arrMoHinh = Array.from(moHinhSet).map(m => [m, '']);
      if (arrMoHinh.length > 0) tsSheet.getRange(4, 4, arrMoHinh.length, 2).setValues(arrMoHinh);

      // Ghi Cửa hàng ban đầu
      const arrStore = Array.from(storeMap.entries()).map(([k, v]) => [k, v]);
      if (arrStore.length > 0) tsSheet.getRange(4, 7, arrStore.length, 2).setValues(arrStore);

      // Ghi Nhân sự & Địa bàn ban đầu (Mặc định gắn Chưa phân bổ)
      const arrStaff = Array.from(staffMap.entries()).map(([mNv, info]) => ['Chưa phân bổ KV', info.maCH, mNv, info.tenNV]);
      if (arrStaff.length > 0) tsSheet.getRange(4, 10, arrStaff.length, 4).setValues(arrStaff);

      autoSyncCount = arrStaff.length;
    }
  }

  // 4. ĐỊNH DẠNG ĐỘ RỘNG CÁC CỘT HỢP LÝ
  tsSheet.setColumnWidth(1, 160);
  tsSheet.setColumnWidth(2, 110);
  tsSheet.setColumnWidth(3, 25);  // Cột đệm
  tsSheet.setColumnWidth(4, 160);
  tsSheet.setColumnWidth(5, 110);
  tsSheet.setColumnWidth(6, 25);  // Cột đệm
  tsSheet.setColumnWidth(7, 100);
  tsSheet.setColumnWidth(8, 160);
  tsSheet.setColumnWidth(9, 25);  // Cột đệm
  tsSheet.setColumnWidth(10, 140);
  tsSheet.setColumnWidth(11, 100);
  tsSheet.setColumnWidth(12, 110);
  tsSheet.setColumnWidth(13, 160);

  const duration = ((new Date().getTime() - startTime) / 1000).toFixed(2);
  Logger.log(`⚡ Khởi tạo cấu trúc tham số chuẩn hóa hoàn tất trong ${duration}s!`);
  ss.setActiveSheet(tsSheet);
  ui.alert(`✅ Đã khởi tạo hoàn tất Bảng tham số chuẩn hóa (Thời gian: ${duration}s)! Hãy bổ sung tên Khu vực ở Cột J.`);
}

/**
 * HÀM DÙNG CHUNG: ĐỌC NHANH BẢN ĐỒ THAM SỐ VÀO RAM CHO CÁC MODULE KHÁC
 */
function getThamSoMappingCached() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tsSheet = ss.getSheetByName(THAMSO_SHEET_NAME);
  const result = {
    staffRegion: new Map(), // maNV -> { khuVuc, maCH, tenNV }
    stores: new Map()       // maCH -> tenCH
  };

  if (!tsSheet || tsSheet.getLastRow() < 4) return result;

  const lastRow = tsSheet.getLastRow();

  // Đọc danh mục Cửa hàng (Cột G:H)
  const storeData = tsSheet.getRange(4, 7, lastRow - 3, 2).getValues();
  storeData.forEach(r => {
    const maCH = String(r[0] || '').trim();
    const tenCH = String(r[1] || '').trim();
    if (maCH) result.stores.set(normalizeHeader(maCH), tenCH);
  });

  // Đọc phân bổ Nhân sự & Địa bàn (Cột J:M)
  const staffData = tsSheet.getRange(4, 10, lastRow - 3, 4).getValues();
  staffData.forEach(r => {
    const khuVuc = String(r[0] || '').trim();
    const maCH = String(r[1] || '').trim();
    const maNV = normalizeHeader(String(r[2] || ''));
    const tenNV = String(r[3] || '').trim();

    if (maNV) {
      result.staffRegion.set(maNV, {
        khuVuc: khuVuc || 'Chưa phân bổ KV',
        maCH: maCH,
        tenNV: tenNV
      });
    }
  });

  return result;
}