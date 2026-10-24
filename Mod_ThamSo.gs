/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * MODULE       : QUẢN LÝ THAM SỐ CÀI ĐẶT & CHUẨN HÓA DANH MỤC
 * PHIÊN BẢN    : 1.0 (Tự động khởi tạo 4 Bảng tham số, quét Master Data thực tế)
 * NGÀY CẬP NHẬT: 20/09/2026 23:00:00 (GMT+7)
 * TÁC GIẢ      : Trịnh Vũ Thế Vinh & AI Assistant
 * MÔ TẢ        : 
 *   - Bảng 1 (A-B): Chuẩn hóa Loại khách hàng
 *   - Bảng 2 (D-E): Chuẩn hóa Mô hình triển khai
 *   - Bảng 3 (G-H): Chuẩn hóa Tên cửa hàng
 *   - Bảng 4 (J-M): Phân bổ Khu vực, Mã CH, Mã NV, Tên NV phục vụ thống kê Cụm
 * ============================================================================
 */

const THAMSO_CONFIG = {
  SHEET_NAME: 'thamso_caidat'
};

/**
 * Khởi tạo hoặc cập nhật các giá trị mới từ raw_mobifiber vào sheet thamso_caidat
 */
function initThamSoCaiDat() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  const ui = SpreadsheetApp.getUi();

  if (!rawSheet) {
    ui.alert(`❌ Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);
    return;
  }

  let tsSheet = ss.getSheetByName(THAMSO_CONFIG.SHEET_NAME);
  const isNew = !tsSheet;

  if (isNew) {
    tsSheet = ss.insertSheet(THAMSO_CONFIG.SHEET_NAME);
  }

  // 1. Quét dữ liệu thực tế từ raw_mobifiber
  const lastRawRow = rawSheet.getLastRow();
  const lastRawCol = rawSheet.getLastColumn();
  if (lastRawRow < CONFIG.DATA_START_ROW_TARGET) {
    ui.alert('ℹ️ Sheet raw_mobifiber chưa có dữ liệu để quét tham số.');
    return;
  }

  const rawHeaders = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastRawCol)
    .getValues()[0].map(h => normalizeHeader(h));

  const idxLoaiKH   = rawHeaders.findIndex(h => h === 'loại khách hàng');
  const idxMoHinh   = rawHeaders.findIndex(h => h === 'loại mô hình triển khai');
  const idxTenCH    = rawHeaders.findIndex(h => h === 'tên cửa hàng');
  const idxMaCH     = rawHeaders.findIndex(h => h === 'mã cửa hàng');
  const idxMaNV     = rawHeaders.findIndex(h => h === 'mã nhân viên');
  const idxTenNV    = rawHeaders.findIndex(h => h === 'tên nhân viên');

  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, lastRawRow - CONFIG.DATA_START_ROW_TARGET + 1, lastRawCol).getValues();

  // Tập hợp các giá trị duy nhất
  const uniqueCustTypes = new Set();
  const uniqueModels = new Set();
  const uniqueStores = new Set();
  const staffStoreMap = new Map(); // Key: "MaCH|MaNV" -> { maCH, maNV, tenNV }

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    if (idxLoaiKH !== -1 && row[idxLoaiKH]) uniqueCustTypes.add(String(row[idxLoaiKH]).trim());
    if (idxMoHinh !== -1 && row[idxMoHinh]) uniqueModels.add(String(row[idxMoHinh]).trim());
    if (idxTenCH !== -1 && row[idxTenCH]) uniqueStores.add(String(row[idxTenCH]).trim());

    if (idxMaNV !== -1 && row[idxMaNV]) {
      const maNV = String(row[idxMaNV]).trim();
      const maCH = idxMaCH !== -1 ? String(row[idxMaCH] || '').trim() : '';
      const tenNV = idxTenNV !== -1 ? String(row[idxTenNV] || '').trim() : '';
      const key = `${maCH}|${maNV}`;
      if (!staffStoreMap.has(key)) {
        staffStoreMap.set(key, { maCH, maNV, tenNV });
      }
    }
  }

  // 2. Thiết lập cấu trúc giao diện cho sheet thamso_caidat (nếu mới tạo)
  if (isNew) {
    // Dòng 1: Tiêu đề tổng quát
    tsSheet.getRange('A1:M1').merge()
      .setValue('HỆ THỐNG THAM SỐ CÀI ĐẶT & CHUẨN HÓA DỮ LIỆU MOBIFIBER')
      .setFontWeight('bold').setFontSize(13).setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');

    // Dòng 2: Tiêu đề 4 khối Bảng
    tsSheet.getRange('A2:B2').merge().setValue('BẢNG 1: LOẠI KHÁCH HÀNG').setFontWeight('bold').setBackground('#D2E3FC').setFontColor('#174EA6').setHorizontalAlignment('center');
    tsSheet.getRange('D2:E2').merge().setValue('BẢNG 2: MÔ HÌNH TRIỂN KHAI').setFontWeight('bold').setBackground('#D2E3FC').setFontColor('#174EA6').setHorizontalAlignment('center');
    tsSheet.getRange('G2:H2').merge().setValue('BẢNG 3: TÊN CỬA HÀNG').setFontWeight('bold').setBackground('#D2E3FC').setFontColor('#174EA6').setHorizontalAlignment('center');
    tsSheet.getRange('J2:M2').merge().setValue('BẢNG 4: PHÂN BỔ KHU VỰC & NHÂN SỰ').setFontWeight('bold').setBackground('#FCE8E6').setFontColor('#C5221F').setHorizontalAlignment('center');

    // Dòng 3: Tiêu đề chi tiết các cột
    const colHeaders = [
      'LOẠI KHÁCH HÀNG (GỐC)', 'TÊN RÚT GỌN', '', 
      'MÔ HÌNH TRIỂN KHAI (GỐC)', 'TÊN RÚT GỌN', '', 
      'TÊN CỬA HÀNG (GỐC)', 'TÊN RÚT GỌN', '', 
      'KHU VỰC', 'MÃ CỬA HÀNG', 'MÃ NHÂN VIÊN', 'TÊN NHÂN VIÊN'
    ];
    tsSheet.getRange(3, 1, 1, 13).setValues([colHeaders]).setFontWeight('bold').setFontColor('#FFFFFF').setHorizontalAlignment('center');

    tsSheet.getRange('A3:B3').setBackground('#1B365D');
    tsSheet.getRange('D3:E3').setBackground('#00558F');
    tsSheet.getRange('G3:H3').setBackground('#007A78');
    tsSheet.getRange('J3:M3').setBackground('#8A1C14');
    tsSheet.setFrozenRows(3);

    // Điền dữ liệu gợi ý Bảng 1: Loại khách hàng
    const arrCust = Array.from(uniqueCustTypes).sort().map(item => [item, suggestShortCustType(item)]);
    if (arrCust.length > 0) tsSheet.getRange(4, 1, arrCust.length, 2).setValues(arrCust);

    // Điền dữ liệu gợi ý Bảng 2: Mô hình triển khai
    const arrModel = Array.from(uniqueModels).sort().map(item => [item, suggestShortModel(item)]);
    if (arrModel.length > 0) tsSheet.getRange(4, 4, arrModel.length, 2).setValues(arrModel);

    // Điền dữ liệu gợi ý Bảng 3: Cửa hàng
    const arrStore = Array.from(uniqueStores).sort().map(item => [item, suggestShortStore(item)]);
    if (arrStore.length > 0) tsSheet.getRange(4, 7, arrStore.length, 2).setValues(arrStore);

    // Điền dữ liệu Bảng 4: Khu vực & Nhân viên
    const arrStaff = Array.from(staffStoreMap.values()).sort((a,b) => a.maCH.localeCompare(b.maCH)).map(item => [
      '', // Để trống cột Khu vực để người dùng tự nhập (KV1, KV2, KV3, KV4, KV5)
      "'" + item.maCH,
      "'" + item.maNV,
      '  ' + item.tenNV
    ]);
    if (arrStaff.length > 0) {
      tsSheet.getRange(4, 10, arrStaff.length, 4).setValues(arrStaff);
      tsSheet.getRange(4, 10, arrStaff.length, 1).setHorizontalAlignment('center'); // Khu vực
      tsSheet.getRange(4, 11, arrStaff.length, 2).setNumberFormat('@').setHorizontalAlignment('center'); // Mã CH, Mã NV
      tsSheet.getRange(4, 13, arrStaff.length, 1).setHorizontalAlignment('left'); // Tên NV
    }

    // Độ rộng các cột ngăn cách
    tsSheet.setColumnWidth(3, 25);
    tsSheet.setColumnWidth(6, 25);
    tsSheet.setColumnWidth(9, 25);

    // Độ rộng các cột dữ liệu
    tsSheet.setColumnWidth(1, 260); tsSheet.setColumnWidth(2, 160); // Bảng 1
    tsSheet.setColumnWidth(4, 260); tsSheet.setColumnWidth(5, 160); // Bảng 2
    tsSheet.setColumnWidth(7, 280); tsSheet.setColumnWidth(8, 160); // Bảng 3
    tsSheet.setColumnWidth(10, 110); tsSheet.setColumnWidth(11, 120); tsSheet.setColumnWidth(12, 120); tsSheet.setColumnWidth(13, 160); // Bảng 4

    ss.setActiveSheet(tsSheet);
    ui.alert(
      '🎉 KHỞI TẠO THAM SỐ CÀI ĐẶT HOÀN TẤT!\n\n' +
      'Đã thiết lập đầy đủ 4 Bảng tham số tại sheet "thamso_caidat":\n' +
      '  1. Loại khách hàng (A-B)\n' +
      '  2. Mô hình triển khai (D-E)\n' +
      '  3. Tên cửa hàng (G-H)\n' +
      '  4. Phân bổ Khu vực & Nhân viên (J-M)\n\n' +
      '👉 Bạn có thể tự do chỉnh sửa cột "TÊN RÚT GỌN" và điền tên 5 Khu vực vào cột J theo đúng quy chuẩn quản lý.'
    );
  } else {
    // Nếu sheet đã có sẵn: chỉ quét bổ sung các giá trị mới chưa có trong danh mục
    syncNewValuesToThamSo(tsSheet, uniqueCustTypes, uniqueModels, uniqueStores, staffStoreMap);
    ss.setActiveSheet(tsSheet);
    ui.alert('✅ Đã đồng bộ kiểm tra và cập nhật các tham số mới (nếu có) vào sheet "thamso_caidat"!');
  }
}

/**
 * Tự động đồng bộ các giá trị mới phát sinh vào bảng tham số
 */
function syncNewValuesToThamSo(tsSheet, uniqueCustTypes, uniqueModels, uniqueStores, staffStoreMap) {
  const lastRow = Math.max(tsSheet.getLastRow(), 4);

  // 1. Quét Bảng 1
  const existingCust = new Set(tsSheet.getRange(4, 1, lastRow - 3, 1).getValues().map(r => String(r[0]).trim()).filter(v => v));
  const newCust = Array.from(uniqueCustTypes).filter(v => !existingCust.has(v)).map(v => [v, suggestShortCustType(v)]);
  if (newCust.length > 0) {
    const nextR = getNextEmptyRowInCol(tsSheet, 1);
    tsSheet.getRange(nextR, 1, newCust.length, 2).setValues(newCust);
  }

  // 2. Quét Bảng 2
  const existingModel = new Set(tsSheet.getRange(4, 4, lastRow - 3, 1).getValues().map(r => String(r[0]).trim()).filter(v => v));
  const newModel = Array.from(uniqueModels).filter(v => !existingModel.has(v)).map(v => [v, suggestShortModel(v)]);
  if (newModel.length > 0) {
    const nextR = getNextEmptyRowInCol(tsSheet, 4);
    tsSheet.getRange(nextR, 4, newModel.length, 2).setValues(newModel);
  }

  // 3. Quét Bảng 3
  const existingStore = new Set(tsSheet.getRange(4, 7, lastRow - 3, 1).getValues().map(r => String(r[0]).trim()).filter(v => v));
  const newStore = Array.from(uniqueStores).filter(v => !existingStore.has(v)).map(v => [v, suggestShortStore(v)]);
  if (newStore.length > 0) {
    const nextR = getNextEmptyRowInCol(tsSheet, 7);
    tsSheet.getRange(nextR, 7, newStore.length, 2).setValues(newStore);
  }

  // 4. Quét Bảng 4
  const existingStaff = new Set(tsSheet.getRange(4, 12, lastRow - 3, 1).getValues().map(r => String(r[0]).trim()).filter(v => v));
  const newStaff = Array.from(staffStoreMap.values()).filter(v => !existingStaff.has(v.maNV)).map(item => [
    '', "'" + item.maCH, "'" + item.maNV, '  ' + item.tenNV
  ]);
  if (newStaff.length > 0) {
    const nextR = getNextEmptyRowInCol(tsSheet, 11);
    tsSheet.getRange(nextR, 10, newStaff.length, 4).setValues(newStaff);
  }
}

function getNextEmptyRowInCol(sheet, col) {
  const vals = sheet.getRange(4, col, Math.max(sheet.getLastRow() - 3, 1), 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] !== '') return 4 + i + 1;
  }
  return 4;
}

// ==========================================
// CÁC HÀM GỢI Ý TÊN RÚT GỌN TỰ ĐỘNG BAN ĐẦU
// ==========================================
function suggestShortCustType(raw) {
  const n = normalizeHeader(raw);
  if (n.includes('hành chính') || n.includes('vũ trang') || n.includes('bca')) return 'HCSN & LLVT';
  if (n.includes('doanh nghiệp') || n.includes('tnhh')) return 'Doanh nghiệp';
  if (n.includes('vnr500')) return 'Khách hàng VNR500';
  if (n.includes('tư nhân')) return 'Cá nhân';
  if (n.includes('nghiệp vụ')) return 'Thuê bao Nghiệp vụ';
  return raw;
}

function suggestShortModel(raw) {
  const n = normalizeHeader(raw);
  if (n.includes('chi nhánh') || n.includes('cửa hàng')) return 'Nội bộ Chi nhánh';
  if (n.includes('gpon')) return 'Trạm BTS GPON';
  if (n.includes('aon')) return 'Trạm BTS AON';
  if (n.includes('cooperation')) return 'Hợp tác toàn diện';
  if (n.includes('mnv') || n.includes('hđtv')) return 'Nghiệp vụ Lãnh đạo';
  if (n.includes('tư nhân')) return 'Hạ tầng cá nhân';
  return raw;
}

function suggestShortStore(raw) {
  const n = normalizeHeader(raw);
  if (n.includes('chi nhánh tổng công ty') || n.includes('lâm đồng')) return 'MobiFone Lâm Đồng';
  if (n.includes('tổ kd ldo') || n.includes('gia tăng')) return 'Tổ KD Lâm Đồng';
  if (n.includes('đak mil')) return 'CH Đăk Mil';
  if (n.includes('đak song')) return 'CH Đăk Song';
  if (n.includes('krông nô')) return 'CH Krông Nô';
  if (n.includes('cư jút')) return 'CH Cư Jút';
  if (n.includes('gia nghĩa')) return 'Tổ KD Gia Nghĩa';
  if (n.includes('phan thiết')) return 'CH Phan Thiết';
  if (n.includes('bảo lộc')) return 'CH Bảo Lộc';
  return raw;
}