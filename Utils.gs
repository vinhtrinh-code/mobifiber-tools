/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * FILE         : Utils.gs
 * PHIÊN BẢN    : 2.4 (Tự động đồng bộ Master Data phát sinh vào thamso_caidat)
 * NGÀY CẬP NHẬT: 20/09/2026 23:45:00 (GMT+7)
 * TÁC GIẢ      : Trịnh Vũ Thế Vinh & AI Assistant
 * ============================================================================
 */

function normalizeCustomerId(val) {
  if (val === null || val === undefined) return '';
  let str = String(val).trim();
  const stripped = str.replace(/^0+/, '');
  return stripped === '' ? '0' : stripped;
}

function normalizeHeader(str) {
  return String(str || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function parseCurrencyToInteger(rawVal) {
  if (rawVal === undefined || rawVal === null || rawVal === '') return '';
  if (typeof rawVal === 'number') return Math.round(rawVal);
  let str = String(rawVal).trim();
  if (!str) return '';

  if (str.includes('.') && str.includes(',')) {
    if (str.lastIndexOf(',') > str.lastIndexOf('.')) str = str.replace(/\./g, '').replace(',', '.');
    else str = str.replace(/,/g, '');
  } else if (str.includes(',')) {
    const parts = str.split(',');
    if (parts.length === 2 && parts[1].length <= 2) str = str.replace(',', '.');
    else str = str.replace(/,/g, '');
  } else if (str.includes('.')) {
    const parts = str.split('.');
    if (parts.length === 2 && parts[1].length <= 2) {}
    else str = str.replace(/\./g, '');
  }

  const cleanStr = str.replace(/[^0-9\.\-]/g, '');
  const num = parseFloat(cleanStr);
  return isNaN(num) ? '' : Math.round(num);
}

function parseToDate(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  if (typeof val === 'number') {
    const excelEpoch = new Date(1899, 11, 30);
    const parsed = new Date(excelEpoch.getTime() + (val * 86400000));
    if (!isNaN(parsed.getTime())) return parsed;
  }
  const str = String(val).trim();
  if (!str) return null;

  const isoMatch = str.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
  if (isoMatch) {

    return new Date(parseInt(isoMatch[1], 10), parseInt(isoMatch[2], 10) - 1, parseInt(isoMatch[3], 10), parseInt(isoMatch[4] || 0, 10), parseInt(isoMatch[5] || 0, 10), parseInt(isoMatch[6] || 0, 10));
  }
  const dmyMatch = str.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
  if (dmyMatch) {
    return new Date(parseInt(dmyMatch[3], 10), parseInt(dmyMatch[2], 10) - 1, parseInt(dmyMatch[1], 10), parseInt(dmyMatch[4] || 0, 10), parseInt(dmyMatch[5] || 0, 10), parseInt(dmyMatch[6] || 0, 10));
  }
  const timestamp = Date.parse(str);
  if (!isNaN(timestamp)) return new Date(timestamp);
  return null;
}

function formatDateTimeStr(val) {
  const d = parseToDate(val);
  if (d) return Utilities.formatDate(d, 'Asia/Saigon', 'dd/MM/yyyy HH:mm:ss');
  return String(val || '');
}

function formatValForLog(val) {
  if (val === undefined || val === null) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Saigon', 'dd/MM/yyyy HH:mm:ss');
  if (typeof val === 'number') return val.toLocaleString('vi-VN');
  return String(val ?? '').replace(/^'+/, '');
}

function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function deleteTempFile(fileId) {
  try {
    if (fileId) {
      const file = DriveApp.getFileById(fileId);
      file.setTrashed(true);
    }
  } catch (e) {}
}

function cleanupOldTempFiles(fileName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const currentSsId = ss.getId();
    const files = DriveApp.getFilesByName(fileName);
    while (files.hasNext()) {
      const f = files.next();
      if (f.getId() !== currentSsId) {
        f.setTrashed(true);
      }
    }
  } catch (e) {}
}

function cleanAllOrphanedExportFiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentSsId = ss.getId();
  let count = 0;
  
  const query = "mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false";
  const files = DriveApp.searchFiles(query);
  
  while (files.hasNext()) {
    const f = files.next();
    const name = f.getName();
    if (f.getId() !== currentSsId && 
       (name.startsWith('PTM_') || name.startsWith('DacThu_') || name.startsWith('giahan_') || name.startsWith('Temp_') || name.startsWith('Loc_'))) {
      f.setTrashed(true);
      count++;
    }
  }
  
  SpreadsheetApp.getUi().alert(`✅ Đã dọn dẹp sạch sẽ ${count} file tạm tồn đọng trên Google Drive của bạn!`);
}

/**
 * ĐỌC TOÀN BỘ BẢNG THAM SỐ VÀO BỘ ĐỆM ĐỂ TRA CỨU TỨC THÌ CHO CÁC MODULE BÁO CÁO
 */
function loadThamSoMapping() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tsSheet = ss.getSheetByName('thamso_caidat');
  
  const mapping = {
    custTypes: new Map(),
    models: new Map(),
    stores: new Map(),
    staffRegion: new Map()
  };

  if (!tsSheet) return mapping;

  const lastRow = tsSheet.getLastRow();
  if (lastRow < 4) return mapping;

  const data = tsSheet.getRange(4, 1, lastRow - 3, 13).getValues();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    // Bảng 1: Loại KH
    const cGoc = String(row[0] || '').trim();
    const cShort = String(row[1] || '').trim();
    if (cGoc && cShort) mapping.custTypes.set(normalizeHeader(cGoc), cShort);

    // Bảng 2: Mô hình
    const mGoc = String(row[3] || '').trim();
    const mShort = String(row[4] || '').trim();
    if (mGoc && mShort) mapping.models.set(normalizeHeader(mGoc), mShort);

    // Bảng 3: Cửa hàng
    const sGoc = String(row[6] || '').trim();
    const sShort = String(row[7] || '').trim();
    if (sGoc && sShort) mapping.stores.set(normalizeHeader(sGoc), sShort);

    // Bảng 4: Khu vực & Nhân viên
    const khuVuc = String(row[9] || '').trim();
    const maCH = String(row[10] || '').replace(/^'+/, '').trim();
    const maNV = String(row[11] || '').replace(/^'+/, '').trim();
    const tenNV = String(row[12] || '').trim();
    if (maNV) {
      mapping.staffRegion.set(normalizeHeader(maNV), { khuVuc, maCH, tenNV });
    }
  }

  return mapping;
}

function getShortValue(map, originalVal) {
  if (!originalVal) return '';
  const key = normalizeHeader(originalVal);
  if (map && map.has(key)) {
    return map.get(key);
  }
  return String(originalVal).trim();
}

/**
 * HÀM TỰ ĐỘNG BẮT PHÁT SINH THAM SỐ MỚI TỪ DỮ LIỆU ĐẦU VÀO VÀ ĐỒNG BỘ VÀO thamso_caidat
 * @param {Array<Array>} newRowsToScan - Mảng các dòng dữ liệu mới hoặc cập nhật
 * @param {Object} colIndices - Vị trí các cột trên raw_mobifiber
 */
function autoDetectAndSyncNewParameters(newRowsToScan, colIndices) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tsSheet = ss.getSheetByName('thamso_caidat');
  if (!tsSheet || !newRowsToScan || newRowsToScan.length === 0) return { newParamsCount: 0, details: '' };

  const lastRow = Math.max(tsSheet.getLastRow(), 4);
  const tsData = tsSheet.getRange(4, 1, lastRow - 3, 13).getValues();

  // Đọc danh mục hiện có trên thamso_caidat
  const existingCustTypes = new Set();
  const existingModels    = new Set();
  const existingStores    = new Set();
  const existingStaff     = new Set();

  for (let i = 0; i < tsData.length; i++) {
    const r = tsData[i];
    if (r[0]) existingCustTypes.add(normalizeHeader(r[0]));
    if (r[3]) existingModels.add(normalizeHeader(r[3]));
    if (r[6]) existingStores.add(normalizeHeader(r[6]));
    if (r[11]) existingStaff.add(normalizeHeader(r[11])); // Tra theo Mã NV
  }

  // Quét các dòng mới để tìm phát sinh
  const toAddCust = new Map();
  const toAddModel = new Map();
  const toAddStore = new Map();
  const toAddStaff = new Map();

  for (let i = 0; i < newRowsToScan.length; i++) {
    const row = newRowsToScan[i];

    // 1. Kiểm tra Loại KH
    if (colIndices.idxLoaiKH !== -1 && row[colIndices.idxLoaiKH]) {
      const val = String(row[colIndices.idxLoaiKH]).trim();
      const norm = normalizeHeader(val);
      if (val && !existingCustTypes.has(norm) && !toAddCust.has(norm)) {
        toAddCust.set(norm, [val, suggestShortCustType(val)]);
      }
    }

    // 2. Kiểm tra Mô hình
    if (colIndices.idxMoHinh !== -1 && row[colIndices.idxMoHinh]) {
      const val = String(row[colIndices.idxMoHinh]).trim();
      const norm = normalizeHeader(val);
      if (val && !existingModels.has(norm) && !toAddModel.has(norm)) {
        toAddModel.set(norm, [val, suggestShortModel(val)]);
      }
    }

    // 3. Kiểm tra Cửa hàng
    if (colIndices.idxTenCH !== -1 && row[colIndices.idxTenCH]) {
      const val = String(row[colIndices.idxTenCH]).trim();
      const norm = normalizeHeader(val);
      if (val && !existingStores.has(norm) && !toAddStore.has(norm)) {
        toAddStore.set(norm, [val, suggestShortStore(val)]);
      }
    }

    // 4. Kiểm tra Nhân sự (Mã NV)
    if (colIndices.idxMaNV !== -1 && row[colIndices.idxMaNV]) {
      const maNV = String(row[colIndices.idxMaNV]).trim();
      const norm = normalizeHeader(maNV);
      if (maNV && !existingStaff.has(norm) && !toAddStaff.has(norm)) {
        const maCH = colIndices.idxMaCH !== -1 ? String(row[colIndices.idxMaCH] || '').trim() : '';
        const tenNV = colIndices.idxTenNV !== -1 ? String(row[colIndices.idxTenNV] || '').trim() : '';
        toAddStaff.set(norm, {
          khuVuc: '[Chưa phân bổ KV]',
          maCH: maCH,
          maNV: maNV,
          tenNV: tenNV
        });
      }
    }
  }

  let totalNew = 0;
  const detailList = [];

  // Ghi bổ sung Bảng 1: Loại KH
  if (toAddCust.size > 0) {
    const rows = Array.from(toAddCust.values());
    const nextRow = getNextEmptyRowInCol(tsSheet, 1);
    tsSheet.getRange(nextRow, 1, rows.length, 2).setValues(rows);
    totalNew += toAddCust.size;
    detailList.push(`${toAddCust.size} Loại KH mới`);
  }

  // Ghi bổ sung Bảng 2: Mô hình
  if (toAddModel.size > 0) {
    const rows = Array.from(toAddModel.values());
    const nextRow = getNextEmptyRowInCol(tsSheet, 4);
    tsSheet.getRange(nextRow, 4, rows.length, 2).setValues(rows);
    totalNew += toAddModel.size;
    detailList.push(`${toAddModel.size} Mô hình mới`);
  }

  // Ghi bổ sung Bảng 3: Cửa hàng
  if (toAddStore.size > 0) {
    const rows = Array.from(toAddStore.values());
    const nextRow = getNextEmptyRowInCol(tsSheet, 7);
    tsSheet.getRange(nextRow, 7, rows.length, 2).setValues(rows);
    totalNew += toAddStore.size;
    detailList.push(`${toAddStore.size} Cửa hàng mới`);
  }

  // Ghi bổ sung Bảng 4: Nhân sự mới
  if (toAddStaff.size > 0) {
    const staffArr = Array.from(toAddStaff.values());
    const rows = staffArr.map(s => [s.khuVuc, "'" + s.maCH, "'" + s.maNV, '  ' + s.tenNV]);
    const nextRow = getNextEmptyRowInCol(tsSheet, 11);
    const range = tsSheet.getRange(nextRow, 10, rows.length, 4);
    range.setValues(rows);

    // Định dạng nổi bật cảnh báo màu cam cho các nhân sự chưa phân bổ khu vực
    tsSheet.getRange(nextRow, 10, rows.length, 1)
      .setBackground('#FEF7E0')
      .setFontColor('#B06000')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    tsSheet.getRange(nextRow, 11, rows.length, 2).setNumberFormat('@').setHorizontalAlignment('center');
    tsSheet.getRange(nextRow, 13, rows.length, 1).setHorizontalAlignment('left');

    totalNew += toAddStaff.size;
    detailList.push(`${toAddStaff.size} Nhân sự mới [Chưa phân bổ KV]`);
  }

  return {
    newParamsCount: totalNew,
    details: detailList.join(', ')
  };
}