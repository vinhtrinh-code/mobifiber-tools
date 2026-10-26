/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * MODULE       : Mod_CapNhat.gs
 * PHIÊN BẢN    : 2.8 (Quét trực tiếp từ Source Header, Chèn đúng dòng Bảng 4 thamso_caidat)
 * TÁC GIẢ      : Trịnh Vũ Thế Vinh & AI Assistant
 * ============================================================================
 */

function showUploadDialog() {
  const html = HtmlService.createHtmlOutputFromFile('UploadDialog')
    .setWidth(590)
    .setHeight(540)
    .setTitle('Cập nhật dữ liệu MobiFiber');
  SpreadsheetApp.getUi().showModalDialog(html, 'Cập nhật dữ liệu MobiFiber');
}

function suggestShortModel(modelName) {
  if (!modelName) return '';
  const str = String(modelName).trim();
  const lower = str.toLowerCase();

  if (lower.includes('bca')) return 'BCA';
  if (lower.includes('aon-khdn') || lower.includes('aon - khdn')) return 'AON-KHDN';
  if (lower.includes('aon')) return 'AON';
  if (lower.includes('gpon-khdn') || lower.includes('gpon - khdn')) return 'GPON-KHDN';
  if (lower.includes('gpon')) return 'GPON';
  if (lower.includes('full') || lower.includes('cooperation')) return 'FFC';
  if (lower.includes('segment')) return 'FSC';
  if (lower.includes('mca')) return 'MCA';
  if (lower.includes('nghiệp vụ') || lower.includes('nghiep vu')) return 'Nghiệp vụ';
  if (lower.includes('cá nhân') || lower.includes('tư nhân') || lower.includes('vie')) return 'VIE';
  if (lower.includes('dn-cty') || lower.includes('doanh nghiệp')) return 'DN-CTy';

  if (str.length > 15) {
    return str.split(/[\s_-]+/).map(w => w.charAt(0).toUpperCase()).join('');
  }
  return str;
}

function isValidSubscriberRow(rawCode, rowArray) {
  if (rawCode === undefined || rawCode === null) return false;
  let cleanCode = String(rawCode).replace(/[\s\uFEFF\xA0]+/g, '').replace(/^'+/, '').trim();
  if (!cleanCode) return false;
  const lower = cleanCode.toLowerCase();
  if (lower.includes('tổng') || lower.includes('cong') || lower.includes('stt') || lower.includes('mã')) {
    return false;
  }
  const digitsOnly = cleanCode.replace(/[^0-9]/g, '');
  if (digitsOnly.length < 8) return false;

  if (rowArray && Array.isArray(rowArray)) {
    let sequentialCount = 0;
    for (let c = 0; c < Math.min(rowArray.length, 10); c++) {
      const val = parseInt(String(rowArray[c] || '').trim(), 10);
      if (!isNaN(val) && val >= 1 && val <= 15) sequentialCount++;
    }
    if (sequentialCount >= 5) return false;
  }
  return true;
}

function formatValueByFieldType(fieldName, rawVal) {
  if (rawVal === undefined || rawVal === null || rawVal === '') return '';
  const normField = normalizeHeader(fieldName);
  let str = String(rawVal).replace(/[\s\uFEFF\xA0]+/g, ' ').trim();

  if (normField === normalizeHeader(CONFIG.KEY_COLUMN_NAME)) {
    str = str.replace(/[^0-9]/g, '');
    if (str.length === 8) str = '00' + str;
    else if (str.length < 10) str = str.padStart(10, '0');
    return "'" + str;
  }
  if (normField === 'số cmnd/cccd') {
    str = str.replace(/[^0-9]/g, '');
    if (str.length === 11) str = '0' + str;
    return "'" + str;
  }
  if (normField === 'sđt liên hệ') {
    str = str.replace(/[^0-9]/g, '');
    if (str.length === 9) str = '0' + str;
    return "'" + str;
  }
  if (CONFIG.CURRENCY_INTEGER_FIELDS.some(f => normalizeHeader(f) === normField)) {
    return parseCurrencyToInteger(rawVal);
  }
  if (normField === 'chu kỳ sử dụng') {
    const num = parseInt(str, 10);
    return isNaN(num) ? str : num;
  }
  if (CONFIG.DATE_ONLY_FIELDS.some(f => normalizeHeader(f) === normField)) {
    const d = parseToDate(rawVal);
    if (d) return Utilities.formatDate(d, 'Asia/Saigon', 'dd/MM/yyyy');
    return str;
  }
  if (CONFIG.DATETIME_FIELDS.some(f => normalizeHeader(f) === normField)) {
    const d = parseToDate(rawVal);
    if (d) return Utilities.formatDate(d, 'Asia/Saigon', 'dd/MM/yyyy HH:mm:ss');
    return str;
  }
  return rawVal;
}

function isValueChanged(fieldName, val1, val2) {
  const normField = normalizeHeader(fieldName);
  const isDateCol = CONFIG.DATE_ONLY_FIELDS.some(f => normalizeHeader(f) === normField) ||
                    CONFIG.DATETIME_FIELDS.some(f => normalizeHeader(f) === normField);
  if (isDateCol) {
    const d1 = parseToDate(val1);
    const d2 = parseToDate(val2);
    if (!d1 && !d2) return false;
    if (!d1 || !d2) return true;
    return Math.abs(d1.getTime() - d2.getTime()) >= 1000;
  }
  if (typeof val1 === 'number' && typeof val2 === 'number') {
    return Math.abs(val1 - val2) > 0.0001;
  }
  return String(val1 ?? '').replace(/^'+/, '').trim() !== String(val2 ?? '').replace(/^'+/, '').trim();
}

/**
 * XỬ LÝ TRUNG TÂM CẬP NHẬT DỮ LIỆU & QUÉT THAM SỐ TOÀN DIỆN
 */
function processDataUpdate(sourceRawData, headerRowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!targetSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);

  const sourceHeaders = sourceRawData[headerRowIndex].map(h => String(h || '').trim());
  const rawDataRows = sourceRawData.slice(headerRowIndex + 1);

  // Chỉ số cột trên file Excel nguồn (Source Indices)
  const srcKeyIdx = sourceHeaders.findIndex(h => normalizeHeader(h) === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  if (srcKeyIdx === -1) throw new Error(`File nguồn không chứa cột "${CONFIG.KEY_COLUMN_NAME}"!`);

  const srcNameIdx = sourceHeaders.findIndex(h => normalizeHeader(h) === normalizeHeader(CONFIG.NAME_COLUMN_NAME));
  const srcGiaCuocIdx = sourceHeaders.findIndex(h => normalizeHeader(h) === 'giá cước');
  const srcChuKyIdx = sourceHeaders.findIndex(h => normalizeHeader(h) === 'chu kỳ sử dụng');
  const srcDtpkIdx = sourceHeaders.findIndex(h => normalizeHeader(h) === 'dtpk');
  const srcMaNVIdx = sourceHeaders.findIndex(h => normalizeHeader(h) === 'mã nhân viên');
  const srcTenNVIdx = sourceHeaders.findIndex(h => normalizeHeader(h) === 'tên nhân viên');
  const srcMaCHIdx = sourceHeaders.findIndex(h => normalizeHeader(h) === 'mã cửa hàng');
  const srcTenCHIdx = sourceHeaders.findIndex(h => normalizeHeader(h) === 'tên cửa hàng');
  const srcMoHinhIdx = sourceHeaders.findIndex(h => normalizeHeader(h) === 'loại mô hình triển khai');

  // Lọc dòng dữ liệu hợp lệ
  const sourceDataRows = rawDataRows.filter(r => {
    if (!r || r.length === 0) return false;
    const rawCode = r[srcKeyIdx];
    return isValidSubscriberRow(rawCode, r);
  });

  const lastTargetCol = targetSheet.getLastColumn();
  const lastTargetRow = targetSheet.getLastRow();
  const targetHeaders = targetSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastTargetCol).getValues()[0].map(h => String(h || '').trim());
  const targetKeyColIdx = targetHeaders.findIndex(h => normalizeHeader(h) === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const targetNameColIdx = targetHeaders.findIndex(h => normalizeHeader(h) === normalizeHeader(CONFIG.NAME_COLUMN_NAME));
  const giaCuocTargetColIdx = targetHeaders.findIndex(h => normalizeHeader(h) === 'giá cước');
  const dtpkTargetColIdx = targetHeaders.findIndex(h => normalizeHeader(h) === 'dtpk');
  const ngaySinhTargetColIdx = targetHeaders.findIndex(h => normalizeHeader(h) === 'ngày sinh');

  // Bản đồ ánh xạ cột Target -> Source
  const colMapping = {};
  targetHeaders.forEach((tHead, tIdx) => {
    if (tHead) {
      const sIdx = sourceHeaders.findIndex(sHead => normalizeHeader(sHead) === normalizeHeader(tHead));
      if (sIdx !== -1) colMapping[tIdx] = sIdx;
    }
  });

  const criticalFieldsNormalized = new Set(CONFIG.CRITICAL_LOG_FIELDS.map(f => normalizeHeader(f)));
  let targetData = [];
  const customerRowMap = new Map();

  if (lastTargetRow >= CONFIG.DATA_START_ROW_TARGET) {
    targetData = targetSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, lastTargetRow - CONFIG.DATA_START_ROW_TARGET + 1, lastTargetCol).getValues();
    targetData.forEach((row, idx) => {
      const normCode = normalizeCustomerId(row[targetKeyColIdx]);
      if (normCode) customerRowMap.set(normCode, { rowIndex: CONFIG.DATA_START_ROW_TARGET + idx, dataIndex: idx });
    });
  }

  const nowStr = Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm');
  const logs = [];
  const newRowsToAppend = [];

  // Mảng gom các thông tin tham số để quét (Lấy trực tiếp từ các cột nguồn của dòng)
  const itemsForParamSync = [];

  let updatedCustCount = 0, updatedFieldsCount = 0;

  sourceDataRows.forEach(srcRow => {
    const rawSrcCustId = srcRow[srcKeyIdx];
    const normSrcCustId = normalizeCustomerId(rawSrcCustId);
    if (!normSrcCustId) return;

    const custName = (srcNameIdx !== -1 && srcRow[srcNameIdx]) ? String(srcRow[srcNameIdx]).trim() : '';
    let cleanDisplayCustId = String(rawSrcCustId).replace(/[^0-9]/g, '').trim();
    if (cleanDisplayCustId.length === 8) cleanDisplayCustId = '00' + cleanDisplayCustId;
    else if (cleanDisplayCustId.length < 10) cleanDisplayCustId = cleanDisplayCustId.padStart(10, '0');

    let derivedDtpk = '';
    const rawGiaCuoc = (srcGiaCuocIdx !== -1) ? parseCurrencyToInteger(srcRow[srcGiaCuocIdx]) : '';
    const rawChuKy = (srcChuKyIdx !== -1) ? parseInt(srcRow[srcChuKyIdx], 10) : '';
    if (typeof rawGiaCuoc === 'number' && typeof rawChuKy === 'number' && rawChuKy > 0) {
      derivedDtpk = Math.round(rawGiaCuoc / rawChuKy);
    }

    // Thu thập tham số nhân sự từ dòng nguồn
    const itemMaNV = srcMaNVIdx !== -1 ? String(srcRow[srcMaNVIdx] || '').trim() : '';
    const itemTenNV = srcTenNVIdx !== -1 ? String(srcRow[srcTenNVIdx] || '').trim() : '';
    const itemMaCH = srcMaCHIdx !== -1 ? String(srcRow[srcMaCHIdx] || '').trim() : '';
    const itemTenCH = srcTenCHIdx !== -1 ? String(srcRow[srcTenCHIdx] || '').trim() : '';
    const itemMoHinh = srcMoHinhIdx !== -1 ? String(srcRow[srcMoHinhIdx] || '').trim() : '';

    if (itemMaNV || itemTenCH || itemMoHinh) {
      itemsForParamSync.push({
        maNV: itemMaNV,
        tenNV: itemTenNV,
        maCH: itemMaCH,
        tenCH: itemTenCH,
        moHinh: itemMoHinh,
        maKH: cleanDisplayCustId,
        tenTB: custName
      });
    }

    // Xử lý ghi dữ liệu vào bảng tính
    if (customerRowMap.has(normSrcCustId)) {
      // 1. CẬP NHẬT THUÊ BAO ĐÃ TỒN TẠI
      const { rowIndex, dataIndex } = customerRowMap.get(normSrcCustId);
      const existingRow = targetData[dataIndex];
      let custHasChange = false;
      const displayName = custName || (targetNameColIdx !== -1 ? String(existingRow[targetNameColIdx] || '') : '');

      for (const [tColIdxStr, sColIdx] of Object.entries(colMapping)) {
        const tColIdx = parseInt(tColIdxStr, 10);
        const fieldName = targetHeaders[tColIdx];
        const normField = normalizeHeader(fieldName);
        if (tColIdx === targetKeyColIdx) continue;

        const oldVal = existingRow[tColIdx];
        let formattedNewVal = formatValueByFieldType(fieldName, srcRow[sColIdx]);

        if (normField === 'dtpk' && (formattedNewVal === '' || formattedNewVal === null || formattedNewVal === undefined)) {
          formattedNewVal = derivedDtpk;
        }

        if (isValueChanged(fieldName, oldVal, formattedNewVal)) {
          existingRow[tColIdx] = formattedNewVal;
          const cellRange = targetSheet.getRange(rowIndex, tColIdx + 1);

          if (CONFIG.CURRENCY_INTEGER_FIELDS.some(f => normalizeHeader(f) === normField)) {
            cellRange.setValue(formattedNewVal).setNumberFormat('#,##0');
          } else if (normField === 'số cmnd/cccd' || normField === 'sđt liên hệ') {
            cellRange.setNumberFormat('@').setValue(formattedNewVal);
          } else if (normField === 'ngày sinh') {
            cellRange.setValue(formattedNewVal).setNumberFormat('dd/MM/yyyy');
          } else {
            cellRange.setValue(formattedNewVal);
          }

          if (criticalFieldsNormalized.has(normField)) {
            const colLetter = columnToLetter(tColIdx + 1);
            logs.push([
              nowStr, cleanDisplayCustId, displayName, 'Cập nhật', fieldName,
              `${colLetter}${rowIndex}`, formatValForLog(oldVal), formatValForLog(formattedNewVal)
            ]);
            updatedFieldsCount++;
          }
          custHasChange = true;
        }
      }
      if (custHasChange) updatedCustCount++;
    } else {
      // 2. THÊM THUÊ BAO MỚI
      const newRow = new Array(lastTargetCol).fill('');
      for (const [tColIdxStr, sColIdx] of Object.entries(colMapping)) {
        const tColIdx = parseInt(tColIdxStr, 10);
        newRow[tColIdx] = formatValueByFieldType(targetHeaders[tColIdx], srcRow[sColIdx]);
      }
      if (dtpkTargetColIdx !== -1 && (newRow[dtpkTargetColIdx] === '' || newRow[dtpkTargetColIdx] === null)) {
        newRow[dtpkTargetColIdx] = derivedDtpk;
      }
      newRowsToAppend.push({ custId: cleanDisplayCustId, custName: custName, rowData: newRow });
    }
  });

  let insertedCount = 0;
  if (newRowsToAppend.length > 0) {
    const startAppendRow = targetSheet.getLastRow() + 1;
    const appendValues = newRowsToAppend.map(item => item.rowData);
    targetSheet.getRange(startAppendRow, 1, appendValues.length, lastTargetCol).setValues(appendValues);
    insertedCount = appendValues.length;

    [targetKeyColIdx, targetHeaders.findIndex(h => normalizeHeader(h) === 'số cmnd/cccd'), targetHeaders.findIndex(h => normalizeHeader(h) === 'sđt liên hệ')].forEach(cIdx => {
      if (cIdx !== -1) targetSheet.getRange(startAppendRow, cIdx + 1, appendValues.length, 1).setNumberFormat('@');
    });

    if (giaCuocTargetColIdx !== -1) targetSheet.getRange(startAppendRow, giaCuocTargetColIdx + 1, appendValues.length, 1).setNumberFormat('#,##0');
    if (dtpkTargetColIdx !== -1) targetSheet.getRange(startAppendRow, dtpkTargetColIdx + 1, appendValues.length, 1).setNumberFormat('#,##0');
    if (ngaySinhTargetColIdx !== -1) targetSheet.getRange(startAppendRow, ngaySinhTargetColIdx + 1, appendValues.length, 1).setNumberFormat('dd/MM/yyyy');

    newRowsToAppend.forEach((item, idx) => {
      const curRow = startAppendRow + idx;
      logs.push([nowStr, item.custId, item.custName, 'Thêm mới', 'Toàn bộ bản ghi', `A${curRow}:AR${curRow}`, '', 'Cập nhật dữ liệu']);
    });
  }

  // TỰ ĐỘNG QUÉT VÀ ĐỒNG BỘ THAM SỐ VÀO thamso_caidat (DỰA TRÊN TẤT CẢ DÒNG NẠP VÀO)
  let syncResult = { newParamsCount: 0, details: '', newStaffList: [] };
  try {
    if (itemsForParamSync.length > 0) {
      syncResult = autoDetectAndSyncNewParameters(itemsForParamSync);
    }
  } catch (syncErr) {
    console.error('Lỗi tự động đồng bộ tham số:', syncErr);
  }

  // Ghi nhật ký vào log_capnhat
  const logSheet = initLogSheet(ss);
  if (logs.length > 0) {
    const startLogRow = logSheet.getLastRow() + 1;
    logSheet.getRange(startLogRow, 2, logs.length, 1).setNumberFormat('@');
    const safeLogs = logs.map(row => {
      const copy = [...row];
      copy[1] = "'" + copy[1].replace(/^'+/, '');
      return copy;
    });
    logSheet.getRange(startLogRow, 1, safeLogs.length, safeLogs[0].length).setValues(safeLogs);
    logSheet.getRange(startLogRow, 1, safeLogs.length, 1).setHorizontalAlignment('center');
    logSheet.getRange(startLogRow, 2, safeLogs.length, 1).setHorizontalAlignment('center');
    logSheet.getRange(startLogRow, 4, safeLogs.length, 1).setHorizontalAlignment('center');
    logSheet.getRange(startLogRow, 6, safeLogs.length, 1).setHorizontalAlignment('center');
    logSheet.autoResizeColumns(1, 8);
    for (let c = 1; c <= 8; c++) {
      if (logSheet.getColumnWidth(c) > 300) logSheet.setColumnWidth(c, 300);
    }
  }

  return {
    success: true,
    updatedCustCount: updatedCustCount,
    updatedFieldsCount: updatedFieldsCount,
    insertedCount: insertedCount,
    totalValidSourceRows: sourceDataRows.length,
    newParamsCount: syncResult.newParamsCount,
    syncDetails: syncResult.details,
    newStaffList: syncResult.newStaffList || []
  };
}

/**
 * TỰ ĐỘNG ĐỒNG BỘ THAM SỐ VÀO thamso_caidat
 * Đọc và chèn chuẩn xác vào từng bảng danh mục
 */
function autoDetectAndSyncNewParameters(paramItems) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tsSheet = ss.getSheetByName('thamso_caidat');
  if (!tsSheet) return { newParamsCount: 0, details: '', newStaffList: [] };

  const maxRows = tsSheet.getMaxRows();

  // 1. Đọc danh mục Mã NV hiện có từ Cột L (cột 12)
  const staffColVals = tsSheet.getRange(4, 12, Math.max(maxRows - 3, 1), 1).getValues();
  const existingStaffCodes = new Set();
  let lastStaffRow = 3;
  for (let i = 0; i < staffColVals.length; i++) {
    const val = String(staffColVals[i][0] || '').trim();
    if (val) {
      existingStaffCodes.add(normalizeHeader(val));
      lastStaffRow = 4 + i;
    }
  }

  // 2. Đọc danh mục Tên Cửa hàng từ Cột F (cột 6)
  const storeColVals = tsSheet.getRange(4, 6, Math.max(maxRows - 3, 1), 1).getValues();
  const existingStoreNames = new Set();
  let lastStoreRow = 3;
  for (let i = 0; i < storeColVals.length; i++) {
    const val = String(storeColVals[i][0] || '').trim();
    if (val) {
      existingStoreNames.add(normalizeHeader(val));
      lastStoreRow = 4 + i;
    }
  }

  // 3. Đọc danh mục Mô hình triển khai từ Cột C (cột 3)
  const modelColVals = tsSheet.getRange(4, 3, Math.max(maxRows - 3, 1), 1).getValues();
  const existingModels = new Set();
  let lastModelRow = 3;
  for (let i = 0; i < modelColVals.length; i++) {
    const val = String(modelColVals[i][0] || '').trim();
    if (val) {
      existingModels.add(normalizeHeader(val));
      lastModelRow = 4 + i;
    }
  }

  const newStaffToAppend = [];
  const newStoresToAppend = [];
  const newModelsToAppend = [];
  const addedStaffKeys = new Set();
  const addedStoreKeys = new Set();
  const addedModelKeys = new Set();
  const newStaffDetails = [];

  paramItems.forEach(item => {
    const maNV = item.maNV;
    const tenNV = item.tenNV;
    const maCH = item.maCH;
    const tenCH = item.tenCH;
    const moHinh = item.moHinh;

    // 1. Kiểm tra Nhân viên mới
    const normMaNV = normalizeHeader(maNV);
    if (normMaNV && !existingStaffCodes.has(normMaNV) && !addedStaffKeys.has(normMaNV)) {
      addedStaffKeys.add(normMaNV);
      // Cột J: Khu vực | Cột K: Cửa hàng | Cột L: Mã NV | Cột M: Tên NV
      newStaffToAppend.push(['[Chưa phân bổ KV]', tenCH || 'Chưa rõ', "'" + maNV, tenNV]);
      newStaffDetails.push({ maNV: maNV, tenNV: tenNV, tenCH: tenCH, maKH: item.maKH, tenTB: item.tenTB });
    }

    // 2. Kiểm tra Cửa hàng mới
    const normTenCH = normalizeHeader(tenCH);
    if (normTenCH && !existingStoreNames.has(normTenCH) && !addedStoreKeys.has(normTenCH)) {
      addedStoreKeys.add(normTenCH);
      // Cột F: Tên CH gốc | Cột G: Tên rút gọn
      newStoresToAppend.push([tenCH, tenCH]);
    }

    // 3. Kiểm tra Mô hình mới
    const normMoHinh = normalizeHeader(moHinh);
    if (normMoHinh && !existingModels.has(normMoHinh) && !addedModelKeys.has(normMoHinh)) {
      addedModelKeys.add(normMoHinh);
      const shortName = suggestShortModel(moHinh);
      // Cột C: Tên mô hình gốc | Cột D: Tên rút gọn
      newModelsToAppend.push([moHinh, shortName]);
    }
  });

  let totalNew = 0;
  const alertDetails = [];

  // Ghi Nhân viên mới vào sheet thamso_caidat (Cột J:M)
  if (newStaffToAppend.length > 0) {
    const startRowStaff = lastStaffRow + 1;
    if (startRowStaff + newStaffToAppend.length - 1 > tsSheet.getMaxRows()) {
      tsSheet.insertRowsAfter(tsSheet.getMaxRows(), newStaffToAppend.length + 5);
    }
    tsSheet.getRange(startRowStaff, 10, newStaffToAppend.length, 4).setValues(newStaffToAppend);
    tsSheet.getRange(startRowStaff, 10, newStaffToAppend.length, 1)
      .setBackground('#FCE8E6')
      .setFontColor('#C5221F')
      .setFontWeight('bold'); // Bôi đỏ cảnh báo cần phân bổ KV
    totalNew += newStaffToAppend.length;
    alertDetails.push(`${newStaffToAppend.length} Nhân viên mới`);
  }

  // Ghi Cửa hàng mới vào sheet thamso_caidat (Cột F:G)
  if (newStoresToAppend.length > 0) {
    const startRowStore = lastStoreRow + 1;
    if (startRowStore + newStoresToAppend.length - 1 > tsSheet.getMaxRows()) {
      tsSheet.insertRowsAfter(tsSheet.getMaxRows(), newStoresToAppend.length + 5);
    }
    tsSheet.getRange(startRowStore, 6, newStoresToAppend.length, 2).setValues(newStoresToAppend);
    totalNew += newStoresToAppend.length;
    alertDetails.push(`${newStoresToAppend.length} Cửa hàng mới`);
  }

  // Ghi Mô hình mới vào sheet thamso_caidat (Cột C:D)
  if (newModelsToAppend.length > 0) {
    const startRowModel = lastModelRow + 1;
    if (startRowModel + newModelsToAppend.length - 1 > tsSheet.getMaxRows()) {
      tsSheet.insertRowsAfter(tsSheet.getMaxRows(), newModelsToAppend.length + 5);
    }
    tsSheet.getRange(startRowModel, 3, newModelsToAppend.length, 2).setValues(newModelsToAppend);
    totalNew += newModelsToAppend.length;
    alertDetails.push(`${newModelsToAppend.length} Mô hình mới`);
  }

  return {
    newParamsCount: totalNew,
    details: alertDetails.join(', '),
    newStaffList: newStaffDetails
  };
}

function initLogSheet(ss) {
  let logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  const expectedHeaders = [
    'Thời gian xử lý', 'MÃ KHÁCH HÀNG', 'TÊN THUÊ BAO', 'Thao tác',
    'Trường dữ liệu (Cột)', 'Tọa độ ô (Row, Column)', 'Before (Giá trị cũ)', 'After (Giá trị mới)'
  ];
  if (!logSheet) {
    logSheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
    logSheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders])
      .setFontWeight('bold').setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');
    logSheet.setFrozenRows(1);
  }
  return logSheet;
}