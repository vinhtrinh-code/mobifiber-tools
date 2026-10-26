/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * MODULE       : Mod_CapNhat.gs
 * PHIÊN BẢN    : 2.6 (Quét tham số toàn diện cho cả TB Cũ và TB Mới + Cảnh báo)
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

  const sourceKeyColIdx = sourceHeaders.findIndex(h => normalizeHeader(h) === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  if (sourceKeyColIdx === -1) throw new Error(`File nguồn không chứa cột "${CONFIG.KEY_COLUMN_NAME}"!`);

  const sourceNameColIdx = sourceHeaders.findIndex(h => normalizeHeader(h) === normalizeHeader(CONFIG.NAME_COLUMN_NAME));
  const sourceGiaCuocIdx = sourceHeaders.findIndex(h => normalizeHeader(h) === 'giá cước');
  const sourceChuKyIdx = sourceHeaders.findIndex(h => normalizeHeader(h) === 'chu kỳ sử dụng');
  const sourceDtpkIdx = sourceHeaders.findIndex(h => normalizeHeader(h) === 'dtpk');

  // Lọc dòng hợp lệ
  const sourceDataRows = rawDataRows.filter(r => {
    if (!r || r.length === 0) return false;
    const rawCode = r[sourceKeyColIdx];
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

  const colIndicesForSync = {
    idxLoaiKH: targetHeaders.findIndex(h => normalizeHeader(h) === 'loại khách hàng'),
    idxMoHinh: targetHeaders.findIndex(h => normalizeHeader(h) === 'loại mô hình triển khai'),
    idxTenCH: targetHeaders.findIndex(h => normalizeHeader(h) === 'tên cửa hàng'),
    idxMaCH: targetHeaders.findIndex(h => normalizeHeader(h) === 'mã cửa hàng'),
    idxMaNV: targetHeaders.findIndex(h => normalizeHeader(h) === 'mã nhân viên'),
    idxTenNV: targetHeaders.findIndex(h => normalizeHeader(h) === 'tên nhân viên')
  };

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
  const allRowsToScan = []; // Nạp TẤT CẢ các dòng để đối soát tham số

  let updatedCustCount = 0, updatedFieldsCount = 0;

  sourceDataRows.forEach(srcRow => {
    const rawSrcCustId = srcRow[sourceKeyColIdx];
    const normSrcCustId = normalizeCustomerId(rawSrcCustId);
    if (!normSrcCustId) return;

    const custName = (sourceNameColIdx !== -1 && srcRow[sourceNameColIdx]) ? String(srcRow[sourceNameColIdx]).trim() : '';
    let cleanDisplayCustId = String(rawSrcCustId).replace(/[^0-9]/g, '').trim();
    if (cleanDisplayCustId.length === 8) cleanDisplayCustId = '00' + cleanDisplayCustId;
    else if (cleanDisplayCustId.length < 10) cleanDisplayCustId = cleanDisplayCustId.padStart(10, '0');

    let derivedDtpk = '';
    const rawGiaCuoc = (sourceGiaCuocIdx !== -1) ? parseCurrencyToInteger(srcRow[sourceGiaCuocIdx]) : '';
    const rawChuKy = (sourceChuKyIdx !== -1) ? parseInt(srcRow[sourceChuKyIdx], 10) : '';
    if (typeof rawGiaCuoc === 'number' && typeof rawChuKy === 'number' && rawChuKy > 0) {
      derivedDtpk = Math.round(rawGiaCuoc / rawChuKy);
    }

    // Chuyển đổi định dạng các trường chuẩn
    const normalizedTargetRow = new Array(lastTargetCol).fill('');
    for (const [tColIdxStr, sColIdx] of Object.entries(colMapping)) {
      const tColIdx = parseInt(tColIdxStr, 10);
      normalizedTargetRow[tColIdx] = formatValueByFieldType(targetHeaders[tColIdx], srcRow[sColIdx]);
    }

    // Thu thập dòng để quét danh mục tham số (cả TB cũ lẫn TB mới)
    allRowsToScan.push(normalizedTargetRow);

    if (customerRowMap.has(normSrcCustId)) {
      // ĐỐI SOÁT CẬP NHẬT
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
        let formattedNewVal = normalizedTargetRow[tColIdx];

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
      // THÊM MỚI VÀO CUỐI BẢNG
      if (dtpkTargetColIdx !== -1 && (normalizedTargetRow[dtpkTargetColIdx] === '' || normalizedTargetRow[dtpkTargetColIdx] === null)) {
        normalizedTargetRow[dtpkTargetColIdx] = derivedDtpk;
      }
      newRowsToAppend.push({ custId: cleanDisplayCustId, custName: custName, rowData: normalizedTargetRow });
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

  // TỰ ĐỘNG BẮT PHÁT SINH THAM SỐ TOÀN DIỆN (QUÉT TRÊN TOÀN BỘ FILE NẠP VÀO)
  let syncResult = { newParamsCount: 0, details: '', newStaffList: [] };
  if (allRowsToScan.length > 0) {
    syncResult = autoDetectAndSyncNewParameters(allRowsToScan, colIndicesForSync);
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
 * TỰ ĐỘNG BẮT PHÁT SINH THAM SỐ MỚI VÀO thamso_caidat
 */
function autoDetectAndSyncNewParameters(rowsToScan, colIndices) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tsSheet = ss.getSheetByName('thamso_caidat');
  if (!tsSheet) return { newParamsCount: 0, details: '', newStaffList: [] };

  const lastRow = tsSheet.getLastRow();

  // 1. Nạp danh mục Mã NV hiện có từ Cột L (cột 12)
  const existingStaffCodes = new Set();
  if (lastRow >= 4) {
    const staffData = tsSheet.getRange(4, 12, lastRow - 3, 1).getValues();
    staffData.forEach(r => {
      const code = normalizeHeader(String(r[0] || ''));
      if (code) existingStaffCodes.add(code);
    });
  }

  // 2. Nạp danh mục Mã Cửa hàng hiện có từ Cột F (cột 6)
  const existingStoreCodes = new Set();
  if (lastRow >= 4) {
    const storeData = tsSheet.getRange(4, 6, lastRow - 3, 1).getValues();
    storeData.forEach(r => {
      const code = normalizeHeader(String(r[0] || ''));
      if (code) existingStoreCodes.add(code);
    });
  }

  const newStaffToAppend = [];
  const newStoresToAppend = [];
  const addedStaffKeys = new Set();
  const addedStoreKeys = new Set();
  const newStaffDetails = [];

  rowsToScan.forEach(row => {
    const maNV = String(row[colIndices.idxMaNV] || '').trim();
    const tenNV = String(row[colIndices.idxTenNV] || '').trim();
    const maCH = String(row[colIndices.idxMaCH] || '').trim();
    const tenCH = String(row[colIndices.idxTenCH] || '').trim();

    // Kiểm tra Nhân viên mới
    const normMaNV = normalizeHeader(maNV);
    if (normMaNV && !existingStaffCodes.has(normMaNV) && !addedStaffKeys.has(normMaNV)) {
      addedStaffKeys.add(normMaNV);
      // Cột J: Khu vực | Cột K: Cửa hàng | Cột L: Mã NV | Cột M: Tên NV
      newStaffToAppend.push(['[Chưa phân bổ KV]', tenCH || 'Chưa rõ', "'" + maNV, tenNV]);
      newStaffDetails.push({ maNV: maNV, tenNV: tenNV, tenCH: tenCH });
    }

    // Kiểm tra Cửa hàng mới
    const normMaCH = normalizeHeader(maCH);
    if (normMaCH && !existingStoreCodes.has(normMaCH) && !addedStoreKeys.has(normMaCH)) {
      addedStoreKeys.add(normMaCH);
      // Cột F: Mã CH | Cột G: Tên CH | Cột H: Tên CH rút gọn
      newStoresToAppend.push(["'" + maCH, tenCH, tenCH]);
    }
  });

  let totalNew = 0;
  const alertDetails = [];

  // Ghi Nhân viên mới vào sheet thamso_caidat (Cột J:M)
  if (newStaffToAppend.length > 0) {
    const startRowStaff = Math.max(tsSheet.getLastRow() + 1, 4);
    tsSheet.getRange(startRowStaff, 10, newStaffToAppend.length, 4).setValues(newStaffToAppend);
    tsSheet.getRange(startRowStaff, 10, newStaffToAppend.length, 1)
      .setBackground('#FCE8E6')
      .setFontColor('#C5221F')
      .setFontWeight('bold'); // Bôi đỏ cảnh báo cần phân bổ KV
    totalNew += newStaffToAppend.length;
    alertDetails.push(`${newStaffToAppend.length} Nhân viên mới`);
  }

  // Ghi Cửa hàng mới vào sheet thamso_caidat (Cột F:H)
  if (newStoresToAppend.length > 0) {
    const startRowStore = Math.max(tsSheet.getLastRow() + 1, 4);
    tsSheet.getRange(startRowStore, 6, newStoresToAppend.length, 3).setValues(newStoresToAppend);
    totalNew += newStoresToAppend.length;
    alertDetails.push(`${newStoresToAppend.length} Cửa hàng mới`);
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