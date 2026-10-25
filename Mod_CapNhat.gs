/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * MODULE       : Mod_CapNhat.gs
 * PHIÊN BẢN    : 2.5 (Cập nhật dữ liệu từ file Excel và ghi Log đối soát)
 * TÁC GIẢ      : Trịnh Vũ Thế Vinh & AI Assistant
 * ============================================================================
 */

function showUploadDialog() {
  const html = HtmlService.createHtmlOutputFromFile('UploadDialog')
    .setWidth(590)
    .setHeight(520)
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
  const newRawRowsToScan = [];

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

    if (customerRowMap.has(normSrcCustId)) {
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
      const newRow = new Array(lastTargetCol).fill('');
      for (const [tColIdxStr, sColIdx] of Object.entries(colMapping)) {
        const tColIdx = parseInt(tColIdxStr, 10);
        newRow[tColIdx] = formatValueByFieldType(targetHeaders[tColIdx], srcRow[sColIdx]);
      }
      if (dtpkTargetColIdx !== -1 && (newRow[dtpkTargetColIdx] === '' || newRow[dtpkTargetColIdx] === null)) {
        newRow[dtpkTargetColIdx] = derivedDtpk;
      }
      newRowsToAppend.push({ custId: cleanDisplayCustId, custName: custName, rowData: newRow });
      newRawRowsToScan.push(newRow);
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

  let syncResult = { newParamsCount: 0, details: '' };
  if (newRawRowsToScan.length > 0 && typeof autoDetectAndSyncNewParameters === 'function') {
    syncResult = autoDetectAndSyncNewParameters(newRawRowsToScan, colIndicesForSync);
  }

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
    syncDetails: syncResult.details
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