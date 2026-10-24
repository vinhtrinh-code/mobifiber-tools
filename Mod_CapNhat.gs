/**
 * ============================================================================
 * MODULE : CẬP NHẬT DỮ LIỆU & AUDIT LOGGING (HIGH-PERFORMANCE ENGINE)
 * PHIÊN BẢN : 3.0 (Tối ưu In-Memory Batching, Xử lý >10.000 thuê bao dưới 20s)
 * ============================================================================
 */

function processDataUpdateOptimized(sourceRawData, headerRowIndex) {
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  const logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);

  if (!targetSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);

  // 1. CHUẨN BỊ HEADER NGUỒN VÀ DỮ LIỆU ĐẦU VÀO TRONG RAM
  const sourceHeaders = sourceRawData[headerRowIndex].map(h => String(h || '').trim());
  const rawDataRows = sourceRawData.slice(headerRowIndex + 1);

  const sourceKeyIdx = sourceHeaders.findIndex(h => normalizeHeader(h) === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  if (sourceKeyIdx === -1) throw new Error(`File nguồn thiếu cột "${CONFIG.KEY_COLUMN_NAME}"!`);

  const sourceNameIdx = sourceHeaders.findIndex(h => normalizeHeader(h) === normalizeHeader(CONFIG.NAME_COLUMN_NAME));
  const sourceGiaCuocIdx = sourceHeaders.findIndex(h => normalizeHeader(h) === 'giá cước');
  const sourceChuKyIdx = sourceHeaders.findIndex(h => normalizeHeader(h) === 'chu kỳ sử dụng');
  const sourceDtpkIdx = sourceHeaders.findIndex(h => normalizeHeader(h) === 'dtpk');

  // Lọc nhanh dòng rác trong RAM
  const validSourceRows = rawDataRows.filter(r => r && r.length > 0 && isValidSubscriberRow(r[sourceKeyIdx], r));

  // 2. ĐỌC TOÀN BỘ SHEET ĐÍCH VÀO RAM (1 LẦN GỌI DUY NHẤT)
  const lastTargetRow = targetSheet.getLastRow();
  const lastTargetCol = targetSheet.getLastColumn();
  const targetHeaders = targetSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastTargetCol).getValues()[0].map(h => String(h || '').trim());

  const targetKeyIdx = targetHeaders.findIndex(h => normalizeHeader(h) === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const targetNameIdx = targetHeaders.findIndex(h => normalizeHeader(h) === normalizeHeader(CONFIG.NAME_COLUMN_NAME));
  const targetGiaCuocIdx = targetHeaders.findIndex(h => normalizeHeader(h) === 'giá cước');
  const targetDtpkIdx = targetHeaders.findIndex(h => normalizeHeader(h) === 'dtpk');

  // Xây dựng bản đồ ánh xạ cột (Column Mapping) O(1)
  const colMapping = [];
  targetHeaders.forEach((tHead, tIdx) => {
    if (tHead) {
      const sIdx = sourceHeaders.findIndex(sHead => normalizeHeader(sHead) === normalizeHeader(tHead));
      if (sIdx !== -1) colMapping.push({ targetIdx: tIdx, sourceIdx: sIdx, headerName: tHead });
    }
  });

  const criticalSet = new Set(CONFIG.CRITICAL_LOG_FIELDS.map(f => normalizeHeader(f)));

  // Nạp toàn bộ dữ liệu đích vào bộ nhớ
  let targetData = [];
  const targetMap = new Map(); // Key: Mã KH -> { rowIndex, dataIndex }

  if (lastTargetRow >= CONFIG.DATA_START_ROW_TARGET) {
    const numRows = lastTargetRow - CONFIG.DATA_START_ROW_TARGET + 1;
    targetData = targetSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, numRows, lastTargetCol).getValues();
    
    for (let i = 0; i < targetData.length; i++) {
      const code = normalizeCustomerId(targetData[i][targetKeyIdx]);
      if (code) {
        targetMap.set(code, {
          rowIndex: CONFIG.DATA_START_ROW_TARGET + i,
          dataIndex: i
        });
      }
    }
  }

  const nowStr = Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm');
  const pendingLogs = [];
  const newRowsToAppend = [];
  let updateCount = 0;

  // 3. XỬ LÝ ĐỐI SOÁT SIÊU TỐC TRÊN RAM
  validSourceRows.forEach(srcRow => {
    const custId = normalizeCustomerId(srcRow[sourceKeyIdx]);
    const custName = sourceNameIdx !== -1 ? String(srcRow[sourceNameIdx] || '').trim() : '';

    // Tính DTPK tự động nếu nguồn chưa có
    let dtpkVal = (sourceDtpkIdx !== -1) ? srcRow[sourceDtpkIdx] : null;
    if (!dtpkVal && sourceGiaCuocIdx !== -1 && sourceChuKyIdx !== -1) {
      const gc = parseCurrency(srcRow[sourceGiaCuocIdx]);
      const ck = parseInt(srcRow[sourceChuKyIdx], 10) || 1;
      dtpkVal = Math.round(gc / ck);
    }

    if (targetMap.has(custId)) {
      // TRƯỜNG HỢP: KHÁCH HÀNG ĐÃ TỒN TẠI -> ĐỐI SOÁT & CẬP NHẬT TRÊN RAM
      const info = targetMap.get(custId);
      const currentRow = targetData[info.dataIndex];
      let hasRowChanged = false;

      colMapping.forEach(m => {
        let newVal = srcRow[m.sourceIdx];
        if (m.targetIdx === targetDtpkIdx && dtpkVal !== null) newVal = dtpkVal;

        const oldVal = currentRow[m.targetIdx];
        const normHeader = normalizeHeader(m.headerName);

        if (isFieldChanged(normHeader, oldVal, newVal)) {
          hasRowChanged = true;
          // Chỉ ghi log các trường trọng yếu (Critical)
          if (criticalSet.has(normHeader)) {
            pendingLogs.push([
              nowStr,
              custId,
              custName || currentRow[targetNameIdx],
              'Cập nhật',
              m.headerName,
              `R${info.rowIndex}C${m.targetIdx + 1}`,
              String(oldVal ?? ''),
              String(newVal ?? '')
            ]);
          }
          currentRow[m.targetIdx] = newVal; // Ghi đè trực tiếp trong RAM
        }
      });

      if (hasRowChanged) updateCount++;
    } else {
      // TRƯỜNG HỢP: KHÁCH HÀNG MỚI -> TẠO DÒNG MỚI TRONG RAM
      const newRow = new Array(lastTargetCol).fill('');
      colMapping.forEach(m => {
        let val = srcRow[m.sourceIdx];
        if (m.targetIdx === targetDtpkIdx && dtpkVal !== null) val = dtpkVal;
        newRow[m.targetIdx] = val;
      });

      newRowsToAppend.push(newRow);

      pendingLogs.push([
        nowStr,
        custId,
        custName,
        'Thêm mới',
        'Toàn bộ bản ghi',
        `Dòng mới`,
        '',
        'Khách hàng mới hòa mạng'
      ]);
    }
  });

  // 4. BATCH WRITE TRỞ LẠI GOOGLE SHEETS (CHỈ GỌI 2 LẦN GHI DUY NHẤT)
  
  // 4.1 Ghi đè toàn bộ dữ liệu đã cập nhật lên raw_mobifiber
  if (targetData.length > 0) {
    targetSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, targetData.length, lastTargetCol).setValues(targetData);
  }

  // 4.2 Nối các dòng mới (nếu có)
  if (newRowsToAppend.length > 0) {
    const nextStartRow = targetSheet.getLastRow() + 1;
    targetSheet.getRange(nextStartRow, 1, newRowsToAppend.length, lastTargetCol).setValues(newRowsToAppend);
  }

  // 4.3 Ghi toàn bộ logs vào log_capnhat một lần
  if (logSheet && pendingLogs.length > 0) {
    const logStartRow = logSheet.getLastRow() + 1;
    logSheet.getRange(logStartRow, 1, pendingLogs.length, pendingLogs[0].length).setValues(pendingLogs);
  }

  const executionTime = ((new Date().getTime() - startTime) / 1000).toFixed(2);
  Logger.log(`⚡ Xử lý xong trong: ${executionTime}s | Đã cập nhật: ${updateCount} TB | Thêm mới: ${newRowsToAppend.length} TB | Logs: ${pendingLogs.length}`);

  return {
    updated: updateCount,
    added: newRowsToAppend.length,
    logs: pendingLogs.length,
    duration: executionTime
  };
}