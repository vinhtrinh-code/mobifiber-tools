/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * MODULE       : KẾT XUẤT DANH SÁCH GIA HẠN CƯỚC MOBIFIBER
 * PHIÊN BẢN    : 2.4 (Fix lỗi parseCurrency thành parseCurrencyToInteger)
 * TÁC GIẢ      : Trịnh Vũ Thế Vinh & AI Assistant
 * ============================================================================
 */

function showRenewalPromptDialog() {
  const html = HtmlService.createHtmlOutput(getRenewalDialogHtml())
    .setWidth(490)
    .setHeight(430)
    .setTitle('Kết xuất Danh sách Gia hạn');
  SpreadsheetApp.getUi().showModalDialog(html, 'Kết xuất Danh sách Gia hạn Cước');
}

function processRenewalReport(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);
  
  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) throw new Error('Bảng raw_mobifiber chưa có dữ liệu.');
  
  let startDate, endDate, sheetTitleSuffix, sheetName;
  if (params.type === 'month') {
    const month = parseInt(params.month, 10);
    const year = parseInt(params.year, 10);
    const mStr = month < 10 ? '0' + month : '' + month;
    startDate = new Date(year, month - 1, 1, 0, 0, 0);
    endDate = new Date(year, month, 0, 23, 59, 59);
    sheetName = `giahan_T${mStr}_${year}`;
    sheetTitleSuffix = `THÁNG ${mStr}/${year}`;
  } else {
    const startParts = params.from.split('/');
    const endParts = params.to.split('/');
    startDate = new Date(parseInt(startParts[2], 10), parseInt(startParts[1], 10) - 1, parseInt(startParts[0], 10), 0, 0, 0);
    endDate = new Date(parseInt(endParts[2], 10), parseInt(endParts[1], 10) - 1, parseInt(endParts[0], 10), 23, 59, 59);
    const d1 = startParts[0] + startParts[1];
    const d2 = endParts[0] + endParts[1];
    sheetName = `giahan_${d1}_${d2}_${endParts[2]}`;
    sheetTitleSuffix = `TỪ ${params.from} ĐẾN ${params.to}`;
  }
  
  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));
  const idxMaKH        = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxTenTB       = headers.findIndex(h => h === 'tên thuê bao');
  const idxSDT         = headers.findIndex(h => h === 'sđt liên hệ');
  const idxDiaChi      = headers.findIndex(h => h === 'địa chỉ lắp đặt');
  const idxGoiCuoc     = headers.findIndex(h => h === 'gói cước');
  const idxGiaCuoc     = headers.findIndex(h => h === 'giá cước');
  const idxChuKy       = headers.findIndex(h => h === 'chu kỳ sử dụng');
  const idxNgayBatDau  = headers.findIndex(h => h === 'ngày băt đầu gói cước' || h === 'ngày bắt đầu gói cước');
  const idxNgayKetThuc = headers.findIndex(h => h === 'ngày kết thúc gói cước');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxChanCat     = headers.findIndex(h => h === 'tình trạng chặn cắt');
  const idxTenCH       = headers.findIndex(h => h === 'tên cửa hàng');
  const idxTenNV       = headers.findIndex(h => h === 'tên nhân viên');
  
  const numRows = lastRow - CONFIG.DATA_START_ROW_TARGET + 1;
  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, numRows, lastCol).getValues();
  
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  const matchedRows = [];
  let countExpired = 0, countUrgent = 0, countWarning = 0, countInMonth = 0;
  let totalRevenue = 0;
  
  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const contractStatus = normalizeHeader(row[idxTinhTrangHD]);
    if (contractStatus === 'đã hủy' || contractStatus === 'da huy') continue;
  
    const endDateVal = parseToDate(row[idxNgayKetThuc]);
    if (!endDateVal) continue;
    if (endDateVal < startDate || endDateVal > endDate) continue;
  
    const endDayOnly = new Date(endDateVal.getTime());
    endDayOnly.setHours(0, 0, 0, 0);
    const diffDays = Math.ceil((endDayOnly.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
  
    let alertStatus = '', bgColor = '#ffffff', fgColor = '#202124';
  
    if (diffDays < 0) {
      alertStatus = '🔴 Mức 1: Quá hạn'; bgColor = '#FCE8E6'; fgColor = '#C5221F'; countExpired++;
    } else if (diffDays <= 7) {
      alertStatus = '🟠 Mức 2: Khẩn cấp (≤ 7 ngày)'; bgColor = '#FEF7E0'; fgColor = '#B06000'; countUrgent++;
    } else if (diffDays <= 15) {
      alertStatus = '🟡 Mức 3: Cận hạn (8 - 15 ngày)'; bgColor = '#FEF9E7'; fgColor = '#7E5700'; countWarning++;
    } else {
      alertStatus = '🟢 Mức 4: Trong hạn'; bgColor = '#E6F4EA'; fgColor = '#137333'; countInMonth++;
    }
  
    let cleanMaKH = String(row[idxMaKH] || '').replace(/^'+/, '').trim();
    if (/^\d{8}$/.test(cleanMaKH)) cleanMaKH = '00' + cleanMaKH;
    let cleanSDT = String(row[idxSDT] || '').replace(/^'+/, '').trim();
    if (/^\d{9}$/.test(cleanSDT)) cleanSDT = '0' + cleanSDT;
  
    // SỬ DỤNG ĐÚNG HÀM CHUẨN HÓA TIỀN TỆ MỚI NHẤT
    const giaCuoc = parseCurrencyToInteger(row[idxGiaCuoc]) || 0;
    totalRevenue += giaCuoc;
    const chuKy = row[idxChuKy] ? parseInt(row[idxChuKy], 10) : '';
  
    matchedRows.push({
      diffDays: diffDays,
      bgColor: bgColor,
      fgColor: fgColor,
      rowValues: [
        alertStatus,
        diffDays < 0 ? `Quá hạn ${Math.abs(diffDays)} ngày` : `Còn ${diffDays} ngày`,
        "'" + cleanMaKH, row[idxTenTB] || '', "'" + cleanSDT, row[idxDiaChi] || '', row[idxGoiCuoc] || '',
        giaCuoc, chuKy, formatDateTimeStr(row[idxNgayBatDau]), formatDateTimeStr(row[idxNgayKetThuc]),
        row[idxChanCat] || 'Bình thường', row[idxTenCH] || '', row[idxTenNV] || ''
      ]
    });
  }
  
  matchedRows.sort((a, b) => a.diffDays - b.diffDays);
  
  let destSheet = ss.getSheetByName(sheetName);
  if (!destSheet) destSheet = ss.insertSheet(sheetName);
  else destSheet.clear();
  
  const ssId = ss.getId();
  const sheetId = destSheet.getSheetId();
  const directExportUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=xlsx&gid=${sheetId}`;
  
  // DÒNG 1: Tiêu đề lớn
  destSheet.getRange(1, 1, 1, 14).merge()
    .setValue(`DANH SÁCH GIA HẠN MOBIFIBER - ${sheetTitleSuffix}`)
    .setFontWeight('bold').setFontSize(13).setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  
  // DÒNG 2: TỔNG THUÊ BAO
  destSheet.getRange(2, 1, 1, 7).merge()
    .setValue(`📊 TỔNG THUÊ BAO ĐẾN HẠN: ${matchedRows.length} TB`)
    .setFontWeight('bold').setFontSize(11).setBackground('#D2E3FC').setFontColor('#174EA6').setHorizontalAlignment('left');
  
  destSheet.getRange(2, 8, 1, 5).merge()
    .setValue(`🕒 Ngày lập: ${Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm')}`)
    .setFontSize(10).setBackground('#D2E3FC').setFontColor('#5F6368').setHorizontalAlignment('right');
  
  // DÒNG 3: DOANH THU GIA HẠN
  destSheet.getRange(3, 1, 1, 12).merge()
    .setValue(`💰 DOANH THU GIA HẠN: ${totalRevenue.toLocaleString('vi-VN')} VNĐ`)
    .setFontWeight('bold').setFontSize(11).setBackground('#E8F0FE').setFontColor('#174EA6').setHorizontalAlignment('left');
  
  // NÚT TẢI TRÊN SHEET
  destSheet.getRange(2, 13, 2, 2).merge()
    .setFormula(`=HYPERLINK("${directExportUrl}"; "📥 TẢI FILE EXCEL")`)
    .setFontWeight('bold').setFontSize(11).setBackground('#1A73E8').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  
  // DÒNG 4: Thống kê 4 mức màu
  const kpiDetails = [
    `🔴 Quá hạn: ${countExpired} TB`, '', '', `🟠 Khẩn cấp (≤ 7 ngày): ${countUrgent} TB`, '', '',
    `🟡 Cận hạn (8 - 15 ngày): ${countWarning} TB`, '', '', `🟢 Trong hạn: ${countInMonth} TB`, '', '', '', ''
  ];
  destSheet.getRange(4, 1, 1, 14).setValues([kpiDetails]);
  destSheet.getRange(4, 1, 1, 3).merge().setBackground('#FCE8E6').setFontColor('#C5221F').setFontWeight('bold').setHorizontalAlignment('center');
  destSheet.getRange(4, 4, 1, 3).merge().setBackground('#FEF7E0').setFontColor('#B06000').setFontWeight('bold').setHorizontalAlignment('center');
  destSheet.getRange(4, 7, 1, 3).merge().setBackground('#FEF9E7').setFontColor('#7E5700').setFontWeight('bold').setHorizontalAlignment('center');
  destSheet.getRange(4, 10, 1, 5).merge().setBackground('#E6F4EA').setFontColor('#137333').setFontWeight('bold').setHorizontalAlignment('center');
  
  // DÒNG 5: Headers
  const tableHeaders = [
    'Trạng thái cảnh báo', 'Thời hạn còn lại', 'MÃ KHÁCH HÀNG', 'TÊN THUÊ BAO', 'SĐT LIÊN HỆ',
    'ĐỊA CHỈ LẮP ĐẶT', 'GÓI CƯỚC', 'GIÁ CƯỚC', 'CHU KỲ', 'NGÀY BẮT ĐẦU', 'NGÀY KẾT THÚC',
    'TÌNH TRẠNG CHẶN CẮT', 'TÊN CỬA HÀNG', 'TÊN NHÂN VIÊN'
  ];
  destSheet.getRange(5, 1, 1, tableHeaders.length).setValues([tableHeaders])
    .setFontWeight('bold').setBackground('#00558F').setFontColor('#FFFFFF').setHorizontalAlignment('center');
  destSheet.setFrozenRows(5);
  
  if (matchedRows.length > 0) {
    const outputData = matchedRows.map(r => r.rowValues);
    const startRow = 6;
    destSheet.getRange(startRow, 1, outputData.length, tableHeaders.length).setValues(outputData);
  
    for (let r = 0; r < matchedRows.length; r++) {
      destSheet.getRange(startRow + r, 1, 1, 2)
        .setBackground(matchedRows[r].bgColor).setFontColor(matchedRows[r].fgColor).setFontWeight('bold');
    }
    destSheet.getRange(startRow, 1, outputData.length, 2).setHorizontalAlignment('center');
    destSheet.getRange(startRow, 3, outputData.length, 1).setNumberFormat('@').setHorizontalAlignment('center');
    destSheet.getRange(startRow, 5, outputData.length, 1).setNumberFormat('@').setHorizontalAlignment('center');
    destSheet.getRange(startRow, 8, outputData.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');
    destSheet.getRange(startRow, 9, outputData.length, 1).setHorizontalAlignment('center');
    destSheet.getRange(startRow, 10, outputData.length, 2).setHorizontalAlignment('center');
    destSheet.getRange(startRow, 12, outputData.length, 1).setHorizontalAlignment('center');
  
    destSheet.autoResizeColumns(1, tableHeaders.length);
    for (let col = 1; col <= tableHeaders.length; col++) {
      const curWidth = destSheet.getColumnWidth(col);
      let newWidth = curWidth + 28;
      if (newWidth > 300) newWidth = 300;
      if (newWidth < 90) newWidth = 90;
      destSheet.setColumnWidth(col, newWidth);
    }
  }
  
  ss.setActiveSheet(destSheet);
  
  return {
    success: true,
    sheetName: sheetName,
    count: matchedRows.length,
    revenue: totalRevenue.toLocaleString('vi-VN'),
    countExpired: countExpired,
    countUrgent: countUrgent,
    countWarning: countWarning,
    countInMonth: countInMonth
  };
}
  
function prepareCleanDownloadUrl(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(sheetName);
  if (!sourceSheet) throw new Error(`Không tìm thấy sheet "${sheetName}"!`);
  
  cleanupOldTempFiles(sheetName);
  
  const tempSs = SpreadsheetApp.create(sheetName);
  const tempSsId = tempSs.getId();
  
  try {
    const copiedSheet = sourceSheet.copyTo(tempSs);
    copiedSheet.setName(sheetName);
  
    const defaultSheet = tempSs.getSheets()[0];
    if (defaultSheet.getSheetId() !== copiedSheet.getSheetId()) {
      tempSs.deleteSheet(defaultSheet);
    }
  
    copiedSheet.getRange('M2:N3').breakApart().clear().setBackground('#D2E3FC');
    copiedSheet.getRange('H2:N2').merge().setValue(sourceSheet.getRange('H2').getValue()).setHorizontalAlignment('right');
    copiedSheet.getRange('A3:N3').merge().setValue(sourceSheet.getRange('A3').getValue()).setHorizontalAlignment('left');
  
    SpreadsheetApp.flush();
  
    const file = DriveApp.getFileById(tempSsId);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
    const downloadUrl = `https://docs.google.com/spreadsheets/d/${tempSsId}/export?format=xlsx`;
  
    return {
      success: true,
      downloadUrl: downloadUrl,
      fileName: `${sheetName}.xlsx`,
      fileId: tempSsId
    };
  } catch (err) {
    DriveApp.getFileById(tempSsId).setTrashed(true);
    throw err;
  }
}
  
function getRenewalDialogHtml() {
  const now = new Date();
  const curMonth = now.getMonth() + 1;
  const curYear = now.getFullYear();
  
  return `
  <!DOCTYPE html>
  <html>
  <head>
    <base target="_top">
    <style>
      body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; padding: 14px; margin: 0; background: #f8f9fa; font-size: 13px; color: #202124; }
      .opt-card { background: white; border: 1px solid #dadce0; border-radius: 8px; padding: 12px 14px; margin-bottom: 10px; }
      .radio-label { font-weight: 600; display: flex; align-items: center; gap: 8px; cursor: pointer; color: #1a73e8; }
      .inputs { margin-top: 8px; padding-left: 24px; }
      .hint { font-size: 11px; color: #5f6368; margin-top: 4px; line-height: 1.4; font-style: italic; }
      input[type="text"] { padding: 7px 10px; border: 1px solid #dadce0; border-radius: 4px; font-size: 13px; box-sizing: border-box; }
      input[type="text"]:focus { border-color: #1a73e8; outline: none; }
      .row { display: flex; gap: 8px; align-items: center; margin-top: 6px; }
      button { width: 100%; border: none; padding: 10px; font-size: 13px; font-weight: bold; border-radius: 6px; cursor: pointer; margin-top: 6px; display: block; box-sizing: border-box; text-align: center; }
      .btn-primary { background: #1a73e8; color: white; }
      .btn-primary:hover { background: #1557b0; }
      .btn-primary:disabled { background: #dadce0; cursor: not-allowed; }
      .btn-download { background: #0b8043; color: white; }
      .btn-download:hover { background: #086334; }
      .btn-close { background: #f1f3f4; color: #202124; border: 1px solid #dadce0; }
      .btn-close:hover { background: #e8eaed; }
      .alert-banner { display: none; background: #fce8e6; border: 1px solid #f5c2c7; color: #c5221f; padding: 8px 12px; border-radius: 6px; font-size: 12px; margin-bottom: 10px; }
      .status { font-size: 12px; margin-top: 8px; padding: 8px; border-radius: 4px; display: none; text-align: center; }
      .status.info { background: #e8f0fe; color: #174ea6; display: block; }
      #successView { display: none; text-align: center; padding: 6px 4px; }
      .success-title { font-size: 16px; font-weight: bold; color: #137333; margin-bottom: 8px; }
      .kpi-card { background: white; border: 1px solid #dadce0; border-radius: 6px; padding: 10px; margin-bottom: 8px; text-align: left; }
    </style>
  </head>
  <body>
    <div id="formView">
      <div id="errorBanner" class="alert-banner"></div>
      <div class="opt-card">
        <label class="radio-label">
          <input type="radio" name="renewType" value="month" checked onchange="toggleRenewType()">
          <strong>1. Kết xuất theo Tháng hết hạn</strong>
        </label>
        <div class="inputs" id="boxMonth">
          <span>Tháng:</span>
          <input type="text" id="mVal" value="${curMonth < 10 ? '0' + curMonth : curMonth}" style="width: 45px; text-align: center;"> /
          <span>Năm:</span>
          <input type="text" id="yVal" value="${curYear}" style="width: 65px; text-align: center;">
          <div class="hint">💡 Lọc chính xác các thuê bao có hạn kết thúc từ ngày 01 đến ngày cuối tháng.</div>
        </div>
      </div>
      <div class="opt-card">
        <label class="radio-label">
          <input type="radio" name="renewType" value="range" onchange="toggleRenewType()">
          <strong>2. Kết xuất theo Khoảng ngày tùy chọn</strong>
        </label>
        <div class="inputs" id="boxRange" style="opacity: 0.5; pointer-events: none;">
          <div class="row"><span style="width: 65px;">Từ ngày:</span><input type="text" id="fDate" placeholder="01/10/2026" style="width: 110px;"></div>
          <div class="row"><span style="width: 65px;">Đến ngày:</span><input type="text" id="tDate" placeholder="15/10/2026" style="width: 110px;"></div>
          <div class="hint">💡 Thích hợp lọc riêng danh sách quá hạn trong quá khứ hoặc một đợt ngắn ngày.</div>
        </div>
      </div>
      <button id="btnRunRenew" class="btn-primary" onclick="submitRenew()">🚀 Kết xuất Danh sách Gia hạn</button>
      <div id="statusRenew" class="status"></div>
    </div>
  
    <div id="successView">
      <div class="success-title">🎉 KẾT XUẤT THÀNH CÔNG!</div>
      <div class="kpi-card">
        <div style="font-weight: 600; color: #1a73e8;" id="resSheetName">Trang tính: ...</div>
        <div style="margin-top: 6px; font-size: 13px;">
          📊 <b>Tổng thuê bao:</b> <span id="resCount" style="color: #c5221f; font-weight: bold;">0</span> TB<br>
          💰 <b>Doanh thu gia hạn:</b> <span id="resRevenue" style="color: #137333; font-weight: bold;">0</span> VNĐ
        </div>
      </div>
      <div class="kpi-card" style="font-size: 12px; line-height: 1.6;">
        <span id="resExpired">🔴 Quá hạn: 0 TB</span><br>
        <span id="resUrgent">🟠 Khẩn cấp (≤ 7 ngày): 0 TB</span><br>
        <span id="resWarning">🟡 Cận hạn (8 - 15 ngày): 0 TB</span><br>
        <span id="resInMonth">🟢 Trong hạn: 0 TB</span>
      </div>
      <button id="btnDlRenew" class="btn-download" onclick="startCleanDownload()">📥 Tải file Excel (.xlsx) về máy</button>
      <button class="btn-close" onclick="closeDialogCleanly()">Đóng cửa sổ</button>
    </div>
  
    <iframe id="downloadFrame" style="display:none;"></iframe>
  
    <script>
      let currentCreatedSheetName = '';
      let currentTempRenewFileId = null;
  
      function showError(msg) {
        const b = document.getElementById('errorBanner');
        b.innerText = '⚠️ ' + msg;
        b.style.display = 'block';
      }
  
      function hideError() {
        document.getElementById('errorBanner').style.display = 'none';
      }
  
      function toggleRenewType() {
        hideError();
        const isMonth = document.querySelector('input[name="renewType"]:checked').value === 'month';
        document.getElementById('boxMonth').style.opacity = isMonth ? '1' : '0.5';
        document.getElementById('boxMonth').style.pointerEvents = isMonth ? 'auto' : 'none';
        document.getElementById('boxRange').style.opacity = isMonth ? '0.5' : '1';
        document.getElementById('boxRange').style.pointerEvents = isMonth ? 'none' : 'auto';
      }
  
      function isValidDateStr(str) {
        if (!/^\\d{1,2}\\/\\d{1,2}\\/\\d{4}$/.test(str)) return false;
        const p = str.split('/');
        const d = parseInt(p[0], 10), m = parseInt(p[1], 10), y = parseInt(p[2], 10);
        if (m < 1 || m > 12 || d < 1 || d > 31) return false;
        const dt = new Date(y, m - 1, d);
        return dt.getFullYear() === y && (dt.getMonth() + 1) === m && dt.getDate() === d;
      }
  
      function parseDateObj(str) {
        const p = str.split('/');
        return new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10));
      }
  
      function submitRenew() {
        hideError();
        const type = document.querySelector('input[name="renewType"]:checked').value;
        const btn = document.getElementById('btnRunRenew');
        const st = document.getElementById('statusRenew');
  
        let params = { type: type };
        if (type === 'month') {
          const m = document.getElementById('mVal').value.trim();
          const y = document.getElementById('yVal').value.trim();
          if (!m || !y) { showError('Vui lòng nhập đầy đủ Tháng và Năm!'); return; }
          const mNum = parseInt(m, 10), yNum = parseInt(y, 10);
          if (isNaN(mNum) || mNum < 1 || mNum > 12) { showError('Tháng phải là số từ 1 đến 12!'); return; }
          if (isNaN(yNum) || yNum < 2000 || yNum > 2100) { showError('Năm không hợp lệ!'); return; }
          params.month = m; params.year = y;
        } else {
          const f = document.getElementById('fDate').value.trim();
          const t = document.getElementById('tDate').value.trim();
          if (!f || !t) { showError('Vui lòng nhập đầy đủ Từ ngày và Đến ngày!'); return; }
          if (!isValidDateStr(f)) { showError('Định dạng "Từ ngày" không đúng (phải là dd/MM/yyyy)!'); return; }
          if (!isValidDateStr(t)) { showError('Định dạng "Đến ngày" không đúng (phải là dd/MM/yyyy)!'); return; }
          if (parseDateObj(f) > parseDateObj(t)) { showError('"Từ ngày" không được lớn hơn "Đến ngày"!'); return; }
          params.from = f; params.to = t;
        }
  
        btn.disabled = true;
        btn.innerText = '⏳ Đang xử lý...';
        st.className = 'status info';
        st.innerText = 'Hệ thống đang trích xuất dữ liệu và tối ưu bảng...';
  
        google.script.run
          .withSuccessHandler(function(res) {
            currentCreatedSheetName = res.sheetName;
            document.getElementById('formView').style.display = 'none';
            document.getElementById('resSheetName').innerText = '📁 Trang tính: ' + res.sheetName;
            document.getElementById('resCount').innerText = res.count;
            document.getElementById('resRevenue').innerText = res.revenue;
            document.getElementById('resExpired').innerText = '🔴 Quá hạn: ' + res.countExpired + ' TB';
            document.getElementById('resUrgent').innerText = '🟠 Khẩn cấp (≤ 7 ngày): ' + res.countUrgent + ' TB';
            document.getElementById('resWarning').innerText = '🟡 Cận hạn (8 - 15 ngày): ' + res.countWarning + ' TB';
            document.getElementById('resInMonth').innerText = '🟢 Trong hạn: ' + res.countInMonth + ' TB';
            document.getElementById('successView').style.display = 'block';
          })
          .withFailureHandler(function(err) {
            btn.disabled = false;
            btn.innerText = '🚀 Kết xuất Danh sách Gia hạn';
            st.style.display = 'none';
            showError(err.message);
          })
          .processRenewalReport(params);
      }
  
      function startCleanDownload() {
        const btn = document.getElementById('btnDlRenew');
        btn.disabled = true;
        btn.innerText = '⏳ Đang chuẩn bị file Excel sạch...';
  
        google.script.run
          .withSuccessHandler(function(res) {
            btn.disabled = false;
            btn.innerText = '📥 Tải file Excel (.xlsx) về máy';
            currentTempRenewFileId = res.fileId;
            document.getElementById('downloadFrame').src = res.downloadUrl;
  
            setTimeout(function() {
              if (currentTempRenewFileId) {
                google.script.run.deleteTempFile(currentTempRenewFileId);
                currentTempRenewFileId = null;
              }
            }, 5000);
          })
          .withFailureHandler(function(err) {
            btn.disabled = false;
            btn.innerText = '📥 Tải file Excel (.xlsx) về máy';
            showError('Lỗi tải file: ' + err.message);
          })
          .prepareCleanDownloadUrl(currentCreatedSheetName);
      }
  
      function closeDialogCleanly() {
        if (currentTempRenewFileId) {
          google.script.run.deleteTempFile(currentTempRenewFileId);
          currentTempRenewFileId = null;
        }
        google.script.host.close();
      }
    </script>
  </body>
  </html>
  `;
}