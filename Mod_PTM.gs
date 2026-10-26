/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * MODULE       : Mod_PTM.gs
 * PHIÊN BẢN    : 2.8 (Khôi phục nút Tải Excel trên Dialog, Căn trái Dòng 1-2, Nối 2 cột A-B Khu vực)
 * TÁC GIẢ      : Trịnh Vũ Thế Vinh & AI Assistant
 * ============================================================================
 */

function showPtmReportDialog() {
  const html = HtmlService.createHtmlOutput(getPtmDialogHtml())
    .setWidth(500)
    .setHeight(470)
    .setTitle('Báo cáo PTM');
  SpreadsheetApp.getUi().showModalDialog(html, 'Báo cáo Thuê bao Phát triển mới (PTM)');
}

function processPtmReport(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) throw new Error('Bảng raw_mobifiber chưa có dữ liệu.');

  const mapping = loadThamSoMapping();

  let startDate, endDate, sheetTitleSuffix, sheetName;
  if (params.type === 'month') {
    const month = parseInt(params.month, 10);
    const year = parseInt(params.year, 10);
    const mStr = month < 10 ? '0' + month : '' + month;
    startDate = new Date(year, month - 1, 1, 0, 0, 0);
    endDate = new Date(year, month, 0, 23, 59, 59);
    sheetName = `PTM_T${mStr}_${year}`;
    sheetTitleSuffix = `THÁNG ${mStr}/${year}`;
  } else {
    const startParts = params.from.split('/');
    const endParts = params.to.split('/');
    startDate = new Date(parseInt(startParts[2], 10), parseInt(startParts[1], 10) - 1, parseInt(startParts[0], 10), 0, 0, 0);
    endDate = new Date(parseInt(endParts[2], 10), parseInt(endParts[1], 10) - 1, parseInt(endParts[0], 10), 23, 59, 59);
    const d1 = startParts[0] + startParts[1];
    const d2 = endParts[0] + endParts[1];
    sheetName = `PTM_${d1}_${d2}_${endParts[2]}`;
    sheetTitleSuffix = `TỪ ${params.from} ĐẾN ${params.to}`;
  }

  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));
  const idxMaKH = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxTenTB = headers.findIndex(h => h === 'tên thuê bao');
  const idxSDT = headers.findIndex(h => h === 'sđt liên hệ');
  const idxDiaChi = headers.findIndex(h => h === 'địa chỉ lắp đặt');
  const idxGoiCuoc = headers.findIndex(h => h === 'gói cước');
  const idxGiaCuoc = headers.findIndex(h => h === 'giá cước');
  const idxChuKy = headers.findIndex(h => h === 'chu kỳ sử dụng');
  const idxDtpk = headers.findIndex(h => h === 'dtpk');
  const idxNgayDauNoi = headers.findIndex(h => h === 'ngày đấu nối');
  const idxNgayThiCong = headers.findIndex(h => h === 'ngày hoàn tất thi công');
  const idxNgayBatDau = headers.findIndex(h => h === 'ngày băt đầu gói cước' || h === 'ngày bắt đầu gói cước');
  const idxThietBi = headers.findIndex(h => h === 'loại thiết bị ont');
  const idxSerial = headers.findIndex(h => h === 'serial ont');
  const idxLoaiKH = headers.findIndex(h => h === 'loại khách hàng');
  const idxNhuCau = headers.findIndex(h => h === 'nhu cầu hòa mạng');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxMaCH = headers.findIndex(h => h === 'mã cửa hàng');
  const idxTenCH = headers.findIndex(h => h === 'tên cửa hàng');
  const idxMaNV = headers.findIndex(h => h === 'mã nhân viên');
  const idxTenNV = headers.findIndex(h => h === 'tên nhân viên');

  const numRows = lastRow - CONFIG.DATA_START_ROW_TARGET + 1;
  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, numRows, lastCol).getValues();

  const ptmRowsData = [];
  const regionSummaryMap = {};
  let totalRevenue = 0;
  let totalDtpk = 0;

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];

    const nhuCau = normalizeHeader(row[idxNhuCau]);
    if (nhuCau !== 'đăng ký mới' && nhuCau !== 'dang ky moi') continue;

    const loaiKH = normalizeHeader(row[idxLoaiKH]);
    if (loaiKH === 'thuê bao nghiệp vụ' || loaiKH === 'thue bao nghiep vu') continue;

    const tinhTrang = normalizeHeader(row[idxTinhTrangHD]);
    if (tinhTrang !== 'đang hoạt động' && tinhTrang !== 'dang hoat dong') continue;

    const chuKyRaw = row[idxChuKy];
    const chuKy = chuKyRaw ? parseInt(chuKyRaw, 10) : 0;

    let ptmDate = null;
    const dDauNoi = parseToDate(row[idxNgayDauNoi]);
    const dBatDau = parseToDate(row[idxNgayBatDau]);
    const dThiCong = parseToDate(row[idxNgayThiCong]);

    if (chuKy === 1) {
      if (dDauNoi && dThiCong) {
        ptmDate = (dThiCong.getTime() > dDauNoi.getTime()) ? dThiCong : dDauNoi;
      } else {
        ptmDate = dThiCong || dDauNoi || dBatDau;
      }
    } else {
      ptmDate = dBatDau;
    }

    if (!ptmDate) continue;
    if (ptmDate < startDate || ptmDate > endDate) continue;

    const giaCuoc = parseCurrencyToInteger(row[idxGiaCuoc]) || 0;
    const dtpk = parseCurrencyToInteger(row[idxDtpk]) || (chuKy > 0 ? Math.round(giaCuoc / chuKy) : 0);

    totalRevenue += giaCuoc;
    totalDtpk += dtpk;

    const rawMaNV = String(row[idxMaNV] || '').trim();
    let khuVuc = 'Chưa phân bổ KV';
    if (rawMaNV && mapping.staffRegion && mapping.staffRegion.has(normalizeHeader(rawMaNV))) {
      const staffInfo = mapping.staffRegion.get(normalizeHeader(rawMaNV));
      if (staffInfo.khuVuc && !staffInfo.khuVuc.includes('Chưa phân bổ')) {
        khuVuc = staffInfo.khuVuc;
      }
    }

    if (!regionSummaryMap[khuVuc]) regionSummaryMap[khuVuc] = { count: 0, revenue: 0, dtpk: 0 };
    regionSummaryMap[khuVuc].count++;
    regionSummaryMap[khuVuc].revenue += giaCuoc;
    regionSummaryMap[khuVuc].dtpk += dtpk;

    let cleanMaKH = String(row[idxMaKH] || '').replace(/^'+/, '').trim();
    if (/^\d{8}$/.test(cleanMaKH)) cleanMaKH = '00' + cleanMaKH;

    let cleanSDT = String(row[idxSDT] || '').replace(/^'+/, '').trim();
    if (/^\d{9}$/.test(cleanSDT)) cleanSDT = '0' + cleanSDT;

    const paddedTenTB = row[idxTenTB] ? ' ' + String(row[idxTenTB]).trim() : '';
    const paddedDiaChi = row[idxDiaChi] ? ' ' + String(row[idxDiaChi]).trim() : '';
    const paddedThietBi = row[idxThietBi] ? ' ' + String(row[idxThietBi]).trim() : '';
    const cleanSerial = String(row[idxSerial] || '').replace(/^'+/, '').trim();
    const cleanMaCH = String(row[idxMaCH] || '').replace(/^'+/, '').trim();
    const cleanMaNV = String(row[idxMaNV] || '').replace(/^'+/, '').trim();

    const rawTenCH = String(row[idxTenCH] || '').trim();
    const shortTenCH = getShortValue(mapping.stores, rawTenCH);
    const paddedTenCH = shortTenCH ? ' ' + shortTenCH : '';
    const paddedTenNV = row[idxTenNV] ? ' ' + String(row[idxTenNV]).trim() : '';

    const sortTime = ptmDate ? ptmDate.getTime() : (dDauNoi ? dDauNoi.getTime() : 9999999999999);

    ptmRowsData.push({
      khuVuc: khuVuc,
      maCH: cleanMaCH,
      maNV: cleanMaNV,
      sortTime: sortTime,
      rowValues: [
        0,
        "'" + cleanMaKH,
        paddedTenTB,
        "'" + cleanSDT,
        paddedDiaChi,
        row[idxGoiCuoc] || '',
        giaCuoc,
        chuKy > 0 ? chuKy : '',
        dtpk,
        formatDateTimeStr(row[idxNgayDauNoi]),
        formatDateTimeStr(row[idxNgayBatDau]),
        paddedThietBi,
        "'" + cleanSerial,
        khuVuc,
        "'" + cleanMaCH,
        "'" + cleanMaNV,
        paddedTenCH,
        paddedTenNV
      ]
    });
  }

  // Sắp xếp 4 cấp
  ptmRowsData.sort((a, b) => {
    const compKV = a.khuVuc.localeCompare(b.khuVuc, 'vi', { sensitivity: 'base' });
    if (compKV !== 0) return compKV;
    const compCH = a.maCH.localeCompare(b.maCH, 'vi', { numeric: true });
    if (compCH !== 0) return compCH;
    const compNV = a.maNV.localeCompare(b.maNV, 'vi', { numeric: true });
    if (compNV !== 0) return compNV;
    return a.sortTime - b.sortTime;
  });

  const ptmFinalRows = ptmRowsData.map((item, index) => {
    item.rowValues[0] = index + 1;
    return item.rowValues;
  });

  let destSheet = ss.getSheetByName(sheetName);
  if (!destSheet) destSheet = ss.insertSheet(sheetName);
  else destSheet.clear();

  const ssId = ss.getId();
  const sheetId = destSheet.getSheetId();
  const directExportUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=xlsx&gid=${sheetId}`;

  // =========================================================================
  // DÒNG 1: TIÊU ĐỀ LỚN (CĂN TRÁI THEO YÊU CẦU)
  // =========================================================================
  destSheet.getRange(1, 1, 1, 18).merge()
    .setValue(`  BÁO CÁO THUÊ BAO MOBIFIBER PHÁT TRIỂN MỚI (PTM) - ${sheetTitleSuffix}`)
    .setFontWeight('bold')
    .setFontSize(13)
    .setBackground('#1B365D')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  // =========================================================================
  // DÒNG 2: TỔNG QUAN KPI (CĂN TRÁI THEO YÊU CẦU) & NÚT TẢI EXCEL
  // =========================================================================
  destSheet.getRange(2, 1, 1, 15).merge()
    .setValue(`  🎯 Tổng số thuê bao PTM: ${ptmRowsData.length} TB | 💰 Tổng doanh thu: ${totalRevenue.toLocaleString('vi-VN')} VNĐ | 📈 Tổng DTPK: ${totalDtpk.toLocaleString('vi-VN')} VNĐ`)
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground('#D2E3FC')
    .setFontColor('#174EA6')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('middle');

  destSheet.getRange(2, 16, 1, 3).merge()
    .setFormula(`=HYPERLINK("${directExportUrl}"; "📥 TẢI FILE EXCEL")`)
    .setFontWeight('bold')
    .setFontSize(11)
    .setBackground('#1A73E8')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  // =========================================================================
  // BẢNG 1: CƠ CẤU PHÁT TRIỂN MỚI THEO KHU VỰC (GỘP A-B KHU VỰC, C-E LIỀN KỀ)
  // =========================================================================
  destSheet.getRange(4, 1, 1, 5).merge()
    .setValue('1. CƠ CẤU PHÁT TRIỂN MỚI THEO KHU VỰC')
    .setFontWeight('bold')
    .setBackground('#E8F0FE')
    .setFontColor('#174EA6');

  // Tiêu đề dòng 5: Gộp A5:B5 cho 'KHU VỰC', C5 là 'SỐ THUÊ BAO', D5 là 'TỔNG DTPK', E5 là 'TỶ TRỌNG'
  destSheet.getRange('A5:B5').merge().setValue('KHU VỰC');
  destSheet.getRange('C5').setValue('SỐ THUÊ BAO');
  destSheet.getRange('D5').setValue('TỔNG DTPK (VNĐ)');
  destSheet.getRange('E5').setValue('TỶ TRỌNG DTPK');
  
  destSheet.getRange('A5:E5')
    .setFontWeight('bold')
    .setBackground('#1B365D')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  const regKeys = Object.keys(regionSummaryMap).sort((a, b) => regionSummaryMap[b].dtpk - regionSummaryMap[a].dtpk);

  let detailStartRow = 8;
  if (regKeys.length > 0) {
    regKeys.forEach((kv, idx) => {
      const rIdx = 6 + idx;
      const countVal = regionSummaryMap[kv].count;
      const dtpkVal = regionSummaryMap[kv].dtpk;
      const shareVal = totalDtpk > 0 ? (dtpkVal / totalDtpk) : 0;

      // Nối 2 cột A:B cho Khu vực
      destSheet.getRange(rIdx, 1, 1, 2).merge().setValue(' ' + kv).setHorizontalAlignment('left');
      destSheet.getRange(rIdx, 3).setValue(countVal).setHorizontalAlignment('center');
      destSheet.getRange(rIdx, 4).setValue(dtpkVal).setNumberFormat('#,##0').setHorizontalAlignment('right');
      destSheet.getRange(rIdx, 5).setValue(shareVal).setNumberFormat('0.0%').setHorizontalAlignment('right');
    });

    detailStartRow = 6 + regKeys.length + 2;
  }

  // =========================================================================
  // BẢNG 2: DANH SÁCH CHI TIẾT THUÊ BAO PHÁT TRIỂN MỚI (PTM)
  // =========================================================================
  destSheet.getRange(detailStartRow - 1, 1, 1, 18).merge()
    .setValue('2. DANH SÁCH CHI TIẾT THUÊ BAO PHÁT TRIỂN MỚI (PTM)')
    .setFontWeight('bold')
    .setBackground('#E8F0FE')
    .setFontColor('#174EA6');

  const detailHeaders = [
    'STT', 'MÃ KHÁCH HÀNG', 'TÊN THUÊ BAO', 'SĐT LIÊN HỆ', 'ĐỊA CHỈ LẮP ĐẶT', 'GÓI CƯỚC',
    'GIÁ CƯỚC', 'CHU KỲ', 'DTPK', 'NGÀY ĐẤU NỐI', 'NGÀY BẮT ĐẦU', 'LOẠI THIẾT BỊ ONT',
    'SERIAL ONT', 'KHU VỰC', 'MÃ CỬA HÀNG', 'MÃ NHÂN VIÊN', 'TÊN CỬA HÀNG', 'TÊN NHÂN VIÊN'
  ];
  destSheet.getRange(detailStartRow, 1, 1, 18).setValues([detailHeaders])
    .setFontWeight('bold')
    .setBackground('#00558F')
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  destSheet.setFrozenRows(detailStartRow);

  if (ptmFinalRows.length > 0) {
    destSheet.getRange(detailStartRow + 1, 1, ptmFinalRows.length, 18).setValues(ptmFinalRows);
    destSheet.getRange(detailStartRow + 1, 1, ptmFinalRows.length, 1).setHorizontalAlignment('center');
    destSheet.getRange(detailStartRow + 1, 2, ptmFinalRows.length, 1).setNumberFormat('@').setHorizontalAlignment('center');
    destSheet.getRange(detailStartRow + 1, 4, ptmFinalRows.length, 1).setNumberFormat('@').setHorizontalAlignment('center');
    destSheet.getRange(detailStartRow + 1, 7, ptmFinalRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');
    destSheet.getRange(detailStartRow + 1, 8, ptmFinalRows.length, 1).setHorizontalAlignment('center');
    destSheet.getRange(detailStartRow + 1, 9, ptmFinalRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');
    destSheet.getRange(detailStartRow + 1, 10, ptmFinalRows.length, 2).setHorizontalAlignment('center');
    destSheet.getRange(detailStartRow + 1, 13, ptmFinalRows.length, 1).setNumberFormat('@').setHorizontalAlignment('center');
    destSheet.getRange(detailStartRow + 1, 14, ptmFinalRows.length, 1).setHorizontalAlignment('center');
    destSheet.getRange(detailStartRow + 1, 15, ptmFinalRows.length, 2).setNumberFormat('@').setHorizontalAlignment('center');

    const exactWidths = [45, 115, 175, 110, 250, 110, 100, 70, 95, 140, 140, 130, 130, 130, 110, 110, 160, 160];
    for (let c = 0; c < exactWidths.length; c++) destSheet.setColumnWidth(c + 1, exactWidths[c]);
  }

  ss.setActiveSheet(destSheet);

  return {
    success: true,
    sheetName: sheetName,
    count: ptmRowsData.length,
    revenue: totalRevenue.toLocaleString('vi-VN'),
    dtpk: totalDtpk.toLocaleString('vi-VN'),
    downloadUrl: directExportUrl
  };
}

function getPtmDialogHtml() {
  const now = new Date();
  const curMonth = now.getMonth() + 1;
  const curYear = now.getFullYear();

  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; padding: 16px; margin: 0; background: #f8fafc; font-size: 13px; color: #1e293b; }
    .card { background: white; border: 1px solid #e2e8f0; border-radius: 10px; padding: 14px; margin-bottom: 12px; }
    .section-title { font-weight: bold; color: #1e40af; margin-bottom: 8px; font-size: 13px; }
    .radio-label { display: flex; align-items: center; gap: 8px; margin-bottom: 6px; cursor: pointer; }
    .inputs { margin-top: 6px; padding-left: 24px; }
    .hint { font-size: 11px; color: #64748b; margin-top: 4px; font-style: italic; }
    input[type="text"] { padding: 6px 10px; border: 1px solid #cbd5e1; border-radius: 6px; font-size: 13px; outline: none; }
    input[type="text"]:focus { border-color: #2563eb; }
    .btn { width: 100%; border: none; padding: 10px; font-size: 13px; font-weight: 600; border-radius: 8px; cursor: pointer; text-align: center; display: block; box-sizing: border-box; }
    .btn-primary { background: #2563eb; color: white; margin-top: 8px; }
    .btn-primary:hover { background: #1d4ed8; }
    .btn-primary:disabled { background: #94a3b8; cursor: not-allowed; }
    .btn-download { background: #16a34a; color: white; margin-top: 10px; text-decoration: none; }
    .btn-download:hover { background: #15803d; }
    .btn-close { background: #f1f5f9; color: #334155; border: 1px solid #cbd5e1; margin-top: 8px; }
    .btn-close:hover { background: #e2e8f0; }
    .alert-banner { display: none; background: #fee2e2; border: 1px solid #fecaca; color: #b91c1c; padding: 8px 12px; border-radius: 6px; font-size: 12px; margin-bottom: 10px; }
    .status { font-size: 12px; margin-top: 8px; padding: 8px; border-radius: 6px; display: none; text-align: center; }
    .status.info { background: #dbeafe; color: #1e40af; display: block; }
    #successView { display: none; text-align: center; padding: 6px 4px; }
    .success-title { font-size: 15px; font-weight: bold; color: #15803d; margin-bottom: 8px; }
    .kpi-card { background: white; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px; margin-bottom: 6px; text-align: left; }
  </style>
</head>
<body>
  <div id="formView">
    <div id="errorBanner" class="alert-banner"></div>
    <div class="card">
      <div class="section-title">Chọn phạm vi thời gian PTM:</div>
      <label class="radio-label">
        <input type="radio" name="timeType" value="month" checked onchange="toggleType()">
        <span>Theo Tháng</span>
      </label>
      <div class="inputs" id="monthBox">
        <span>Tháng:</span>
        <input type="text" id="monthVal" value="${curMonth < 10 ? '0' + curMonth : curMonth}" style="width: 45px; text-align: center;"> /
        <span>Năm:</span>
        <input type="text" id="yearVal" value="${curYear}" style="width: 65px; text-align: center;">
        <div class="hint">💡 Lọc hợp đồng mới thỏa điều kiện PTM trong tháng.</div>
      </div>

      <label class="radio-label" style="margin-top: 10px;">
        <input type="radio" name="timeType" value="range" onchange="toggleType()">
        <span>Từ ngày... đến ngày...</span>
      </label>
      <div class="inputs" id="rangeBox" style="opacity: 0.5; pointer-events: none;">
        <div style="display: flex; gap: 8px; align-items: center; margin-top: 4px;">
          <span style="width: 65px;">Từ ngày:</span>
          <input type="text" id="fromDate" placeholder="01/10/${curYear}" style="width: 110px;">
        </div>
        <div style="display: flex; gap: 8px; align-items: center; margin-top: 6px;">
          <span style="width: 65px;">Đến ngày:</span>
          <input type="text" id="toDate" placeholder="31/10/${curYear}" style="width: 110px;">
        </div>
      </div>
    </div>

    <button id="btnRun" class="btn btn-primary" onclick="runReport()">🚀 Kết xuất Báo cáo PTM</button>
    <div id="statusBox" class="status"></div>
  </div>

  <div id="successView">
    <div class="success-title">🎉 KẾT XUẤT PTM THÀNH CÔNG!</div>
    <div class="kpi-card">
      <div style="font-weight: bold; color: #1e40af;" id="resSheetName">Trang tính: ...</div>
      <div style="margin-top: 6px; font-size: 13px; line-height: 1.7;">
        🎯 <b>Thuê bao PTM:</b> <span id="resCount" style="color: #b91c1c; font-weight: bold;">0</span> TB<br>
        💰 <b>Tổng doanh thu:</b> <span id="resRevenue" style="color: #15803d; font-weight: bold;">0</span> VNĐ<br>
        📈 <b>Tổng DTPK:</b> <span id="resDtpk" style="color: #b45309; font-weight: bold;">0</span> VNĐ
      </div>
    </div>

    <!-- NÚT TẢI FILE EXCEL TRÊN DIALOG ĐÃ KHÔI PHỤC -->
    <button id="btnDownload" class="btn btn-download" onclick="downloadExcel()">📥 Tải file Excel (.xlsx) về máy</button>
    <button class="btn btn-close" onclick="google.script.host.close()">Đóng cửa sổ</button>
  </div>

  <!-- IFRAME PHỤC VỤ DOWNLOAD SẠCH -->
  <iframe id="downloadFrame" style="display:none;"></iframe>

  <script>
    let currentDownloadUrl = '';

    function showError(msg) {
      const b = document.getElementById('errorBanner');
      b.innerText = '⚠️ ' + msg;
      b.style.display = 'block';
    }
    function hideError() {
      document.getElementById('errorBanner').style.display = 'none';
    }
    function toggleType() {
      hideError();
      const isMonth = document.querySelector('input[name="timeType"]:checked').value === 'month';
      document.getElementById('monthBox').style.opacity = isMonth ? '1' : '0.5';
      document.getElementById('monthBox').style.pointerEvents = isMonth ? 'auto' : 'none';
      document.getElementById('rangeBox').style.opacity = isMonth ? '0.5' : '1';
      document.getElementById('rangeBox').style.pointerEvents = isMonth ? 'none' : 'auto';
    }

    function runReport() {
      hideError();
      const timeType = document.querySelector('input[name="timeType"]:checked').value;
      const btn = document.getElementById('btnRun');
      const status = document.getElementById('statusBox');

      let params = { type: timeType };
      if (timeType === 'month') {
        const m = document.getElementById('monthVal').value.trim();
        const y = document.getElementById('yearVal').value.trim();
        if (!m || !y) { showError('Vui lòng nhập đủ Tháng và Năm!'); return; }
        params.month = m;
        params.year = y;
      } else {
        const f = document.getElementById('fromDate').value.trim();
        const t = document.getElementById('toDate').value.trim();
        if (!f || !t) { showError('Vui lòng nhập đủ Từ ngày và Đến ngày!'); return; }
        params.from = f;
        params.to = t;
      }

      btn.disabled = true;
      btn.innerText = '⏳ Đang quét dữ liệu PTM...';
      status.className = 'status info';
      status.innerText = 'Hệ thống đang trích xuất và sắp xếp 4 cấp...';

      google.script.run
        .withSuccessHandler(function(res) {
          currentDownloadUrl = res.downloadUrl;
          document.getElementById('formView').style.display = 'none';
          document.getElementById('resSheetName').innerText = '📁 Sheet: ' + res.sheetName;
          document.getElementById('resCount').innerText = res.count;
          document.getElementById('resRevenue').innerText = res.revenue;
          document.getElementById('resDtpk').innerText = res.dtpk;
          document.getElementById('successView').style.display = 'block';
        })
        .withFailureHandler(function(err) {
          btn.disabled = false;
          btn.innerText = '🚀 Kết xuất Báo cáo PTM';
          status.style.display = 'none';
          showError(err.message);
        })
        .processPtmReport(params);
    }

    function downloadExcel() {
      if (currentDownloadUrl) {
        document.getElementById('downloadFrame').src = currentDownloadUrl;
      } else {
        alert('Không tìm thấy đường link tải file!');
      }
    }
  </script>
</body>
</html>
  `;
}