/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * MODULE       : Mod_DoanhThuPhanBo.gs
 * PHIÊN BẢN    : 2.2 (Khôi phục tùy chọn Khoảng ngày, Fix lỗi parseCurrency & Nút Tải Excel)
 * TÁC GIẢ      : Trịnh Vũ Thế Vinh & AI Assistant
 * ============================================================================
 */

function showRevenueAllocationDialog() {
  const html = HtmlService.createHtmlOutput(getRevenueAllocationDialogHtml())
    .setWidth(500)
    .setHeight(480)
    .setTitle('Báo cáo Doanh thu Phân bổ (DTPK)');
  SpreadsheetApp.getUi().showModalDialog(html, 'Báo cáo Doanh thu Phân bổ MobiFiber');
}

/**
 * XỬ LÝ KẾT XUẤT BẢNG TỔNG HỢP DOANH THU PHÂN BỔ (DTPK)
 */
function processRevenueAllocationSummary(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) throw new Error('Bảng raw_mobifiber chưa có dữ liệu.');

  const mapping = loadThamSoMapping();

  let startDate, endDate, timeTitle, timeCode, summarySheetName;
  if (params.type === 'month') {
    const month = parseInt(params.month, 10);
    const year = parseInt(params.year, 10);
    const mStr = month < 10 ? '0' + month : '' + month;
    startDate = new Date(year, month - 1, 1, 0, 0, 0);
    endDate = new Date(year, month, 0, 23, 59, 59);
    timeTitle = `THÁNG ${mStr}/${year}`;
    timeCode = `T${mStr}_${year}`;
    summarySheetName = `DTPB_${timeCode}`;
  } else {
    const startParts = params.from.split('/');
    const endParts = params.to.split('/');
    startDate = new Date(parseInt(startParts[2], 10), parseInt(startParts[1], 10) - 1, parseInt(startParts[0], 10), 0, 0, 0);
    endDate = new Date(parseInt(endParts[2], 10), parseInt(endParts[1], 10) - 1, parseInt(endParts[0], 10), 23, 59, 59);
    const d1 = startParts[0] + startParts[1];
    const d2 = endParts[0] + endParts[1];
    timeTitle = `TỪ ${params.from} ĐẾN ${params.to}`;
    timeCode = `${d1}_${d2}_${endParts[2]}`;
    summarySheetName = `DTPB_${timeCode}`;
  }

  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol)
    .getValues()[0].map(h => normalizeHeader(h));

  const idxMaKH = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxLoaiKH = headers.findIndex(h => h === 'loại khách hàng');
  const idxGiaCuoc = headers.findIndex(h => h === 'giá cước');
  const idxChuKy = headers.findIndex(h => h === 'chu kỳ sử dụng');
  const idxDtpk = headers.findIndex(h => h === 'dtpk');
  const idxNgayDauNoi = headers.findIndex(h => h === 'ngày đấu nối');
  const idxNgayThiCong = headers.findIndex(h => h === 'ngày hoàn tất thi công');
  const idxNgayBatDau = headers.findIndex(h => h === 'ngày băt đầu gói cước' || h === 'ngày bắt đầu gói cước');
  const idxNgayKetThuc = headers.findIndex(h => h === 'ngày kết thúc gói cước');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxChanCat = headers.findIndex(h => h === 'tình trạng chặn cắt');
  const idxNhuCau = headers.findIndex(h => h === 'nhu cầu hòa mạng');
  const idxMaNV = headers.findIndex(h => h === 'mã nhân viên');

  const numRows = lastRow - CONFIG.DATA_START_ROW_TARGET + 1;
  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, numRows, lastCol).getValues();

  const regionMap = {};
  let grandTotalOldTB = 0;
  let grandTotalOldDtpk = 0;
  let grandTotalNewTB = 0;
  let grandTotalNewRevenue = 0;

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];

    const loaiKH = normalizeHeader(row[idxLoaiKH]);
    if (loaiKH === 'thuê bao nghiệp vụ' || loaiKH === 'thue bao nghiep vu') continue;

    const tinhTrangHD = normalizeHeader(row[idxTinhTrangHD]);
    if (tinhTrangHD !== 'đang hoạt động' && tinhTrangHD !== 'dang hoat dong') continue;

    const chanCat = normalizeHeader(row[idxChanCat]);
    if (chanCat !== 'bình thường' && chanCat !== 'binh thuong' && chanCat !== '') continue;

    const chuKyRaw = row[idxChuKy];
    const chuKy = chuKyRaw ? parseInt(chuKyRaw, 10) : 0;
    const giaCuoc = parseCurrencyToInteger(row[idxGiaCuoc]) || 0;
    const dtpk = parseCurrencyToInteger(row[idxDtpk]) || (chuKy > 0 ? Math.round(giaCuoc / chuKy) : 0);

    const dDauNoi = parseToDate(row[idxNgayDauNoi]);
    const dThiCong = parseToDate(row[idxNgayThiCong]);
    const dBatDau = parseToDate(row[idxNgayBatDau]);
    const dKetThuc = parseToDate(row[idxNgayKetThuc]);
    const nhuCau = normalizeHeader(row[idxNhuCau]);

    let isPtmInPeriod = false;
    if (nhuCau === 'đăng ký mới' || nhuCau === 'dang ky moi') {
      let ptmDate = null;
      if (chuKy === 1) {
        if (dDauNoi && dThiCong) {
          ptmDate = (dThiCong.getTime() > dDauNoi.getTime()) ? dThiCong : dDauNoi;
        } else {
          ptmDate = dThiCong || dDauNoi || dBatDau;
        }
      } else {
        ptmDate = dBatDau;
      }

      if (ptmDate && ptmDate >= startDate && ptmDate <= endDate) {
        isPtmInPeriod = true;
      }
    }

    let isRetainedInPeriod = false;
    if (!isPtmInPeriod) {
      if (chuKy === 1) {
        isRetainedInPeriod = true;
      } else {
        if (dKetThuc && dKetThuc >= startDate) {
          isRetainedInPeriod = true;
        }
      }
    }

    if (!isPtmInPeriod && !isRetainedInPeriod) continue;

    const rawMaNV = String(row[idxMaNV] || '').trim();
    let khuVuc = 'Chưa phân bổ KV';
    if (rawMaNV && mapping.staffRegion && mapping.staffRegion.has(normalizeHeader(rawMaNV))) {
      const staffInfo = mapping.staffRegion.get(normalizeHeader(rawMaNV));
      if (staffInfo.khuVuc && !staffInfo.khuVuc.includes('Chưa phân bổ')) {
        khuVuc = staffInfo.khuVuc;
      }
    }

    if (!regionMap[khuVuc]) {
      regionMap[khuVuc] = { oldTB: 0, oldDtpk: 0, newTB: 0, newRevenue: 0 };
    }

    if (isPtmInPeriod) {
      regionMap[khuVuc].newTB += 1;
      regionMap[khuVuc].newRevenue += dtpk;
      grandTotalNewTB += 1;
      grandTotalNewRevenue += dtpk;
    } else if (isRetainedInPeriod) {
      regionMap[khuVuc].oldTB += 1;
      regionMap[khuVuc].oldDtpk += dtpk;
      grandTotalOldTB += 1;
      grandTotalOldDtpk += dtpk;
    }
  }

  const grandTotalDtpk = grandTotalOldDtpk + grandTotalNewRevenue;
  const grandTotalAllTB = grandTotalOldTB + grandTotalNewTB;

  // Tạo Sheet Tổng hợp
  let destSheet = ss.getSheetByName(summarySheetName);
  if (!destSheet) destSheet = ss.insertSheet(summarySheetName);
  else destSheet.clear();

  const ssId = ss.getId();
  const sheetId = destSheet.getSheetId();
  const directExportUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=xlsx&gid=${sheetId}`;

  // Tiêu đề & KPI
  destSheet.getRange('A1:H1').merge()
    .setValue(`  BÁO CÁO DOANH THU PHÂN BỔ MOBIFIBER - ${timeTitle}`)
    .setFontWeight('bold').setFontSize(13).setBackground('#1B365D').setFontColor('#FFFFFF')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');

  destSheet.getRange('A2:E2').merge()
    .setValue(`  🎯 TỔNG DTPK TOÀN TỈNH: ${grandTotalDtpk.toLocaleString('vi-VN')} VNĐ`)
    .setFontWeight('bold').setFontSize(11).setBackground('#D2E3FC').setFontColor('#174EA6')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');

  destSheet.getRange('F2:H2').merge()
    .setValue(`🕒 Ngày lập: ${Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm')}`)
    .setFontSize(10).setBackground('#D2E3FC').setFontColor('#5F6368')
    .setHorizontalAlignment('right').setVerticalAlignment('middle');

  destSheet.getRange('A3:H3').merge()
    .setValue(`  📊 TỔNG THUÊ BAO: ${grandTotalAllTB} TB (Hiện hữu: ${grandTotalOldTB} TB | PTM mới: ${grandTotalNewTB} TB)`)
    .setFontWeight('bold').setFontSize(11).setBackground('#E8F0FE').setFontColor('#174EA6')
    .setHorizontalAlignment('left').setVerticalAlignment('middle');

  // Nút Thao tác trên Sheet
  destSheet.getRange('A4:D4').merge()
    .setFormula(`=HYPERLINK("#gid=${sheetId}&action=detail"; "📋 KẾT XUẤT DANH SÁCH CHI TIẾT")`)
    .setFontWeight('bold').setFontSize(11).setBackground('#00558F').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  destSheet.getRange('E4:H4').merge()
    .setFormula(`=HYPERLINK("${directExportUrl}"; "📥 TẢI FILE EXCEL TỔNG HỢP")`)
    .setFontWeight('bold').setFontSize(11).setBackground('#1A73E8').setFontColor('#FFFFFF')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');

  const tableHeaders = [
    'KHU VỰC', 'TB HIỆN HỮU', 'DTPK HIỆN HỮU', 'TB PTM', 'DOANH THU PTM', 'TỔNG SỐ TB', 'TỔNG DTPK (VNĐ)', 'TỶ TRỌNG (%)'
  ];

  destSheet.getRange(6, 1, 1, tableHeaders.length).setValues([tableHeaders])
    .setFontWeight('bold').setBackground('#1B365D').setFontColor('#FFFFFF').setHorizontalAlignment('center').setVerticalAlignment('middle');
  destSheet.setFrozenRows(6);

  const regionRows = Object.keys(regionMap).map(kv => {
    const d = regionMap[kv];
    const totalTbKv = d.oldTB + d.newTB;
    const totalDtpkKv = d.oldDtpk + d.newRevenue;
    const share = grandTotalDtpk > 0 ? (totalDtpkKv / grandTotalDtpk) : 0;
    return [
      ' ' + kv,
      d.oldTB,
      d.oldDtpk,
      d.newTB,
      d.newRevenue,
      totalTbKv,
      totalDtpkKv,
      share
    ];
  }).sort((a, b) => b[6] - a[6]);

  let curRow = 7;
  if (regionRows.length > 0) {
    destSheet.getRange(curRow, 1, regionRows.length, tableHeaders.length).setValues(regionRows);
    destSheet.getRange(curRow, 1, regionRows.length, 1).setHorizontalAlignment('left');
    destSheet.getRange(curRow, 2, regionRows.length, 1).setHorizontalAlignment('center');
    destSheet.getRange(curRow, 3, regionRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');
    destSheet.getRange(curRow, 4, regionRows.length, 1).setHorizontalAlignment('center');
    destSheet.getRange(curRow, 5, regionRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');
    destSheet.getRange(curRow, 6, regionRows.length, 1).setHorizontalAlignment('center');
    destSheet.getRange(curRow, 7, regionRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');
    destSheet.getRange(curRow, 8, regionRows.length, 1).setNumberFormat('0.0%').setHorizontalAlignment('right');
    curRow += regionRows.length;
  }

  // Dòng Tổng Cộng
  const summaryTotalRow = [
    ' TỔNG CỘNG TOÀN TỈNH',
    grandTotalOldTB,
    grandTotalOldDtpk,
    grandTotalNewTB,
    grandTotalNewRevenue,
    grandTotalAllTB,
    grandTotalDtpk,
    1.0
  ];

  destSheet.getRange(curRow, 1, 1, tableHeaders.length).setValues([summaryTotalRow])
    .setFontWeight('bold').setBackground('#E8F0FE').setFontColor('#174EA6');
  destSheet.getRange(curRow, 1).setHorizontalAlignment('left');
  destSheet.getRange(curRow, 2).setHorizontalAlignment('center');
  destSheet.getRange(curRow, 3).setNumberFormat('#,##0').setHorizontalAlignment('right');
  destSheet.getRange(curRow, 4).setHorizontalAlignment('center');
  destSheet.getRange(curRow, 5).setNumberFormat('#,##0').setHorizontalAlignment('right');
  destSheet.getRange(curRow, 6).setHorizontalAlignment('center');
  destSheet.getRange(curRow, 7).setNumberFormat('#,##0').setHorizontalAlignment('right');
  destSheet.getRange(curRow, 8).setNumberFormat('0.0%').setHorizontalAlignment('right');

  const exactWidths = [160, 110, 130, 100, 130, 110, 140, 124];
  for (let c = 0; c < exactWidths.length; c++) destSheet.setColumnWidth(c + 1, exactWidths[c]);

  ss.setActiveSheet(destSheet);

  return {
    success: true,
    sheetName: summarySheetName,
    timeCode: timeCode,
    totalDtpk: grandTotalDtpk.toLocaleString('vi-VN'),
    totalTB: grandTotalAllTB,
    oldTB: grandTotalOldTB,
    newTB: grandTotalNewTB,
    downloadUrl: directExportUrl
  };
}

/**
 * GIAO DIỆN DIALOG DOANH THU PHÂN BỔ (DTPK)
 */
function getRevenueAllocationDialogHtml() {
  const now = new Date();
  const curMonth = now.getMonth() + 1;
  const curYear = now.getFullYear();

  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; padding: 14px; margin: 0; background: #f8fafc; font-size: 13px; color: #1e293b; }
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
      <div class="section-title">Chọn kỳ phân bổ doanh thu:</div>
      
      <!-- LỰA CHỌN 1: THEO THÁNG (KHÔI PHỤC ĐẦY ĐỦ) -->
      <label class="radio-label">
        <input type="radio" name="reportType" value="month" checked onchange="toggleType()">
        <span>Theo Tháng phân bổ (MM/YYYY)</span>
      </label>
      <div class="inputs" id="monthBox">
        <span>Tháng:</span>
        <input type="text" id="monthVal" value="${curMonth < 10 ? '0' + curMonth : curMonth}" style="width: 45px; text-align: center;"> /
        <span>Năm:</span>
        <input type="text" id="yearVal" value="${curYear}" style="width: 65px; text-align: center;">
        <div class="hint">💡 Báo cáo sẽ tổng hợp TB Hiện hữu, PTM và tính tỷ trọng DTPK theo từng Khu vực.</div>
      </div>

      <!-- LỰA CHỌN 2: TỪ NGÀY ĐẾN NGÀY (KHÔI PHỤC ĐẦY ĐỦ) -->
      <label class="radio-label" style="margin-top: 10px;">
        <input type="radio" name="reportType" value="range" onchange="toggleType()">
        <span>Theo Khoảng ngày tùy chọn</span>
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

    <button id="btnRun" class="btn btn-primary" onclick="runReport()">🚀 Khởi tạo báo cáo DTPK</button>
    <div id="statusBox" class="status"></div>
  </div>

  <div id="successView">
    <div class="success-title">🎉 KHỞI TẠO BÁO CÁO THÀNH CÔNG!</div>
    <div class="kpi-card">
      <div style="font-weight: bold; color: #1e40af;" id="resSheetName">Bảng tổng hợp: ...</div>
      <div style="margin-top: 6px; font-size: 13px; line-height: 1.7;">
        📈 <b>Tổng DTPK toàn tỉnh:</b> <span id="resDtpk" style="color: #b45309; font-weight: bold;">0</span> VNĐ<br>
        📊 <b>Tổng số thuê bao:</b> <span id="resTotalTB" style="color: #1e40af; font-weight: bold;">0</span> TB<br>
        🏛️ <b>Thuê bao Hiện hữu:</b> <span id="resOldTB" style="color: #334155; font-weight: bold;">0</span> TB<br>
        🚀 <b>Thuê bao PTM mới:</b> <span id="resNewTB" style="color: #15803d; font-weight: bold;">0</span> TB
      </div>
    </div>

    <!-- NÚT TẢI FILE EXCEL TRỰC TIẾP TRÊN DIALOG -->
    <button id="btnDownload" class="btn btn-download" onclick="downloadExcel()">📥 Tải file Excel (.xlsx) về máy</button>
    <button class="btn btn-close" onclick="google.script.host.close()">Đóng cửa sổ</button>
  </div>

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
      const isMonth = document.querySelector('input[name="reportType"]:checked').value === 'month';
      document.getElementById('monthBox').style.opacity = isMonth ? '1' : '0.5';
      document.getElementById('monthBox').style.pointerEvents = isMonth ? 'auto' : 'none';
      document.getElementById('rangeBox').style.opacity = isMonth ? '0.5' : '1';
      document.getElementById('rangeBox').style.pointerEvents = isMonth ? 'none' : 'auto';
    }

    function runReport() {
      hideError();
      const type = document.querySelector('input[name="reportType"]:checked').value;
      const btn = document.getElementById('btnRun');
      const status = document.getElementById('statusBox');

      let params = { type: type };
      if (type === 'month') {
        const m = document.getElementById('monthVal').value.trim();
        const y = document.getElementById('yearVal').value.trim();
        if (!m || !y) { showError('Vui lòng nhập đầy đủ Tháng và Năm!'); return; }
        params.month = m;
        params.year = y;
      } else {
        const f = document.getElementById('fromDate').value.trim();
        const t = document.getElementById('toDate').value.trim();
        if (!f || !t) { showError('Vui lòng nhập đầy đủ Từ ngày và Đến ngày!'); return; }
        params.from = f;
        params.to = t;
      }

      btn.disabled = true;
      btn.innerText = '⏳ Đang tính toán doanh thu...';
      status.className = 'status info';
      status.innerText = 'Hệ thống đang tổng hợp DTPK theo từng Khu vực...';

      google.script.run
        .withSuccessHandler(function(res) {
          currentDownloadUrl = res.downloadUrl;
          document.getElementById('formView').style.display = 'none';
          document.getElementById('resSheetName').innerText = '📁 Bảng tổng hợp: ' + res.sheetName;
          document.getElementById('resDtpk').innerText = res.totalDtpk;
          document.getElementById('resTotalTB').innerText = res.totalTB;
          document.getElementById('resOldTB').innerText = res.oldTB;
          document.getElementById('resNewTB').innerText = res.newTB;
          document.getElementById('successView').style.display = 'block';
        })
        .withFailureHandler(function(err) {
          btn.disabled = false;
          btn.innerText = '🚀 Khởi tạo báo cáo DTPK';
          status.style.display = 'none';
          showError(err.message);
        })
        .processRevenueAllocationSummary(params);
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

/**
 * TIỆN ÍCH TỰ ĐỘNG TÌM VÀ BẬT BẢNG CẢNH BÁO 2 THUÊ BAO CHƯA PHÂN BỔ KV
 */
function findUnmappedSubscribers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  const tsSheet = ss.getSheetByName('thamso_caidat');
  const ui = SpreadsheetApp.getUi();

  if (!rawSheet || !tsSheet) return;

  const mapping = loadThamSoMapping();
  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));

  const idxMaKH = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxTenTB = headers.findIndex(h => h === 'tên thuê bao');
  const idxMaNV = headers.findIndex(h => h === 'mã nhân viên');
  const idxTenNV = headers.findIndex(h => h === 'tên nhân viên');
  const idxMaCH = headers.findIndex(h => h === 'mã cửa hàng');
  const idxTenCH = headers.findIndex(h => h === 'tên cửa hàng');
  const idxLoaiKH = headers.findIndex(h => h === 'loại khách hàng');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');

  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, lastRow - CONFIG.DATA_START_ROW_TARGET + 1, lastCol).getValues();
  const unmappedList = [];

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const loaiKH = normalizeHeader(row[idxLoaiKH]);
    if (loaiKH === 'thuê bao nghiệp vụ' || loaiKH === 'thue bao nghiep vu') continue;
    const tinhTrangHD = normalizeHeader(row[idxTinhTrangHD]);
    if (tinhTrangHD !== 'đang hoạt động' && tinhTrangHD !== 'dang hoat dong') continue;

    const rawMaNV = String(row[idxMaNV] || '').trim();
    let khuVuc = 'Chưa phân bổ KV';
    if (rawMaNV && mapping.staffRegion && mapping.staffRegion.has(normalizeHeader(rawMaNV))) {
      const staffInfo = mapping.staffRegion.get(normalizeHeader(rawMaNV));
      if (staffInfo.khuVuc && !staffInfo.khuVuc.includes('Chưa phân bổ')) {
        khuVuc = staffInfo.khuVuc;
      }
    }

    if (khuVuc === 'Chưa phân bổ KV') {
      unmappedList.push({
        row: CONFIG.DATA_START_ROW_TARGET + i,
        maKH: row[idxMaKH],
        tenTB: row[idxTenTB],
        maNV: rawMaNV || '(Trống Mã NV)',
        tenNV: row[idxTenNV] || '(Trống)',
        tenCH: row[idxTenCH] || '(Trống)'
      });
    }
  }

  if (unmappedList.length === 0) {
    ui.alert('✅ Tất cả thuê bao đang hoạt động đều đã được gán Khu vực đầy đủ!');
  } else {
    let msg = `⚠️ Phát hiện ${unmappedList.length} thuê bao chưa được phân bổ Khu vực:\n\n`;
    unmappedList.forEach((item, idx) => {
      msg += `${idx + 1}. Dòng ${item.row}: Mã KH [${item.maKH}] - ${item.tenTB}\n   👉 Mã NV: [${item.maNV}] | Tên NV: ${item.tenNV} | CH: ${item.tenCH}\n\n`;
    });
    msg += `Vui lòng vào sheet "thamso_caidat" để bổ sung Khu vực cho các Mã NV trên, sau đó chạy lại báo cáo.`;
    ui.alert(msg);
  }
}