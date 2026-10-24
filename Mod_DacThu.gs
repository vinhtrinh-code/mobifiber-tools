/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * FILE        : Mod_DacThu.gs
 * PHIÊN BẢN   : 3.0 (High-Performance Engine: In-Memory Filtering & Matrix Formatting)
 * NGÀY TẠO    : 20/09/2026 15:00:00 (GMT+7)
 * NGÀY CẬP NHẬT: 24/10/2026 19:50:00 (GMT+7)
 * TÁC GIẢ     : Trịnh Vũ Thế Vinh & AI Assistant
 * MÔ TẢ       : Kết xuất báo cáo thuê bao đặc thù (Thuê bao Nghiệp vụ & BCA / Vũ trang)
 * ============================================================================
 */

function showSpecialReportDialog() {
  const todayStr = Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy');
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 18px; margin: 0; font-size: 13px; color: #202124; background: #fff; }
        .form-group { margin-bottom: 12px; }
        label { display: block; font-weight: 600; margin-bottom: 5px; color: #3c4043; }
        select, input[type="text"] { width: 100%; box-sizing: border-box; padding: 8px 10px; border: 1px solid #dadce0; border-radius: 6px; font-size: 13px; outline: none; }
        select:focus, input[type="text"]:focus { border-color: #1a73e8; box-shadow: 0 0 0 2px rgba(26,115,232,0.2); }
        .radio-group { margin-bottom: 12px; }
        .radio-item { margin-bottom: 6px; display: flex; align-items: center; gap: 6px; }
        .date-range-box { display: none; gap: 8px; margin-top: 6px; }
        .btn-container { display: flex; justify-content: flex-end; gap: 8px; margin-top: 16px; }
        button { padding: 8px 16px; border-radius: 6px; font-weight: 600; cursor: pointer; border: none; font-size: 13px; }
        .btn-cancel { background: #f1f3f4; color: #3c4043; }
        .btn-submit { background: #1a73e8; color: #fff; }
        .btn-submit:disabled { background: #dadce0; cursor: not-allowed; }
      </style>
    </head>
    <body>
      <div class="form-group">
        <label>Nhóm đối tượng đặc thù:</label>
        <select id="targetGroup">
          <option value="all">Tất cả (Thuê bao Nghiệp vụ & BCA/Vũ trang)</option>
          <option value="nv">Chỉ Thuê bao Nghiệp vụ</option>
          <option value="bca">Chỉ Thuê bao BCA / Lực lượng vũ trang</option>
        </select>
      </div>

      <div class="radio-group">
        <label>Thời gian ghi nhận:</label>
        <div class="radio-item">
          <input type="radio" id="timeAll" name="timeType" value="all" checked onchange="toggleTimeBox()">
          <label for="timeAll" style="font-weight:normal; margin:0;">Toàn bộ lũy kế đang hoạt động</label>
        </div>
        <div class="radio-item">
          <input type="radio" id="timeRange" name="timeType" value="range" onchange="toggleTimeBox()">
          <label for="timeRange" style="font-weight:normal; margin:0;">Theo khoảng thời gian phát triển</label>
        </div>
      </div>

      <div class="date-range-box" id="rangeBox">
        <div style="flex: 1;">
          <label>Từ ngày (dd/mm/yyyy):</label>
          <input type="text" id="fromVal" value="${todayStr}">
        </div>
        <div style="flex: 1;">
          <label>Đến ngày (dd/mm/yyyy):</label>
          <input type="text" id="toVal" value="${todayStr}">
        </div>
      </div>

      <div class="btn-container">
        <button class="btn-cancel" onclick="google.script.host.close()">Hủy</button>
        <button class="btn-submit" id="btnRun" onclick="submitSpecial()">Kết xuất báo cáo</button>
      </div>

      <script>
        function toggleTimeBox() {
          const isRange = document.getElementById('timeRange').checked;
          document.getElementById('rangeBox').style.display = isRange ? 'flex' : 'none';
        }

        function submitSpecial() {
          const group = document.getElementById('targetGroup').value;
          const isRange = document.getElementById('timeRange').checked;
          const params = { targetGroup: group, timeType: isRange ? 'range' : 'all' };

          if (isRange) {
            params.from = document.getElementById('fromVal').value.trim();
            params.to = document.getElementById('toVal').value.trim();
            if (!params.from || !params.to) {
              alert('Vui lòng nhập đủ cả 2 ngày!');
              return;
            }
          }

          document.getElementById('btnRun').disabled = true;
          document.getElementById('btnRun').innerText = 'Đang xử lý...';

          google.script.run
            .withSuccessHandler(function() {
              google.script.host.close();
            })
            .withFailureHandler(function(err) {
              alert('Lỗi: ' + err.message);
              document.getElementById('btnRun').disabled = false;
              document.getElementById('btnRun').innerText = 'Kết xuất báo cáo';
            })
            .generateSpecialReport(params);
        }
      </script>
    </body>
    </html>
  `;
  const ui = HtmlService.createHtmlOutput(html).setWidth(420).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(ui, '🛡️ Báo cáo Thuê bao Nghiệp vụ & BCA');
}

/**
 * XỬ LÝ LỌC & KẾT XUẤT BÁO CÁO ĐẶC THÙ TRÊN BỘ NHỚ
 */
function generateSpecialReport(params) {
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);

  // 1. XÁC ĐỊNH PHẠM VI THỜI GIAN VÀ TÊN SHEET
  const timeType = params.timeType;
  const targetGroup = params.targetGroup;
  let startDate = null, endDate = null, timeTitle = '', timeCode = '';

  if (timeType === 'range') {
    const pStart = params.from.split('/');
    const pEnd = params.to.split('/');
    startDate = new Date(parseInt(pStart[2], 10), parseInt(pStart[1], 10) - 1, parseInt(pStart[0], 10), 0, 0, 0);
    endDate = new Date(parseInt(pEnd[2], 10), parseInt(pEnd[1], 10) - 1, parseInt(pEnd[0], 10), 23, 59, 59);
    timeTitle = `TỪ ${params.from} ĐẾN ${params.to}`;
    timeCode = `${pStart[0]}${pStart[1]}_${pEnd[0]}${pEnd[1]}_${pEnd[2]}`;
  } else {
    timeTitle = 'TOÀN BỘ DANH SÁCH LŨY KẾ';
    timeCode = 'LuyKe';
  }

  let groupTitle = 'NGHIỆP VỤ & BCA';
  if (targetGroup === 'nv') groupTitle = 'THUÊ BAO NGHIỆP VỤ';
  if (targetGroup === 'bca') groupTitle = 'THUÊ BAO BCA / VŨ TRANG';

  const sheetName = `DacThu_${targetGroup.toUpperCase()}_${timeCode}`;

  // 2. ĐỌC DỮ LIỆU ĐÍCH VÀ THAM SỐ VÀO RAM
  const lastCol = rawSheet.getLastColumn();
  const lastRow = rawSheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) throw new Error('Sheet raw_mobifiber chưa có dữ liệu!');

  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));
  const idxMaKH = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxLoaiKH = headers.findIndex(h => h === 'loại khách hàng');
  const idxTenTB = headers.findIndex(h => h === 'tên thuê bao');
  const idxSDT = headers.findIndex(h => h === 'sđt liên hệ');
  const idxDiaChi = headers.findIndex(h => h === 'địa chỉ lắp đặt');
  const idxGoiCuoc = headers.findIndex(h => h === 'gói cước');
  const idxGiaCuoc = headers.findIndex(h => h === 'giá cước');
  const idxChuKy = headers.findIndex(h => h === 'chu kỳ sử dụng');
  const idxDtpk = headers.findIndex(h => h === 'dtpk');
  const idxThietBi = headers.findIndex(h => h === 'loại thiết bị ont');
  const idxSerial = headers.findIndex(h => h === 'serial ont');
  const idxNgayDauNoi = headers.findIndex(h => h === 'ngày đấu nối');
  const idxNgayBatDau = headers.findIndex(h => h === 'ngày băt đầu gói cước' || h === 'ngày bắt đầu gói cước');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxMaCH = headers.findIndex(h => h === 'mã cửa hàng');
  const idxTenCH = headers.findIndex(h => h === 'tên cửa hàng');
  const idxMaNV = headers.findIndex(h => h === 'mã nhân viên');
  const idxTenNV = headers.findIndex(h => h === 'tên nhân viên');

  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, lastRow - CONFIG.DATA_START_ROW_TARGET + 1, lastCol).getValues();

  // Nạp danh mục Khu vực từ thamso_caidat
  const staffRegionMap = new Map();
  const thamSoSheet = ss.getSheetByName('thamso_caidat');
  if (thamSoSheet && thamSoSheet.getLastRow() >= 4) {
    const tsData = thamSoSheet.getRange(4, 10, thamSoSheet.getLastRow() - 3, 4).getValues();
    tsData.forEach(r => {
      const kv = String(r[0] || '').trim();
      const maNV = normalizeHeader(String(r[2] || ''));
      if (maNV && kv) staffRegionMap.set(maNV, kv);
    });
  }

  // 3. LỌC VÀ THỐNG KÊ TRÊN RAM
  const specRowsData = [];
  const regionSummaryMap = new Map();
  let totalRevenue = 0;
  let totalDtpk = 0;

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const normStatusHD = normalizeHeader(String(row[idxTinhTrangHD] || ''));

    // Loại trừ hợp đồng đã hủy/chấm dứt
    if (normStatusHD.includes('hủy') || normStatusHD.includes('huy') || normStatusHD.includes('thanh lý') || normStatusHD.includes('chấm dứt')) {
      continue;
    }

    const normLoaiKH = normalizeHeader(String(row[idxLoaiKH] || ''));
    const isNghiepVu = normLoaiKH.includes('thuê bao nghiệp vụ') || normLoaiKH.includes('thue bao nghiep vu');
    const isBCA = normLoaiKH.includes('bca') || normLoaiKH.includes('vũ trang') || normLoaiKH.includes('vu trang') || normLoaiKH.includes('hành chính') || normLoaiKH.includes('công an');

    if (targetGroup === 'nv' && !isNghiepVu) continue;
    if (targetGroup === 'bca' && !isBCA) continue;
    if (targetGroup === 'all' && !isNghiepVu && !isBCA) continue;

    // Kiểm tra thời gian nếu lọc theo khoảng
    const dDauNoi = parseToDate(row[idxNgayDauNoi]);
    const dBatDau = parseToDate(row[idxNgayBatDau]);
    const checkDate = dDauNoi || dBatDau;

    if (timeType === 'range') {
      if (!checkDate || checkDate < startDate || checkDate > endDate) continue;
    }

    const giaCuoc = parseCurrency(row[idxGiaCuoc]);
    const chuKy = parseInt(row[idxChuKy], 10) || 1;
    let dtpk = parseCurrency(row[idxDtpk]);
    if (!dtpk && giaCuoc > 0) dtpk = Math.round(giaCuoc / chuKy);

    totalRevenue += giaCuoc;
    totalDtpk += dtpk;

    // Xác định Khu vực
    const rawMaNV = normalizeHeader(String(row[idxMaNV] || ''));
    let khuVuc = 'Chưa phân bổ KV';
    if (rawMaNV && staffRegionMap.has(rawMaNV)) {
      khuVuc = staffRegionMap.get(rawMaNV);
    }

    if (!regionSummaryMap.has(khuVuc)) {
      regionSummaryMap.set(khuVuc, { count: 0, revenue: 0, dtpk: 0 });
    }
    const stat = regionSummaryMap.get(khuVuc);
    stat.count++;
    stat.revenue += giaCuoc;
    stat.dtpk += dtpk;

    let cleanMaKH = String(row[idxMaKH] || '').replace(/^'+/, '').trim();
    if (/^\d{8}$/.test(cleanMaKH)) cleanMaKH = '00' + cleanMaKH;

    const ngayDauNoiStr = dDauNoi ? Utilities.formatDate(dDauNoi, 'Asia/Saigon', 'dd/MM/yyyy') : '';
    const ngayBatDauStr = dBatDau ? Utilities.formatDate(dBatDau, 'Asia/Saigon', 'dd/MM/yyyy') : '';

    specRowsData.push([
      specRowsData.length + 1,
      khuVuc,
      row[idxLoaiKH],
      cleanMaKH,
      row[idxTenTB],
      row[idxSDT],
      row[idxDiaChi],
      row[idxGoiCuoc],
      giaCuoc,
      chuKy,
      dtpk,
      ngayDauNoiStr,
      ngayBatDauStr,
      row[idxThietBi],
      row[idxSerial],
      row[idxMaCH],
      row[idxTenCH],
      row[idxMaNV],
      row[idxTenNV]
    ]);
  }

  // 4. KẾT XUẤT VÀ MATRIX FORMATTING
  let reportSheet = ss.getSheetByName(sheetName);
  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet(sheetName);
  }

  const nowFormatted = Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm');

  // Khối Bảng 1: Thống kê Khu vực
  const b1Header = [
    [`BÁO CÁO THUÊ BAO ĐẶC THÙ (${groupTitle}) - ${timeTitle}`, '', '', '', ''],
    [`🎯 TỔNG CỘNG: ${specRowsData.length} TB | DOANH THU: ${totalRevenue.toLocaleString('vi-VN')} đ | DTPK: ${totalDtpk.toLocaleString('vi-VN')} đ`, '', '', '', `🕒 Lập lúc: ${nowFormatted}`],
    ['KHU VỰC', 'SỐ LƯỢNG TB', 'DOANH THU (VNĐ)', 'DTPK (VNĐ)', 'TỶ TRỌNG (%)']
  ];

  const b1Data = [];
  const sortedRegions = Array.from(regionSummaryMap.keys()).sort();
  sortedRegions.forEach(kv => {
    const s = regionSummaryMap.get(kv);
    const ratio = totalDtpk > 0 ? (s.dtpk / totalDtpk) : 0;
    b1Data.push([kv, s.count, s.revenue, s.dtpk, ratio]);
  });
  b1Data.push(['TỔNG CỘNG', specRowsData.length, totalRevenue, totalDtpk, totalDtpk > 0 ? 1 : 0]);

  reportSheet.getRange(1, 1, 3, 5).setValues(b1Header);
  reportSheet.getRange(4, 1, b1Data.length, 5).setValues(b1Data);

  // Định dạng Bảng 1
  reportSheet.getRange('A1:E1').merge().setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  reportSheet.getRange('A2:D2').merge().setBackground('#e8f0fe').setFontColor('#174ea6').setFontWeight('bold');
  reportSheet.getRange('A3:E3').setBackground('#34a853').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  reportSheet.getRange(4, 2, b1Data.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('center');
  reportSheet.getRange(4, 3, b1Data.length, 2).setNumberFormat('#,##0" đ"');
  reportSheet.getRange(4, 5, b1Data.length, 1).setNumberFormat('0.00%').setHorizontalAlignment('right');
  reportSheet.getRange(3 + b1Data.length, 1, 1, 5).setBackground('#ea4335').setFontColor('#ffffff').setFontWeight('bold');

  // Khối Bảng 2: Danh sách chi tiết
  const b2StartRow = 5 + b1Data.length;
  const b2Header = [
    'STT', 'KHU VỰC', 'PHÂN LOẠI KH', 'MÃ KHÁCH HÀNG', 'TÊN THUÊ BAO', 'SĐT LIÊN HỆ', 'ĐỊA CHỈ LẮP ĐẶT',
    'GÓI CƯỚC', 'GIÁ CƯỚC', 'CHU KỲ', 'DTPK', 'NGÀY ĐẤU NỐI', 'NGÀY BẮT ĐẦU', 'LOẠI ONT', 'SERIAL ONT',
    'MÃ CỬA HÀNG', 'TÊN CỬA HÀNG', 'MÃ NHÂN VIÊN', 'NHÂN VIÊN QUẢN LÝ'
  ];

  reportSheet.getRange(b2StartRow, 1, 1, b2Header.length).setValues([b2Header])
    .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  if (specRowsData.length > 0) {
    reportSheet.getRange(b2StartRow + 1, 1, specRowsData.length, b2Header.length).setValues(specRowsData);

    // Format Bảng 2 hàng loạt
    reportSheet.getRange(b2StartRow + 1, 1, specRowsData.length, 1).setHorizontalAlignment('center');
    reportSheet.getRange(b2StartRow + 1, 4, specRowsData.length, 1).setNumberFormat('@').setHorizontalAlignment('center');
    reportSheet.getRange(b2StartRow + 1, 9, specRowsData.length, 1).setNumberFormat('#,##0" đ"');
    reportSheet.getRange(b2StartRow + 1, 10, specRowsData.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('center');
    reportSheet.getRange(b2StartRow + 1, 11, specRowsData.length, 1).setNumberFormat('#,##0" đ"');
    reportSheet.getRange(b2StartRow + 1, 12, specRowsData.length, 2).setHorizontalAlignment('center');
  }

  // Tối ưu độ rộng cột
  reportSheet.setColumnWidth(1, 50);
  reportSheet.setColumnWidth(2, 130);
  reportSheet.setColumnWidth(3, 140);
  reportSheet.setColumnWidth(4, 110);
  reportSheet.setColumnWidth(5, 180);
  reportSheet.setColumnWidth(6, 100);
  reportSheet.setColumnWidth(7, 230);
  reportSheet.setColumnWidth(8, 100);
  reportSheet.setColumnWidth(9, 110);
  reportSheet.setColumnWidth(11, 110);

  const duration = ((new Date().getTime() - startTime) / 1000).toFixed(2);
  Logger.log(`⚡ Kết xuất Báo cáo Đặc thù hoàn tất trong ${duration}s!`);
  ss.setActiveSheet(reportSheet);
}