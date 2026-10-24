/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * FILE        : Mod_GiaHan.gs
 * PHIÊN BẢN   : 3.0 (High-Performance Engine: In-Memory Expiry Engine & Matrix Highlighting)
 * NGÀY TẠO    : 20/09/2026 10:00:00 (GMT+7)
 * NGÀY CẬP NHẬT: 24/10/2026 19:45:00 (GMT+7)
 * TÁC GIẢ     : Trịnh Vũ Thế Vinh & AI Assistant
 * MÔ TẢ       : Kết xuất danh sách thuê bao đến hạn gia hạn gói cước, phân bổ theo địa bàn & nhân viên
 * ============================================================================
 */

function showRenewalPromptDialog() {
  const currentMonth = Utilities.formatDate(new Date(), 'Asia/Saigon', 'MM/yyyy');
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
        input[type="text"] { width: 100%; box-sizing: border-box; padding: 8px 10px; border: 1px solid #dadce0; border-radius: 6px; font-size: 13px; outline: none; }
        input[type="text"]:focus { border-color: #1a73e8; box-shadow: 0 0 0 2px rgba(26,115,232,0.2); }
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
      <div class="radio-group">
        <label>Chọn phạm vi thời gian hết hạn gói:</label>
        <div class="radio-item">
          <input type="radio" id="typeMonth" name="timeType" value="month" checked onchange="toggleTimeType()">
          <label for="typeMonth" style="font-weight:normal; margin:0;">Theo tháng hết hạn</label>
        </div>
        <div class="radio-item">
          <input type="radio" id="typeRange" name="timeType" value="range" onchange="toggleTimeType()">
          <label for="typeRange" style="font-weight:normal; margin:0;">Từ ngày... đến ngày...</label>
        </div>
      </div>

      <div class="form-group" id="monthBox">
        <label>Tháng hết hạn (MM/YYYY):</label>
        <input type="text" id="monthVal" value="${currentMonth}">
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
        <button class="btn-submit" id="btnRun" onclick="submitRenewal()">Xuất danh sách Gia hạn</button>
      </div>

      <script>
        function toggleTimeType() {
          const isMonth = document.getElementById('typeMonth').checked;
          document.getElementById('monthBox').style.display = isMonth ? 'block' : 'none';
          document.getElementById('rangeBox').style.display = isMonth ? 'none' : 'flex';
        }

        function submitRenewal() {
          const isMonth = document.getElementById('typeMonth').checked;
          const params = { isMonth: isMonth };
          if (isMonth) {
            params.monthStr = document.getElementById('monthVal').value.trim();
            if (!/^\\d{2}\\/\\d{4}$/.test(params.monthStr)) {
              alert('Vui lòng nhập đúng định dạng MM/YYYY (Ví dụ: 10/2026)');
              return;
            }
          } else {
            params.from = document.getElementById('fromVal').value.trim();
            params.to = document.getElementById('toVal').value.trim();
            if (!params.from || !params.to) {
              alert('Vui lòng nhập đầy đủ cả 2 ngày!');
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
              document.getElementById('btnRun').innerText = 'Xuất danh sách Gia hạn';
            })
            .generateRenewalReport(params);
        }
      </script>
    </body>
    </html>
  `;
  const ui = HtmlService.createHtmlOutput(html).setWidth(400).setHeight(270);
  SpreadsheetApp.getUi().showModalDialog(ui, '⏰ Kết xuất Danh sách Gia hạn Gói cước');
}

/**
 * XỬ LÝ LỌC & KẾT XUẤT DANH SÁCH GIA HẠN SIÊU TỐC
 */
function generateRenewalReport(params) {
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);

  // 1. XÁC ĐỊNH KHOẢNG THỜI GIAN
  let startDate, endDate, sheetName, titleSuffix;
  if (params.isMonth) {
    const parts = params.monthStr.split('/');
    const m = parseInt(parts[0], 10);
    const y = parseInt(parts[1], 10);
    startDate = new Date(y, m - 1, 1, 0, 0, 0);
    endDate = new Date(y, m, 0, 23, 59, 59);
    sheetName = `GiaHan_T${parts[0]}_${parts[1]}`;
    titleSuffix = `THÁNG ${params.monthStr}`;
  } else {
    const pStart = params.from.split('/');
    const pEnd = params.to.split('/');
    startDate = new Date(parseInt(pStart[2], 10), parseInt(pStart[1], 10) - 1, parseInt(pStart[0], 10), 0, 0, 0);
    endDate = new Date(parseInt(pEnd[2], 10), parseInt(pEnd[1], 10) - 1, parseInt(pEnd[0], 10), 23, 59, 59);
    sheetName = `GiaHan_${pStart[0]}${pStart[1]}_${pEnd[0]}${pEnd[1]}_${pEnd[2]}`;
    titleSuffix = `TỪ ${params.from} ĐẾN ${params.to}`;
  }

  // 2. ĐỌC DỮ LIỆU ĐÍCH VÀ THAM SỐ VÀO RAM
  const lastCol = rawSheet.getLastColumn();
  const lastRow = rawSheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) throw new Error('Sheet raw_mobifiber chưa có dữ liệu!');

  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));
  const idxMaKH = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxTenTB = headers.findIndex(h => h === 'tên thuê bao');
  const idxSDT = headers.findIndex(h => h === 'sđt liên hệ');
  const idxDiaChi = headers.findIndex(h => h === 'địa chỉ lắp đặt');
  const idxGoiCuoc = headers.findIndex(h => h === 'gói cước');
  const idxGiaCuoc = headers.findIndex(h => h === 'giá cước');
  const idxChuKy = headers.findIndex(h => h === 'chu kỳ sử dụng');
  const idxNgayBatDau = headers.findIndex(h => h === 'ngày băt đầu gói cước' || h === 'ngày bắt đầu gói cước');
  const idxNgayKetThuc = headers.findIndex(h => h === 'ngày kết thúc gói cước');
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

  // 3. LỌC VÀ PHÂN LOẠI IN-MEMORY
  const renewalRows = [];
  const rowColors = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let totalRenewalRevenue = 0;
  let countExpired = 0; // Đã quá hạn
  let countNearExp = 0; // Sắp hết hạn (<= 7 ngày)
  let countNormal = 0;  // Còn hạn bình thường

  for (let i = 0; i < rawData.length; i++) {
    const r = rawData[i];
    const statusHD = normalizeHeader(String(r[idxTinhTrangHD] || ''));

    // Loại trừ hợp đồng đã hủy/chấm dứt
    if (statusHD.includes('hủy') || statusHD.includes('huy') || statusHD.includes('thanh lý') || statusHD.includes('chấm dứt')) {
      continue;
    }

    const dKetThuc = parseToDate(r[idxNgayKetThuc]);
    if (!dKetThuc) continue;

    // Kiểm tra nằm trong phạm vi cần lọc
    if (dKetThuc < startDate || dKetThuc > endDate) continue;

    const giaCuoc = parseCurrency(r[idxGiaCuoc]);
    const chuKy = parseInt(r[idxChuKy], 10) || 1;
    totalRenewalRevenue += giaCuoc;

    // Tính trạng thái hạn gói cước
    const diffDays = Math.ceil((dKetThuc.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
    let statusNote = '';
    let rowColor = '#ffffff';

    if (diffDays < 0) {
      statusNote = `Quá hạn ${Math.abs(diffDays)} ngày`;
      rowColor = '#fce8e6'; // Đỏ nhạt
      countExpired++;
    } else if (diffDays <= 7) {
      statusNote = `Sắp hết hạn (${diffDays} ngày)`;
      rowColor = '#fef7e0'; // Vàng nhạt
      countNearExp++;
    } else {
      statusNote = `Còn ${diffDays} ngày`;
      rowColor = '#ffffff';
      countNormal++;
    }

    // Xác định Khu vực
    const maNV = normalizeHeader(String(r[idxMaNV] || ''));
    let khuVuc = 'Chưa phân bổ KV';
    if (maNV && staffRegionMap.has(maNV)) {
      khuVuc = staffRegionMap.get(maNV);
    }

    let cleanMaKH = String(r[idxMaKH] || '').replace(/^'+/, '').trim();
    if (/^\d{8}$/.test(cleanMaKH)) cleanMaKH = '00' + cleanMaKH;

    const dBatDau = parseToDate(r[idxNgayBatDau]);
    const batDauStr = dBatDau ? Utilities.formatDate(dBatDau, 'Asia/Saigon', 'dd/MM/yyyy') : '';
    const ketThucStr = Utilities.formatDate(dKetThuc, 'Asia/Saigon', 'dd/MM/yyyy');

    renewalRows.push([
      renewalRows.length + 1,
      khuVuc,
      cleanMaKH,
      r[idxTenTB],
      r[idxSDT],
      r[idxDiaChi],
      r[idxGoiCuoc],
      giaCuoc,
      chuKy,
      batDauStr,
      ketThucStr,
      statusNote,
      r[idxMaCH],
      r[idxTenCH],
      r[idxMaNV],
      r[idxTenNV]
    ]);

    rowColors.push(new Array(16).fill(rowColor));
  }

  // 4. KẾT XUẤT VÀ MATRIX FORMATTING
  let reportSheet = ss.getSheetByName(sheetName);
  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet(sheetName);
  }

  const nowFormatted = Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm');

  // Khối Header & KPI Card
  const headerMatrix = [
    [`DANH SÁCH THUÊ BAO ĐẾN HẠN GIA HẠN GÓI CƯỚC - ${titleSuffix}`, '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    [
      `🎯 TỔNG CẦN GIA HẠN: ${renewalRows.length} TB | DOANH THU KỲ VỌNG: ${totalRenewalRevenue.toLocaleString('vi-VN')} đ | 🔴 Quá hạn: ${countExpired} | 🟡 Sắp hết hạn (<=7 ngày): ${countNearExp}`,
      '', '', '', '', '', '', '', '', '', '', '', '', '', `🕒 Lập lúc: ${nowFormatted}`, ''
    ],
    [
      'STT', 'KHU VỰC', 'MÃ KHÁCH HÀNG', 'TÊN THUÊ BAO', 'SĐT LIÊN HỆ', 'ĐỊA CHỈ LẮP ĐẶT',
      'GÓI CƯỚC', 'GIÁ CƯỚC', 'CHU KỲ', 'NGÀY BẮT ĐẦU', 'NGÀY KẾT THÚC', 'TÌNH TRẠNG HẠN',
      'MÃ CỬA HÀNG', 'TÊN CỬA HÀNG', 'MÃ NHÂN VIÊN', 'NHÂN VIÊN QUẢN LÝ'
    ]
  ];

  reportSheet.getRange(1, 1, 3, 16).setValues(headerMatrix);

  // Định dạng Tiêu đề
  reportSheet.getRange('A1:P1').merge().setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  reportSheet.getRange('A2:N2').merge().setBackground('#e8f0fe').setFontColor('#174ea6').setFontWeight('bold');
  reportSheet.getRange('O2:P2').merge().setBackground('#e8f0fe').setFontColor('#5f6368').setHorizontalAlignment('right');
  reportSheet.getRange('A3:P3').setBackground('#34a853').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  // 5. GHI DỮ LIỆU VÀ TÔ MÀU MATRIX
  if (renewalRows.length > 0) {
    const dataRange = reportSheet.getRange(4, 1, renewalRows.length, 16);
    dataRange.setValues(renewalRows);
    dataRange.setBackgrounds(rowColors);

    // Định dạng số & căn lề đồng loạt
    reportSheet.getRange(4, 1, renewalRows.length, 1).setHorizontalAlignment('center');
    reportSheet.getRange(4, 3, renewalRows.length, 1).setNumberFormat('@').setHorizontalAlignment('center');
    reportSheet.getRange(4, 8, renewalRows.length, 1).setNumberFormat('#,##0" đ"');
    reportSheet.getRange(4, 9, renewalRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('center');
    reportSheet.getRange(4, 10, renewalRows.length, 2).setHorizontalAlignment('center');
    reportSheet.getRange(4, 12, renewalRows.length, 1).setHorizontalAlignment('center').setFontWeight('bold');
  }

  // Tối ưu độ rộng cột
  reportSheet.setColumnWidth(1, 50);
  reportSheet.setColumnWidth(2, 130);
  reportSheet.setColumnWidth(3, 110);
  reportSheet.setColumnWidth(4, 180);
  reportSheet.setColumnWidth(5, 100);
  reportSheet.setColumnWidth(6, 230);
  reportSheet.setColumnWidth(7, 100);
  reportSheet.setColumnWidth(8, 110);
  reportSheet.setColumnWidth(9, 65);
  reportSheet.setColumnWidth(10, 95);
  reportSheet.setColumnWidth(11, 95);
  reportSheet.setColumnWidth(12, 140);
  reportSheet.setColumnWidth(13, 90);
  reportSheet.setColumnWidth(14, 150);
  reportSheet.setColumnWidth(15, 90);
  reportSheet.setColumnWidth(16, 140);

  const duration = ((new Date().getTime() - startTime) / 1000).toFixed(2);
  Logger.log(`⚡ Kết xuất Danh sách Gia hạn hoàn tất trong ${duration}s!`);
  ss.setActiveSheet(reportSheet);
}