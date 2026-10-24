/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * FILE        : Mod_PTM.gs
 * PHIÊN BẢN   : 3.0 (Tối ưu In-Memory Aggregation, Matrix Formatting & Fast Reporting)
 * NGÀY TẠO    : 19/09/2026 09:15:00 (GMT+7)
 * NGÀY CẬP NHẬT: 24/10/2026 19:25:00 (GMT+7)
 * TÁC GIẢ     : Trịnh Vũ Thế Vinh & AI Assistant
 * MÔ TẢ       : Khởi tạo và kết xuất Báo cáo Phát triển mới (PTM) theo chu kỳ & địa bàn
 * ============================================================================
 */

function showPtmReportDialog() {
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
        input[type="text"] { width: 100%; box-sizing: border-box; padding: 8px 10px; border: 1px solid #dadce0; border-radius: 6px; font-size: 13px; }
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
        <label>Chọn phạm vi thời gian PTM:</label>
        <div class="radio-item">
          <input type="radio" id="typeMonth" name="timeType" value="month" checked onchange="toggleTimeType()">
          <label for="typeMonth" style="font-weight:normal; margin:0;">Theo tháng</label>
        </div>
        <div class="radio-item">
          <input type="radio" id="typeRange" name="timeType" value="range" onchange="toggleTimeType()">
          <label for="typeRange" style="font-weight:normal; margin:0;">Từ ngày... đến ngày...</label>
        </div>
      </div>

      <div class="form-group" id="monthBox">
        <label>Tháng PTM (MM/YYYY):</label>
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
        <button class="btn-submit" id="btnRun" onclick="submitPtm()">Xuất báo cáo PTM</button>
      </div>

      <script>
        function toggleTimeType() {
          const isMonth = document.getElementById('typeMonth').checked;
          document.getElementById('monthBox').style.display = isMonth ? 'block' : 'none';
          document.getElementById('rangeBox').style.display = isMonth ? 'none' : 'flex';
        }

        function submitPtm() {
          const isMonth = document.getElementById('typeMonth').checked;
          const params = { isMonth: isMonth };
          if (isMonth) {
            params.monthStr = document.getElementById('monthVal').value.trim();
            if (!/^\\d{2}\\/\\d{4}$/.test(params.monthStr)) {
              alert('Vui lòng nhập định dạng MM/YYYY (Ví dụ: 10/2026)');
              return;
            }
          } else {
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
              document.getElementById('btnRun').innerText = 'Xuất báo cáo PTM';
            })
            .generatePtmReport(params);
        }
      </script>
    </body>
    </html>
  `;
  const ui = HtmlService.createHtmlOutput(html).setWidth(400).setHeight(270);
  SpreadsheetApp.getUi().showModalDialog(ui, '🎯 Kết xuất Báo cáo Phát triển mới (PTM)');
}

/**
 * XỬ LÝ LỌC & KẾT XUẤT BÁO CÁO PTM TỐI ƯU TRÊN BỘ NHỚ
 */
function generatePtmReport(params) {
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);

  // 1. XÁC ĐỊNH KHOẢNG THỜI GIAN LỌC
  let startDate, endDate, sheetName, titleSuffix;
  if (params.isMonth) {
    const parts = params.monthStr.split('/');
    const m = parseInt(parts[0], 10);
    const y = parseInt(parts[1], 10);
    startDate = new Date(y, m - 1, 1, 0, 0, 0);
    endDate = new Date(y, m, 0, 23, 59, 59);
    sheetName = `PTM_T${parts[0]}_${parts[1]}`;
    titleSuffix = `THÁNG ${params.monthStr}`;
  } else {
    const pStart = params.from.split('/');
    const pEnd = params.to.split('/');
    startDate = new Date(parseInt(pStart[2], 10), parseInt(pStart[1], 10) - 1, parseInt(pStart[0], 10), 0, 0, 0);
    endDate = new Date(parseInt(pEnd[2], 10), parseInt(pEnd[1], 10) - 1, parseInt(pEnd[0], 10), 23, 59, 59);
    sheetName = `PTM_${pStart[0]}${pStart[1]}_${pEnd[0]}${pEnd[1]}_${pEnd[2]}`;
    titleSuffix = `TỪ ${params.from} ĐẾN ${params.to}`;
  }

  // 2. ĐỌC DỮ LIỆU ĐÍCH VÀ DANH MỤC THAM SỐ VÀO RAM
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

  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, lastRow - CONFIG.DATA_START_ROW_TARGET + 1, lastCol).getValues();

  // Nạp bảng tham số ánh xạ Khu vực
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

  // 3. XỬ LÝ LỌC IN-MEMORY
  const ptmRows = [];
  const regionStats = new Map();
  let totalRev = 0;
  let totalDtpk = 0;

  for (let i = 0; i < rawData.length; i++) {
    const r = rawData[i];

    // Tiêu chí 1: Nhu cầu = "Đăng ký mới"
    const nhuCau = normalizeHeader(String(r[idxNhuCau] || ''));
    if (nhuCau !== 'đăng ký mới' && nhuCau !== 'dang ky moi') continue;

    // Tiêu chí 2: Loại trừ "Thuê bao nghiệp vụ"
    const loaiKH = normalizeHeader(String(r[idxLoaiKH] || ''));
    if (loaiKH === 'thuê bao nghiệp vụ' || loaiKH === 'thue bao nghiep vu') continue;

    // Tiêu chí 3: Tình trạng hợp đồng = "Đang hoạt động"
    const tinhTrang = normalizeHeader(String(r[idxTinhTrangHD] || ''));
    if (tinhTrang !== 'đang hoạt động' && tinhTrang !== 'dang hoat dong') continue;

    const chuKy = parseInt(r[idxChuKy], 10) || 1;
    const dDauNoi = parseToDate(r[idxNgayDauNoi]);
    const dThiCong = parseToDate(r[idxNgayThiCong]);
    const dBatDau = parseToDate(r[idxNgayBatDau]);

    // Tiêu chí 4: Xác định Ngày PTM chuẩn
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

    if (!ptmDate || ptmDate < startDate || ptmDate > endDate) continue;

    const giaCuoc = parseCurrency(r[idxGiaCuoc]);
    let dtpkVal = parseCurrency(r[idxDtpk]);
    if (!dtpkVal && giaCuoc > 0) dtpkVal = Math.round(giaCuoc / chuKy);

    totalRev += giaCuoc;
    totalDtpk += dtpkVal;

    // Xác định khu vực
    const maNV = normalizeHeader(String(r[idxMaNV] || ''));
    let khuVuc = 'Chưa phân bổ KV';
    if (maNV && staffRegionMap.has(maNV)) {
      khuVuc = staffRegionMap.get(maNV);
    }

    if (!regionStats.has(khuVuc)) {
      regionStats.set(khuVuc, { count: 0, revenue: 0, dtpk: 0 });
    }
    const stat = regionStats.get(khuVuc);
    stat.count++;
    stat.revenue += giaCuoc;
    stat.dtpk += dtpkVal;

    // Chuẩn hóa mã KH hiển thị
    let maKH = String(r[idxMaKH] || '').replace(/^'+/, '').trim();
    if (/^\d{8}$/.test(maKH)) maKH = '00' + maKH;

    const ptmDateStr = Utilities.formatDate(ptmDate, 'Asia/Saigon', 'dd/MM/yyyy');

    ptmRows.push([
      ptmRows.length + 1,
      khuVuc,
      maKH,
      r[idxTenTB],
      r[idxSDT],
      r[idxDiaChi],
      r[idxGoiCuoc],
      giaCuoc,
      chuKy,
      dtpkVal,
      ptmDateStr,
      r[idxThietBi],
      r[idxSerial],
      r[idxMaCH],
      r[idxTenCH],
      r[idxMaNV],
      r[idxTenNV]
    ]);
  }

  // 4. KẾT XUẤT VÀ MATRIX FORMATTING
  let reportSheet = ss.getSheetByName(sheetName);
  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet(sheetName);
  }

  const nowStr = Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm');

  // Khối Bảng 1: Tổng hợp KPI theo Khu vực
  const b1Header = [
    [`BÁO CÁO PHÁT TRIỂN MỚI MOBIFIBER - ${titleSuffix}`, '', '', '', ''],
    [`🎯 TỔNG PTM: ${ptmRows.length} TB | DOANH THU: ${totalRev.toLocaleString('vi-VN')} đ | DTPK: ${totalDtpk.toLocaleString('vi-VN')} đ`, '', '', '', `🕒 Lập lúc: ${nowStr}`],
    ['KHU VỰC', 'SỐ LƯỢNG PTM', 'DOANH THU PTM (VNĐ)', 'DTPK (VNĐ)', 'TỶ TRỌNG (%)']
  ];

  const b1Data = [];
  const sortedRegions = Array.from(regionStats.keys()).sort();
  sortedRegions.forEach(kv => {
    const s = regionStats.get(kv);
    const ratio = totalRev > 0 ? (s.revenue / totalRev) : 0;
    b1Data.push([kv, s.count, s.revenue, s.dtpk, ratio]);
  });
  b1Data.push(['TỔNG CỘNG', ptmRows.length, totalRev, totalDtpk, totalRev > 0 ? 1 : 0]);

  // Ghi Bảng 1
  reportSheet.getRange(1, 1, 3, 5).setValues(b1Header);
  reportSheet.getRange(4, 1, b1Data.length, 5).setValues(b1Data);

  // Merge tiêu đề Bảng 1
  reportSheet.getRange('A1:E1').merge().setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  reportSheet.getRange('A2:D2').merge().setBackground('#e8f0fe').setFontColor('#174ea6').setFontWeight('bold');
  reportSheet.getRange('A3:E3').setBackground('#34a853').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  // Format số liệu Bảng 1
  reportSheet.getRange(4, 2, b1Data.length, 1).setNumberFormat('#,##0');
  reportSheet.getRange(4, 3, b1Data.length, 2).setNumberFormat('#,##0" đ"');
  reportSheet.getRange(4, 5, b1Data.length, 1).setNumberFormat('0.00%');
  reportSheet.getRange(3 + b1Data.length, 1, 1, 5).setBackground('#ea4335').setFontColor('#ffffff').setFontWeight('bold');

  // Khối Bảng 2: Danh sách chi tiết thuê bao PTM
  const b2StartRow = 5 + b1Data.length;
  const b2Header = [
    'STT', 'KHU VỰC', 'MÃ KHÁCH HÀNG', 'TÊN THUÊ BAO', 'SĐT LIÊN HỆ', 'ĐỊA CHỈ LẮP ĐẶT',
    'GÓI CƯỚC', 'GIÁ CƯỚC', 'CHU KỲ', 'DTPK', 'NGÀY PTM', 'LOẠI ONT', 'SERIAL ONT',
    'MÃ CỬA HÀNG', 'TÊN CỬA HÀNG', 'MÃ NHÂN VIÊN', 'TÊN NHÂN VIÊN'
  ];

  reportSheet.getRange(b2StartRow, 1, 1, b2Header.length).setValues([b2Header])
    .setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  if (ptmRows.length > 0) {
    reportSheet.getRange(b2StartRow + 1, 1, ptmRows.length, b2Header.length).setValues(ptmRows);
    
    // Format Bảng 2 hàng loạt
    reportSheet.getRange(b2StartRow + 1, 1, ptmRows.length, 1).setHorizontalAlignment('center');
    reportSheet.getRange(b2StartRow + 1, 3, ptmRows.length, 1).setNumberFormat('@').setHorizontalAlignment('center');
    reportSheet.getRange(b2StartRow + 1, 8, ptmRows.length, 1).setNumberFormat('#,##0" đ"');
    reportSheet.getRange(b2StartRow + 1, 9, ptmRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('center');
    reportSheet.getRange(b2StartRow + 1, 10, ptmRows.length, 1).setNumberFormat('#,##0" đ"');
    reportSheet.getRange(b2StartRow + 1, 11, ptmRows.length, 1).setHorizontalAlignment('center');
  }

  // Tự động căn chỉnh chiều rộng các cột chính
  reportSheet.setColumnWidth(1, 50);
  reportSheet.setColumnWidth(2, 130);
  reportSheet.setColumnWidth(3, 110);
  reportSheet.setColumnWidth(4, 180);
  reportSheet.setColumnWidth(5, 100);
  reportSheet.setColumnWidth(6, 220);
  reportSheet.setColumnWidth(7, 100);
  reportSheet.setColumnWidth(8, 110);
  reportSheet.setColumnWidth(10, 110);
  reportSheet.setColumnWidth(11, 100);

  const duration = ((new Date().getTime() - startTime) / 1000).toFixed(2);
  Logger.log(`⚡ Kết xuất Báo cáo PTM hoàn tất trong ${duration}s!`);
  ss.setActiveSheet(reportSheet);
}