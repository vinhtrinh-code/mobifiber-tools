/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * FILE        : Mod_DoanhThuPhanBo.gs
 * PHIÊN BẢN   : 3.0 (High-Performance Engine: In-Memory Aggregation & Matrix Formatting)
 * NGÀY TẠO    : 21/09/2026 08:30:00 (GMT+7)
 * NGÀY CẬP NHẬT: 24/10/2026 19:15:00 (GMT+7)
 * TÁC GIẢ     : Trịnh Vũ Thế Vinh & AI Assistant
 * MÔ TẢ       : Module tổng hợp Báo cáo Doanh thu Phân bổ (DTPK) theo Khu vực & Địa bàn
 * ============================================================================
 */

function showRevenueAllocationDialog() {
  const currentMonth = Utilities.formatDate(new Date(), 'Asia/Saigon', 'MM/yyyy');
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 18px; margin: 0; font-size: 13px; color: #202124; background: #fff; }
        .form-group { margin-bottom: 14px; }
        label { display: block; font-weight: 600; margin-bottom: 6px; color: #3c4043; }
        input[type="text"] { width: 100%; box-sizing: border-box; padding: 9px 12px; border: 1px solid #dadce0; border-radius: 6px; font-size: 14px; outline: none; }
        input[type="text"]:focus { border-color: #1a73e8; box-shadow: 0 0 0 2px rgba(26,115,232,0.2); }
        .hint { font-size: 11px; color: #70757a; margin-top: 4px; }
        .btn-container { display: flex; justify-content: flex-end; gap: 8px; margin-top: 18px; }
        button { padding: 8px 18px; border-radius: 6px; font-weight: 600; cursor: pointer; border: none; font-size: 13px; }
        .btn-cancel { background: #f1f3f4; color: #3c4043; }
        .btn-submit { background: #1a73e8; color: #fff; }
        .btn-submit:disabled { background: #dadce0; cursor: not-allowed; }
      </style>
    </head>
    <body>
      <div class="form-group">
        <label>Tháng phân bổ doanh thu (MM/YYYY):</label>
        <input type="text" id="monthInput" value="${currentMonth}" placeholder="Ví dụ: 10/2026">
        <div class="hint">Báo cáo sẽ tổng hợp TB Hiện hữu, PTM và tính tỷ trọng DTPK theo từng Khu vực.</div>
      </div>
      <div class="btn-container">
        <button class="btn-cancel" onclick="google.script.host.close()">Hủy</button>
        <button class="btn-submit" id="btnSubmit" onclick="runAllocation()">Khởi tạo báo cáo</button>
      </div>
      <script>
        function runAllocation() {
          const val = document.getElementById('monthInput').value.trim();
          if (!/^\\d{2}\\/\\d{4}$/.test(val)) {
            alert('Vui lòng nhập đúng định dạng MM/YYYY (Ví dụ: 10/2026)');
            return;
          }
          document.getElementById('btnSubmit').disabled = true;
          document.getElementById('btnSubmit').innerText = 'Đang xử lý...';
          google.script.run
            .withSuccessHandler(function(res) {
              google.script.host.close();
            })
            .withFailureHandler(function(err) {
              alert('Lỗi: ' + err.message);
              document.getElementById('btnSubmit').disabled = false;
              document.getElementById('btnSubmit').innerText = 'Khởi tạo báo cáo';
            })
            .generateRevenueAllocationReport(val);
        }
      </script>
    </body>
    </html>
  `;
  const ui = HtmlService.createHtmlOutput(html).setWidth(380).setHeight(190);
  SpreadsheetApp.getUi().showModalDialog(ui, '📊 Báo cáo Doanh thu Phân bổ (DTPK)');
}

/**
 * XỬ LÝ TỔNG HỢP VÀ TẠO BÁO CÁO DTPB TỐI ƯU TRONG BỘ NHỚ
 */
function generateRevenueAllocationReport(targetMonthStr) {
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);

  const parts = targetMonthStr.split('/');
  const targetMonth = parseInt(parts[0], 10);
  const targetYear = parseInt(parts[1], 10);
  const sheetTargetName = `DTPB_T${parts[0]}_${parts[1]}`;

  // 1. ĐỌC DỮ LIỆU ĐẦU VÀO VÀ THAM SỐ (1 LẦN TRONG RAM)
  const lastCol = rawSheet.getLastColumn();
  const lastRow = rawSheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) throw new Error('Sheet raw_mobifiber chưa có dữ liệu!');

  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));
  const idxMaKH = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxTenTB = headers.findIndex(h => h === 'tên thuê bao');
  const idxGiaCuoc = headers.findIndex(h => h === 'giá cước');
  const idxChuKy = headers.findIndex(h => h === 'chu kỳ sử dụng');
  const idxDtpk = headers.findIndex(h => h === 'dtpk');
  const idxNgayDauNoi = headers.findIndex(h => h === 'ngày đấu nối');
  const idxNgayBatDau = headers.findIndex(h => h === 'ngày băt đầu gói cước' || h === 'ngày bắt đầu gói cước');
  const idxNgayKetThuc = headers.findIndex(h => h === 'ngày kết thúc gói cước');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxMaNV = headers.findIndex(h => h === 'mã nhân viên');

  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, lastRow - CONFIG.DATA_START_ROW_TARGET + 1, lastCol).getValues();

  // Đọc danh mục phân bổ khu vực
  const staffRegionMap = new Map();
  const thamSoSheet = ss.getSheetByName('thamso_caidat');
  if (thamSoSheet && thamSoSheet.getLastRow() >= 4) {
    const tsData = thamSoSheet.getRange(4, 10, thamSoSheet.getLastRow() - 3, 4).getValues(); // Cột J:M (Khu vực, Mã CH, Mã NV, Tên NV)
    tsData.forEach(r => {
      const kv = String(r[0] || '').trim();
      const maNV = normalizeHeader(String(r[2] || ''));
      if (maNV && kv) staffRegionMap.set(maNV, kv);
    });
  }

  // 2. KHỞI TẠO MẢNG THỐNG KÊ TRÊN RAM
  const regionMap = new Map(); // Key: Tên Khu vực -> { currentSub: 0, currentDtpk: 0, ptmSub: 0, ptmRevenue: 0, totalDtpk: 0 }
  
  let totalCurrentSub = 0;
  let totalCurrentDtpk = 0;
  let totalPtmSub = 0;
  let totalPtmRevenue = 0;
  let grandTotalDtpk = 0;

  const monthStartDate = new Date(targetYear, targetMonth - 1, 1, 0, 0, 0);
  const monthEndDate = new Date(targetYear, targetMonth, 0, 23, 59, 59);

  // 3. TÍNH TOÁN IN-MEMORY
  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const statusHD = normalizeHeader(String(row[idxTinhTrangHD] || ''));

    // Bỏ qua hợp đồng đã hủy/chấm dứt
    if (statusHD.includes('hủy') || statusHD.includes('huy') || statusHD.includes('thanh lý') || statusHD.includes('chấm dứt')) {
      continue;
    }

    const dDauNoi = parseToDate(row[idxNgayDauNoi]);
    const dBatDau = parseToDate(row[idxNgayBatDau]);
    const dKetThuc = parseToDate(row[idxNgayKetThuc]);
    const checkDate = dDauNoi || dBatDau;

    const giaCuoc = parseCurrency(row[idxGiaCuoc]);
    const chuKy = parseInt(row[idxChuKy], 10) || 1;
    let dtpk = parseCurrency(row[idxDtpk]);
    if (!dtpk && giaCuoc > 0) dtpk = Math.round(giaCuoc / chuKy);

    // Xác định khu vực
    const maNV = normalizeHeader(String(row[idxMaNV] || ''));
    let khuVuc = 'Chưa phân bổ KV';
    if (maNV && staffRegionMap.has(maNV)) {
      khuVuc = staffRegionMap.get(maNV);
    }

    if (!regionMap.has(khuVuc)) {
      regionMap.set(khuVuc, { currentSub: 0, currentDtpk: 0, ptmSub: 0, ptmRevenue: 0, totalDtpk: 0 });
    }
    const regStat = regionMap.get(khuVuc);

    // Phân loại PTM hay Hiện hữu
    const isPtm = checkDate && checkDate >= monthStartDate && checkDate <= monthEndDate;

    if (isPtm) {
      regStat.ptmSub++;
      regStat.ptmRevenue += giaCuoc;
      regStat.totalDtpk += dtpk;
      totalPtmSub++;
      totalPtmRevenue += giaCuoc;
      grandTotalDtpk += dtpk;
    } else {
      // Thuê bao hiện hữu còn hiệu lực gói cước
      const isValidCurrent = !dKetThuc || dKetThuc >= monthStartDate;
      if (isValidCurrent) {
        regStat.currentSub++;
        regStat.currentDtpk += dtpk;
        regStat.totalDtpk += dtpk;
        totalCurrentSub++;
        totalCurrentDtpk += dtpk;
        grandTotalDtpk += dtpk;
      }
    }
  }

  // 4. XÂY DỰNG SHEET BÁO CÁO MỚI (BULK WRITE & BULK FORMAT)
  let reportSheet = ss.getSheetByName(sheetTargetName);
  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet(sheetTargetName);
  }

  const nowFormatted = Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm');

  // Khối Header & KPI Cards
  const headerMatrix = [
    [`BÁO CÁO DOANH THU PHÂN BỔ MOBIFIBER - THÁNG ${targetMonthStr}`, '', '', '', '', '', '', ''],
    [`🎯 TỔNG DTPK TOÀN TỈNH: ${grandTotalDtpk.toLocaleString('vi-VN')} VNĐ`, '', '', '', '', `🕒 Ngày lập: ${nowFormatted}`, '', ''],
    [`📊 TỔNG THUÊ BAO: ${totalCurrentSub + totalPtmSub} TB (Hiện hữu: ${totalCurrentSub} | PTM: ${totalPtmSub})`, '', '', '', '', '', '', ''],
    ['KHU VỰC', 'TB HIỆN HỮU', 'DTPK HIỆN HỮU', 'TB PTM', 'DOANH THU PTM', 'TỔNG SỐ TB', 'TỔNG DTPK (VNĐ)', 'TỶ TRỌNG (%)']
  ];

  // Khối Dữ liệu từng Khu vực
  const dataMatrix = [];
  const sortedRegions = Array.from(regionMap.keys()).sort();

  sortedRegions.forEach(regName => {
    const s = regionMap.get(regName);
    const totalSub = s.currentSub + s.ptmSub;
    const ratio = grandTotalDtpk > 0 ? (s.totalDtpk / grandTotalDtpk) : 0;

    dataMatrix.push([
      regName,
      s.currentSub,
      s.currentDtpk,
      s.ptmSub,
      s.ptmRevenue,
      totalSub,
      s.totalDtpk,
      ratio
    ]);
  });

  // Hàng Tổng Cộng
  const totalRatio = grandTotalDtpk > 0 ? 1 : 0;
  dataMatrix.push([
    'TỔNG CỘNG',
    totalCurrentSub,
    totalCurrentDtpk,
    totalPtmSub,
    totalPtmRevenue,
    totalCurrentSub + totalPtmSub,
    grandTotalDtpk,
    totalRatio
  ]);

  // Ghi toàn bộ dữ liệu 1 lần
  reportSheet.getRange(1, 1, 4, 8).setValues(headerMatrix);
  reportSheet.getRange(5, 1, dataMatrix.length, 8).setValues(dataMatrix);

  // Merge các ô tiêu đề
  reportSheet.getRange('A1:H1').merge();
  reportSheet.getRange('A2:E2').merge();
  reportSheet.getRange('F2:H2').merge();
  reportSheet.getRange('A3:H3').merge();

  // Định dạng Matrix Format 1 lần duy nhất
  reportSheet.getRange('A1:H1').setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  reportSheet.getRange('A2:H2').setBackground('#e8f0fe').setFontColor('#174ea6').setFontWeight('bold');
  reportSheet.getRange('A3:H3').setBackground('#f1f3f4').setFontColor('#3c4043').setFontStyle('italic');
  reportSheet.getRange('A4:H4').setBackground('#34a853').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  // Định dạng số tiền và tỷ lệ %
  reportSheet.getRange(5, 2, dataMatrix.length, 1).setNumberFormat('#,##0');
  reportSheet.getRange(5, 3, dataMatrix.length, 1).setNumberFormat('#,##0" đ"');
  reportSheet.getRange(5, 4, dataMatrix.length, 1).setNumberFormat('#,##0');
  reportSheet.getRange(5, 5, dataMatrix.length, 1).setNumberFormat('#,##0" đ"');
  reportSheet.getRange(5, 6, dataMatrix.length, 1).setNumberFormat('#,##0');
  reportSheet.getRange(5, 7, dataMatrix.length, 1).setNumberFormat('#,##0" đ"');
  reportSheet.getRange(5, 8, dataMatrix.length, 1).setNumberFormat('0.00%');

  // Hàng Tổng Cộng cuối cùng
  const lastReportRow = 4 + dataMatrix.length;
  reportSheet.getRange(lastReportRow, 1, 1, 8).setBackground('#ea4335').setFontColor('#ffffff').setFontWeight('bold');

  // Căn chỉnh độ rộng cột
  reportSheet.setColumnWidth(1, 160);
  reportSheet.setColumnWidth(2, 110);
  reportSheet.setColumnWidth(3, 140);
  reportSheet.setColumnWidth(4, 100);
  reportSheet.setColumnWidth(5, 140);
  reportSheet.setColumnWidth(6, 110);
  reportSheet.setColumnWidth(7, 150);
  reportSheet.setColumnWidth(8, 110);

  const duration = ((new Date().getTime() - startTime) / 1000).toFixed(2);
  Logger.log(`⚡ Tạo báo cáo DTPB hoàn tất trong ${duration}s!`);
  ss.setActiveSheet(reportSheet);
}