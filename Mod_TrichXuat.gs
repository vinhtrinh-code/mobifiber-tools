/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * FILE        : Mod_TrichXuat.gs
 * PHIÊN BẢN   : 3.0 (High-Performance Engine: Multi-Criteria Filter & Matrix Export)
 * NGÀY TẠO    : 20/09/2026 17:00:00 (GMT+7)
 * NGÀY CẬP NHẬT: 24/10/2026 20:00:00 (GMT+7)
 * TÁC GIẢ     : Trịnh Vũ Thế Vinh & AI Assistant
 * MÔ TẢ       : Trích xuất danh sách thuê bao theo đa điều kiện động và kết xuất bảng dữ liệu sạch
 * ============================================================================
 */

function showQueryDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  const thamSoSheet = ss.getSheetByName('thamso_caidat');

  // Lấy nhanh danh sách Cửa hàng & Khu vực duy nhất để đưa vào Dropdown
  const storeList = [];
  const regionList = [];

  if (thamSoSheet && thamSoSheet.getLastRow() >= 4) {
    const tsData = thamSoSheet.getRange(4, 10, thamSoSheet.getLastRow() - 3, 2).getValues(); // Cột J:K (Khu vực, Mã/Tên CH)
    const regSet = new Set();
    const stSet = new Set();

    tsData.forEach(r => {
      const kv = String(r[0] || '').trim();
      const ch = String(r[1] || '').trim();
      if (kv && !kv.includes('Chưa phân bổ')) regSet.add(kv);
      if (ch) stSet.add(ch);
    });

    regionList.push(...Array.from(regSet).sort());
    storeList.push(...Array.from(stSet).sort());
  }

  const regionOptions = regionList.map(r => `<option value="${r}">${r}</option>`).join('');
  const storeOptions = storeList.map(s => `<option value="${s}">${s}</option>`).join('');

  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 16px; margin: 0; font-size: 13px; color: #202124; background: #fff; }
        .form-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-bottom: 12px; }
        .form-group { margin-bottom: 10px; }
        label { display: block; font-weight: 600; margin-bottom: 4px; color: #3c4043; font-size: 12px; }
        select, input[type="text"] { width: 100%; box-sizing: border-box; padding: 7px 9px; border: 1px solid #dadce0; border-radius: 6px; font-size: 13px; outline: none; }
        select:focus, input[type="text"]:focus { border-color: #1a73e8; box-shadow: 0 0 0 2px rgba(26,115,232,0.2); }
        .btn-container { display: flex; justify-content: flex-end; gap: 8px; margin-top: 14px; }
        button { padding: 8px 16px; border-radius: 6px; font-weight: 600; cursor: pointer; border: none; font-size: 13px; }
        .btn-cancel { background: #f1f3f4; color: #3c4043; }
        .btn-submit { background: #1a73e8; color: #fff; }
        .btn-submit:disabled { background: #dadce0; cursor: not-allowed; }
      </style>
    </head>
    <body>
      <div class="form-grid">
        <div class="form-group">
          <label>Khu vực địa bàn:</label>
          <select id="filterRegion">
            <option value="ALL">-- Tất cả khu vực --</option>
            ${regionOptions}
          </select>
        </div>
        <div class="form-group">
          <label>Cửa hàng / Điểm bán:</label>
          <select id="filterStore">
            <option value="ALL">-- Tất cả cửa hàng --</option>
            ${storeOptions}
          </select>
        </div>
      </div>

      <div class="form-grid">
        <div class="form-group">
          <label>Tình trạng hợp đồng:</label>
          <select id="filterContract">
            <option value="ALL">-- Tất cả trạng thái --</option>
            <option value="active">Đang hoạt động</option>
            <option value="blocked">Chặn cắt / Khóa cước</option>
            <option value="cancelled">Hủy / Chấm dứt</option>
          </select>
        </div>
        <div class="form-group">
          <label>Nhóm gói cước (Từ khóa):</label>
          <input type="text" id="filterPackage" placeholder="Ví dụ: 3MF, 6MF, FIBER...">
        </div>
      </div>

      <div class="btn-container">
        <button class="btn-cancel" onclick="google.script.host.close()">Hủy</button>
        <button class="btn-submit" id="btnSubmit" onclick="runExport()">Trích xuất danh sách</button>
      </div>

      <script>
        function runExport() {
          const params = {
            region: document.getElementById('filterRegion').value,
            store: document.getElementById('filterStore').value,
            contract: document.getElementById('filterContract').value,
            pkg: document.getElementById('filterPackage').value.trim()
          };

          document.getElementById('btnSubmit').disabled = true;
          document.getElementById('btnSubmit').innerText = 'Đang trích xuất...';

          google.script.run
            .withSuccessHandler(function(count) {
              google.script.host.close();
            })
            .withFailureHandler(function(err) {
              alert('Lỗi: ' + err.message);
              document.getElementById('btnSubmit').disabled = false;
              document.getElementById('btnSubmit').innerText = 'Trích xuất danh sách';
            })
            .exportSubscribersByCriteria(params);
        }
      </script>
    </body>
    </html>
  `;
  const ui = HtmlService.createHtmlOutput(html).setWidth(460).setHeight(260);
  SpreadsheetApp.getUi().showModalDialog(ui, '🎯 Trích xuất Danh sách Đa điều kiện');
}

/**
 * XỬ LÝ LỌC VÀ KẾT XUẤT ĐA ĐIỀU KIỆN SIÊU TỐC TRÊN RAM
 */
function exportSubscribersByCriteria(params) {
  const startTime = new Date().getTime();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) throw new Error('Sheet raw_mobifiber chưa có dữ liệu!');

  // 1. ĐỌC DỮ LIỆU ĐÍCH VÀ THAM SỐ VÀO RAM
  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol).getValues()[0].map(h => normalizeHeader(h));
  const idxMaKH = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxTenTB = headers.findIndex(h => h === 'tên thuê bao');
  const idxSDT = headers.findIndex(h => h === 'sđt liên hệ');
  const idxDiaChi = headers.findIndex(h => h === 'địa chỉ lắp đặt');
  const idxGoiCuoc = headers.findIndex(h => h === 'gói cước');
  const idxGiaCuoc = headers.findIndex(h => h === 'giá cước');
  const idxChuKy = headers.findIndex(h => h === 'chu kỳ sử dụng');
  const idxDtpk = headers.findIndex(h => h === 'dtpk');
  const idxNgayBatDau = headers.findIndex(h => h === 'ngày băt đầu gói cước' || h === 'ngày bắt đầu gói cước');
  const idxNgayKetThuc = headers.findIndex(h => h === 'ngày kết thúc gói cước');
  const idxTinhTrangHD = headers.findIndex(h => h === 'tình trạng hợp đồng');
  const idxChanCat = headers.findIndex(h => h === 'tình trạng chặn cắt');
  const idxLoaiONT = headers.findIndex(h => h === 'loại thiết bị ont');
  const idxSerial = headers.findIndex(h => h === 'serial ont');
  const idxMaCH = headers.findIndex(h => h === 'mã cửa hàng');
  const idxTenCH = headers.findIndex(h => h === 'tên cửa hàng');
  const idxMaNV = headers.findIndex(h => h === 'mã nhân viên');
  const idxTenNV = headers.findIndex(h => h === 'tên nhân viên');

  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, lastRow - CONFIG.DATA_START_ROW_TARGET + 1, lastCol).getValues();

  // Nạp bản đồ Khu vực từ thamso_caidat
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

  // 2. LỌC ĐA ĐIỀU KIỆN IN-MEMORY
  const filterPkg = params.pkg ? normalizeHeader(params.pkg) : '';
  const filterReg = params.region !== 'ALL' ? params.region : null;
  const filterSt = params.store !== 'ALL' ? normalizeHeader(params.store) : null;
  const filterCt = params.contract;

  const matchedRows = [];
  let totalRev = 0;
  let totalDtpk = 0;

  for (let i = 0; i < rawData.length; i++) {
    const r = rawData[i];

    // Lọc Hợp đồng
    const statusHD = normalizeHeader(String(r[idxTinhTrangHD] || ''));
    const isCancelled = statusHD.includes('hủy') || statusHD.includes('huy') || statusHD.includes('thanh lý') || statusHD.includes('chấm dứt');
    const statusChan = normalizeHeader(String(r[idxChanCat] || ''));
    const isBlocked = statusChan.includes('nợ') || statusChan.includes('khóa') || statusChan.includes('chặn');

    if (filterCt === 'active' && isCancelled) continue;
    if (filterCt === 'blocked' && !isBlocked) continue;
    if (filterCt === 'cancelled' && !isCancelled) continue;

    // Lọc Gói cước
    const goiCuoc = String(r[idxGoiCuoc] || '');
    if (filterPkg && !normalizeHeader(goiCuoc).includes(filterPkg)) continue;

    // Lọc Cửa hàng
    const tenCH = String(r[idxTenCH] || '');
    if (filterSt && !normalizeHeader(tenCH).includes(filterSt)) continue;

    // Lọc Khu vực
    const maNV = normalizeHeader(String(r[idxMaNV] || ''));
    let khuVuc = 'Chưa phân bổ KV';
    if (maNV && staffRegionMap.has(maNV)) {
      khuVuc = staffRegionMap.get(maNV);
    }
    if (filterReg && khuVuc !== filterReg) continue;

    const giaCuoc = parseCurrency(r[idxGiaCuoc]);
    const chuKy = parseInt(r[idxChuKy], 10) || 1;
    let dtpk = parseCurrency(r[idxDtpk]);
    if (!dtpk && giaCuoc > 0) dtpk = Math.round(giaCuoc / chuKy);

    totalRev += giaCuoc;
    totalDtpk += dtpk;

    let maKH = String(r[idxMaKH] || '').replace(/^'+/, '').trim();
    if (/^\d{8}$/.test(maKH)) maKH = '00' + maKH;

    const dBatDau = parseToDate(r[idxNgayBatDau]);
    const dKetThuc = parseToDate(r[idxNgayKetThuc]);
    const batDauStr = dBatDau ? Utilities.formatDate(dBatDau, 'Asia/Saigon', 'dd/MM/yyyy') : '';
    const ketThucStr = dKetThuc ? Utilities.formatDate(dKetThuc, 'Asia/Saigon', 'dd/MM/yyyy') : '';

    matchedRows.push([
      matchedRows.length + 1,
      khuVuc,
      maKH,
      r[idxTenTB],
      r[idxSDT],
      r[idxDiaChi],
      goiCuoc,
      giaCuoc,
      chuKy,
      dtpk,
      batDauStr,
      ketThucStr,
      r[idxTinhTrangHD],
      r[idxChanCat],
      r[idxLoaiONT],
      r[idxSerial],
      r[idxMaCH],
      tenCH,
      r[idxMaNV],
      r[idxTenNV]
    ]);
  }

  // 3. KẾT XUẤT RA SHEET MỚI
  const timeSuffix = Utilities.formatDate(new Date(), 'Asia/Saigon', 'ddMM_HHmm');
  const targetSheetName = `TrichXuat_${timeSuffix}`;
  const reportSheet = ss.insertSheet(targetSheetName);

  const nowFormatted = Utilities.formatDate(new Date(), 'Asia/Saigon', 'dd/MM/yyyy HH:mm');

  // Header & Card KPI
  const headerMatrix = [
    [`DANH SÁCH THUÊ BAO TRÍCH XUẤT ĐA ĐIỀU KIỆN (${timeSuffix})`, '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    [
      `🎯 TỔNG KẾT QUẢ: ${matchedRows.length} TB | DOANH THU: ${totalRev.toLocaleString('vi-VN')} đ | DTPK: ${totalDtpk.toLocaleString('vi-VN')} đ`,
      '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', `🕒 Lập lúc: ${nowFormatted}`, ''
    ],
    [
      'STT', 'KHU VỰC', 'MÃ KHÁCH HÀNG', 'TÊN THUÊ BAO', 'SĐT LIÊN HỆ', 'ĐỊA CHỈ LẮP ĐẶT',
      'GÓI CƯỚC', 'GIÁ CƯỚC', 'CHU KỲ', 'DTPK', 'NGÀY BẮT ĐẦU', 'NGÀY KẾT THÚC',
      'TÌNH TRẠNG HĐ', 'TÌNH TRẠNG CHẶN', 'LOẠI ONT', 'SERIAL ONT',
      'MÃ CỬA HÀNG', 'TÊN CỬA HÀNG', 'MÃ NHÂN VIÊN', 'NHÂN VIÊN QUẢN LÝ'
    ]
  ];

  reportSheet.getRange(1, 1, 3, 20).setValues(headerMatrix);

  // Định dạng Header
  reportSheet.getRange('A1:T1').merge().setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center');
  reportSheet.getRange('A2:R2').merge().setBackground('#e8f0fe').setFontColor('#174ea6').setFontWeight('bold');
  reportSheet.getRange('S2:T2').merge().setBackground('#e8f0fe').setFontColor('#5f6368').setHorizontalAlignment('right');
  reportSheet.getRange('A3:T3').setBackground('#34a853').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');

  // Ghi dữ liệu
  if (matchedRows.length > 0) {
    reportSheet.getRange(4, 1, matchedRows.length, 20).setValues(matchedRows);

    // Format số & căn lề
    reportSheet.getRange(4, 1, matchedRows.length, 1).setHorizontalAlignment('center');
    reportSheet.getRange(4, 3, matchedRows.length, 1).setNumberFormat('@').setHorizontalAlignment('center');
    reportSheet.getRange(4, 8, matchedRows.length, 1).setNumberFormat('#,##0" đ"');
    reportSheet.getRange(4, 9, matchedRows.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('center');
    reportSheet.getRange(4, 10, matchedRows.length, 1).setNumberFormat('#,##0" đ"');
    reportSheet.getRange(4, 11, matchedRows.length, 2).setHorizontalAlignment('center');
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
  reportSheet.setColumnWidth(10, 110);

  const duration = ((new Date().getTime() - startTime) / 1000).toFixed(2);
  Logger.log(`⚡ Trích xuất hoàn tất ${matchedRows.length} TB trong ${duration}s!`);
  ss.setActiveSheet(reportSheet);

  return matchedRows.length;
}