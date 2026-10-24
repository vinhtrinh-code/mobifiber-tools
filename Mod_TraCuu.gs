/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * MODULE       : TRA CỨU THÔNG TIN THUÊ BAO (360°)
 * PHIÊN BẢN    : 2.2 (Sửa lỗi cú pháp Unexpected token '<', Tra cứu nhanh Mã KH / SĐT / CCCD)
 * NGÀY CẬP NHẬT: 20/09/2026 20:00:00 (GMT+7)
 * TÁC GIẢ      : Trịnh Vũ Thế Vinh & AI Assistant
 * ============================================================================
 */

function showCustomerLookupDialog() {
  const html = HtmlService.createHtmlOutput(getCustomerLookupHtml())
    .setWidth(620)
    .setHeight(540)
    .setTitle('Tra cứu thông tin thuê bao MobiFiber');
  SpreadsheetApp.getUi().showModalDialog(html, 'Tra cứu thuê bao');
}

function searchCustomerRecord(keyword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) return { status: 'error', message: 'Không tìm thấy sheet raw_mobifiber!' };

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) return { status: 'empty' };

  const headers = rawSheet.getRange(CONFIG.HEADER_ROW_TARGET, 1, 1, lastCol)
    .getValues()[0].map(h => normalizeHeader(h));

  const idxMaKH = headers.findIndex(h => h === normalizeHeader(CONFIG.KEY_COLUMN_NAME));
  const idxSDT  = headers.findIndex(h => h === 'sđt liên hệ');
  const idxCCCD = headers.findIndex(h => h === 'số cmnd/cccd');

  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, lastRow - CONFIG.DATA_START_ROW_TARGET + 1, lastCol).getValues();
  const cleanKw = String(keyword || '').trim().toLowerCase().replace(/^0+/, '');

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const maKH = normalizeCustomerId(row[idxMaKH]);
    const sdt  = String(row[idxSDT] || '').replace(/^0+/, '').trim();
    const cccd = String(row[idxCCCD] || '').replace(/^0+/, '').trim();

    if (maKH === cleanKw || sdt === cleanKw || cccd === cleanKw) {
      const record = {};
      headers.forEach((h, colIdx) => {
        let val = row[colIdx];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, 'Asia/Saigon', 'dd/MM/yyyy HH:mm:ss');
        }
        record[h] = val;
      });
      return { status: 'found', data: record, rowNumber: CONFIG.DATA_START_ROW_TARGET + i };
    }
  }
  return { status: 'not_found' };
}

function getCustomerLookupHtml() {
  return `<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; padding: 14px; margin: 0; background: #f8f9fa; color: #202124; font-size: 13px; }
    .search-box { display: flex; gap: 8px; margin-bottom: 12px; }
    input { flex: 1; padding: 10px; border: 1px solid #dadce0; border-radius: 6px; font-size: 13px; outline: none; }
    input:focus { border-color: #1a73e8; }
    button { background: #1a73e8; color: white; border: none; padding: 10px 18px; border-radius: 6px; font-weight: bold; cursor: pointer; }
    button:hover { background: #1557b0; }
    .card { background: white; border: 1px solid #dadce0; border-radius: 8px; padding: 14px; margin-top: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
    .row { display: flex; margin-bottom: 8px; font-size: 12px; }
    .label { width: 140px; font-weight: 600; color: #5f6368; }
    .value { flex: 1; font-weight: 500; color: #202124; }
    .badge { display: inline-block; padding: 2px 8px; border-radius: 12px; font-size: 11px; font-weight: bold; }
    .badge-success { background: #e6f4ea; color: #137333; }
    .badge-danger { background: #fce8e6; color: #c5221f; }
  </style>
</head>
<body>
  <div class="search-box">
    <input type="text" id="keyword" placeholder="Nhập Mã khách hàng, SĐT hoặc số CCCD..." onkeydown="if(event.key==='Enter') doSearch()">
    <button type="button" onclick="doSearch()">Tìm kiếm</button>
  </div>
  <div id="result"></div>

  <script>
    window.onload = function() {
      const inp = document.getElementById('keyword');
      if (inp) inp.focus();
    };

    function doSearch() {
      const kw = document.getElementById('keyword').value.trim();
      if (!kw) return;
      document.getElementById('result').innerHTML = '<div style="text-align:center; padding:20px; color:#666;">⏳ Đang tra cứu dữ liệu...</div>';

      google.script.run
        .withSuccessHandler(function(res) {
          if (res.status === 'not_found' || res.status === 'empty') {
            document.getElementById('result').innerHTML = '<div style="text-align:center; padding:20px; color:#c5221f;">❌ Không tìm thấy thông tin thuê bao khớp với từ khóa!</div>';
            return;
          }
          const d = res.data;
          const statusVal = String(d['tình trạng hợp đồng'] || '').toLowerCase();
          const statusBadge = (statusVal.includes('hủy') || statusVal.includes('huy')) ? 
            '<span class="badge badge-danger">Đã hủy</span>' : '<span class="badge badge-success">Đang hoạt động</span>';

          const giaCuocFormatted = Number(d['giá cước'] || 0).toLocaleString('vi-VN');

          document.getElementById('result').innerHTML = 
            '<div class="card">' +
              '<div style="font-size:15px; font-weight:bold; color:#1a73e8; margin-bottom:10px; border-bottom:1px solid #eee; padding-bottom:6px;">' +
                '👤 ' + (d['tên thuê bao'] || 'Chưa có tên') + ' ' + statusBadge +
              '</div>' +
              '<div class="row"><div class="label">Mã khách hàng:</div><div class="value" style="color:#c5221f; font-weight:bold;">' + (d['mã khách hàng'] || '') + '</div></div>' +
              '<div class="row"><div class="label">Số điện thoại:</div><div class="value">' + (d['sđt liên hệ'] || '') + '</div></div>' +
              '<div class="row"><div class="label">Địa chỉ lắp đặt:</div><div class="value">' + (d['địa chỉ lắp đặt'] || '') + '</div></div>' +
              '<div class="row"><div class="label">Gói cước / Giá:</div><div class="value">' + (d['gói cước'] || '') + ' (Chu kỳ ' + (d['chu kỳ sử dụng'] || '') + 'T) - ' + giaCuocFormatted + ' đ</div></div>' +
              '<div class="row"><div class="label">Hạn gói cước:</div><div class="value" style="color:#d93025; font-weight:bold;">' + (d['ngày kết thúc gói cước'] || '') + '</div></div>' +
              '<div class="row"><div class="label">Tình trạng chặn cắt:</div><div class="value">' + (d['tình trạng chặn cắt'] || 'Bình thường') + '</div></div>' +
              '<div class="row"><div class="label">Thiết bị ONT:</div><div class="value">' + (d['loại thiết bị ont'] || '') + ' (Serial: ' + (d['serial ont'] || 'Chưa có') + ')</div></div>' +
              '<div class="row"><div class="label">Đơn vị quản lý:</div><div class="value">' + (d['tên cửa hàng'] || '') + ' - ' + (d['tên nhân viên'] || '') + '</div></div>' +
            '</div>';
        })
        .withFailureHandler(function(err) {
          document.getElementById('result').innerHTML = '<div style="text-align:center; padding:20px; color:#c5221f;">❌ Lỗi tra cứu: ' + err.message + '</div>';
        })
        .searchCustomerRecord(kw);
    }
  </script>
</body>
</html>`;
}