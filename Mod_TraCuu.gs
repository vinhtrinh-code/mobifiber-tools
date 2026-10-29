/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * MODULE       : TRA CỨU THÔNG TIN THUÊ BAO (360°)
 * PHIÊN BẢN    : 3.0 (Đồng bộ UI/UX với WebApp Dashboard, Fix lỗi parseCurrency)
 * NGÀY CẬP NHẬT: 29/10/2026
 * TÁC GIẢ      : Trịnh Vũ Thế Vinh & AI Assistant
 * ============================================================================
 */

function showCustomerLookupDialog() {
  const html = HtmlService.createHtmlOutput(getCustomerLookupHtml())
    .setWidth(750)
    .setHeight(560)
    .setTitle('Tra cứu Hồ sơ Khách hàng 360°');
  SpreadsheetApp.getUi().showModalDialog(html, 'Hồ sơ Khách hàng MobiFiber 360°');
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
  const idxMaNV = headers.findIndex(h => h === 'mã nhân viên'); // Dùng để xác định Khu Vực
  const idxUser = headers.findIndex(h => h === 'user');
  
  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, lastRow - CONFIG.DATA_START_ROW_TARGET + 1, lastCol).getValues();
  const cleanKw = String(keyword || '').trim().toLowerCase().replace(/^0+/, '');
  
  // Tải tham số Khu vực để gán cho chi tiết tra cứu
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
  
  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];
    const maKH = normalizeCustomerId(row[idxMaKH]);
    const sdt  = String(row[idxSDT] || '').replace(/^0+/, '').trim();
    const cccd = String(row[idxCCCD] || '').replace(/^0+/, '').trim();
    const user = idxUser !== -1 ? String(row[idxUser] || '').trim().toLowerCase() : '';
  
    // Hỗ trợ tra cứu bằng Mã KH, SĐT, CCCD hoặc User
    if (maKH === cleanKw || sdt === cleanKw || cccd === cleanKw || (user && user.includes(cleanKw))) {
      const record = {};
      
      headers.forEach((h, colIdx) => {
        let val = row[colIdx];
        
        // Chuẩn hóa định dạng hiển thị cho Tiền tệ và Ngày tháng
        if (h === 'giá cước' || h === 'dtpk') {
          val = parseCurrencyToInteger(val) || 0;
        } else if (CONFIG.DATETIME_FIELDS && CONFIG.DATETIME_FIELDS.map(f=>normalizeHeader(f)).includes(h)) {
          const d = parseToDate(val);
          val = d ? Utilities.formatDate(d, 'Asia/Saigon', 'dd/MM/yyyy') : '---';
        } else if (val instanceof Date) {
          val = Utilities.formatDate(val, 'Asia/Saigon', 'dd/MM/yyyy');
        }
        
        record[h] = val;
      });
      
      // Ánh xạ Khu Vực
      let khuVuc = 'TT.KDVT';
      if (idxMaNV !== -1) {
        const rawMaNV = normalizeHeader(String(row[idxMaNV] || ''));
        if (rawMaNV && staffRegionMap.has(rawMaNV)) {
          khuVuc = staffRegionMap.get(rawMaNV);
        }
      }
      record['khu vực'] = khuVuc;
      
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
  <!-- Tích hợp Tailwind CSS để đồng bộ Giao diện với WebApp -->
  <link href="https://cdnjs.cloudflare.com/ajax/libs/tailwindcss/2.2.19/tailwind.min.css" rel="stylesheet">
  <style>
    body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; padding: 20px; margin: 0; background: #f8f9fa; color: #202124; font-size: 13px; }
    
    .search-big-input { width: 100%; padding: 14px 20px; font-size: 14px; border: 2px solid #cbd5e1; border-radius: 12px; outline: none; transition: 0.2s; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); }
    .search-big-input:focus { border-color: #2563eb; box-shadow: 0 0 0 4px rgba(37,99,235,0.1); }
    
    .detail-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 16px; padding: 20px; }
    .info-group { background: #f8fafc; padding: 16px; border-radius: 12px; border: 1px solid #e2e8f0; }
    .info-item { margin-bottom: 12px; display: flex; flex-direction: column; gap: 4px; }
    .info-item:last-child { margin-bottom: 0; }
    .info-label { font-size: 11px; color: #64748b; font-weight: 600; text-transform: uppercase; }
    .info-value { font-size: 14px; color: #0f172a; font-weight: 500; }
    
    .badge { display: inline-block; padding: 4px 10px; border-radius: 20px; font-size: 12px; font-weight: 600; }
    .badge-green { background: #dcfce7; color: #166534; }
    .badge-red { background: #fee2e2; color: #991b1b; }
    .badge-blue { background: #dbeafe; color: #1e40af; }
    
    .spinner-border { width: 30px; height: 30px; border: 3px solid #cbd5e1; border-top-color: #2563eb; border-radius: 50%; animation: spin 1s linear infinite; margin: 0 auto; }
    @keyframes spin { to { transform: rotate(360deg); } }
  </style>
</head>
<body>
  <div style="max-width: 100%; margin: 0 auto;">
    <h2 style="font-size: 20px; font-weight: bold; color: #0f172a; margin-bottom: 8px; text-align: center;">Hồ sơ Khách hàng (360°)</h2>
    <p style="color: #64748b; margin-bottom: 20px; text-align: center;">Hỗ trợ tìm kiếm theo Mã KH, Số điện thoại, Số CCCD hoặc User.</p>
    
    <input type="text" id="keyword" class="search-big-input" placeholder="Nhập từ khóa tìm kiếm (nhấn Enter để tìm)..." onkeypress="handleSearchMain(event)">
    
    <div id="searchResultsMain" style="margin-top: 24px; text-align: left;"></div>
  </div>

  <script>
    window.onload = function() {
      const inp = document.getElementById('keyword');
      if (inp) inp.focus();
    };

    function formatVND(val) {
      if (!val) return '0 đ';
      return Math.round(val).toLocaleString('vi-VN') + ' đ';
    }

    function handleSearchMain(e) {
      if (e.key === 'Enter') {
        const val = document.getElementById('keyword').value.trim();
        if (val.length < 3) {
          alert('Vui lòng nhập ít nhất 3 ký tự (Mã KH, Số ĐT, User hoặc CCCD).');
          return;
        }
        
        document.getElementById('searchResultsMain').innerHTML = '<div style="text-align: center; padding: 40px;"><div class="spinner-border"></div><div style="margin-top:12px; color:#64748b; font-weight: 500;">Đang tra cứu dữ liệu máy chủ...</div></div>';
        
        google.script.run
          .withSuccessHandler(function(res) {
            if (res.status === 'not_found' || res.status === 'empty') {
              document.getElementById('searchResultsMain').innerHTML = '<div style="text-align: center; color: #dc2626; padding: 40px; font-weight: bold;">❌ Không tìm thấy thông tin thuê bao khớp với từ khóa!</div>';
              return;
            }
            
            const item = res.data;
            let badgeChan = (item['tình trạng chặn cắt'] || 'Bình thường') === 'Bình thường' ? 'badge-green' : 'badge-red';
            let badgeHD = String(item['tình trạng hợp đồng'] || '').includes('Đang hoạt động') ? 'badge-blue' : 'badge-red';
            
            // Xử lý giá trị trống
            const maKH = item['mã khách hàng'] || '';
            const user = item['user'] || '';
            const tenTB = item['tên thuê bao'] || 'Chưa cập nhật';
            const sdt = item['sđt liên hệ'] || '---';
            const khuVuc = item['khu vực'] || '---';
            const diaChi = item['địa chỉ lắp đặt'] || '---';
            const goiCuoc = item['gói cước'] || '---';
            const chuKy = item['chu kỳ sử dụng'] || '1';
            const giaCuoc = item['giá cước'] || 0;
            const dtpk = item['dtpk'] || 0;
            const dBatDau = item['ngày bắt đầu gói cước'] || item['ngày băt đầu gói cước'] || '---';
            const dKetThuc = item['ngày kết thúc gói cước'] || '---';
            const loaiONT = item['loại thiết bị ont'] || '---';
            const serial = item['serial ont'] || '---';
            const tenCH = item['tên cửa hàng'] || '---';
            const tenNV = item['tên nhân viên'] || '---';

            // Dựng Layout đồng bộ hoàn toàn với bản WebApp
            let html = \`
            <div style="background: white; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
              <div style="background: #f8fafc; padding: 16px 20px; border-bottom: 1px solid #e2e8f0; display: flex; justify-content: space-between; align-items: center;">
                <div>
                  <div style="font-weight: bold; color: #0f172a; font-size: 16px;">\${tenTB}</div>
                  <div style="font-family: monospace; color: #2563eb; font-size: 13px; font-weight: 600; margin-top: 4px;">KH: \${maKH} \${user ? ' | User: ' + user : ''}</div>
                </div>
                <div style="text-align: right; display: flex; flex-direction: column; gap: 6px; align-items: flex-end;">
                  <span class="badge \${badgeHD}">\${item['tình trạng hợp đồng'] || '---'}</span>
                  <span class="badge \${badgeChan}">\${item['tình trạng chặn cắt'] || 'Bình thường'}</span>
                </div>
              </div>
              
              <div class="detail-grid">
                <div class="info-group">
                  <div class="info-item"><span class="info-label">Khu vực & Địa chỉ</span><span class="info-value">📍 \${khuVuc} - \${diaChi}</span></div>
                  <div class="info-item"><span class="info-label">Liên hệ</span><span class="info-value">📞 \${sdt}</span></div>
                  <div class="info-item"><span class="info-label">Gói cước</span><span class="info-value">📦 \${goiCuoc} (\${chuKy}T) - \${formatVND(giaCuoc)}</span></div>
                  <div class="info-item"><span class="info-label">Chu kỳ cước</span><span class="info-value">🗓️ \${dBatDau} ➔ \${dKetThuc}</span></div>
                </div>
                
                <div class="info-group">
                  <div class="info-item"><span class="info-label">Doanh thu phân kỳ</span><span class="info-value" style="color: #059669; font-size: 16px; font-weight: bold;">💰 \${formatVND(dtpk)} / tháng</span></div>
                  <div class="info-item"><span class="info-label">Thiết bị ONT</span><span class="info-value">⚙️ \${loaiONT} (SN: \${serial})</span></div>
                  <div class="info-item"><span class="info-label">Phát triển bởi</span><span class="info-value">🧑‍💼 \${tenNV}</span></div>
                  <div class="info-item"><span class="info-label">Điểm bán</span><span class="info-value">🏪 \${tenCH}</span></div>
                </div>
              </div>
            </div>\`;
            
            document.getElementById('searchResultsMain').innerHTML = html;
          })
          .withFailureHandler(function(err) {
            document.getElementById('searchResultsMain').innerHTML = '<div style="color: #dc2626; padding: 20px; text-align:center; font-weight:bold;">Lỗi tra cứu: ' + err.message + '</div>';
          })
          .searchCustomerRecord(val);
      }
    }
  </script>
</body>
</html>`;
}