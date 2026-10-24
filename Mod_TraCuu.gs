/**
 * ============================================================================
 * HỆ THỐNG QUẢN LÝ DỮ LIỆU MOBIFIBER - TTKD VIỄN THÔNG LÂM ĐỒNG
 * FILE        : Mod_TraCuu.gs
 * PHIÊN BẢN   : 3.0 (High-Performance Engine: In-Memory Multi-Field Search 360°)
 * NGÀY TẠO    : 20/09/2026 16:30:00 (GMT+7)
 * NGÀY CẬP NHẬT: 24/10/2026 19:55:00 (GMT+7)
 * TÁC GIẢ     : Trịnh Vũ Thế Vinh & AI Assistant
 * MÔ TẢ       : Tra cứu hồ sơ thuê bao 360° siêu tốc theo Mã KH, Số điện thoại hoặc Serial ONT
 * ============================================================================
 */

function showCustomerLookupDialog() {
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; padding: 18px; margin: 0; font-size: 13px; color: #202124; background: #fff; }
        .search-box { display: flex; gap: 8px; margin-bottom: 14px; }
        input[type="text"] { flex: 1; padding: 9px 12px; border: 1px solid #dadce0; border-radius: 6px; font-size: 14px; outline: none; }
        input[type="text"]:focus { border-color: #1a73e8; box-shadow: 0 0 0 2px rgba(26,115,232,0.2); }
        button { padding: 8px 18px; border-radius: 6px; font-weight: 600; cursor: pointer; border: none; font-size: 13px; }
        .btn-search { background: #1a73e8; color: #fff; }
        .btn-search:disabled { background: #dadce0; cursor: not-allowed; }
        .hint { font-size: 11px; color: #70757a; margin-top: -8px; margin-bottom: 12px; }
        #resultContainer { max-height: 420px; overflow-y: auto; display: none; }
        .card { border: 1px solid #dadce0; border-radius: 8px; padding: 14px; margin-bottom: 12px; background: #fafafa; }
        .card-header { display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #e8eaed; padding-bottom: 8px; margin-bottom: 10px; }
        .cust-name { font-weight: bold; font-size: 15px; color: #1a73e8; }
        .badge { font-size: 11px; padding: 3px 8px; border-radius: 4px; font-weight: bold; }
        .badge-active { background: #e6f4ea; color: #137333; }
        .badge-inactive { background: #fce8e6; color: #c5221f; }
        .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 8px; font-size: 12px; }
        .info-item span:first-child { color: #5f6368; display: inline-block; width: 110px; }
        .info-item span:last-child { font-weight: 600; color: #202124; }
        .empty-msg { text-align: center; color: #70757a; padding: 20px; font-style: italic; }
      </style>
    </head>
    <body>
      <div class="search-box">
        <input type="text" id="queryInput" placeholder="Nhập Mã KH, Số điện thoại hoặc Serial ONT..." onkeydown="handleKey(event)">
        <button class="btn-search" id="btnSearch" onclick="doSearch()">Tra cứu</button>
      </div>
      <div class="hint">⚡ Tìm kiếm nhanh không phân biệt chữ hoa, chữ thường hoặc khoảng trắng thừa.</div>

      <div id="loading" style="display:none; text-align:center; padding: 20px; color: #1a73e8; font-weight:600;">
        ⏳ Đang quét cơ sở dữ liệu...
      </div>

      <div id="resultContainer"></div>

      <script>
        function handleKey(e) {
          if (e.key === 'Enter') doSearch();
        }

        function doSearch() {
          const q = document.getElementById('queryInput').value.trim();
          if (!q) {
            alert('Vui lòng nhập từ khóa cần tra cứu!');
            return;
          }

          document.getElementById('btnSearch').disabled = true;
          document.getElementById('loading').style.display = 'block';
          document.getElementById('resultContainer').style.display = 'none';

          google.script.run
            .withSuccessHandler(function(results) {
              document.getElementById('loading').style.display = 'none';
              document.getElementById('btnSearch').disabled = false;
              renderResults(results);
            })
            .withFailureHandler(function(err) {
              document.getElementById('loading').style.display = 'none';
              document.getElementById('btnSearch').disabled = false;
              alert('Lỗi tra cứu: ' + err.message);
            })
            .searchSubscriber360(q);
        }

        function renderResults(list) {
          const box = document.getElementById('resultContainer');
          box.innerHTML = '';
          box.style.display = 'block';

          if (!list || list.length === 0) {
            box.innerHTML = '<div class="empty-msg">❌ Không tìm thấy thuê bao nào khớp với từ khóa đã nhập.</div>';
            return;
          }

          list.forEach(item => {
            const isActive = !item.tinhTrangHD.toLowerCase().includes('hủy') && !item.tinhTrangHD.toLowerCase().includes('huy');
            const badgeClass = isActive ? 'badge-active' : 'badge-inactive';

            const card = document.createElement('div');
            card.className = 'card';
            card.innerHTML = \`
              <div class="card-header">
                <span class="cust-name">\${item.tenTB || 'Chưa có tên'}</span>
                <span class="badge \${badgeClass}">\${item.tinhTrangHD || 'Chưa rõ'}</span>
              </div>
              <div class="info-grid">
                <div class="info-item"><span>Mã KH:</span> <span>\${item.maKH}</span></div>
                <div class="info-item"><span>Số điện thoại:</span> <span>\${item.sdt || '---'}</span></div>
                <div class="info-item"><span>Gói cước:</span> <span>\${item.goiCuoc || '---'}</span></div>
                <div class="info-item"><span>Giá cước:</span> <span>\${Number(item.giaCuoc || 0).toLocaleString('vi-VN')} đ</span></div>
                <div class="info-item"><span>Chu kỳ:</span> <span>\${item.chuKy || 1} tháng</span></div>
                <div class="info-item"><span>DTPK:</span> <span>\${Number(item.dtpk || 0).toLocaleString('vi-VN')} đ</span></div>
                <div class="info-item"><span>Bắt đầu gói:</span> <span>\${item.ngayBatDau || '---'}</span></div>
                <div class="info-item"><span>Hết hạn gói:</span> <span>\${item.ngayKetThuc || '---'}</span></div>
                <div class="info-item"><span>Tình trạng chặn:</span> <span>\${item.tinhTrangChan || 'Bình thường'}</span></div>
                <div class="info-item"><span>Serial ONT:</span> <span>\${item.serial || '---'}</span></div>
                <div class="info-item" style="grid-column: 1 / -1;"><span>Địa chỉ:</span> <span>\${item.diaChi || '---'}</span></div>
                <div class="info-item" style="grid-column: 1 / -1;"><span>Phụ trách:</span> <span>\${item.tenNV || '---'} (\${item.tenCH || '---'})</span></div>
              </div>
            \`;
            box.appendChild(card);
          });
        }
      </script>
    </body>
    </html>
  `;
  const ui = HtmlService.createHtmlOutput(html).setWidth(600).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(ui, '🔍 Tra cứu thông tin thuê bao 360°');
}

/**
 * THUẬT TOÁN TÌM KIẾM SIÊU TỐC TRÊN RAM THEO ĐA TRƯỜNG
 */
function searchSubscriber360(keyword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(CONFIG.TARGET_SHEET_NAME);
  if (!rawSheet) throw new Error(`Không tìm thấy sheet "${CONFIG.TARGET_SHEET_NAME}"!`);

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();
  if (lastRow < CONFIG.DATA_START_ROW_TARGET) return [];

  // Chuẩn hóa từ khóa tìm kiếm
  const qClean = String(keyword || '').replace(/[\s\uFEFF\xA0]+/g, '').toLowerCase();
  const qDigits = qClean.replace(/[^0-9]/g, '');

  // Đọc dữ liệu vào RAM
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
  const idxSerial = headers.findIndex(h => h === 'serial ont');
  const idxTenCH = headers.findIndex(h => h === 'tên cửa hàng');
  const idxTenNV = headers.findIndex(h => h === 'tên nhân viên');

  const rawData = rawSheet.getRange(CONFIG.DATA_START_ROW_TARGET, 1, lastRow - CONFIG.DATA_START_ROW_TARGET + 1, lastCol).getValues();
  const matched = [];

  for (let i = 0; i < rawData.length; i++) {
    const row = rawData[i];

    let maKHRaw = String(row[idxMaKH] || '').replace(/[\s\uFEFF\xA0]+/g, '').replace(/^'+/, '');
    if (/^\d{8}$/.test(maKHRaw)) maKHRaw = '00' + maKHRaw;

    const sdtRaw = String(row[idxSDT] || '').replace(/[^0-9]/g, '');
    const serialRaw = String(row[idxSerial] || '').replace(/[\s\uFEFF\xA0]+/g, '').toLowerCase();

    // Khớp 1 trong 3 tiêu chí: Mã KH, Số điện thoại hoặc Serial ONT
    const matchMaKH = qDigits && maKHRaw.includes(qDigits);
    const matchSDT = qDigits && sdtRaw.includes(qDigits);
    const matchSerial = qClean && serialRaw.includes(qClean);

    if (matchMaKH || matchSDT || matchSerial) {
      const dBatDau = parseToDate(row[idxNgayBatDau]);
      const dKetThuc = parseToDate(row[idxNgayKetThuc]);
      const giaCuoc = parseCurrency(row[idxGiaCuoc]);
      const chuKy = parseInt(row[idxChuKy], 10) || 1;
      let dtpk = parseCurrency(row[idxDtpk]);
      if (!dtpk && giaCuoc > 0) dtpk = Math.round(giaCuoc / chuKy);

      matched.push({
        maKH: maKHRaw,
        tenTB: row[idxTenTB],
        sdt: row[idxSDT],
        diaChi: row[idxDiaChi],
        goiCuoc: row[idxGoiCuoc],
        giaCuoc: giaCuoc,
        chuKy: chuKy,
        dtpk: dtpk,
        ngayBatDau: dBatDau ? Utilities.formatDate(dBatDau, 'Asia/Saigon', 'dd/MM/yyyy') : '',
        ngayKetThuc: dKetThuc ? Utilities.formatDate(dKetThuc, 'Asia/Saigon', 'dd/MM/yyyy') : '',
        tinhTrangHD: String(row[idxTinhTrangHD] || 'Đang hoạt động'),
        tinhTrangChan: String(row[idxChanCat] || 'Bình thường'),
        serial: row[idxSerial],
        tenCH: row[idxTenCH],
        tenNV: row[idxTenNV]
      });

      // Giới hạn hiển thị 20 kết quả gần nhất để giao diện tải nhanh
      if (matched.length >= 20) break;
    }
  }

  return matched;
}