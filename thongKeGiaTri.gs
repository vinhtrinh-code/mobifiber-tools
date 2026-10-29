function thongKeGiaTri() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('raw_mobifiber');
  const data = sh.getDataRange().getValues();
  const headers = data[1].map(h => String(h).trim().toUpperCase());
  const cols = ['NHU CẦU HÒA MẠNG','TÌNH TRẠNG HỢP ĐỒNG','TÌNH TRẠNG CHẶN CẮT',
                'CHU KỲ SỬ DỤNG','LOẠI KHÁCH HÀNG','LOẠI MÔ HÌNH TRIỂN KHAI'];
  const out = [['CỘT','GIÁ TRỊ (có dấu | để thấy khoảng trắng thừa)','SỐ DÒNG']];
  cols.forEach(c => {
    const i = headers.indexOf(c);
    if (i === -1) { out.push([c,'(không tìm thấy cột)','']); return; }
    const cnt = {};
    for (let r = 2; r < data.length; r++) {
      const v = '|' + String(data[r][i]) + '|';
      cnt[v] = (cnt[v] || 0) + 1;
    }
    Object.keys(cnt).sort().forEach(v => out.push([c, v, cnt[v]]));
  });
  const t = ss.getSheetByName('tmp_giatri') || ss.insertSheet('tmp_giatri');
  t.clear();
  t.getRange(1,1,out.length,3).setValues(out);
}