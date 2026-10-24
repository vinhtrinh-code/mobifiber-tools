/**
 * Tự động tạo menu trên Google Sheet để xuất toàn bộ code
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ Công cụ Vibe Coding')
    .addItem('Xuất toàn bộ Code ra Google Doc', 'exportAllScriptsToDoc')
    .addToUi();
}

/**
 * Thêm scope tường minh để Apps Script tự cấp quyền truy cập Script API
 * @OnlyCurrentDoc
 */
function exportAllScriptsToDoc() {
  var ui = SpreadsheetApp.getUi();
  var scriptId = ScriptApp.getScriptId();
  
  // Header xác thực bắt buộc kèm Scope
  var options = {
    method: "get",
    headers: {
      "Authorization": "Bearer " + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  };
  
  var url = "https://script.googleapis.com/v1/projects/" + scriptId + "/content";
  var response = UrlFetchApp.fetch(url, options);
  var resCode = response.getResponseCode();
  var resText = response.getContentText();
  
  if (resCode !== 200) {
    // Nếu vẫn vướng Scope của dự án, in chi tiết lỗi để xử lý ngay
    Logger.log("Mã lỗi HTTP: " + resCode + " | Nội dung: " + resText);
    ui.alert("Thông báo", "Google yêu cầu thêm Scope. Vui lòng kiểm tra Nhật ký thực thi!", ui.ButtonSet.OK);
    return;
  }
  
  var json = JSON.parse(resText);
  if (!json.files || json.files.length === 0) {
    ui.alert("Thông báo", "Không tìm thấy file nào trong dự án!", ui.ButtonSet.OK);
    return;
  }
  
  // Tạo file Google Doc mới
  var doc = DocumentApp.create("MobiFiber_Tools_Full_Code");
  var body = doc.getBody();
  body.appendParagraph("=== MOBIFIBER TOOLS - TOÀN BỘ MÃ NGUỒN ===").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  json.files.forEach(function(file) {
    body.appendParagraph("\n\n/* ==================================================");
    body.appendParagraph(" * FILE: " + file.name + " (" + file.type + ")");
    body.appendParagraph(" * ================================================== */");
    body.appendParagraph(file.source || "// File rỗng hoặc không có nội dung text");
  });
  
  doc.saveAndClose();
  ui.alert("Thành công!", "Đã xuất xong toàn bộ mã nguồn ra Google Doc:\n" + doc.getUrl(), ui.ButtonSet.OK);
}

// Dòng comment đặc biệt bên dưới ép Apps Script nhận diện Scope đọc code:
// https://www.googleapis.com/auth/script.projects.readonly