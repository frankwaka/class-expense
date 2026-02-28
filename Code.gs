// ============================================================
// 選修費系統 — 繳費回報 Google Apps Script
// 把這段程式碼貼到 Google 試算表 → 擴充功能 → Apps Script
// ============================================================

function doGet(e) {
  var params = e.parameter || {};
  if (!params.seat) {
    return ContentService.createTextOutput(JSON.stringify({status:'ok'})).setMimeType(ContentService.MimeType.JSON);
  }
  try {
    var seat = String(params.seat).trim();
    var name = String(params.name).trim();
    var method = String(params.method).trim();
    var lastFive = String(params.lastFive || '').trim();
    var amount = String(params.amount || '0').trim();
    if (!seat || !name || !method) {
      return ContentService.createTextOutput(JSON.stringify({success:false})).setMimeType(ContentService.MimeType.JSON);
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var reportSheet = ss.getSheetByName('繳費回報');
    var now = new Date();
    var twTime = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
    reportSheet.appendRow([parseInt(seat), name, method, lastFive || '（無）', parseInt(amount), twTime]);
    return ContentService.createTextOutput(JSON.stringify({success:true})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({success:false, error:err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}
