// ============================================================
// 選修費系統 — 繳費回報 Google Apps Script
// 把這段程式碼貼到 Google 試算表 → 擴充功能 → Apps Script
// ============================================================

/**
 * doGet(e)
 * 如果帶有 seat 參數 → 處理繳費回報
 * 如果沒帶參數 → 健康檢查
 */
function doGet(e) {
  var params = e.parameter || {};

  // 沒帶參數 = 健康檢查
  if (!params.seat) {
    return jsonResponse({ status: 'ok', message: '繳費回報 API 運作中' });
  }

  // 有帶參數 = 繳費回報
  try {
    var seat = String(params.seat || '').trim();
    var name = String(params.name || '').trim();
    var method = String(params.method || '').trim();
    var lastFive = String(params.lastFive || '').trim();

    // 基本驗證
    if (!seat || !name || !method) {
      return jsonResponse({ success: false, message: '資料不完整，請確認座號、姓名和付款方式' });
    }

    // 驗證座號是否存在（從「選修費」頁籤比對）
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName('選修費');
    if (!mainSheet) {
      return jsonResponse({ success: false, message: '找不到「選修費」頁籤' });
    }

    var mainData = mainSheet.getDataRange().getValues();
    var headers = mainData[0];
    var seatCol = -1, nameCol = -1;
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]);
      if (h.includes('座號')) seatCol = i;
      if (h.includes('姓名')) nameCol = i;
    }
    if (seatCol === -1 || nameCol === -1) {
      return jsonResponse({ success: false, message: '試算表格式異常，找不到座號或姓名欄' });
    }

    // 比對座號+姓名
    var verified = false;
    for (var r = 1; r < mainData.length; r++) {
      var rowSeat = String(Math.round(Number(mainData[r][seatCol])));
      var rowName = String(mainData[r][nameCol]).trim();
      if (rowSeat === seat && rowName === name) {
        verified = true;
        break;
      }
    }
    if (!verified) {
      return jsonResponse({ success: false, message: '座號與姓名不符，請確認後重新送出' });
    }

    // 寫入「繳費回報」頁籤
    var reportSheet = ss.getSheetByName('繳費回報');
    if (!reportSheet) {
      reportSheet = ss.insertSheet('繳費回報');
      reportSheet.getRange('A1').setValue('12/5').setFontWeight('bold').setBackground('#FAEDEB');
      reportSheet.appendRow(['座號', '姓名', '付款方式', '匯款後五碼', '回報時間']);
    }

    // 取得台灣時間
    var now = new Date();
    var twTime = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');

    // 寫入新一行
    reportSheet.appendRow([
      parseInt(seat),
      name,
      method,
      lastFive || '（無）',
      twTime
    ]);

    return jsonResponse({
      success: true,
      message: name + '（座號 ' + seat + '）的繳費回報已成功送出！'
    });

  } catch (err) {
    return jsonResponse({ success: false, message: '系統錯誤：' + err.message });
  }
}

/**
 * 回傳 JSON 格式的回應
 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
