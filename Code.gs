// ============================================================
// 選修費系統 — 繳費回報 Google Apps Script
// 把這段程式碼貼到 Google 試算表 → 擴充功能 → Apps Script
// ============================================================

/**
 * doPost(e)
 * 接收前端 POST 過來的繳費回報資料，寫入「繳費回報」頁籤
 */
function doPost(e) {
  try {
    // 解析 POST 資料
    const data = JSON.parse(e.postData.contents);
    const seat = String(data.seat || '').trim();
    const name = String(data.name || '').trim();
    const method = String(data.method || '').trim();
    const lastFive = String(data.lastFive || '').trim();

    // 基本驗證
    if (!seat || !name || !method) {
      return jsonResponse({ success: false, message: '資料不完整，請確認座號、姓名和付款方式' });
    }

    // 驗證座號是否存在（從「選修費」頁籤比對）
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = ss.getSheetByName('選修費');
    if (!mainSheet) {
      return jsonResponse({ success: false, message: '找不到「選修費」頁籤' });
    }

    const mainData = mainSheet.getDataRange().getValues();
    // 找座號欄和姓名欄
    const headers = mainData[0];
    let seatCol = -1, nameCol = -1;
    for (let i = 0; i < headers.length; i++) {
      const h = String(headers[i]);
      if (h.includes('座號')) seatCol = i;
      if (h.includes('姓名')) nameCol = i;
    }
    if (seatCol === -1 || nameCol === -1) {
      return jsonResponse({ success: false, message: '試算表格式異常，找不到座號或姓名欄' });
    }

    // 比對座號+姓名
    let verified = false;
    for (let r = 1; r < mainData.length; r++) {
      const rowSeat = String(Math.round(Number(mainData[r][seatCol])));
      const rowName = String(mainData[r][nameCol]).trim();
      if (rowSeat === seat && rowName === name) {
        verified = true;
        break;
      }
    }
    if (!verified) {
      return jsonResponse({ success: false, message: '座號與姓名不符，請確認後重新送出' });
    }

    // 寫入「繳費回報」頁籤
    let reportSheet = ss.getSheetByName('繳費回報');
    if (!reportSheet) {
      // 自動建立頁籤和標題列
      reportSheet = ss.insertSheet('繳費回報');
      reportSheet.appendRow(['座號', '姓名', '付款方式', '匯款後五碼', '回報時間', '', '12/5']);
      // 設定標題格式
      const headerRange = reportSheet.getRange(1, 1, 1, 5);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#F5E6D8');
      // G1 = 繳費截止日（改這格就能更新網頁上的截止日）
      const deadlineCell = reportSheet.getRange('G1');
      deadlineCell.setFontWeight('bold');
      deadlineCell.setBackground('#FAEDEB');
      reportSheet.getRange('F1').setValue('繳費截止日→').setFontColor('#999999');
    }

    // 取得台灣時間
    const now = new Date();
    const twTime = Utilities.formatDate(now, 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');

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
      message: `${name}（座號 ${seat}）的繳費回報已成功送出！`
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

/**
 * doGet(e) — 簡單的健康檢查
 */
function doGet(e) {
  return jsonResponse({ status: 'ok', message: '繳費回報 API 運作中' });
}
