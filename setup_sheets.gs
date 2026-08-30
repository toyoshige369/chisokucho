function autoResizeAllSheets() {
  var ss = SpreadsheetApp.openById('1WBKWY3SP6zcJn9mYGTrK3sgFL4EXcmOVTL_yVxVcZTQ');
  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    var lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      sheet.autoResizeColumns(1, lastCol);
      Logger.log('調整完了: ' + sheet.getName());
    }
  });
  Logger.log('全シート調整完了');
}

/**
 * 毎日0時に autoResizeAllSheets を実行するトリガーを設定する
 * 既存の同名トリガーは削除してから再登録（重複防止）
 */
function setDailyResizeTrigger() {
  // 既存の同名トリガーを削除（重複防止）
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'autoResizeAllSheets') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  // 毎日0時に実行するトリガーを新規作成
  ScriptApp.newTrigger('autoResizeAllSheets')
    .timeBased()
    .atHour(0)
    .everyDays(1)
    .create();
  Logger.log('自動調整トリガー設定完了');
}
