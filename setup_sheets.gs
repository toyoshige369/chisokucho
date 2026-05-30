/**
 * 知足帳 スプレッドシート初期化スクリプト
 * 実行方法：GASエディタで setupSheets() を選択して実行
 * 警告：既存データはすべて削除されます
 */
function setupSheets() {
  var ss = SpreadsheetApp.openById('1WBKWY3SP6zcJn9mYGTrK3sgFL4EXcmOVTL_yVxVcZTQ');

  var schemas = {
    'mood':     ['id', 'recorded_at', 'score'],
    'memo':     ['id', 'created_at', 'updated_at', 'memo_date', 'cat', 'text', 'excludeDiary'],
    'diary':    ['id', 'created_at', 'updated_at', 'diary_date', 'prompt', 'answer', 'good1', 'good2', 'good3'],
    'task':     ['id', 'created_at', 'updated_at', 'name', 'priority', 'due_date', 'person', 'done', 'completed_at'],
    'schedule': ['id', 'created_at', 'updated_at', 'event_date', 'start_time', 'end_time', 'all_day', 'person', 'name', 'sub', 'color'],
    'deed':     ['id', 'created_at', 'updated_at', 'deed_date', 'year', 'text'],
    'love':     ['id', 'created_at', 'updated_at', 'name', 'cat', 'color', 'textColor', 'pinned', 'taps', 'memo'],
    'mind':     ['id', 'created_at', 'updated_at', 'title', 'sub', 'url', 'cat', 'icon', 'taps', 'states'],
    'settings': ['key', 'value', 'updated_at']
  };

  Object.keys(schemas).forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      Logger.log('シート作成: ' + name);
    } else {
      sheet.clearContents();
      Logger.log('シートクリア: ' + name);
    }
    var headers = schemas[name];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    // ヘッダー行を太字・背景色で見やすく
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    // 列幅を自動調整
    sheet.autoResizeColumns(1, headers.length);
  });

  Logger.log('セットアップ完了');
  return {ok: true};
}

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
