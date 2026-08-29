var SS_ID = '1WBKWY3SP6zcJn9mYGTrK3sgFL4EXcmOVTL_yVxVcZTQ';
var BACKUP_FOLDER_ID = '1IjQvJAtTQXWgdvfo5cXfH94MD4xO-Lys';
var SHEETS = ['memo','diary','mood','task','deed','love','mind','settings','schedule'];

/* タイムスタンプ生成ヘルパー */
function now() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}

function doGet(e) {
  if (!e.parameter.action) {
    return HtmlService.createHtmlOutputFromFile('index')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setHeight(812);
  }

  var action = e.parameter.action;
  var ss = SpreadsheetApp.openById(SS_ID);
  var result;

  try {
    if (action === 'getMood') {
      result = getMood(ss);
    } else if (action === 'saveMood') {
      result = saveMood(ss, e.parameter.score);
    } else if (action === 'getMemos') {
      result = getMemos(ss);
    } else if (action === 'saveMemo') {
      result = saveMemo(ss, e.parameter.id, e.parameter.memo_date || e.parameter.date, e.parameter.time, e.parameter.cat, e.parameter.text, e.parameter.excludeDiary, e.parameter.mood);
    } else if (action === 'getTasks') {
      result = getTasks(ss);
    } else if (action === 'saveTask') {
      result = saveTask(ss, e.parameter.id, e.parameter.name, e.parameter.priority, e.parameter.due_date, e.parameter.due_time, e.parameter.done);
    } else if (action === 'updateTask') {
      result = updateTask(ss, e.parameter.id, e.parameter.name, e.parameter.priority, e.parameter.due_date, e.parameter.due_time, e.parameter.done);
    } else if (action === 'deleteTask') {
      result = deleteTask(ss, e.parameter.id);
    } else if (action === 'getDiary') {
      result = getDiary(ss, e.parameter.date);
    } else if (action === 'saveDiary') {
      result = saveDiary(ss, e.parameter.date, e.parameter.prompt, e.parameter.answer, e.parameter.good1, e.parameter.good2, e.parameter.good3);
    } else if (action === 'getLove') {
      result = getLove(ss);
    } else if (action === 'saveLove') {
      result = saveLove(ss, e.parameter.id, e.parameter.name, e.parameter.cat, e.parameter.color, e.parameter.textColor, e.parameter.pinned, e.parameter.love_level);
    } else if (action === 'deleteLove') {
      result = deleteLove(ss, e.parameter.id);
    } else if (action === 'getEvents') {
      result = getEvents(ss, e.parameter.startDate, e.parameter.endDate);
    } else if (action === 'saveEvent') {
      result = saveEvent(ss, e.parameter.id, e.parameter.event_date, e.parameter.start_time, e.parameter.end_time, e.parameter.name, e.parameter.sub, e.parameter.color);
    } else if (action === 'deleteEvent') {
      result = deleteEvent(ss, e.parameter.id);
    } else if (action === 'getMind') {
      result = getMind(ss);
    } else if (action === 'saveMind') {
      result = saveMind(ss, e.parameter.id, e.parameter.title, e.parameter.sub, e.parameter.url, e.parameter.cat, e.parameter.icon, e.parameter.taps, e.parameter.states, e.parameter.from_ai);
    } else if (action === 'fetchYouTubeInfo') {
      result = fetchYouTubeInfo(e.parameter.url);
    } else if (action === 'deleteMemo') {
      result = deleteMemo(ss, e.parameter.id);
    } else if (action === 'deleteMind') {
      result = deleteMind(ss, e.parameter.id);
    } else if (action === 'getDeed') {
      result = getDeed(ss, e.parameter.deed_date);
    } else if (action === 'saveDeed') {
      result = saveDeed(ss, e.parameter.id, e.parameter.deed_date, e.parameter.year, e.parameter.text);
    } else if (action === 'deleteDeed') {
      result = deleteDeed(ss, e.parameter.id);
    } else {
      result = {error: 'unknown action'};
    }
  } catch(err) {
    result = {error: err.message};
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── mood: id / recorded_at / score ── */
function getMood(ss) {
  var sheet = ss.getSheetByName('mood');
  var rows = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    result.push({
      id: Number(rows[i][0]),
      recorded_at: String(rows[i][1]||''),
      score: Number(rows[i][2])
    });
  }
  result.sort(function(a,b){return Number(b.id)-Number(a.id);});
  return result;
}

function saveMood(ss, score) {
  var sheet = ss.getSheetByName('mood');
  sheet.insertRowBefore(2);
  sheet.getRange(2,1,1,3).setValues([[Date.now(), now(), Number(score)]]);
  return {ok: true, action: 'inserted'};
}

/* ── memo: id / created_at / updated_at / memo_date / cat / text / excludeDiary / mood ── */
function getMemos(ss) {
  var sheet = ss.getSheetByName('memo');
  var rows = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var moodRaw = rows[i][7];
    result.push({
      id: Number(rows[i][0]),
      created_at: String(rows[i][1]||''),
      updated_at: String(rows[i][2]||''),
      memo_date: (function(){ var r=rows[i][3]; return r instanceof Date ? Utilities.formatDate(r,'Asia/Tokyo','yyyy-MM-dd') : String(r||''); })(),
      cat: String(rows[i][4]||''),
      text: String(rows[i][5]||''),
      excludeDiary: rows[i][6]===true||rows[i][6]==='TRUE',
      mood: (moodRaw!==''&&moodRaw!==null&&moodRaw!==undefined) ? Number(moodRaw) : null
    });
  }
  result.sort(function(a,b){return Number(b.id)-Number(a.id);});
  return result;
}

function deleteMemo(ss, id) {
  var sheet = ss.getSheetByName('memo');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (Number(rows[i][0]) === Number(id)) {
      sheet.deleteRow(i + 1);
      return {ok: true};
    }
  }
  return {ok: false, error: 'not found'};
}

function saveMemo(ss, id, date, time, cat, text, excludeDiary, mood) {
  var sheet = ss.getSheetByName('memo');
  /* H1ヘッダーがmoodでなければ設定 */
  if (sheet.getRange(1, 8).getValue() !== 'mood') {
    sheet.getRange(1, 8).setValue('mood');
  }
  /* memo_dateが空の場合は今日の日付（Asia/Tokyo）を自動設定 */
  var memoDate = date || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  var moodVal = (mood !== undefined && mood !== '' && mood !== null) ? Number(mood) : '';
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (Number(rows[i][0]) === Number(id)) {
      sheet.getRange(i+1, 3).setValue(now());
      sheet.getRange(i+1, 4).setValue(memoDate);
      sheet.getRange(i+1, 5).setValue(cat||'');
      sheet.getRange(i+1, 6).setValue(text||'');
      sheet.getRange(i+1, 7).setValue(excludeDiary==='true');
      if (moodVal !== '') sheet.getRange(i+1, 8).setValue(moodVal);
      return {ok: true, action: 'updated'};
    }
  }
  sheet.insertRowBefore(2);
  sheet.getRange(2,1,1,8).setValues([[Number(id), now(), now(), memoDate, cat||'', text||'', excludeDiary==='true', moodVal]]);
  return {ok: true, action: 'inserted'};
}

/* ── task: id / created_at / updated_at / name / priority / due_date / due_time / person / done / completed_at ── */
function getTasks(ss) {
  var sheet = ss.getSheetByName('task');
  var rows = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    result.push({
      id: Number(rows[i][0]),
      created_at: String(rows[i][1]||''),
      updated_at: String(rows[i][2]||''),
      name: String(rows[i][3]||''),
      priority: String(rows[i][4]||'★★'),
      due_date: rows[i][5] instanceof Date ? Utilities.formatDate(rows[i][5],'Asia/Tokyo','yyyy-MM-dd') : String(rows[i][5]||''),
      due_time: rows[i][6] instanceof Date ? Utilities.formatDate(rows[i][6],'Asia/Tokyo','HH:mm') : String(rows[i][6]||''),
      person: String(rows[i][7]||'self'),
      done: rows[i][8]===true||rows[i][8]==='TRUE',
      completed_at: String(rows[i][9]||'')
    });
  }
  result.sort(function(a,b){return Number(b.id)-Number(a.id);});
  return result;
}

function saveTask(ss, id, name, priority, due_date, due_time) {
  var sheet = ss.getSheetByName('task');
  sheet.insertRowBefore(2);
  sheet.getRange(2,1,1,10).setValues([[Number(id), now(), now(), name, priority||'★★', due_date||'', due_time||'', 'self', false, '']]);
  return {ok: true, action: 'inserted'};
}

function updateTask(ss, id, name, priority, due_date, due_time, done) {
  var sheet = ss.getSheetByName('task');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (Number(rows[i][0]) === Number(id)) {
      var isDone = done === 'true';
      sheet.getRange(i+1, 3).setValue(now());
      if (name !== undefined)     sheet.getRange(i+1, 4).setValue(name);
      if (priority !== undefined) sheet.getRange(i+1, 5).setValue(priority);
      if (due_date !== undefined) sheet.getRange(i+1, 6).setValue(due_date);
      if (due_time !== undefined) sheet.getRange(i+1, 7).setValue(due_time);
      if (done !== undefined) {
        sheet.getRange(i+1, 9).setValue(isDone);
        sheet.getRange(i+1, 10).setValue(isDone ? now() : '');
      }
      return {ok: true, action: 'updated'};
    }
  }
  return {ok: false, error: 'not found'};
}

function deleteTask(ss, id) {
  var sheet = ss.getSheetByName('task');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (Number(rows[i][0]) === Number(id)) {
      sheet.deleteRow(i + 1);
      return {ok: true};
    }
  }
  return {ok: false, error: 'not found'};
}

/* ── diary: id / created_at / updated_at / diary_date / prompt / answer / good1 / good2 / good3 ── */
function getDiary(ss, date) {
  var sheet = ss.getSheetByName('diary');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var rawDiaryDate = rows[i][3];
    var diaryDateStr = rawDiaryDate instanceof Date ? Utilities.formatDate(rawDiaryDate,'Asia/Tokyo','yyyy-MM-dd') : String(rawDiaryDate||'');
    if (diaryDateStr === date) {
      return {
        id: Number(rows[i][0]),
        created_at: String(rows[i][1]||''),
        updated_at: String(rows[i][2]||''),
        diary_date: diaryDateStr,
        prompt: String(rows[i][4]||''),
        answer: String(rows[i][5]||''),
        good1: String(rows[i][6]||''),
        good2: String(rows[i][7]||''),
        good3: String(rows[i][8]||'')
      };
    }
  }
  return {id: null, created_at: '', updated_at: '', diary_date: date, prompt: '', answer: '', good1: '', good2: '', good3: ''};
}

function saveDiary(ss, date, prompt, answer, good1, good2, good3) {
  var sheet = ss.getSheetByName('diary');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    var rawDate = rows[i][3];
    var dateStr = rawDate instanceof Date ? Utilities.formatDate(rawDate,'Asia/Tokyo','yyyy-MM-dd') : String(rawDate||'');
    if (dateStr === date) {
      sheet.getRange(i+1, 3).setValue(now());
      sheet.getRange(i+1, 5).setValue(prompt||'');
      sheet.getRange(i+1, 6).setValue(answer||'');
      sheet.getRange(i+1, 7).setValue(good1||'');
      sheet.getRange(i+1, 8).setValue(good2||'');
      sheet.getRange(i+1, 9).setValue(good3||'');
      return {ok: true, action: 'updated'};
    }
  }
  sheet.insertRowBefore(2);
  sheet.getRange(2,1,1,9).setValues([[Date.now(), now(), now(), date, prompt||'', answer||'', good1||'', good2||'', good3||'']]);
  return {ok: true, action: 'inserted'};
}

/* ── love: id / created_at / updated_at / name / cat / color / textColor / pinned / love_level / memo ── */
function getLove(ss) {
  var sheet = ss.getSheetByName('love');
  var rows = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    result.push({
      id: String(rows[i][0]),
      created_at: String(rows[i][1]||''),
      updated_at: String(rows[i][2]||''),
      name: String(rows[i][3]||''),
      cat: String(rows[i][4]||''),
      color: String(rows[i][5]||'rgba(192,80,72,.15)'),
      textColor: String(rows[i][6]||'#c05048'),
      pinned: rows[i][7]===true||rows[i][7]==='TRUE',
      love_level: Number(rows[i][8])||1,
      memo: String(rows[i][9]||'')
    });
  }
  return result;
}

function saveLove(ss, id, name, cat, color, textColor, pinned, love_level) {
  var sheet = ss.getSheetByName('love');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.getRange(i+1, 3).setValue(now());
      sheet.getRange(i+1, 4).setValue(name||'');
      sheet.getRange(i+1, 5).setValue(cat||'');
      sheet.getRange(i+1, 6).setValue(color||'');
      sheet.getRange(i+1, 7).setValue(textColor||'');
      sheet.getRange(i+1, 8).setValue(pinned==='true');
      sheet.getRange(i+1, 9).setValue(Number(love_level)||1);
      return {ok: true, action: 'updated'};
    }
  }
  sheet.insertRowBefore(2);
  sheet.getRange(2,1,1,10).setValues([[String(id), now(), now(), name||'', cat||'', color||'', textColor||'', false, Number(love_level)||1, '']]);
  return {ok: true, action: 'inserted'};
}

function deleteLove(ss, id) {
  var sheet = ss.getSheetByName('love');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return {ok: true};
    }
  }
  return {ok: false, error: 'not found'};
}

/* ── schedule: id / created_at / updated_at / event_date / start_time / end_time / all_day / person / name / sub / color ── */
function getEvents(ss, startDate, endDate) {
  var sheet = ss.getSheetByName('schedule');
  var rows = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var raw = rows[i][3];
    var event_date = raw instanceof Date
      ? Utilities.formatDate(raw, 'Asia/Tokyo', 'yyyy-MM-dd')
      : String(raw);
    if (startDate && event_date < startDate) continue;
    if (endDate && event_date > endDate) continue;
    var rawStart = rows[i][4];
    var start_time = rawStart instanceof Date
      ? Utilities.formatDate(rawStart, 'Asia/Tokyo', 'HH:mm')
      : String(rawStart || '');
    var rawEnd = rows[i][5];
    var end_time = rawEnd instanceof Date
      ? Utilities.formatDate(rawEnd, 'Asia/Tokyo', 'HH:mm')
      : String(rawEnd || '');
    result.push({
      id: Number(rows[i][0]),
      created_at: String(rows[i][1]||''),
      updated_at: String(rows[i][2]||''),
      event_date: event_date,
      start_time: start_time,
      end_time: end_time,
      all_day: rows[i][6]===true||rows[i][6]==='TRUE',
      person: String(rows[i][7]||'self'),
      name: String(rows[i][8]||''),
      sub: String(rows[i][9]||''),
      color: String(rows[i][10]||''),
      type: 'ev'
    });
  }
  result.sort(function(a,b){
    if(a.event_date!==b.event_date)return a.event_date.localeCompare(b.event_date);
    return (a.start_time||'').localeCompare(b.start_time||'');
  });
  return result;
}

function saveEvent(ss, id, date, time, endTime, name, sub, color) {
  var sheet = ss.getSheetByName('schedule');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (Number(rows[i][0]) === Number(id)) {
      sheet.getRange(i+1, 3).setValue(now());
      sheet.getRange(i+1, 4).setValue(date||'');
      sheet.getRange(i+1, 5).setValue(time||'');
      sheet.getRange(i+1, 6).setValue(endTime||'');
      sheet.getRange(i+1, 9).setValue(name||'');
      sheet.getRange(i+1, 10).setValue(sub||'');
      sheet.getRange(i+1, 11).setValue(color||'');
      return {ok: true, action: 'updated'};
    }
  }
  sheet.insertRowBefore(2);
  sheet.getRange(2,1,1,11).setValues([[Number(id), now(), now(), date||'', time||'', endTime||'', false, 'self', name||'', sub||'', color||'']]);
  return {ok: true, action: 'inserted'};
}

function deleteEvent(ss, id) {
  var sheet = ss.getSheetByName('schedule');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (Number(rows[i][0]) === Number(id)) {
      sheet.deleteRow(i + 1);
      return {ok: true};
    }
  }
  return {ok: false, error: 'not found'};
}

/* ── mind: id / created_at / updated_at / title / sub / url / cat / icon / taps / states ── */
function getMind(ss) {
  var sheet = ss.getSheetByName('mind');
  var rows = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var statesRaw = String(rows[i][9]||'');
    var states = statesRaw ? statesRaw.split(',') : [];
    result.push({
      id: Number(rows[i][0]),
      created_at: String(rows[i][1]||''),
      updated_at: String(rows[i][2]||''),
      title: String(rows[i][3]||''),
      sub: String(rows[i][4]||''),
      url: String(rows[i][5]||''),
      cat: String(rows[i][6]||'audio'),
      icon: String(rows[i][7]||''),
      taps: Number(rows[i][8])||0,
      states: states,
      from_ai: rows[i][10] === 1 || rows[i][10] === true || String(rows[i][10]) === '1',
    });
  }
  result.sort(function(a,b){return Number(b.id)-Number(a.id);});
  return result;
}

function saveMind(ss, id, title, sub, url, cat, icon, taps, states, fromAi) {
  var sheet = ss.getSheetByName('mind');
  var rows = sheet.getDataRange().getValues();
  var statesStr = states||'';
  for (var i = 1; i < rows.length; i++) {
    if (Number(rows[i][0]) === Number(id)) {
      sheet.getRange(i+1, 3).setValue(now());
      sheet.getRange(i+1, 4).setValue(title||'');
      sheet.getRange(i+1, 5).setValue(sub||'');
      sheet.getRange(i+1, 6).setValue(url||'');
      sheet.getRange(i+1, 7).setValue(cat||'audio');
      sheet.getRange(i+1, 8).setValue(icon||'');
      sheet.getRange(i+1, 9).setValue(Number(taps)||0);
      sheet.getRange(i+1, 10).setValue(statesStr);
      sheet.getRange(i+1, 11).setValue(fromAi ? 1 : 0); // from_ai
      return {ok: true, action: 'updated'};
    }
  }
  sheet.insertRowBefore(2);
  sheet.getRange(2,1,1,11).setValues([[Number(id), now(), now(), title||'', sub||'', url||'', cat||'audio', icon||'', Number(taps)||0, statesStr, fromAi ? 1 : 0]]);
  return {ok: true, action: 'inserted'};
}

function deleteMind(ss, id) {
  var sheet = ss.getSheetByName('mind');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (Number(rows[i][0]) === Number(id)) {
      sheet.deleteRow(i + 1);
      return {ok: true};
    }
  }
  return {ok: false, error: 'not found'};
}

function backupAllSheets() {
  var ss = SpreadsheetApp.openById(SS_ID);
  var data = {};
  SHEETS.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    data[name] = sheet.getDataRange().getValues();
  });
  var dateStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd');
  var filename = 'backup_' + dateStr + '.json';
  var folder = DriveApp.getFolderById(BACKUP_FOLDER_ID);
  var existing = folder.getFilesByName(filename);
  while (existing.hasNext()) existing.next().setTrashed(true);
  folder.createFile(filename, JSON.stringify(data), MimeType.PLAIN_TEXT);
  var files = [];
  var all = folder.getFilesByType(MimeType.PLAIN_TEXT);
  while (all.hasNext()) files.push(all.next());
  files.sort(function(a,b){ return b.getName() > a.getName() ? 1 : -1; });
  for (var i = 4; i < files.length; i++) files[i].setTrashed(true);
  return {ok: true, file: filename};
}

/* ── deed: id / created_at / updated_at / deed_date(M/D) / year / text ── */
function getDeed(ss, deed_date) {
  var sheet = ss.getSheetByName('deed');
  var rows = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    var cellDate = rows[i][3];
    /* Dateオブジェクトの場合はM/d形式に変換（例：5/11） */
    var dateStr = cellDate instanceof Date
      ? Utilities.formatDate(cellDate, 'Asia/Tokyo', 'M/d')
      : String(cellDate||'');
    /* deed_date指定時はフィルタ、未指定時は全件返す */
    if (deed_date && dateStr !== deed_date) continue;
    result.push({
      id: Number(rows[i][0]),
      deed_date: dateStr,
      year: Number(rows[i][4]||0),
      text: String(rows[i][5]||'')
    });
  }
  result.sort(function(a,b){return Number(b.id)-Number(a.id);});
  return result;
}

function saveDeed(ss, id, deed_date, year, text) {
  var sheet = ss.getSheetByName('deed');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (Number(rows[i][0]) === Number(id)) {
      sheet.getRange(i+1, 3).setValue(now());
      sheet.getRange(i+1, 4).setValue(deed_date||'');
      sheet.getRange(i+1, 5).setValue(Number(year)||0);
      sheet.getRange(i+1, 6).setValue(text||'');
      return {ok: true, action: 'updated'};
    }
  }
  sheet.insertRowBefore(2);
  sheet.getRange(2,1,1,6).setValues([[Number(id), now(), now(), deed_date||'', Number(year)||0, text||'']]);
  return {ok: true, action: 'inserted'};
}

function deleteDeed(ss, id) {
  var sheet = ss.getSheetByName('deed');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (Number(rows[i][0]) === Number(id)) {
      sheet.deleteRow(i+1);
      return {ok: true};
    }
  }
  return {ok: false, error: 'not found'};
}

function restoreFromBackup(filename) {
  var folder = DriveApp.getFolderById(BACKUP_FOLDER_ID);
  var files = folder.getFilesByName(filename);
  if (!files.hasNext()) throw new Error('バックアップファイルが見つかりません: ' + filename);
  var data = JSON.parse(files.next().getBlob().getDataAsString());
  var ss = SpreadsheetApp.openById(SS_ID);
  SHEETS.forEach(function(name) {
    if (!data[name]) return;
    var sheet = ss.getSheetByName(name);
    if (!sheet) return;
    sheet.clearContents();
    sheet.getRange(1, 1, data[name].length, data[name][0].length).setValues(data[name]);
  });
  return {ok: true, restored: filename};
}

function fetchYouTubeInfo(videoUrl) {
  videoUrl = videoUrl || 'https://youtu.be/ptiK8U4WlSc';
  var apiKey = PropertiesService.getScriptProperties().getProperty('YOUTUBE_API_KEY');
  if (!apiKey) return {ok: false, error: 'YOUTUBE_API_KEY not set'};
  var videoId = '';
  var match = videoUrl.match(/(?:youtube\.com\/watch\?v=|youtu\.be\/)([^&?]+)/);
  if (match) videoId = match[1];
  if (!videoId) return {ok: false, error: 'YouTube URLではありません'};
  var url = 'https://www.googleapis.com/youtube/v3/videos?part=snippet&id=' + videoId + '&key=' + apiKey;
  var res = UrlFetchApp.fetch(url);
  var data = JSON.parse(res.getContentText());
  if (!data.items || data.items.length === 0) return {ok: false, error: '動画が見つかりません'};
  var snippet = data.items[0].snippet;
  return {ok: true, title: snippet.title, channel: snippet.channelTitle};
}

function setWeeklyBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'backupAllSheets') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('backupAllSheets')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(0)
    .create();
  return {ok: true};
}

function doPost(e) {
  var out = {
    ok: true,
    contentType: (e && e.postData) ? e.postData.type : null,
    body: (e && e.postData) ? e.postData.contents : null
  };
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
