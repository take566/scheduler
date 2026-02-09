/**
 * Setup.gs - 初期セットアップ・トリガー設定・メニュー用関数
 */

var SHEET_NAMES = ['設定', '参加者マスタ', 'カレンダー同期', 'フォーム回答', '空き時間算出', '候補日時'];

/**
 * シート作成・ヘッダ・初期値を設定する
 */
function initialSetup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var endDefault = new Date(today);
  endDefault.setDate(endDefault.getDate() + 14);

  for (var i = 0; i < SHEET_NAMES.length; i++) {
    var name = SHEET_NAMES[i];
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
  }

  var settingsSheet = ss.getSheetByName('設定');
  settingsSheet.clear();
  settingsSheet.getRange(1, 1, 11, 1).setValues([
    ['検索開始日'],
    ['検索終了日'],
    ['スロット間隔（分）'],
    ['営業時間（開始）'],
    ['営業時間（終了）'],
    ['土日除外'],
    ['祝日除外'],
    ['TENTATIVE 除外'],
    ['最大候補数'],
    ['最終同期日時'],
    ['フォームURL']
  ]);
  settingsSheet.getRange(1, 2, 9, 2).setValues([
    [today],
    [endDefault],
    [30],
    ['09:00'],
    ['18:00'],
    [true],
    [true],
    [false],
    [10]
  ]);
  settingsSheet.getRange(10, 2).setValue('');
  settingsSheet.getRange(11, 2).setValue('');

  var memberSheet = ss.getSheetByName('参加者マスタ');
  memberSheet.clear();
  memberSheet.getRange(1, 1, 1, 2).setValues([['参加者メール', '参加者名']]);

  var syncSheet = ss.getSheetByName('カレンダー同期');
  syncSheet.clear();
  syncSheet.getRange(1, 1, 1, 8).setValues([[
    '参加者メール', '参加者名', '予定タイトル', '開始日時', '終了日時', '終日イベント', 'ステータス', '同期日時'
  ]]);

  var formSheet = ss.getSheetByName('フォーム回答');
  formSheet.clear();
  formSheet.getRange(1, 1, 1, 12).setValues([[
    'タイムスタンプ', '会議タイトル', '主催者メール', '参加者（カンマ区切り）', '希望開始日', '希望終了日',
    '所要時間（分）', '希望時間帯（開始）', '希望時間帯（終了）', '優先条件', '備考', '処理ステータス'
  ]]);

  var slotSheet = ss.getSheetByName('空き時間算出');
  slotSheet.clear();
  slotSheet.getRange(1, 1, 1, 8).setValues([[
    'リクエストID', '会議タイトル', '候補日時（開始）', '候補日時（終了）', 'スコア', '順位', 'TENTATIVE警告', 'ステータス'
  ]]);

  var candidateSheet = ss.getSheetByName('候補日時');
  candidateSheet.clear();
  candidateSheet.getRange(1, 1, 1, 7).setValues([[
    '順位', '日付', '時間帯', '曜日', 'スコア', '備考', '決定'
  ]]);

  SpreadsheetApp.getUi().alert('初期セットアップが完了しました。');
}

/**
 * 時間主導型トリガーとフォーム送信トリガーを設定する
 */
function setupTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger('syncAllCalendars')
    .timeBased()
    .everyHours(1)
    .create();
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
  SpreadsheetApp.getUi().alert('トリガーを設定しました。（syncAllCalendars: 毎時 / onFormSubmit: フォーム送信時）');
}

/**
 * 設定シートのフォームURLを開く（ダイアログでリンク表示）
 */
function openForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = getSettings(ss);
  var url = settings.formUrl || '';
  if (!url) {
    SpreadsheetApp.getUi().alert('「設定」シートのフォームURL（11行目）にフォームのURLを入力してください。');
    return;
  }
  var html = HtmlService.createHtmlOutput(
    '<p><a href="' + url + '" target="_blank" rel="noopener">フォームを開く</a></p>'
  ).setWidth(300).setHeight(80);
  SpreadsheetApp.getUi().showModalDialog(html, 'フォーム');
}

/**
 * 最新のフォーム回答をもとに空き時間を再算出し、結果を書き出す
 */
function recalculateSlots() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var responseSheet = ss.getSheetByName('フォーム回答');
  if (!responseSheet) {
    SpreadsheetApp.getUi().alert('フォーム回答シートがありません。');
    return;
  }
  var lastRow = responseSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('フォーム回答がありません。');
    return;
  }
  syncAllCalendars();
  var request = parseFormResponse(responseSheet, lastRow);
  var availableSlots = findAvailableSlots(request);
  writeResults(ss, request, availableSlots);
  SpreadsheetApp.getUi().alert('空き時間の再算出が完了しました。');
}

/**
 * 候補日時シートで「決定」にチェックが入った行の会議をカレンダーに登録する
 */
function createCalendarEvent() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var candidateSheet = ss.getSheetByName('候補日時');
  var responseSheet = ss.getSheetByName('フォーム回答');
  if (!candidateSheet || !responseSheet) {
    SpreadsheetApp.getUi().alert('候補日時シートまたはフォーム回答シートがありません。');
    return;
  }
  var lastRow = candidateSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('候補日時がありません。');
    return;
  }
  var data = candidateSheet.getRange(2, 1, lastRow, 7).getValues();
  var timeZone = Session.getScriptTimeZone() || 'Asia/Tokyo';
  var decidedRow = -1;
  for (var i = 0; i < data.length; i++) {
    if (data[i][6] === true) {
      decidedRow = i;
      break;
    }
  }
  if (decidedRow === -1) {
    SpreadsheetApp.getUi().alert('「決定」列にチェックを入れた行がありません。');
    return;
  }
  var dateStr = data[decidedRow][1];
  var timeRangeStr = (data[decidedRow][2] || '').toString().trim();
  var title = '';
  var participants = [];
  var formLastRow = responseSheet.getLastRow();
  if (formLastRow >= 2) {
    var formRow = responseSheet.getRange(formLastRow, 1, formLastRow, 4).getValues()[0];
    title = (formRow[1] || '会議').toString().trim();
    var partStr = (formRow[3] || '').toString().trim();
    if (partStr) {
      var parts = partStr.split(/[,、]/);
      for (var p = 0; p < parts.length; p++) {
        var email = parts[p].trim();
        if (email) participants.push(email);
      }
    }
  }
  if (!title) title = '会議';

  var match = timeRangeStr.match(/(\d{1,2}):(\d{2})\s*[〜~-]\s*(\d{1,2}):(\d{2})/);
  var startHour = 9;
  var startMin = 0;
  var endHour = 10;
  var endMin = 0;
  if (match) {
    startHour = parseInt(match[1], 10);
    startMin = parseInt(match[2], 10);
    endHour = parseInt(match[3], 10);
    endMin = parseInt(match[4], 10);
  }
  var startDate = new Date(dateStr);
  startDate.setHours(startHour, startMin, 0, 0);
  var endDate = new Date(dateStr);
  endDate.setHours(endHour, endMin, 0, 0);

  try {
    var calendar = CalendarApp.getDefaultCalendar();
    var event = calendar.createEvent(title, startDate, endDate);
    for (var j = 0; j < participants.length; j++) {
      event.addGuest(participants[j]);
    }
    SpreadsheetApp.getUi().alert('カレンダーに会議を登録しました: ' + title + ' ' + dateStr + ' ' + timeRangeStr);
  } catch (e) {
    SpreadsheetApp.getUi().alert('登録に失敗しました: ' + e.message);
  }
}
