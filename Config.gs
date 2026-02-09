/**
 * Config.gs - 定数・設定値・設定シート／参加者マスタ／カレンダー同期データ取得
 */

/** 設定シートの行番号（項目名はA列、値はB列） */
var SETTING_ROW_SEARCH_START = 1;
var SETTING_ROW_SEARCH_END = 2;
var SETTING_ROW_SLOT_INTERVAL = 3;
var SETTING_ROW_BUSINESS_START = 4;
var SETTING_ROW_BUSINESS_END = 5;
var SETTING_ROW_EXCLUDE_WEEKENDS = 6;
var SETTING_ROW_EXCLUDE_HOLIDAYS = 7;
var SETTING_ROW_TENTATIVE_EXCLUDED = 8;
var SETTING_ROW_MAX_CANDIDATES = 9;
var SETTING_ROW_LAST_SYNC = 10;
var SETTING_ROW_FORM_URL = 11;

/** スコア計算用 */
var SLOT_SCORE_MAX = 10000;

/**
 * 設定シートから設定値を取得する
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシート
 * @returns {Object} 設定オブジェクト
 */
function getSettings(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('設定');
  if (!sheet) {
    return getDefaultSettings();
  }
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var endDefault = new Date(today);
  endDefault.setDate(endDefault.getDate() + 14);

  var searchStart = sheet.getRange(SETTING_ROW_SEARCH_START, 2).getValue();
  var searchEnd = sheet.getRange(SETTING_ROW_SEARCH_END, 2).getValue();
  if (!searchStart || !(searchStart instanceof Date)) searchStart = today;
  if (!searchEnd || !(searchEnd instanceof Date)) searchEnd = endDefault;
  if (typeof searchStart === 'object' && searchStart.getHours) searchStart.setHours(0, 0, 0, 0);
  if (typeof searchEnd === 'object' && searchEnd.getHours) searchEnd.setHours(23, 59, 59, 999);

  var slotInterval = sheet.getRange(SETTING_ROW_SLOT_INTERVAL, 2).getValue();
  if (slotInterval === '' || slotInterval === null) slotInterval = 30;

  var businessStart = sheet.getRange(SETTING_ROW_BUSINESS_START, 2).getValue();
  var businessEnd = sheet.getRange(SETTING_ROW_BUSINESS_END, 2).getValue();
  if (!businessStart) businessStart = '09:00';
  if (!businessEnd) businessEnd = '18:00';
  if (typeof businessStart === 'object' && businessStart.getHours !== undefined) {
    businessStart = Utilities.formatDate(businessStart, Session.getScriptTimeZone() || 'Asia/Tokyo', 'HH:mm');
  }
  if (typeof businessEnd === 'object' && businessEnd.getHours !== undefined) {
    businessEnd = Utilities.formatDate(businessEnd, Session.getScriptTimeZone() || 'Asia/Tokyo', 'HH:mm');
  }

  var excludeWeekends = sheet.getRange(SETTING_ROW_EXCLUDE_WEEKENDS, 2).getValue();
  var excludeHolidays = sheet.getRange(SETTING_ROW_EXCLUDE_HOLIDAYS, 2).getValue();
  var tentativeExcluded = sheet.getRange(SETTING_ROW_TENTATIVE_EXCLUDED, 2).getValue();
  var maxCandidates = sheet.getRange(SETTING_ROW_MAX_CANDIDATES, 2).getValue();
  var lastSync = sheet.getRange(SETTING_ROW_LAST_SYNC, 2).getValue();

  if (excludeWeekends === '') excludeWeekends = true;
  if (excludeHolidays === '') excludeHolidays = true;
  if (tentativeExcluded === '') tentativeExcluded = false;
  if (maxCandidates === '' || maxCandidates === null) maxCandidates = 10;

  return {
    searchStartDate: searchStart,
    searchEndDate: searchEnd,
    slotInterval: Number(slotInterval),
    businessStart: String(businessStart),
    businessEnd: String(businessEnd),
    excludeWeekends: !!excludeWeekends,
    excludeHolidays: !!excludeHolidays,
    tentativeExcluded: !!tentativeExcluded,
    maxCandidates: Number(maxCandidates),
    lastSyncDate: lastSync,
    formUrl: sheet.getRange(SETTING_ROW_FORM_URL, 2).getValue() || ''
  };
}

/**
 * 設定シートが無い場合のデフォルト設定
 * @returns {Object}
 */
function getDefaultSettings() {
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var endDefault = new Date(today);
  endDefault.setDate(endDefault.getDate() + 14);
  return {
    searchStartDate: today,
    searchEndDate: endDefault,
    slotInterval: 30,
    businessStart: '09:00',
    businessEnd: '18:00',
    excludeWeekends: true,
    excludeHolidays: true,
    tentativeExcluded: false,
    maxCandidates: 10,
    lastSyncDate: null,
    formUrl: ''
  };
}

/**
 * 参加者マスタシートから参加者一覧を取得する
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシート
 * @returns {Array<{email: string, name: string}>}
 */
function getMembers(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('参加者マスタ');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  var data = sheet.getRange(2, 1, lastRow, 2).getValues();
  var members = [];
  for (var i = 0; i < data.length; i++) {
    var email = (data[i][0] || '').toString().trim();
    if (!email) continue;
    members.push({
      email: email,
      name: (data[i][1] || '').toString().trim() || email
    });
  }
  return members;
}

/**
 * カレンダー同期シートから予定データを取得する
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - スプレッドシート
 * @returns {Array<Array>} 各行 [参加者メール, 参加者名, 予定タイトル, 開始日時, 終了日時, 終日, ステータス, 同期日時]
 */
function getCalendarData(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('カレンダー同期');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow, 8).getValues();
}
