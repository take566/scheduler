/**
 * Code.gs - メインスクリプト（エントリポイント）
 */

/**
 * スプレッドシートを開いたときにカスタムメニューを追加する
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📅 会議調整ツール')
    .addItem('カレンダー同期（全参加者）', 'syncAllCalendars')
    .addItem('空き時間を再算出', 'recalculateSlots')
    .addSeparator()
    .addItem('フォームを開く', 'openForm')
    .addItem('確定した会議をカレンダーに登録', 'createCalendarEvent')
    .addSeparator()
    .addItem('初期セットアップ', 'initialSetup')
    .addItem('トリガー設定', 'setupTriggers')
    .addToUi();
}

/**
 * フォーム送信時に自動実行されるトリガー関数
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e - フォーム送信イベント（使用しない場合は省略可）
 */
function onFormSubmit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var responseSheet = ss.getSheetByName('フォーム回答');
  if (!responseSheet) return;
  var lastRow = responseSheet.getLastRow();
  if (lastRow < 2) return;

  syncAllCalendars();

  var request = parseFormResponse(responseSheet, lastRow);
  var availableSlots = findAvailableSlots(request);
  writeResults(ss, request, availableSlots);

  responseSheet.getRange(lastRow, 12).setValue('処理完了');

  sendResultNotification(request, availableSlots);
}
