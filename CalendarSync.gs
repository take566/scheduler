/**
 * CalendarSync.gs - 参加者カレンダー取得と「カレンダー同期」シートへの書き出し
 */

/**
 * 全参加者のカレンダー予定を同期する
 * メニューまたはトリガーから実行
 */
function syncAllCalendars() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var settings = getSettings(ss);
  var members = getMembers(ss);
  var syncSheet = ss.getSheetByName('カレンダー同期');
  if (!syncSheet) {
    Logger.log('カレンダー同期シートがありません。初期セットアップを実行してください。');
    return;
  }

  if (syncSheet.getLastRow() >= 2) {
    syncSheet.getRange(2, 1, syncSheet.getLastRow(), 8).clearContent();
  }

  var startDate = settings.searchStartDate;
  var endDate = settings.searchEndDate;
  var allEvents = [];

  for (var i = 0; i < members.length; i++) {
    var member = members[i];
    try {
      var calendar = CalendarApp.getCalendarById(member.email);
      if (!calendar) {
        Logger.log('カレンダー未取得: ' + member.email);
        continue;
      }
      var events = calendar.getEvents(startDate, endDate);
      for (var j = 0; j < events.length; j++) {
        var event = events[j];
        var status = event.getMyStatus ? event.getMyStatus() : 'BUSY';
        if (!status) status = 'BUSY';
        allEvents.push([
          member.email,
          member.name,
          event.getTitle(),
          event.getStartTime(),
          event.getEndTime(),
          event.isAllDayEvent(),
          status,
          new Date()
        ]);
      }
    } catch (e) {
      Logger.log('エラー (' + member.email + '): ' + e.message);
    }
  }

  if (allEvents.length > 0) {
    syncSheet.getRange(2, 1, 1 + allEvents.length, 8).setValues(allEvents);
  }

  var settingsSheet = ss.getSheetByName('設定');
  if (settingsSheet) {
    settingsSheet.getRange(SETTING_ROW_LAST_SYNC, 2).setValue(new Date());
  }
}
