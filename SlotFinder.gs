/**
 * SlotFinder.gs - 空き時間算出ロジック
 */

/**
 * 空き時間スロットを算出する
 * @param {Object} request - フォーム入力から生成されたリクエストオブジェクト
 * @returns {Array<{start: Date, end: Date, score: number, hasTentative: boolean, tentativeMembers: string[]}>}
 */
function findAvailableSlots(request) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calendarData = getCalendarData(ss);
  var settings = getSettings(ss);
  var slots = [];
  var slotInterval = settings.slotInterval || 30;

  var busyMap = buildBusyMap(calendarData, request.participants);

  var currentDate = new Date(request.dateRange.start);
  currentDate.setHours(0, 0, 0, 0);
  var rangeEnd = new Date(request.dateRange.end);
  rangeEnd.setHours(23, 59, 59, 999);

  while (currentDate.getTime() <= rangeEnd.getTime()) {
    if (isBusinessDay(currentDate, settings)) {
      var slotStart = combineDateAndTime(currentDate, request.timeRange.start);
      var dayEnd = combineDateAndTime(currentDate, request.timeRange.end);

      while (addMinutes(slotStart, request.duration).getTime() <= dayEnd.getTime()) {
        var slotEnd = addMinutes(slotStart, request.duration);
        var conflict = checkConflicts(busyMap, request.participants, slotStart, slotEnd, currentDate, settings);

        if (!conflict.hasConflict) {
          var baseScore = calculateScore(slotStart, request.priority);
          var tentativePenalty = (conflict.hasTentative ? 5 : 0);
          slots.push({
            start: new Date(slotStart.getTime()),
            end: new Date(slotEnd.getTime()),
            score: baseScore - tentativePenalty,
            hasTentative: conflict.hasTentative,
            tentativeMembers: conflict.tentativeMembers || []
          });
        }
        slotStart = addMinutes(slotStart, slotInterval);
      }
    }
    currentDate.setDate(currentDate.getDate() + 1);
  }

  slots.sort(function (a, b) { return b.score - a.score; });
  return slots;
}

/**
 * カレンダーデータから参加者ごとの予定マップを構築する
 * @param {Array<Array>} calendarData - getCalendarData() の戻り値
 * @param {string[]} participants - 参加者メールの配列
 * @returns {Object} email -> [{ start, end, status, isAllDay }]
 */
function buildBusyMap(calendarData, participants) {
  var busyMap = {};
  var participantSet = {};
  for (var p = 0; p < participants.length; p++) {
    participantSet[participants[p]] = true;
    busyMap[participants[p]] = [];
  }
  for (var i = 0; i < calendarData.length; i++) {
    var row = calendarData[i];
    var email = (row[0] || '').toString().trim();
    if (!participantSet[email]) continue;
    var start = row[3] instanceof Date ? row[3] : new Date(row[3]);
    var end = row[4] instanceof Date ? row[4] : new Date(row[4]);
    var isAllDay = !!row[5];
    var status = (row[6] || 'BUSY').toString();
    busyMap[email].push({ start: start, end: end, status: status, isAllDay: isAllDay });
  }
  return busyMap;
}

/**
 * 予定の衝突を検出する
 * @param {Object} busyMap - buildBusyMap の戻り値
 * @param {string[]} participants
 * @param {Date} slotStart
 * @param {Date} slotEnd
 * @param {Date} slotDate - スロットの日付（終日判定用）
 * @param {Object} settings - getSettings() の戻り値
 * @returns {{ hasConflict: boolean, hasTentative: boolean, tentativeMembers: string[] }}
 */
function checkConflicts(busyMap, participants, slotStart, slotEnd, slotDate, settings) {
  var hasConflict = false;
  var hasTentative = false;
  var tentativeMembers = [];

  for (var i = 0; i < participants.length; i++) {
    var email = participants[i];
    var events = busyMap[email] || [];
    for (var j = 0; j < events.length; j++) {
      var event = events[j];
      if (event.isAllDay) {
        var eventDay = new Date(event.start);
        eventDay.setHours(0, 0, 0, 0);
        var slotDay = new Date(slotDate);
        slotDay.setHours(0, 0, 0, 0);
        if (eventDay.getTime() === slotDay.getTime()) {
          hasConflict = true;
          break;
        }
        continue;
      }
      if (event.start.getTime() < slotEnd.getTime() && event.end.getTime() > slotStart.getTime()) {
        if (event.status === 'TENTATIVE') {
          if (settings && settings.tentativeExcluded) {
            hasConflict = true;
            break;
          }
          hasTentative = true;
          if (tentativeMembers.indexOf(email) === -1) tentativeMembers.push(email);
        } else {
          hasConflict = true;
          break;
        }
      }
    }
    if (hasConflict) break;
  }
  return { hasConflict: hasConflict, hasTentative: hasTentative, tentativeMembers: tentativeMembers };
}


/**
 * スロットの適合度スコアを計算する（優先条件を反映）
 * @param {Date} slotStart
 * @param {string} priority - 最短日時優先 / 午前優先 / 午後優先 など
 * @returns {number}
 */
function calculateScore(slotStart, priority) {
  if (!priority) priority = '';
  var now = new Date();
  if (priority.indexOf('最短') !== -1 || priority.indexOf('日時') !== -1) {
    var diffMs = slotStart.getTime() - now.getTime();
    var diffMinutes = Math.floor(diffMs / (60 * 1000));
    return SLOT_SCORE_MAX - Math.max(0, diffMinutes);
  }
  if (priority.indexOf('午前') !== -1) {
    return 100 + (slotStart.getHours() < 12 ? 10 : 0);
  }
  if (priority.indexOf('午後') !== -1) {
    return 100 + (slotStart.getHours() >= 12 ? 10 : 0);
  }
  return 100;
}
