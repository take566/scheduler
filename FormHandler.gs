/**
 * FormHandler.gs - フォーム回答のパースと結果出力
 */

/**
 * フォーム回答シートの指定行からリクエストオブジェクトを生成する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} responseSheet - フォーム回答シート
 * @param {number} row - 行番号
 * @returns {Object} request
 */
function parseFormResponse(responseSheet, row) {
  var data = responseSheet.getRange(row, 1, row, 11).getValues()[0];
  var timestamp = data[0];
  var title = (data[1] || '').toString().trim();
  var organizer = (data[2] || '').toString().trim();
  var participantsStr = (data[3] || '').toString().trim();
  var participants = [];
  if (participantsStr) {
    var parts = participantsStr.split(/[,、]/);
    for (var i = 0; i < parts.length; i++) {
      var p = parts[i].trim();
      if (p) participants.push(p);
    }
  }
  var startDate = data[4];
  var endDate = data[5];
  if (startDate && !(startDate instanceof Date)) startDate = new Date(startDate);
  if (endDate && !(endDate instanceof Date)) endDate = new Date(endDate);
  if (startDate) startDate.setHours(0, 0, 0, 0);
  if (endDate) endDate.setHours(23, 59, 59, 999);

  var duration = data[6];
  if (duration === '' || duration === null) duration = 60;
  duration = Number(duration);

  var timeStart = data[7];
  var timeEnd = data[8];
  var timeZone = Session.getScriptTimeZone() || 'Asia/Tokyo';
  if (timeStart && timeStart instanceof Date) {
    timeStart = Utilities.formatDate(timeStart, timeZone, 'HH:mm');
  }
  if (!timeStart || timeStart === '') timeStart = '09:00';
  if (timeEnd && timeEnd instanceof Date) {
    timeEnd = Utilities.formatDate(timeEnd, timeZone, 'HH:mm');
  }
  if (!timeEnd || timeEnd === '') timeEnd = '18:00';

  var priority = (data[9] || '').toString().trim();
  var requestId = timestamp instanceof Date
    ? Utilities.formatDate(timestamp, timeZone, 'yyyyMMdd_HHmmss')
    : String(timestamp);

  return {
    requestId: requestId,
    title: title,
    organizer: organizer,
    participants: participants,
    dateRange: { start: startDate, end: endDate },
    timeRange: { start: timeStart, end: timeEnd },
    duration: duration,
    priority: priority,
    notes: (data[10] || '').toString().trim()
  };
}

/**
 * 空き時間算出結果と候補日時をシートに出力する
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object} request
 * @param {Array} availableSlots
 */
function writeResults(ss, request, availableSlots) {
  var settings = getSettings(ss);
  var maxCandidates = settings.maxCandidates || 10;
  var timeZone = Session.getScriptTimeZone() || 'Asia/Tokyo';

  var slotSheet = ss.getSheetByName('空き時間算出');
  var candidateSheet = ss.getSheetByName('候補日時');
  if (!slotSheet || !candidateSheet) return;

  if (slotSheet.getLastRow() >= 2) {
    slotSheet.getRange(2, 1, slotSheet.getLastRow(), 8).clearContent();
  }
  if (candidateSheet.getLastRow() >= 2) {
    candidateSheet.getRange(2, 1, candidateSheet.getLastRow(), 7).clearContent();
  }

  var slotRows = [];
  for (var i = 0; i < availableSlots.length; i++) {
    var slot = availableSlots[i];
    var tentativeStr = slot.hasTentative && slot.tentativeMembers && slot.tentativeMembers.length
      ? slot.tentativeMembers.join(', ')
      : '';
    slotRows.push([
      request.requestId,
      request.title,
      slot.start,
      slot.end,
      slot.score,
      i + 1,
      tentativeStr,
      '候補'
    ]);
  }
  if (slotRows.length > 0) {
    slotSheet.getRange(2, 1, 1 + slotRows.length, 8).setValues(slotRows);
  }

  var topSlots = availableSlots.slice(0, maxCandidates);
  var candidateRows = [];
  for (var j = 0; j < topSlots.length; j++) {
    var s = topSlots[j];
    var dateStr = Utilities.formatDate(s.start, timeZone, 'yyyy/MM/dd');
    var startStr = Utilities.formatDate(s.start, timeZone, 'HH:mm');
    var endStr = Utilities.formatDate(s.end, timeZone, 'HH:mm');
    var timeRangeStr = startStr + '〜' + endStr;
    var dowStr = Utilities.formatDate(s.start, timeZone, 'E');
    var noteStr = s.hasTentative && s.tentativeMembers && s.tentativeMembers.length
      ? '暫定予定あり: ' + s.tentativeMembers.join(', ')
      : '';
    candidateRows.push([
      j + 1,
      dateStr,
      timeRangeStr,
      dowStr,
      s.score,
      noteStr,
      false
    ]);
  }
  if (candidateRows.length > 0) {
    candidateSheet.getRange(2, 1, 1 + candidateRows.length, 7).setValues(candidateRows);
  }
}
