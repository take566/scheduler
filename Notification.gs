/**
 * Notification.gs - メール通知
 */

/**
 * 結果を主催者にメール通知する
 * @param {Object} request - リクエストオブジェクト
 * @param {Array<{start: Date, end: Date, hasTentative: boolean, tentativeMembers: string[]}>} slots - 空きスロット配列
 */
function sendResultNotification(request, slots) {
  if (!request || !request.organizer) return;
  var topSlots = slots.slice(0, 5);
  var timeZone = Session.getScriptTimeZone() || 'Asia/Tokyo';

  var body = '会議「' + request.title + '」の日程候補が見つかりました。\n\n';
  body += '参加者: ' + (request.participants ? request.participants.join(', ') : '') + '\n';
  body += '所要時間: ' + request.duration + '分\n\n';
  body += '--- 候補日時 TOP5 ---\n\n';

  for (var i = 0; i < topSlots.length; i++) {
    var slot = topSlots[i];
    var dateStr = Utilities.formatDate(slot.start, timeZone, 'yyyy/MM/dd (E)');
    var startStr = Utilities.formatDate(slot.start, timeZone, 'HH:mm');
    var endStr = Utilities.formatDate(slot.end, timeZone, 'HH:mm');
    body += (i + 1) + '. ' + dateStr + ' ' + startStr + '〜' + endStr;
    if (slot.hasTentative && slot.tentativeMembers && slot.tentativeMembers.length) {
      body += ' ⚠ 暫定予定あり: ' + slot.tentativeMembers.join(', ');
    }
    body += '\n';
  }

  body += '\n詳細はスプレッドシートをご確認ください:\n';
  body += SpreadsheetApp.getActiveSpreadsheet().getUrl();

  try {
    MailApp.sendEmail({
      to: request.organizer,
      subject: '[会議調整] 「' + request.title + '」の候補日時',
      body: body
    });
  } catch (e) {
    Logger.log('メール送信エラー: ' + e.message);
  }
}
