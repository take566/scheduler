/**
 * WebApp_Test.gs - WebApp.gs の手動テスト関数
 *
 * Apps Script エディタで直接実行して動作確認する。
 * `runAllWebAppTests` を選択して "実行" → ログを確認。
 */

function runAllWebAppTests() {
  var results = [];
  results.push(testValidateWebAppPayload_OK_());
  results.push(testValidateWebAppPayload_Empty_());
  results.push(testValidateWebAppPayload_NoParticipants_());
  results.push(testValidateWebAppPayload_BadDate_());
  results.push(testValidateWebAppPayload_DateRangeReversed_());
  results.push(testValidateWebAppPayload_SlotLenOutOfRange_());
  results.push(testValidateWebAppPayload_TooManyParticipants_());
  results.push(testNormalizeRequest_DefaultsApplied_());
  results.push(testDoPost_NoSecret_Unauthorized_());
  results.push(testDoPost_BadJSON_());
  results.push(testDoPost_HappyPath_RequiresFixtures_());

  var passed = results.filter(function (r) { return r.pass; }).length;
  var failed = results.filter(function (r) { return !r.pass; });
  Logger.log('Total: ' + results.length + ', Passed: ' + passed + ', Failed: ' + failed.length);
  failed.forEach(function (r) { Logger.log('FAIL: ' + r.name + ' - ' + r.message); });
}

function assert_(name, cond, message) {
  if (cond) {
    Logger.log('PASS: ' + name);
    return { name: name, pass: true };
  }
  Logger.log('FAIL: ' + name + ' - ' + message);
  return { name: name, pass: false, message: message };
}

function testValidateWebAppPayload_OK_() {
  var err = validateWebAppPayload_({
    participants: ['a@example.com', 'b@example.com'],
    dateRange: { start: '2026-05-08', end: '2026-05-22' },
    slotLengthMinutes: 30,
    count: 5
  });
  return assert_('validate happy path', err === null, 'expected null, got: ' + err);
}

function testValidateWebAppPayload_Empty_() {
  var err = validateWebAppPayload_(null);
  return assert_('validate null', typeof err === 'string', 'expected error string');
}

function testValidateWebAppPayload_NoParticipants_() {
  var err = validateWebAppPayload_({ participants: [], dateRange: { start: '2026-05-08', end: '2026-05-22' } });
  return assert_('validate empty participants', err && err.indexOf('participants') !== -1, 'expected participants error, got: ' + err);
}

function testValidateWebAppPayload_BadDate_() {
  var err = validateWebAppPayload_({ participants: ['a@x.com'], dateRange: { start: 'not-a-date', end: '2026-05-22' } });
  return assert_('validate bad date', err && err.indexOf('invalid date') !== -1, 'expected date error, got: ' + err);
}

function testValidateWebAppPayload_DateRangeReversed_() {
  var err = validateWebAppPayload_({ participants: ['a@x.com'], dateRange: { start: '2026-05-22', end: '2026-05-08' } });
  return assert_('validate reversed range', err && err.indexOf('<= dateRange.end') !== -1, 'expected order error, got: ' + err);
}

function testValidateWebAppPayload_SlotLenOutOfRange_() {
  var err = validateWebAppPayload_({
    participants: ['a@x.com'],
    dateRange: { start: '2026-05-08', end: '2026-05-22' },
    slotLengthMinutes: 5
  });
  return assert_('validate slot len min', err && err.indexOf('slotLengthMinutes') !== -1, 'expected range error, got: ' + err);
}

function testValidateWebAppPayload_TooManyParticipants_() {
  var pp = [];
  for (var i = 0; i < 25; i++) pp.push('user' + i + '@example.com');
  var err = validateWebAppPayload_({
    participants: pp,
    dateRange: { start: '2026-05-08', end: '2026-05-22' }
  });
  return assert_('validate too many participants', err && err.indexOf('too many') !== -1, 'expected too many error, got: ' + err);
}

function testNormalizeRequest_DefaultsApplied_() {
  var req = normalizeRequestForLegacy_({
    participants: ['a@x.com'],
    dateRange: { start: '2026-05-08', end: '2026-05-22' }
  });
  var ok =
    req.timeRange.start === '09:00' &&
    req.timeRange.end === '18:00' &&
    req.duration === 30 &&
    req.priority === '' &&
    req.participants.length === 1 &&
    req.dateRange.start instanceof Date;
  return assert_('normalize defaults', ok, 'unexpected normalized request: ' + JSON.stringify(req));
}

function testDoPost_NoSecret_Unauthorized_() {
  var resp = doPost({
    parameter: {},
    postData: {
      contents: JSON.stringify({
        participants: ['a@x.com'],
        dateRange: { start: '2026-05-08', end: '2026-05-22' }
      })
    }
  });
  var body = JSON.parse(resp.getContent());
  return assert_('doPost no secret', body.ok === false && body.error.indexOf('unauthorized') !== -1, 'expected unauthorized, got: ' + JSON.stringify(body));
}

function testDoPost_BadJSON_() {
  // 共有シークレットを script properties に一時設定
  var props = PropertiesService.getScriptProperties();
  var orig = props.getProperty('SHARED_SECRET');
  props.setProperty('SHARED_SECRET', 'test-secret');
  try {
    var resp = doPost({
      parameter: { secret: 'test-secret' },
      postData: { contents: 'not-json{{{' }
    });
    var body = JSON.parse(resp.getContent());
    return assert_('doPost bad JSON', body.ok === false && body.error.indexOf('invalid JSON') !== -1, 'expected JSON error, got: ' + JSON.stringify(body));
  } finally {
    if (orig) props.setProperty('SHARED_SECRET', orig);
    else props.deleteProperty('SHARED_SECRET');
  }
}

/**
 * findAvailableSlots はスプレッドシートに依存するため
 * "カレンダー同期" シートが存在し空でないことを前提とする。
 * 実環境でしか動かない (mock 困難)。
 */
function testDoPost_HappyPath_RequiresFixtures_() {
  var ss;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch (_) {
    return assert_('doPost happy path (skipped: no spreadsheet)', true, 'skipped');
  }
  if (!ss || !ss.getSheetByName('カレンダー同期')) {
    return assert_('doPost happy path (skipped: no fixtures)', true, 'skipped');
  }
  var props = PropertiesService.getScriptProperties();
  var orig = props.getProperty('SHARED_SECRET');
  props.setProperty('SHARED_SECRET', 'test-secret');
  try {
    var resp = doPost({
      parameter: { secret: 'test-secret' },
      postData: {
        contents: JSON.stringify({
          participants: ['nobody@example.com'],
          dateRange: { start: '2026-05-08', end: '2026-05-09' },
          slotLengthMinutes: 30,
          count: 3
        })
      }
    });
    var body = JSON.parse(resp.getContent());
    return assert_('doPost happy path', body.ok === true && Array.isArray(body.slots), 'unexpected response: ' + JSON.stringify(body));
  } finally {
    if (orig) props.setProperty('SHARED_SECRET', orig);
    else props.deleteProperty('SHARED_SECRET');
  }
}
