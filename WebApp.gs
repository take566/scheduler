/**
 * WebApp.gs - HTTP Web App エンドポイント (外部システム連携)
 *
 * OpenClaw Slack bot 等の外部システムから POST で会議候補日時を取得する。
 *
 * デプロイ手順は README.md "Web App デプロイ手順" を参照。
 *
 * 認証:
 *   Web App は ANYONE_ANONYMOUS で公開するが、Script Properties の
 *   `SHARED_SECRET` を header `X-Shared-Secret` で検証することで実質的に保護する。
 *
 * リクエスト例:
 *   POST {webappUrl}
 *   X-Shared-Secret: <SHARED_SECRET>
 *   Content-Type: application/json
 *   {
 *     "participants": ["a@example.com", "b@example.com"],
 *     "dateRange": {"start": "2026-05-08", "end": "2026-05-22"},
 *     "slotLengthMinutes": 30,
 *     "count": 5,
 *     "timezone": "Asia/Tokyo",
 *     "timeRange": {"start": "09:00", "end": "18:00"},   // 任意 (省略時 09:00-18:00)
 *     "priority": "最短日時優先"                          // 任意
 *   }
 *
 * レスポンス例:
 *   {
 *     "ok": true,
 *     "slots": [
 *       {"start": "2026-05-08T14:00:00+09:00", "end": "2026-05-08T14:30:00+09:00",
 *        "score": 1.0, "available_count": 2, "has_tentative": false, "tentative_members": []}
 *     ],
 *     "request_echo": { ... }
 *   }
 *
 * エラー時:
 *   {"ok": false, "error": "..."}
 */

/**
 * doPost - HTTP POST handler. Apps Script ランタイムが Web App リクエストで自動呼出。
 * @param {GoogleAppsScript.Events.DoPost} e
 * @returns {GoogleAppsScript.Content.TextOutput}
 */
function doPost(e) {
  try {
    // 1. 共有シークレット検証
    var providedSecret = extractSharedSecret_(e);
    var expectedSecret = PropertiesService.getScriptProperties().getProperty('SHARED_SECRET');
    if (!expectedSecret) {
      return jsonResponse_({
        ok: false,
        error: 'server misconfigured: SHARED_SECRET not set in Script Properties'
      });
    }
    if (!providedSecret || providedSecret !== expectedSecret) {
      return jsonResponse_({
        ok: false,
        error: 'unauthorized: invalid or missing X-Shared-Secret'
      });
    }

    // 2. body parse
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse_({ ok: false, error: 'empty request body' });
    }
    var payload;
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return jsonResponse_({ ok: false, error: 'invalid JSON: ' + parseErr.message });
    }

    // 3. validation
    var validationErr = validateWebAppPayload_(payload);
    if (validationErr) {
      return jsonResponse_({ ok: false, error: validationErr });
    }

    // 4. 既存 findAvailableSlots(request) スキーマに正規化
    var legacyRequest = normalizeRequestForLegacy_(payload);

    // 5. 空き時間算出 (既存ロジック)
    //    syncAllCalendars は呼ばない (実行時間 6 分制限・カレンダー同期は別 trigger)
    var allSlots = findAvailableSlots(legacyRequest);

    // 6. 上位 count 件に切り詰めて整形
    var count = payload.count || 5;
    var topSlots = allSlots.slice(0, count);
    var timezone = payload.timezone || Session.getScriptTimeZone() || 'Asia/Tokyo';
    var formattedSlots = topSlots.map(function (s) {
      return {
        start: Utilities.formatDate(s.start, timezone, "yyyy-MM-dd'T'HH:mm:ssXXX"),
        end: Utilities.formatDate(s.end, timezone, "yyyy-MM-dd'T'HH:mm:ssXXX"),
        score: s.score,
        available_count: legacyRequest.participants.length - (s.tentativeMembers ? s.tentativeMembers.length : 0),
        has_tentative: !!s.hasTentative,
        tentative_members: s.tentativeMembers || []
      };
    });

    return jsonResponse_({
      ok: true,
      slots: formattedSlots,
      request_echo: payload,
      total_candidates_before_truncation: allSlots.length
    });
  } catch (err) {
    return jsonResponse_({
      ok: false,
      error: 'internal error: ' + (err && err.message ? err.message : String(err))
    });
  }
}

/**
 * doGet - 動作確認用 (返値はバージョン情報のみ。秘匿情報を含めない)
 * @param {GoogleAppsScript.Events.DoGet} e
 */
function doGet(e) {
  return jsonResponse_({
    ok: true,
    service: 'meeting-scheduler-webapp',
    method: 'POST required',
    docs: 'https://github.com/take566/scheduler#web-app-デプロイ手順'
  });
}

/**
 * 共有シークレットを request から取り出す。
 * Apps Script の DoPost イベントは header をそのままは公開しないため、
 * 以下の優先順位で取得する:
 *   1. e.parameter['secret']         (URL ?secret=XXX)
 *   2. payload._sharedSecret         (body 内 magic field — 互換用)
 *   3. e.postData.headers['X-Shared-Secret']  (将来の Apps Script header API 対応)
 *
 * 推奨は 1 (URL クエリパラメタ): クライアント側は HTTPS で URL を使うため漏洩リスクは body と同等。
 *
 * @param {GoogleAppsScript.Events.DoPost} e
 * @returns {string|null}
 */
function extractSharedSecret_(e) {
  if (!e) return null;
  // 1. URL query parameter (?secret=...)
  if (e.parameter && e.parameter.secret) {
    return e.parameter.secret;
  }
  // 2. body 内 _sharedSecret field
  if (e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      if (body && body._sharedSecret) return body._sharedSecret;
    } catch (_) { /* fallthrough */ }
  }
  // 3. headers (将来の API 対応)
  if (e.postData && e.postData.headers) {
    var h = e.postData.headers;
    return h['X-Shared-Secret'] || h['x-shared-secret'] || null;
  }
  return null;
}

/**
 * payload 検証。エラー文字列を返す (正常なら null)。
 */
function validateWebAppPayload_(payload) {
  if (!payload || typeof payload !== 'object') return 'payload must be a JSON object';

  if (!Array.isArray(payload.participants) || payload.participants.length === 0) {
    return 'participants: non-empty array of email strings required';
  }
  for (var i = 0; i < payload.participants.length; i++) {
    if (typeof payload.participants[i] !== 'string' || !payload.participants[i].trim()) {
      return 'participants[' + i + ']: must be non-empty email string';
    }
  }
  if (payload.participants.length > 20) {
    return 'participants: too many (max 20)';
  }

  if (!payload.dateRange || typeof payload.dateRange !== 'object') {
    return 'dateRange: object with start/end required';
  }
  var startDate = new Date(payload.dateRange.start);
  var endDate = new Date(payload.dateRange.end);
  if (isNaN(startDate.getTime())) return 'dateRange.start: invalid date';
  if (isNaN(endDate.getTime())) return 'dateRange.end: invalid date';
  if (startDate.getTime() > endDate.getTime()) return 'dateRange.start must be <= dateRange.end';
  // 期間上限 60 日 (GAS 6 分制限対策)
  var days = (endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24);
  if (days > 60) return 'dateRange: span too long (max 60 days)';

  var slotLen = payload.slotLengthMinutes;
  if (slotLen != null) {
    if (typeof slotLen !== 'number' || slotLen < 15 || slotLen > 240) {
      return 'slotLengthMinutes: must be number in [15, 240]';
    }
  }

  var count = payload.count;
  if (count != null) {
    if (typeof count !== 'number' || count < 1 || count > 20) {
      return 'count: must be number in [1, 20]';
    }
  }

  if (payload.timeRange) {
    if (typeof payload.timeRange !== 'object') return 'timeRange: object with start/end required';
    if (payload.timeRange.start && !/^\d{2}:\d{2}$/.test(payload.timeRange.start)) {
      return 'timeRange.start: must be "HH:mm" format';
    }
    if (payload.timeRange.end && !/^\d{2}:\d{2}$/.test(payload.timeRange.end)) {
      return 'timeRange.end: must be "HH:mm" format';
    }
  }

  return null;
}

/**
 * webapp payload を既存 findAvailableSlots(request) のスキーマに変換する。
 * findAvailableSlots は以下を期待:
 *   request.participants:   string[]
 *   request.dateRange:      { start: Date, end: Date }
 *   request.timeRange:      { start: 'HH:mm', end: 'HH:mm' }
 *   request.duration:       number (minutes)
 *   request.priority:       string (任意)
 *   request.requestId:      string (任意)
 *   request.title:          string (任意)
 */
function normalizeRequestForLegacy_(payload) {
  return {
    requestId: 'webapp-' + new Date().getTime(),
    title: payload.title || 'Web App 経由のリクエスト',
    organizer: payload.organizer || '',
    participants: payload.participants.slice(),
    dateRange: {
      start: new Date(payload.dateRange.start),
      end: new Date(payload.dateRange.end)
    },
    timeRange: {
      start: (payload.timeRange && payload.timeRange.start) || '09:00',
      end: (payload.timeRange && payload.timeRange.end) || '18:00'
    },
    duration: payload.slotLengthMinutes || 30,
    priority: payload.priority || '',
    notes: payload.notes || ''
  };
}

/**
 * JSON レスポンスを ContentService で返す。
 */
function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
