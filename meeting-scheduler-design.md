# 会議日程調整ツール 設計書

## 1. 概要

### 1.1 目的

Google スプレッドシートを基盤とした会議日程調整ツールを構築する。
参加者の Google カレンダーから予定を自動同期し、全員が空いている時間帯を自動で算出・提案することで、会議調整の手間を大幅に削減する。

### 1.2 対象ユーザー

- 社内チームの会議オーガナイザー（幹事）
- 会議参加者（候補者）

### 1.3 技術スタック

| 要素 | 技術 |
|------|------|
| フロントエンド | Google スプレッドシート + Google フォーム |
| バックエンド | Google Apps Script (GAS) |
| カレンダー連携 | Google Calendar API（GAS 組み込み） |
| データストア | Google スプレッドシート（各シート） |
| トリガー | GAS 時間主導型トリガー / フォーム送信トリガー |

---

## 2. システム構成

### 2.1 全体アーキテクチャ

```
┌──────────────┐     ┌──────────────────┐     ┌───────────────────┐
│ Google Form   │────▶│ Google Apps Script │────▶│ Google Spreadsheet │
│ (入力UI)      │     │ (ロジック層)       │     │ (データストア)      │
└──────────────┘     └────────┬─────────┘     └───────────────────┘
                              │
                     ┌────────▼─────────┐
                     │ Google Calendar   │
                     │ API               │
                     │ (予定取得)         │
                     └──────────────────┘
```

### 2.2 スプレッドシート構成（シート一覧）

| シート名 | 用途 |
|----------|------|
| `設定` | ツール全体の設定値（検索期間、スロット長など） |
| `参加者マスタ` | 会議候補者のメールアドレス・名前一覧 |
| `カレンダー同期` | Google カレンダーから取得した予定データ |
| `フォーム回答` | Google フォームからの回答データ（自動連携） |
| `空き時間算出` | 全参加者の空き時間の算出結果 |
| `候補日時` | 最終的な会議候補日時のランキング |

---

## 3. 機能詳細

### 3.1 カレンダー同期機能

#### 概要

指定した参加者の Google カレンダーから、指定期間内の予定（Busy 情報）を取得し、「カレンダー同期」シートに書き出す。

#### 処理フロー

```
1. 「参加者マスタ」シートから対象メールアドレスを取得
2. 「設定」シートから検索期間（開始日〜終了日）を取得
3. 各参加者について CalendarApp.getEvents() で予定を取得
4. 取得した予定を「カレンダー同期」シートに書き出し
5. 実行ログ（最終同期日時）を「設定」シートに記録
```

#### カレンダー同期シート スキーマ

| 列 | カラム名 | 型 | 説明 |
|----|----------|------|------|
| A | 参加者メール | String | カレンダーオーナーのメール |
| B | 参加者名 | String | 表示名 |
| C | 予定タイトル | String | イベント名（プライバシー設定に応じて「予定あり」） |
| D | 開始日時 | DateTime | イベント開始 |
| E | 終了日時 | DateTime | イベント終了 |
| F | 終日イベント | Boolean | TRUE / FALSE |
| G | ステータス | String | BUSY / TENTATIVE / FREE |
| H | 同期日時 | DateTime | データ取得タイムスタンプ |

#### 主要関数

```javascript
/**
 * 全参加者のカレンダー予定を同期する
 * メニューまたはトリガーから実行
 */
function syncAllCalendars() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings(ss);          // 設定シートから取得
  const members = getMembers(ss);            // 参加者マスタから取得
  const syncSheet = ss.getSheetByName('カレンダー同期');

  // 既存データクリア
  syncSheet.getRange(2, 1, syncSheet.getLastRow(), 8).clearContent();

  const startDate = settings.searchStartDate;
  const endDate = settings.searchEndDate;
  let allEvents = [];

  members.forEach(member => {
    try {
      const calendar = CalendarApp.getCalendarById(member.email);
      if (!calendar) {
        Logger.log(`カレンダー未取得: ${member.email}`);
        return;
      }
      const events = calendar.getEvents(startDate, endDate);
      events.forEach(event => {
        allEvents.push([
          member.email,
          member.name,
          event.getTitle(),
          event.getStartTime(),
          event.getEndTime(),
          event.isAllDayEvent(),
          event.getMyStatus() || 'BUSY',
          new Date()
        ]);
      });
    } catch (e) {
      Logger.log(`エラー (${member.email}): ${e.message}`);
    }
  });

  if (allEvents.length > 0) {
    syncSheet.getRange(2, 1, allEvents.length, 8).setValues(allEvents);
  }
}
```

#### 権限とアクセス制御

- スクリプト実行者がカレンダーへのアクセス権限を持つ必要あり
- 組織内ユーザーの場合、Google Workspace 管理者設定で「空き時間の共有」を有効化していること
- 他ユーザーのカレンダーへのアクセスは `CalendarApp.getCalendarById()` を使用（閲覧権限が必要）
- 権限不足の場合は FreeBusy API（`Calendar.Freebusy.query()`）にフォールバック

#### 同期スケジュール

| トリガー種別 | 間隔 | 説明 |
|-------------|------|------|
| 時間主導型トリガー | 毎時 | 定期的にカレンダーを自動同期 |
| スプレッドシートメニュー | 手動 | 「会議調整 > カレンダー同期」から即時実行 |

---

### 3.2 フォーム入力機能

#### 概要

Google フォームを通じて、会議の条件（参加者、希望日範囲、所要時間など）を入力する。フォーム送信時にトリガーが起動し、空き時間の算出を自動実行する。

#### フォーム項目設計

| # | 項目名 | 入力タイプ | 必須 | 説明 |
|---|--------|-----------|------|------|
| 1 | 会議タイトル | テキスト（短文） | ○ | 会議の名称 |
| 2 | 主催者メール | テキスト（短文） | ○ | フォーム回答者のメール |
| 3 | 参加者 | チェックボックス | ○ | 参加者マスタから選択（複数可） |
| 4 | 希望日の範囲（開始） | 日付 | ○ | 候補日の検索開始日 |
| 5 | 希望日の範囲（終了） | 日付 | ○ | 候補日の検索終了日 |
| 6 | 会議の所要時間 | プルダウン | ○ | 30分 / 60分 / 90分 / 120分 |
| 7 | 希望時間帯（開始） | プルダウン | ○ | 9:00〜18:00（30分刻み） |
| 8 | 希望時間帯（終了） | プルダウン | ○ | 9:00〜18:00（30分刻み） |
| 9 | 優先条件 | ラジオボタン | - | 最短日時優先 / 午前優先 / 午後優先 |
| 10 | 備考 | テキスト（長文） | - | 追加の要望など |

#### フォーム回答シート スキーマ

| 列 | カラム名 | 型 |
|----|----------|------|
| A | タイムスタンプ | DateTime |
| B | 会議タイトル | String |
| C | 主催者メール | String |
| D | 参加者（カンマ区切り） | String |
| E | 希望開始日 | Date |
| F | 希望終了日 | Date |
| G | 所要時間（分） | Number |
| H | 希望時間帯（開始） | Time |
| I | 希望時間帯（終了） | Time |
| J | 優先条件 | String |
| K | 備考 | String |
| L | 処理ステータス | String |

#### フォーム送信トリガー

```javascript
/**
 * フォーム送信時に自動実行されるトリガー関数
 */
function onFormSubmit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const responseSheet = ss.getSheetByName('フォーム回答');
  const lastRow = responseSheet.getLastRow();

  // 最新のカレンダーデータで同期
  syncAllCalendars();

  // フォーム回答をパース
  const request = parseFormResponse(responseSheet, lastRow);

  // 空き時間を算出
  const availableSlots = findAvailableSlots(request);

  // 結果を出力
  writeResults(ss, request, availableSlots);

  // 処理ステータスを更新
  responseSheet.getRange(lastRow, 12).setValue('処理完了');

  // 主催者に結果通知メールを送信
  sendResultNotification(request, availableSlots);
}
```

---

### 3.3 空き時間算出ロジック

#### 概要

カレンダー同期データとフォーム入力条件を突合し、全参加者が空いている時間スロットを算出する。

#### アルゴリズム

```
入力:
  - participants: 参加者メールリスト
  - dateRange: { start: Date, end: Date }
  - timeRange: { start: "09:00", end: "18:00" }
  - duration: 会議所要時間（分）
  - slotInterval: スロット間隔（デフォルト30分）

処理:
  1. dateRange 内の各日付について:
     a. timeRange 内で slotInterval 刻みの候補スロットを生成
        例: 9:00-10:00, 9:30-10:30, 10:00-11:00, ...
     b. 各候補スロットについて:
        - 全参加者のカレンダー予定と照合
        - 1人でも BUSY な予定と重複していれば除外
        - 終日イベントがある参加者は当日すべて除外
        - TENTATIVE は設定に応じて除外/警告
     c. 全員が空いているスロットを「空き」として記録

出力:
  - 空きスロットのリスト（日時昇順）
  - 各スロットの適合度スコア（優先条件を反映）
```

#### スコアリングロジック

| 優先条件 | スコア計算 |
|----------|-----------|
| 最短日時優先 | `score = MAX_SCORE - (スロット日時 - 現在日時)` |
| 午前優先 | 午前スロット: +10pt / 午後スロット: +0pt |
| 午後優先 | 午前スロット: +0pt / 午後スロット: +10pt |
| TENTATIVE あり | -5pt（暫定予定との重複がある場合） |

#### 主要関数

```javascript
/**
 * 空き時間スロットを算出する
 * @param {Object} request - フォーム入力から生成されたリクエストオブジェクト
 * @returns {Array} 空きスロットの配列
 */
function findAvailableSlots(request) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calendarData = getCalendarData(ss);
  const slots = [];
  const slotInterval = 30; // 分

  // 参加者ごとの予定をマップ化（高速検索用）
  const busyMap = buildBusyMap(calendarData, request.participants);

  // 日付ループ
  let currentDate = new Date(request.dateRange.start);
  while (currentDate <= request.dateRange.end) {
    // 土日を除外（設定による）
    if (isBusinessDay(currentDate)) {
      // 時間スロットループ
      let slotStart = combineDateAndTime(currentDate, request.timeRange.start);
      const dayEnd = combineDateAndTime(currentDate, request.timeRange.end);

      while (addMinutes(slotStart, request.duration) <= dayEnd) {
        const slotEnd = addMinutes(slotStart, request.duration);
        const conflict = checkConflicts(busyMap, request.participants, slotStart, slotEnd);

        if (!conflict.hasConflict) {
          slots.push({
            start: new Date(slotStart),
            end: new Date(slotEnd),
            score: calculateScore(slotStart, request.priority),
            hasTentative: conflict.hasTentative,
            tentativeMembers: conflict.tentativeMembers
          });
        }

        slotStart = addMinutes(slotStart, slotInterval);
      }
    }
    currentDate.setDate(currentDate.getDate() + 1);
  }

  // スコア降順でソート
  slots.sort((a, b) => b.score - a.score);
  return slots;
}

/**
 * 予定の衝突を検出する
 */
function checkConflicts(busyMap, participants, slotStart, slotEnd) {
  let hasConflict = false;
  let hasTentative = false;
  let tentativeMembers = [];

  for (const email of participants) {
    const events = busyMap[email] || [];
    for (const event of events) {
      // 時間の重複判定
      if (event.start < slotEnd && event.end > slotStart) {
        if (event.status === 'TENTATIVE') {
          hasTentative = true;
          tentativeMembers.push(email);
        } else {
          hasConflict = true;
          break;
        }
      }
    }
    if (hasConflict) break;
  }

  return { hasConflict, hasTentative, tentativeMembers };
}
```

#### 空き時間算出シート スキーマ

| 列 | カラム名 | 型 | 説明 |
|----|----------|------|------|
| A | リクエストID | String | フォーム回答のタイムスタンプから生成 |
| B | 会議タイトル | String | 会議名 |
| C | 候補日時（開始） | DateTime | スロット開始 |
| D | 候補日時（終了） | DateTime | スロット終了 |
| E | スコア | Number | 適合度スコア |
| F | 順位 | Number | スコアに基づくランク |
| G | TENTATIVE警告 | String | 暫定予定のある参加者名 |
| H | ステータス | String | 候補 / 確定 / 却下 |

---

### 3.4 結果出力・通知機能

#### 候補日時シート

空き時間算出結果から上位候補をピックアップし、見やすいフォーマットで「候補日時」シートに出力する。

| 列 | カラム名 | 説明 |
|----|----------|------|
| A | 順位 | 推奨順位 |
| B | 日付 | 候補日 |
| C | 時間帯 | 例: 10:00〜11:00 |
| D | 曜日 | 月〜金 |
| E | スコア | 適合度 |
| F | 備考 | TENTATIVE 警告など |
| G | 決定 | チェックボックス（確定時に使用） |

#### メール通知

```javascript
/**
 * 結果を主催者にメール通知する
 */
function sendResultNotification(request, slots) {
  const topSlots = slots.slice(0, 5); // 上位5候補

  let body = `会議「${request.title}」の日程候補が見つかりました。\n\n`;
  body += `参加者: ${request.participants.join(', ')}\n`;
  body += `所要時間: ${request.duration}分\n\n`;
  body += `--- 候補日時 TOP5 ---\n\n`;

  topSlots.forEach((slot, i) => {
    const dateStr = Utilities.formatDate(slot.start, 'Asia/Tokyo', 'yyyy/MM/dd (E)');
    const startStr = Utilities.formatDate(slot.start, 'Asia/Tokyo', 'HH:mm');
    const endStr = Utilities.formatDate(slot.end, 'Asia/Tokyo', 'HH:mm');
    body += `${i + 1}. ${dateStr} ${startStr}〜${endStr}`;
    if (slot.hasTentative) {
      body += ` ⚠ 暫定予定あり: ${slot.tentativeMembers.join(', ')}`;
    }
    body += '\n';
  });

  body += `\n詳細はスプレッドシートをご確認ください:\n`;
  body += SpreadsheetApp.getActiveSpreadsheet().getUrl();

  MailApp.sendEmail({
    to: request.organizer,
    subject: `[会議調整] 「${request.title}」の候補日時`,
    body: body
  });
}
```

---

## 4. カスタムメニュー

スプレッドシート上にカスタムメニューを追加し、各機能を手動実行できるようにする。

```javascript
function onOpen() {
  const ui = SpreadsheetApp.getUi();
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
```

---

## 5. 設定シート

| 行 | 設定項目 | デフォルト値 | 説明 |
|----|----------|-------------|------|
| 1 | 検索開始日 | （当日） | カレンダー検索の開始日 |
| 2 | 検索終了日 | （当日+14日） | カレンダー検索の終了日 |
| 3 | スロット間隔（分） | 30 | 候補スロットの刻み幅 |
| 4 | 営業時間（開始） | 09:00 | デフォルトの検索時間帯 |
| 5 | 営業時間（終了） | 18:00 | デフォルトの検索時間帯 |
| 6 | 土日除外 | TRUE | 土日を候補から除外するか |
| 7 | 祝日除外 | TRUE | 日本の祝日を除外するか |
| 8 | TENTATIVE 除外 | FALSE | 暫定予定を BUSY 扱いにするか |
| 9 | 最大候補数 | 10 | 出力する候補日時の上限 |
| 10 | 最終同期日時 | （自動更新） | 最後にカレンダー同期した日時 |

---

## 6. セットアップ手順

### 6.1 初期構築

```
1. Google スプレッドシートを新規作成
2. 上記6シートを作成し、ヘッダ行を設定
3. Apps Script エディタを開き、コードをデプロイ
4. Google フォームを作成し、スプレッドシートにリンク
5. 初期セットアップ関数を実行（シート作成・初期値設定）
```

### 6.2 権限設定

```
1. 初回実行時に OAuth 認可画面が表示される
2. 以下のスコープを許可:
   - Google Calendar（読み取り / 書き込み）
   - Gmail（メール送信）
   - Google Spreadsheet（読み書き）
   - Google Forms（フォーム連携）
3. 組織のカレンダー共有設定を確認
```

### 6.3 トリガー設定

```javascript
function setupTriggers() {
  // 既存トリガーを削除
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // カレンダー定期同期（毎時）
  ScriptApp.newTrigger('syncAllCalendars')
    .timeBased()
    .everyHours(1)
    .create();

  // フォーム送信トリガー
  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onFormSubmit()
    .create();
}
```

---

## 7. 制約事項・注意点

### 7.1 GAS の実行制限

| 制限項目 | 値 |
|----------|------|
| スクリプト実行時間 | 最大 6分（無料） / 30分（Workspace） |
| CalendarApp.getEvents() 呼び出し | 1日あたり制限あり |
| MailApp.sendEmail() | 1日100通（無料）/ 1,500通（Workspace） |
| トリガー数 | プロジェクトあたり最大20個 |

### 7.2 カレンダーアクセスの制約

- 他ユーザーのカレンダーは、閲覧権限が付与されている場合のみ取得可能
- 権限がない場合は FreeBusy API を使用するが、予定タイトルは取得不可（BUSY/FREE のみ）
- 外部ドメインのカレンダーは組織のポリシーにより制限される場合がある

### 7.3 データ量に関する注意

- 参加者が多い場合（10名以上）、同期処理に時間がかかる可能性あり
- 検索期間が長い場合（1ヶ月以上）はバッチ分割を推奨
- スプレッドシートの行数上限は約1,000万セル

---

## 8. 今後の拡張案

| 項目 | 概要 | 優先度 |
|------|------|--------|
| Slack 通知連携 | 候補日時を Slack チャンネルに投稿 | 高 |
| 投票機能 | 参加者が候補日時に対して投票できるフォーム | 高 |
| カレンダー自動登録 | 確定した日時で Google カレンダーにイベント作成 | 中 |
| 会議室予約連携 | Google Workspace の会議室リソースと連携 | 中 |
| 繰り返し会議対応 | 週次・月次の定例会議に対応 | 低 |
| ダッシュボード | 調整履歴・統計情報の可視化 | 低 |

---

## 9. ファイル構成

```
meeting-scheduler/
├── Code.gs                  # メインスクリプト（エントリポイント）
├── CalendarSync.gs          # カレンダー同期処理
├── SlotFinder.gs            # 空き時間算出ロジック
├── FormHandler.gs           # フォーム送信ハンドラ
├── Notification.gs          # メール通知
├── Utils.gs                 # ユーティリティ関数
├── Config.gs                # 定数・設定値
└── Setup.gs                 # 初期セットアップ・トリガー設定
```
