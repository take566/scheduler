# 会議日程調整ツール（Meeting Scheduler）

Google スプレッドシート＋Apps Script で動かす会議日程調整ツールです。設計書は [meeting-scheduler-design.md](meeting-scheduler-design.md) を参照してください。

## clasp での管理

このプロジェクトは [clasp](https://github.com/google/clasp) で Google Apps Script をローカル開発・デプロイします。

### 必要な環境

- Node.js（推奨: 18 以上）
- npm で clasp をグローバルインストール: `npm install -g @google/clasp`

### 初回セットアップ

1. **ログイン**
   ```bash
   clasp login
   ```
   ブラウザで Google アカウントを認可します。

2. **新規 GAS プロジェクトとして作成する場合**
   ```bash
   clasp create --type sheets --title "会議調整ツール"
   ```
   実行すると `.clasp.json` に `scriptId` が書き込まれます。

3. **既存の GAS プロジェクトに紐付ける場合**
   - スプレッドシートを開く → 拡張機能 → Apps Script → プロジェクトの設定 → スクリプト ID をコピー
   - `.clasp.json` の `"scriptId": ""` にその ID を貼り付け

4. **プッシュ（GAS にアップロード）**
   ```bash
   clasp push
   ```

### よく使うコマンド

| コマンド | 説明 |
|----------|------|
| `clasp push` | ローカルの .gs / appsscript.json を GAS にアップロード |
| `clasp pull` | GAS のコードをローカルに取得（上書き注意） |
| `clasp open` | ブラウザで Apps Script エディタを開く |
| `clasp deploy` | 新しいバージョンとしてデプロイ（テスト用など） |
| `clasp logs` | 実行ログを表示 |

### ファイル構成（clasp が push するもの）

- `*.gs` … スクリプト（Code.gs, Config.gs, CalendarSync.gs など）
- `appsscript.json` … プロジェクト設定（タイムゾーン等）
- `.clasp.json` … スクリプト ID など（リポジトリに含めて共有可。別環境用なら .gitignore に追加しても可）

`.claspignore` により `*.md` などは push されません。

## デプロイ後の動作確認

スプレッドシートにスクリプトをデプロイしたあとの確認手順は [VERIFICATION.md](VERIFICATION.md) を参照してください。

## Web App デプロイ手順

外部システム (例: Slack bot) から POST で会議候補日時を取得するための Web App エンドポイント (`WebApp.gs`) を提供しています。デプロイ手順は以下:

1. **コードを GAS にアップロード**
   ```bash
   clasp push
   ```
2. **共有シークレットを生成して Script Properties に登録**
   ```bash
   openssl rand -hex 16    # 例: 32 文字の乱数
   ```
   - Apps Script エディタを開く: `clasp open`
   - 左サイドバーの ⚙️ "プロジェクトの設定" → "スクリプト プロパティ" → "スクリプト プロパティを追加"
   - キー: `SHARED_SECRET`、値: 上で生成した乱数
3. **Web App としてデプロイ**
   - エディタ右上の "デプロイ" → "新しいデプロイ"
   - 種類: ウェブアプリ
   - 説明: 任意 (例: `meeting-scheduler-webapp v1`)
   - 次のユーザーとして実行: **自分**（参加者カレンダー読み取りのため）
   - アクセスできるユーザー: **全員**（ANYONE_ANONYMOUS／共有シークレットで保護）
   - "デプロイ" → 表示される Web アプリ URL を控える
     ```
     https://script.google.com/macros/s/AKfycb.../exec
     ```
4. **呼び出し側 (OpenClaw) に環境変数を設定**
   ```bash
   # .env (OpenClaw リポジトリ側)
   MEETING_SCHEDULER_WEBAPP_URL=https://script.google.com/macros/s/AKfycb.../exec
   MEETING_SCHEDULER_SHARED_SECRET=<step 2 で生成した値>
   ```
5. **動作確認 (curl)**
   ```bash
   curl -L -X POST "${MEETING_SCHEDULER_WEBAPP_URL}?secret=${MEETING_SCHEDULER_SHARED_SECRET}" \
     -H "Content-Type: application/json" \
     -d '{
       "participants": ["a@example.com", "b@example.com"],
       "dateRange": {"start": "2026-05-08", "end": "2026-05-22"},
       "slotLengthMinutes": 30,
       "count": 5,
       "timezone": "Asia/Tokyo"
     }'
   ```
   `-L` は GAS の 302 redirect (`script.googleusercontent.com`) を追跡するため必須。

> **コード変更後の再デプロイ**: `clasp push` の後に "デプロイの管理" → 既存デプロイの ✏️ "編集" → "新しいバージョン" を選んでデプロイ。URL は変わりません (新規デプロイすると URL が変わるので注意)。

## API リファレンス: POST `{webappUrl}`

### 認証

URL クエリパラメータ `?secret=<SHARED_SECRET>` で渡します。
互換のため body 内 `_sharedSecret` でも受け付けます。

### リクエスト

| フィールド | 型 | 必須 | 説明 |
|---|---|---|---|
| `participants` | string[] | ✓ | 参加者メールアドレス (1〜20 件) |
| `dateRange.start` | string (YYYY-MM-DD or ISO) | ✓ | 検索期間 開始 |
| `dateRange.end` | string (YYYY-MM-DD or ISO) | ✓ | 検索期間 終了 (start ≤ end、最大 60 日) |
| `slotLengthMinutes` | number | (default 30) | スロット長 [15, 240] |
| `count` | number | (default 5) | 返却候補数 [1, 20] |
| `timezone` | string | (default Asia/Tokyo) | レスポンスの ISO 8601 文字列に使用 |
| `timeRange.start` | string ("HH:mm") | (default "09:00") | 1 日の検索開始時刻 |
| `timeRange.end` | string ("HH:mm") | (default "18:00") | 1 日の検索終了時刻 |
| `priority` | string | (任意) | "最短日時優先" / "午前優先" / "午後優先" など |
| `title` | string | (任意) | 候補返却に echo されるタイトル |

### レスポンス (成功時)

```json
{
  "ok": true,
  "slots": [
    {
      "start": "2026-05-08T14:00:00+09:00",
      "end": "2026-05-08T14:30:00+09:00",
      "score": 100,
      "available_count": 2,
      "has_tentative": false,
      "tentative_members": []
    }
  ],
  "request_echo": { ... },
  "total_candidates_before_truncation": 23
}
```

### レスポンス (エラー時)

```json
{ "ok": false, "error": "unauthorized: invalid or missing X-Shared-Secret" }
```

### 制約

- カレンダー同期データ (シート `カレンダー同期`) が事前に同期されている必要があります (別 trigger で `syncAllCalendars()` を回してください)。`doPost` 自体は同期を呼びません (GAS の 6 分実行制限のため)。
- 参加者の Google カレンダーに対する読み取り権限が、デプロイユーザーに付与されている必要があります (同一 Workspace ドメイン内が前提)。

