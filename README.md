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
