# 申請書自動入力アプリ セットアップ手順

このドキュメントは、ローカル開発・本番デプロイ（Render）・Google Cloud 設定を一括で参照できる統合ガイドです。

## 目次

1. [前提条件](#1-前提条件)
2. [バックエンド環境構築](#2-バックエンド環境構築)
3. [環境変数](#3-環境変数)
4. [テンプレートファイル](#4-テンプレートファイル)
5. [Google Cloud 設定](#5-google-cloud-設定)
6. [フロントエンド環境構築](#6-フロントエンド環境構築)
7. [アプリケーションの起動](#7-アプリケーションの起動)
8. [Render へのデプロイ](#8-render-へのデプロイ)
9. [メール添付（オプション）](#9-メール添付オプション)
10. [トラブルシューティング](#10-トラブルシューティング)

---

## 1. 前提条件

- **Python 3.8以上**（バックエンド）
- **Node.js 18以上** および **npm**（フロントエンド）
- **Git**（任意）

---

## 2. バックエンド環境構築

```bash
cd backend
python -m venv venv
```

**Windows:**
```bash
venv\Scripts\activate
```

**macOS/Linux:**
```bash
source venv/bin/activate
```

```bash
pip install -r requirements.txt
```

`.env` を `backend/.env.example` を元に作成し、必要な値を設定してください（[環境変数](#3-環境変数)を参照）。

---

## 3. 環境変数

`backend/.env` に以下を設定します。`backend/.env.example` をコピーして編集してください。

### 必須（Drive / Sheets 利用時）

| 変数名 | 説明 |
|--------|------|
| `GOOGLE_DRIVE_CREDENTIALS_PATH` | サービスアカウント JSON のパス。ローカル例: `./companyestablishsupport-a633c02517ff.json` |
| `GOOGLE_DRIVE_FOLDER_ID` | 生成ファイルをアップロードする**自分で作成した** Google ドライブのフォルダ ID（URL の `folders/` の後） |
| `GOOGLE_SHEETS_SPREADSHEET_ID` | 申請内容を記録するスプレッドシートの ID（URL の `d/` の後） |

### 任意（翻訳・住所変換）

| 変数名 | 説明 |
|--------|------|
| `GOOGLE_TRANSLATE_API_KEY` | ひらがな/カタカナ変換・翻訳用 |
| `GOOGLE_MAPS_API_KEY` | 住所の日本語変換用 |

### 任意（Slack）

| 変数名 | 説明 |
|--------|------|
| `SLACK_WEBHOOK_URL` | テキスト通知用 |
| `SLACK_BOT_TOKEN` / `SLACK_CHANNEL_ID` / `SLACK_USER_ID` | ファイルアップロード用 |

### その他

| 変数名 | 説明 |
|--------|------|
| `TEMPLATE_DIR` | テンプレートの配置ディレクトリ（未設定時は `backend` のカレントディレクトリ） |
| `PORT` | サーバーポート（未設定時は `10000`。Render では自動設定） |

`.env` は Git にコミットしないでください。

---

## 4. テンプレートファイル

以下のファイルを `backend` ディレクトリに配置してください。

- `template-word-registration-application.docx`（登記申請書）
- `template-word-article-of-incorporation.docx`（定款）
- `template-excel-seal-registration.xlsx`（印鑑届出書）

`TEMPLATE_DIR` を設定している場合は、そのディレクトリに配置します。

---

## 5. Google Cloud 設定

### 5.1 API の有効化

[Google Cloud Console](https://console.cloud.google.com/) で以下を有効化します。

- **Google Drive API**
- **Google Sheets API**

（翻訳・住所変換を使う場合は Google Translate API・Maps 関連 API も有効化）

### 5.2 サービスアカウント

1. [IAMと管理] → [サービスアカウント] → [サービスアカウントを作成]
2. 名前は任意（例: `company-form-automation`）
3. [キー] タブから [キーを追加] → [新しいキーを作成] → **JSON** を選択してダウンロード
4. ダウンロードした JSON を `backend` に配置し、`GOOGLE_DRIVE_CREDENTIALS_PATH` でそのパスを指定

### 5.3 Google ドライブのフォルダ

1. **自分**の Google ドライブで、アップロード用のフォルダを 1 つ作成
2. そのフォルダを**共有**し、サービスアカウントのメール（`xxx@xxx.iam.gserviceaccount.com`）を**編集者**で追加
3. フォルダを開いた状態の URL の `folders/` の後ろの文字列がフォルダ ID → `GOOGLE_DRIVE_FOLDER_ID` に設定

重要: サービスアカウントにはストレージクォータがありません。**必ず自分で作成したフォルダ**を共有して、その ID を指定してください。

### 5.4 Google スプレッドシート

1. [Google スプレッドシート](https://sheets.google.com/) で新規作成
2. 1 行目にヘッダーを設定（例: 生成日時, 会社名, 代表者名, …）
3. [共有] でサービスアカウントのメールを**編集者**で追加
4. URL の `d/` の後ろがスプレッドシート ID → `GOOGLE_SHEETS_SPREADSHEET_ID` に設定

### 5.5 動作確認

```bash
cd backend
# 仮想環境を有効化した状態で
python test_google_connection.py
```

認証・フォルダ・スプレッドシートへのアクセスが成功するか確認できます。

---

## 6. フロントエンド環境構築

```bash
cd frontend
npm install
```

バックエンドの URL を変える場合のみ、`frontend/.env` に `VITE_API_BASE_URL=http://localhost:10000` などを設定します。通常は `vite.config.ts` のプロキシで足ります。

---

## 7. アプリケーションの起動

1. **バックエンド**
   ```bash
   cd backend
   # 仮想環境を有効化
   python main.py
   ```
   → `http://localhost:10000`（API ドキュメント: `http://localhost:10000/docs`）

2. **フロントエンド**（別ターミナル）
   ```bash
   cd frontend
   npm run dev
   ```
   → `http://localhost:3000`

フォーム送信後、申請は `POST /submit-application` で受理され、バックグラウンドで登記申請書・定款・印鑑届出書の生成 → Drive アップロード → スプレッドシート記録が行われます。

---

## 8. Render へのデプロイ

### 8.1 前提

- リポジトリが GitHub に push 済み
- [Google Cloud 設定](#5-google-cloud-設定) が完了していること

### 8.2 サービス作成

1. [Render](https://render.com) で [New +] → **Web Service**
2. リポジトリを選択
3. **Root Directory**: `backend`
4. **Build Command**: `pip install -r requirements.txt`
5. **Start Command**: `python main.py`
6. **Environment**: Python 3

### 8.3 認証情報（Secret File）

1. 対象 Web Service → **Environment** → **Secret Files**
2. [Add Secret File]
   - **Name**: 例 `google-drive-credentials.json`
   - **Path**: `/etc/secrets/google-drive-credentials.json`
   - **Content**: サービスアカウント JSON の中身を貼り付け

### 8.4 環境変数

**Environment** → **Environment Variables** に追加:

- `GOOGLE_DRIVE_CREDENTIALS_PATH` = `/etc/secrets/google-drive-credentials.json`（Secret File の Path と一致させる）
- `GOOGLE_DRIVE_FOLDER_ID` = 共有したフォルダの ID
- `GOOGLE_SHEETS_SPREADSHEET_ID` = スプレッドシートの ID

必要に応じて `GOOGLE_TRANSLATE_API_KEY`・`GOOGLE_MAPS_API_KEY` も設定。

### 8.5 デプロイ後確認

- Render の **Logs** でエラーがないか確認
- フロントからフォーム送信し、指定フォルダに 3 ファイルがアップロードされ、スプレッドシートに 1 行追加されることを確認

### 8.6 よくあるエラー

- **認証情報ファイルが見つかりません**  
  Secret File の Path と `GOOGLE_DRIVE_CREDENTIALS_PATH` が一致しているか確認
- **403 accessNotConfigured**  
  Google Cloud で Drive API / Sheets API が有効か確認
- **403 storageQuotaExceeded**  
  `GOOGLE_DRIVE_FOLDER_ID` に**自分で作成し、サービスアカウントを共有した**フォルダの ID を指定しているか確認

---

## 9. メール添付（オプション）

申請受付後、生成した 3 ファイルを申請者メールに添付して送る機能を追加する場合は、別ドキュメントを参照してください。

- **[docs/email-attachment-setup.md](email-attachment-setup.md)** … SMTP / Gmail API の選び方、環境変数、`_background_submit_task` からの呼び出し方

---

## 10. トラブルシューティング

### モジュールが見つからない

- 仮想環境を有効化しているか確認
- `pip install -r requirements.txt` を再実行

### 環境変数が読み込まれない

- `.env` が `backend` ディレクトリにあるか確認
- `python-dotenv` が入っているか確認（`requirements.txt` に含まれています）

### テンプレートファイルが見つからない

- ファイル名が `template-word-registration-application.docx` / `template-word-article-of-incorporation.docx` / `template-excel-seal-registration.xlsx` か確認
- `TEMPLATE_DIR` を設定している場合は、そのディレクトリに配置

### 認証情報ファイルが見つかりません

- `GOOGLE_DRIVE_CREDENTIALS_PATH` のパスが正しいか、ファイルが存在するか確認
- 相対パスの場合、`backend` をカレントにして実行しているか確認

### 接続タイムアウト（WinError 10060 等）

- 企業ネットワーク・プロキシ・VPN の影響の可能性
- `backend/test_google_connection.py` で詳細を確認
- ローカルで続く場合は Render にデプロイして試すと解消することがあります

### API リクエストが失敗する（フロント）

- バックエンドが起動しているか確認
- `vite.config.ts` のプロキシ設定と CORS を確認

---

## 参考

- API 仕様: 起動中のバックエンドの `http://localhost:10000/docs`（Swagger UI）
- メール添付の実装: [docs/email-attachment-setup.md](email-attachment-setup.md)
