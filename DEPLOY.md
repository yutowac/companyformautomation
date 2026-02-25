# GitHub と Render でのデプロイ手順

このドキュメントでは、本リポジトリを **GitHub** にプッシュし、**Render** で API をデプロイするまでの手順を説明します。

---

## 1. GitHub の準備

### 1.1 リポジトリの作成

1. [GitHub](https://github.com) にログインし、**New repository** で新規リポジトリを作成します。
2. リポジトリ名は任意（例: `companyformautomation`）。
3. **Private** または **Public** を選択（認証情報を扱うため Private 推奨）。
4. **Add a README** 等は不要（既存のローカルを push するため）。

### 1.2 ローカルを Git リポジトリにして push

プロジェクトルート（`companyformautomation-main`）で実行します。

```powershell
cd C:\Users\WACHI.YUTO.P\Desktop\companyformautomation-main

# 既に .git がある場合はスキップ
git init

# リモートを追加（URL は自分のリポジトリに置き換え）
git remote add origin https://github.com/<あなたのユーザー名>/<リポジトリ名>.git

# 全ファイルを追加（.gitignore で除外されたものは含まれない）
git add .
git status   # .env / venv / node_modules 等が含まれていないことを確認

git commit -m "Initial commit: company form automation API + frontend"
git branch -M main
git push -u origin main
```

### 1.3 コミット前に確認すること

- **`.env`** がコミットされていないこと（`.gitignore` で除外済み）
- **`backend/venv/`** がコミットされていないこと
- **`backend/companyestablishsupport-*.json`**（Google 認証情報）がコミットされていないこと
- **`frontend/node_modules/`** がコミットされていないこと
- テンプレートファイル（`template-word-*.docx`, `template-excel-*.xlsx`）は **コミットする**（リポジトリに含める）

---

## 2. Render でのデプロイ

### 2.1 Render にサインアップ・ログイン

1. [Render](https://render.com) にアクセスし、**Get Started** でサインアップ（GitHub 連携が便利です）。
2. ダッシュボードで **New +** → **Web Service** を選択。

### 2.2 リポジトリの接続

1. **Connect a repository** で、先ほど push した GitHub リポジトリを選択。
2. 接続後、Render が `render.yaml` を検出する場合は **Apply** で Blueprint をそのまま使えます。
3. Blueprint を使わない場合は、手動で次のように設定します。

| 項目 | 値 |
|------|-----|
| **Name** | `companyformautomation-api`（任意） |
| **Region** | Singapore（または希望のリージョン） |
| **Root Directory** | `backend` |
| **Runtime** | Python 3 |
| **Build Command** | `pip install -r requirements.txt` |
| **Start Command** | `uvicorn main:app --host 0.0.0.0 --port $PORT` |

### 2.3 環境変数の設定

Render の **Environment** で、次の環境変数を **Secret** として設定します。

| キー | 説明 | 必須 |
|-----|------|------|
| `GOOGLE_TRANSLATE_API_KEY` | Google Translate API キー | ✅ |
| `GOOGLE_MAPS_API_KEY` | Google Maps API キー | ✅ |
| `SLACK_WEBHOOK_URL` | Slack Incoming Webhook URL | 任意 |
| `SLACK_BOT_TOKEN` | Slack Bot Token | 任意 |
| `SLACK_CHANNEL_ID` | Slack チャンネル ID | 任意 |
| `SLACK_USER_ID` | Slack ユーザー ID（DM 用等） | 任意 |
| `GOOGLE_DRIVE_CREDENTIALS_PATH` | 認証 JSON のパス（下記参照） | Drive 利用時 |
| `GOOGLE_DRIVE_FOLDER_ID` | アップロード先フォルダ ID | 任意 |
| `GOOGLE_SHEETS_SPREADSHEET_ID` | 記録先スプレッドシート ID | 任意 |

- **PORT** は Render が自動で設定するため、追加不要です。

### 2.4 Google サービスアカウント JSON の渡し方（Render）

認証情報ファイルは **Git に含めません**。Render では次のいずれかで渡します。

**方法 A: Secret File（推奨）**

1. Render の **Environment** で **Secret Files** を開く。
2. **Filename** に `credentials.json` など任意の名前を指定。
3. **Contents** に、ローカルの `companyestablishsupport-*.json` の内容をそのまま貼り付け。
4. 環境変数で `GOOGLE_DRIVE_CREDENTIALS_PATH=/etc/secrets/credentials.json` のように、Render が表示するパスを設定する。

**方法 B: 環境変数に JSON を入れる**

- キー名を `GOOGLE_DRIVE_CREDENTIALS_JSON` などにし、値に JSON 文字列を貼り付ける方法もあります。その場合は `config.py` 側で「環境変数が設定されていればファイルの代わりにその文字列を使う」処理を追加する必要があります（今回は未実装）。

### 2.5 デプロイ実行

1. **Create Web Service** でデプロイを開始します。
2. ビルド・起動が成功すると、`https://<サービス名>.onrender.com` で API にアクセスできます。
3. 動作確認: ブラウザで `https://<サービス名>.onrender.com/docs` を開き、Swagger UI が表示されれば OK です。

---

## 3. フロントエンドから本番 API を参照する場合

- ローカル開発: `frontend/vite.config.ts` の proxy で `http://localhost:10000` を参照。
- 本番: フロントエンドの API ベース URL を、Render の URL（例: `https://companyformautomation-api.onrender.com`）に変更してください（例: `frontend/.env.production` やビルド時の環境変数 `VITE_API_BASE_URL` など）。

---

## 4. トラブルシューティング

| 現象 | 確認すること |
|------|----------------|
| ビルド失敗 | Root Directory が `backend` になっているか。`requirements.txt` が `backend` にあるか。 |
| 起動失敗 | 環境変数（特に `GOOGLE_*`）が設定されているか。Start Command の `$PORT` がそのまま使われているか。 |
| テンプレート not found | テンプレート（.docx / .xlsx）がリポジトリの `backend/` に含まれているか。`TEMPLATE_DIR` は未設定でよい（デフォルトで backend のカレントディレクトリを使用）。 |
| 認証エラー | `GOOGLE_DRIVE_CREDENTIALS_PATH` が Secret File のパスと一致しているか。JSON の内容が正しいか。 |

---

## 5. 参考

- [Render Blueprint Spec](https://render.com/docs/blueprint-spec)
- [Render Web Services](https://render.com/docs/web-services)
- ローカル環境の詳細: `SETUP.md` / `LOCAL_SETUP.md`
