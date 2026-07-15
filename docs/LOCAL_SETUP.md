# ローカル環境での起動方法

> **PostgreSQL（Docker）とログイン／サンプルデータ**: 詳細は [local-database-setup.md](./local-database-setup.md) を参照してください。

## 前提条件

1. Python 3.8以上がインストールされていること
2. Node.js と npm がインストールされていること
3. `.env`ファイルに必要な環境変数が設定されていること

## 環境変数の設定

`backend/`ディレクトリに`.env`ファイルを作成し、以下の環境変数を設定してください：

```env
# Google API キー
GOOGLE_TRANSLATE_API_KEY=your-translate-api-key
GOOGLE_MAPS_API_KEY=your-maps-api-key

# Google Drive API設定
GOOGLE_DRIVE_CREDENTIALS_PATH=./google-drive-credentials.json
GOOGLE_DRIVE_FOLDER_ID=your-folder-id  # オプション
GOOGLE_SHEETS_SPREADSHEET_ID=your-spreadsheet-id

# Slack設定（既存）
SLACK_WEBHOOK_URL=your-slack-webhook-url
SLACK_BOT_TOKEN=your-slack-bot-token
SLACK_USER_ID=your-slack-user-id
```

## バックエンドの起動

1. バックエンドディレクトリに移動：
```bash
cd backend
```

2. 仮想環境を有効化（Windows）：
```bash
venv\Scripts\activate
```

3. 依存パッケージをインストール（初回のみ）：
```bash
pip install -r requirements.txt
```

4. バックエンドサーバーを起動：
```bash
python main.py
```

バックエンドサーバーは `http://localhost:10000` で起動します。

## フロントエンドの起動

別のターミナルウィンドウで：

1. フロントエンドディレクトリに移動：
```bash
cd frontend
```

2. 依存パッケージをインストール（初回のみ）：
```bash
npm install
```

3. 開発サーバーを起動：
```bash
npm run dev
```

フロントエンドは通常 `http://localhost:5173` で起動します（Viteのデフォルトポート）。

## 動作確認

1. ブラウザで `http://localhost:5173` にアクセス
2. フォームに必要な情報を入力
3. 「送信」ボタンをクリック
4. 以下を確認：
   - 3つのファイル（登記申請、定款、印鑑届出）が生成される
   - Googleドライブに会社名のフォルダが作成され、ファイルがアップロードされる
   - Google Sheetsにデータが記録される

## トラブルシューティング

### バックエンドが起動しない

- `.env`ファイルが正しく設定されているか確認
- `google-drive-credentials.json`ファイルが`backend/`ディレクトリに存在するか確認
- 仮想環境が有効化されているか確認

### フロントエンドが起動しない

- `node_modules`がインストールされているか確認（`npm install`を実行）
- ポート5173が使用中でないか確認

### Googleドライブへのアップロードが失敗する

- `GOOGLE_DRIVE_CREDENTIALS_PATH`が正しく設定されているか確認
- サービスアカウントのJSONキーファイルが正しいパスに配置されているか確認
- サービスアカウントに適切な権限が付与されているか確認

### Spreadsheetへの記録が失敗する

- `GOOGLE_SHEETS_SPREADSHEET_ID`が正しく設定されているか確認
- スプレッドシートにサービスアカウントが共有されているか確認
- スプレッドシートのヘッダー行が正しく設定されているか確認

## スプレッドシートの列名

スプレッドシートのヘッダー行は以下の順序で設定してください：

```
CreatedDate | CompanyName | RepresentativeName | RepresentativeBirthDay | RepresentativeAddress | BusinessPurpose1 | BusinessPurpose2 | BusinessPurpose3 | BusinessPurpose4 | BusinessPurpose5 | Email Address
```




