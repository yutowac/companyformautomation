# 生成ファイルを添付したメール送信機能の実装手順

申請受付後、バックエンドでファイル生成・Drive アップロード・Spreadsheet 記録が完了したあと、申請者（フォームのメールアドレス）に生成した 3 ファイルを添付したメールを送信する機能を追加する場合の手順です。

---

## 1. 送信方式の選択

### A. SMTP（推奨: 設定が簡単）

- 送信専用のメールアドレスを用意する（例: Gmail のアプリパスワード、または社内 SMTP サーバー）。
- Python 標準の `smtplib` と `email.mime` で実装可能。追加ライブラリは不要。

### B. Gmail API

- Google Cloud で Gmail API を有効化し、OAuth 2.0 またはサービスアカウント（ドメイン委任）の認証情報を用意する。
- 既存の Drive 用サービスアカウントとは別に、Gmail 送信用の認証・スコープ（`https://www.googleapis.com/auth/gmail.send`）が必要。
- `google-api-python-client` で `users().messages().send()` を利用する。

---

## 2. 環境変数

### SMTP の場合（backend/.env および Render の Environment）

```env
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=your-sender@example.com
SMTP_PASSWORD=your-app-password
MAIL_FROM=your-sender@example.com
MAIL_FROM_NAME=One-Stop Inc.
# メール送信を有効にするフラグ（任意）
ENABLE_EMAIL_ATTACHMENTS=true
```

### Gmail API の場合

- 認証 JSON のパス（例: `GMAIL_CREDENTIALS_PATH`）または、トークン用のパス。
- スコープ: `https://www.googleapis.com/auth/gmail.send`

---

## 3. バックエンド実装の流れ

### 3.1 メール送信関数の追加（backend/main.py または別モジュール）

1. 次のようなシグネチャの関数を 1 つ用意する。

   ```python
   def send_email_with_attachments(
       to_email: str,
       subject: str,
       body: str,
       attachment_paths: list[str],
   ) -> None:
       ...
   ```

2. **SMTP の場合**
   - `smtplib.SMTP`（または `SMTP_SSL`）で接続し、`starttls()` のあと `login(SMTP_USER, SMTP_PASSWORD)`。
   - `email.mime.multipart.MIMEMultipart` でメッセージを作成し、本文を `MIMEText` で追加。
   - 各 `attachment_paths` に対して `email.mime.base.MIMEBase` で添付（`application/octet-stream`）、ファイル名は `os.path.basename(path)` で設定。
   - `msg.attach()` で添付を追加し、`smtp.sendmail(MAIL_FROM, to_email, msg.as_string())` で送信。

3. **Gmail API の場合**
   - 認証済みの `build('gmail', 'v1', ...)` でサービスを取得。
   - MIME メッセージを組み立て（`MIMEMultipart` + 本文 + 添付）、`base64.urlsafe_b64encode` でエンコード。
   - `service.users().messages().send(userId='me', body={'raw': raw})` で送信。

### 3.2 呼び出しタイミング

- `_background_submit_task`（`POST /submit-application` のバックグラウンドタスク）の**最後**で呼ぶ。
- 処理順: 登記申請書生成 → 定款生成 → 印鑑届出書生成 → `append_to_spreadsheet` → **メール送信**。
- 3 つの生成処理で得た出力ファイルパス（`output_path`）をリストにまとめ、`send_email_with_attachments(data.email, 件名, 本文, [path1, path2, path3])` に渡す。
- メール送信は例外をキャッチし、失敗時はログのみ残して Drive/Spreadsheet の結果には影響させない。
- 環境変数 `ENABLE_EMAIL_ATTACHMENTS` が `true` のときだけ送信するようにすると運用しやすい。

---

## 4. 本番環境（Render）での注意

- **SMTP**: Render の **Environment** に上記の SMTP 関連変数を **Secret** で設定する。
- **Gmail API**: 認証 JSON やトークンは **Secret File** でマウントし、そのパスを環境変数で指定する。スコープに `gmail.send` を含めること。

---

## 5. 参考

- Python `smtplib`: https://docs.python.org/3/library/smtplib.html
- Python `email.mime`: https://docs.python.org/3/library/email.mime.html
- Gmail API Send Mail: https://developers.google.com/gmail/api/guides/sending
