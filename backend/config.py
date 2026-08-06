# config.py
import os
from dotenv import load_dotenv

# .env ファイルから環境変数を読み込む
load_dotenv()

# Google API キー
GOOGLE_TRANSLATE_API_KEY = os.getenv("GOOGLE_TRANSLATE_API_KEY")
GOOGLE_MAPS_API_KEY = os.getenv("GOOGLE_MAPS_API_KEY")

# Slack Webhook（テキスト通知用）
SLACK_WEBHOOK_URL = os.getenv("SLACK_WEBHOOK_URL")

# Slack Bot Token（ファイルアップロード用）
SLACK_BOT_TOKEN = os.getenv("SLACK_BOT_TOKEN")
SLACK_CHANNEL_ID = os.getenv("SLACK_CHANNEL_ID")
SLACK_USER_ID = os.getenv("SLACK_USER_ID")

# Google Drive API設定
# ローカル環境のデフォルトパス（本番環境では.envで上書き）
GOOGLE_DRIVE_CREDENTIALS_PATH = os.getenv(
    "GOOGLE_DRIVE_CREDENTIALS_PATH", 
    "./companyestablishsupport-a633c02517ff.json"
)
GOOGLE_DRIVE_FOLDER_ID = os.getenv("GOOGLE_DRIVE_FOLDER_ID")  # 親フォルダID（オプション）
GOOGLE_SHEETS_SPREADSHEET_ID = os.getenv("GOOGLE_SHEETS_SPREADSHEET_ID")  # Spreadsheet ID

# ==========
# Google Drive (OAuth ユーザー認証用)
# ==========
GOOGLE_DRIVE_OAUTH_CLIENT_ID = os.getenv("GOOGLE_DRIVE_OAUTH_CLIENT_ID")
GOOGLE_DRIVE_OAUTH_CLIENT_SECRET = os.getenv("GOOGLE_DRIVE_OAUTH_CLIENT_SECRET")
GOOGLE_DRIVE_OAUTH_REFRESH_TOKEN = os.getenv("GOOGLE_DRIVE_OAUTH_REFRESH_TOKEN")

# 通常は固定で OK
GOOGLE_DRIVE_OAUTH_TOKEN_URI = os.getenv(
    "GOOGLE_DRIVE_OAUTH_TOKEN_URI",
    "https://oauth2.googleapis.com/token",
)

# Drive で使うスコープ
GOOGLE_DRIVE_OAUTH_SCOPES = [
    "https://www.googleapis.com/auth/drive.file",
]


# その他（例：Notionなど）※使っていないなら削除可
NOTION_API_KEY = os.getenv("NOTION_API_KEY")
NOTION_DATABASE_ID = os.getenv("NOTION_DATABASE_ID")

# テンプレート・生成ファイルの保存先（未設定時はカレントディレクトリ。Render では設定不要）
TEMPLATE_DIR = os.getenv("TEMPLATE_DIR", os.path.dirname(os.path.abspath(__file__)) or ".")

# OpenAI（カタカナ変換用・任意）
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_KATAKANA_MODEL = os.getenv("OPENAI_KATAKANA_MODEL", "gpt-4o-mini")

# PostgreSQL + JWT（ログイン・申請保存用）
DATABASE_URL = os.getenv("DATABASE_URL")
JWT_SECRET_KEY = os.getenv("JWT_SECRET_KEY", "change-me-in-production")
JWT_ALGORITHM = os.getenv("JWT_ALGORITHM", "HS256")
ACCESS_TOKEN_EXPIRE_MINUTES = int(os.getenv("ACCESS_TOKEN_EXPIRE_MINUTES", "43200"))  # 30日相当（テスト用。本番は短く推奨）

# テストアカウント（起動時に DB に未存在なら作成）
AUTH_TEST_EMAIL = os.getenv("AUTH_TEST_EMAIL", "test@example.com")
AUTH_TEST_PASSWORD = os.getenv("AUTH_TEST_PASSWORD", "testpassword123")

# ローカル用サンプルデータ（申請済みユーザー）。本番では true にしないこと
SEED_SAMPLE_DATA = os.getenv("SEED_SAMPLE_DATA", "false").strip().lower() in (
    "1",
    "true",
    "yes",
    "on",
)
SAMPLE_APPLIED_EMAIL = os.getenv("SAMPLE_APPLIED_EMAIL", "applied@example.com")
SAMPLE_APPLIED_PASSWORD = os.getenv("SAMPLE_APPLIED_PASSWORD", "testpassword123")

# Monthly Payment Requests → Google Sheets tab name
GOOGLE_SHEETS_PAYMENT_SHEET = os.getenv("GOOGLE_SHEETS_PAYMENT_SHEET", "PaymentRequests")

# Admin API for marking payment requests paid (header X-Admin-Key)
ADMIN_API_KEY = os.getenv("ADMIN_API_KEY", "")

# Max active (non-paid) payment requests per user
PAYMENT_REQUEST_MAX_ACTIVE = int(os.getenv("PAYMENT_REQUEST_MAX_ACTIVE", "10"))
