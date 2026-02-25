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

# その他（例：Notionなど）※使っていないなら削除可
NOTION_API_KEY = os.getenv("NOTION_API_KEY")
NOTION_DATABASE_ID = os.getenv("NOTION_DATABASE_ID")

# テンプレート・生成ファイルの保存先（未設定時はカレントディレクトリ。Render では設定不要）
TEMPLATE_DIR = os.getenv("TEMPLATE_DIR", os.path.dirname(os.path.abspath(__file__)) or ".")
