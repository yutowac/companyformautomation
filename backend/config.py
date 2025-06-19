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

# その他（例：Notionなど）※使っていないなら削除可
NOTION_API_KEY = os.getenv("NOTION_API_KEY")
NOTION_DATABASE_ID = os.getenv("NOTION_DATABASE_ID")
