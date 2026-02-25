# test_google_connection.py
import os
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
import httplib2
from google_auth_httplib2 import AuthorizedHttp

load_dotenv()

def test_google_drive_connection():
    """Google Drive APIへの接続をテスト"""
    print("=" * 60)
    print("Google Drive API 接続テスト")
    print("=" * 60)
    
    # 設定値の確認
    credentials_path = os.getenv("GOOGLE_DRIVE_CREDENTIALS_PATH", "./companyestablishsupport-a633c02517ff.json")
    folder_id = os.getenv("GOOGLE_DRIVE_FOLDER_ID")
    
    print(f"\n📋 設定値:")
    print(f"  - 認証情報ファイル: {credentials_path}")
    print(f"  - 存在確認: {os.path.exists(credentials_path)}")
    print(f"  - GOOGLE_DRIVE_FOLDER_ID: {folder_id if folder_id else '(未設定)'}")
    
    if not folder_id:
        print("\n❌ GOOGLE_DRIVE_FOLDER_IDが設定されていません")
        print("   .envファイルに以下を追加してください：")
        print("   GOOGLE_DRIVE_FOLDER_ID=your-folder-id")
        print("   ※フォルダIDは、Googleドライブでフォルダを開いた際のURLから取得できます")
        print("   ※例: https://drive.google.com/drive/folders/1xxxxxxxxxxxxxxxxxxxxxxxxxxxxx")
        return False
    
    try:
        # 認証情報を読み込み
        print(f"\n🔍 認証情報を読み込み中...")
        creds = service_account.Credentials.from_service_account_file(
            credentials_path,
            scopes=['https://www.googleapis.com/auth/drive']
        )
        print(f"✅ 認証情報の読み込み成功")
        
        # HTTPクライアントを作成
        print(f"🔍 HTTPクライアントを作成中...")
        http = httplib2.Http(timeout=300)
        authorized_http = AuthorizedHttp(creds, http=http)
        print(f"✅ HTTPクライアントの作成成功")
        
        # Google Drive APIサービスを取得
        print(f"🔍 Google Drive APIサービスを取得中...")
        drive_service = build('drive', 'v3', http=authorized_http)
        print(f"✅ Google Drive APIサービスの取得成功")
        
        # フォルダ情報を取得してテスト
        print(f"\n🔍 フォルダ情報を取得中...")
        folder = drive_service.files().get(fileId=folder_id).execute()
        print(f"✅ 接続成功!")
        print(f"  - フォルダ名: {folder.get('name')}")
        print(f"  - フォルダID: {folder.get('id')}")
        print(f"  - 作成日時: {folder.get('createdTime')}")
        
        return True
    except Exception as e:
        print(f"\n❌ 接続エラー: {e}")
        import traceback
        print(traceback.format_exc())
        return False

def test_google_sheets_connection():
    """Google Sheets APIへの接続をテスト"""
    print("\n" + "=" * 60)
    print("Google Sheets API 接続テスト")
    print("=" * 60)
    
    # 設定値の確認
    credentials_path = os.getenv("GOOGLE_DRIVE_CREDENTIALS_PATH", "./companyestablishsupport-a633c02517ff.json")
    spreadsheet_id = os.getenv("GOOGLE_SHEETS_SPREADSHEET_ID")
    
    print(f"\n📋 設定値:")
    print(f"  - 認証情報ファイル: {credentials_path}")
    print(f"  - 存在確認: {os.path.exists(credentials_path)}")
    print(f"  - GOOGLE_SHEETS_SPREADSHEET_ID: {spreadsheet_id if spreadsheet_id else '(未設定)'}")
    
    if not spreadsheet_id:
        print("\n❌ GOOGLE_SHEETS_SPREADSHEET_IDが設定されていません")
        print("   .envファイルに以下を追加してください：")
        print("   GOOGLE_SHEETS_SPREADSHEET_ID=your-spreadsheet-id")
        return False
    
    try:
        # 認証情報を読み込み
        print(f"\n🔍 認証情報を読み込み中...")
        creds = service_account.Credentials.from_service_account_file(
            credentials_path,
            scopes=['https://www.googleapis.com/auth/spreadsheets']
        )
        print(f"✅ 認証情報の読み込み成功")
        
        # HTTPクライアントを作成
        print(f"🔍 HTTPクライアントを作成中...")
        http = httplib2.Http(timeout=300)
        authorized_http = AuthorizedHttp(creds, http=http)
        print(f"✅ HTTPクライアントの作成成功")
        
        # Google Sheets APIサービスを取得
        print(f"🔍 Google Sheets APIサービスを取得中...")
        sheets_service = build('sheets', 'v4', http=authorized_http)
        print(f"✅ Google Sheets APIサービスの取得成功")
        
        # スプレッドシート情報を取得してテスト
        print(f"\n🔍 スプレッドシート情報を取得中...")
        spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        print(f"✅ 接続成功!")
        print(f"  - スプレッドシート名: {spreadsheet.get('properties', {}).get('title')}")
        print(f"  - スプレッドシートID: {spreadsheet_id}")
        
        return True
    except Exception as e:
        print(f"\n❌ 接続エラー: {e}")
        import traceback
        print(traceback.format_exc())
        return False

if __name__ == "__main__":
    print("Google API 接続テストを開始します...\n")
    
    drive_result = test_google_drive_connection()
    sheets_result = test_google_sheets_connection()
    
    print("\n" + "=" * 60)
    print("テスト結果")
    print("=" * 60)
    print(f"Google Drive API: {'✅ 成功' if drive_result else '❌ 失敗'}")
    print(f"Google Sheets API: {'✅ 成功' if sheets_result else '❌ 失敗'}")
    
    if drive_result and sheets_result:
        print("\n✅ すべての接続テストが成功しました！")
    else:
        print("\n⚠️ 一部の接続テストが失敗しました。上記のエラーメッセージを確認してください。")




