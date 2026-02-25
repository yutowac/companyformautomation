# backend/oauth_drive_refresh_token.py

from __future__ import annotations

import json
import pathlib

from google_auth_oauthlib.flow import InstalledAppFlow

# Drive で必要なスコープ
SCOPES = [
    "https://www.googleapis.com/auth/drive.file",
]

BASE_DIR = pathlib.Path(__file__).resolve().parent
# Google Cloud Console でダウンロードした OAuth クライアントの JSON
CLIENT_SECRET_FILE = BASE_DIR / "drive_oauth_client_secret.json"


def main() -> None:
    """
    1回だけ実行して、ユーザーの Drive アクセス用 refresh_token を取得するスクリプト。
    """
    if not CLIENT_SECRET_FILE.exists():
        raise FileNotFoundError(
            f"OAuth client secret JSON が見つかりません: {CLIENT_SECRET_FILE}\n"
            "Google Cloud Console から OAuth クライアント (デスクトップアプリ) を作成して "
            "JSON をこのパスに保存してください。"
        )

    flow = InstalledAppFlow.from_client_secrets_file(
        str(CLIENT_SECRET_FILE),
        scopes=SCOPES,
    )

    # ブラウザを立ち上げて Google アカウントでログイン → 同意
    creds = flow.run_local_server(port=8080, prompt="consent")

    print("\n=== OAuth 認証完了 ===")
    print(f"access_token: {creds.token!r}")
    print(f"refresh_token: {creds.refresh_token!r}")
    print(f"token_uri: {creds.token_uri!r}")
    print(f"client_id: {creds.client_id!r}")
    print(f"client_secret: {creds.client_secret!r}")
    print(f"scopes: {creds.scopes!r}")

    # 必要なら JSON でも保存（ローカル検証用）
    out_path = BASE_DIR / "drive_oauth_creds_sample.json"
    data = {
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": list(creds.scopes),
    }
    out_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nサンプル資格情報を {out_path} に書き出しました。")


if __name__ == "__main__":
    main()