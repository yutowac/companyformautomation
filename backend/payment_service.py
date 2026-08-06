"""Helpers for Monthly Payment Requests: freeze window, validation, Sheets, Drive."""
from __future__ import annotations

import io
import os
import re
from datetime import datetime
from typing import Any
from zoneinfo import ZoneInfo

from config import (
    GOOGLE_DRIVE_CREDENTIALS_PATH,
    GOOGLE_DRIVE_FOLDER_ID,
    GOOGLE_DRIVE_OAUTH_CLIENT_ID,
    GOOGLE_DRIVE_OAUTH_CLIENT_SECRET,
    GOOGLE_DRIVE_OAUTH_REFRESH_TOKEN,
    GOOGLE_DRIVE_OAUTH_SCOPES,
    GOOGLE_DRIVE_OAUTH_TOKEN_URI,
    GOOGLE_SHEETS_PAYMENT_SHEET,
    GOOGLE_SHEETS_SPREADSHEET_ID,
)

JST = ZoneInfo("Asia/Tokyo")

ACCOUNT_TYPES = ("checking", "ordinary", "savings")
# Half-width katakana, prolonged sound, half-width dakuten/handakuten, space
HALF_WIDTH_KANA_RE = re.compile(r"^[\uFF65-\uFF9F\u0020]+$")
ACCOUNT_NUMBER_RE = re.compile(r"^\d{7}$")
AMOUNT_RE = re.compile(r"^\d+$")


def now_jst() -> datetime:
    return datetime.now(JST)


def is_payment_editable_now(when: datetime | None = None) -> bool:
    """Editable through the 20th of each month 23:59:59 JST (inclusive)."""
    dt = when or now_jst()
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=JST)
    else:
        dt = dt.astimezone(JST)
    if dt.day < 20:
        return True
    if dt.day > 20:
        return False
    # day == 20: allowed until end of day
    return True


def validate_payment_payload(data: dict[str, Any]) -> dict[str, Any]:
    payee = str(data.get("payeeName") or "").strip()
    bank = str(data.get("bankName") or "").strip()
    branch = str(data.get("branchName") or "").strip()
    account_type = str(data.get("accountType") or "").strip().lower()
    account_number = str(data.get("accountNumber") or "").strip()
    holder = str(data.get("accountHolderKana") or "").strip()
    amount = str(data.get("amountJpy") or "").strip().replace(",", "")
    invoice = str(data.get("invoiceNumber") or "").strip()

    if not payee:
        raise ValueError("Payee name is required")
    if not bank:
        raise ValueError("Bank name is required")
    if not branch:
        raise ValueError("Branch name is required")
    if account_type not in ACCOUNT_TYPES:
        raise ValueError("Account type must be checking, ordinary, or savings")
    if not ACCOUNT_NUMBER_RE.match(account_number):
        raise ValueError("Account number must be exactly 7 digits")
    if not holder or len(holder) > 30 or not HALF_WIDTH_KANA_RE.match(holder):
        raise ValueError("Account holder name must be half-width kana (max 30 characters)")
    if not AMOUNT_RE.match(amount):
        raise ValueError("Amount must be half-width digits (JPY)")

    return {
        "payeeName": payee,
        "bankName": bank,
        "branchName": branch,
        "accountType": account_type,
        "accountNumber": account_number,
        "accountHolderKana": holder,
        "amountJpy": amount,
        "invoiceNumber": invoice,
    }


def _get_sheets_service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    import httplib2
    from google_auth_httplib2 import AuthorizedHttp

    if not GOOGLE_DRIVE_CREDENTIALS_PATH or not os.path.exists(GOOGLE_DRIVE_CREDENTIALS_PATH):
        return None
    credentials = service_account.Credentials.from_service_account_file(
        GOOGLE_DRIVE_CREDENTIALS_PATH,
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    http = httplib2.Http(timeout=60)
    authorized_http = AuthorizedHttp(credentials, http=http)
    return build("sheets", "v4", http=authorized_http)


def append_payment_to_spreadsheet(
    *,
    user_email: str,
    payload: dict[str, Any],
    request_id: int,
    status: str,
    attachment_url: str = "",
) -> None:
    if not GOOGLE_SHEETS_SPREADSHEET_ID:
        print("⚠️ Spreadsheet ID not set; skipping payment sheet append.")
        return
    try:
        sheets_service = _get_sheets_service()
        if sheets_service is None:
            print("⚠️ Sheets credentials missing; skipping payment sheet append.")
            return
        row = [
            now_jst().strftime("%Y-%m-%d %H:%M:%S"),
            user_email,
            payload.get("payeeName", ""),
            payload.get("bankName", ""),
            payload.get("branchName", ""),
            payload.get("accountType", ""),
            payload.get("accountNumber", ""),
            payload.get("accountHolderKana", ""),
            payload.get("amountJpy", ""),
            payload.get("invoiceNumber", ""),
            attachment_url or "",
            str(request_id),
            status,
        ]
        sheet = GOOGLE_SHEETS_PAYMENT_SHEET or "PaymentRequests"
        sheets_service.spreadsheets().values().append(
            spreadsheetId=GOOGLE_SHEETS_SPREADSHEET_ID,
            range=f"'{sheet}'!A2",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": [row]},
        ).execute()
        print(f"✅ Payment spreadsheet append ok (request_id={request_id})")
    except Exception as e:
        print(f"❌ Payment spreadsheet append error: {e}")


def upload_payment_attachment(content: bytes, filename: str, mime_type: str) -> str:
    """Upload image to Drive; return webViewLink or empty string on failure."""
    try:
        from google.oauth2.credentials import Credentials as UserCredentials
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseUpload
        import httplib2
        from google_auth_httplib2 import AuthorizedHttp

        if not (
            GOOGLE_DRIVE_OAUTH_CLIENT_ID
            and GOOGLE_DRIVE_OAUTH_CLIENT_SECRET
            and GOOGLE_DRIVE_OAUTH_REFRESH_TOKEN
        ):
            print("⚠️ Drive OAuth not configured; skipping attachment upload.")
            return ""

        creds = UserCredentials(
            token=None,
            refresh_token=GOOGLE_DRIVE_OAUTH_REFRESH_TOKEN,
            token_uri=GOOGLE_DRIVE_OAUTH_TOKEN_URI,
            client_id=GOOGLE_DRIVE_OAUTH_CLIENT_ID,
            client_secret=GOOGLE_DRIVE_OAUTH_CLIENT_SECRET,
            scopes=GOOGLE_DRIVE_OAUTH_SCOPES,
        )
        http = httplib2.Http(timeout=60)
        authorized_http = AuthorizedHttp(creds, http=http)
        drive = build("drive", "v3", http=authorized_http)

        metadata: dict[str, Any] = {"name": filename}
        if GOOGLE_DRIVE_FOLDER_ID:
            metadata["parents"] = [GOOGLE_DRIVE_FOLDER_ID]

        media = MediaIoBaseUpload(io.BytesIO(content), mimetype=mime_type or "image/jpeg", resumable=False)
        created = (
            drive.files()
            .create(body=metadata, media_body=media, fields="id, webViewLink")
            .execute()
        )
        link = created.get("webViewLink") or ""
        if not link and created.get("id"):
            link = f"https://drive.google.com/file/d/{created['id']}/view"
        print(f"✅ Payment attachment uploaded: {filename}")
        return link
    except Exception as e:
        print(f"❌ Payment attachment upload error: {e}")
        return ""
