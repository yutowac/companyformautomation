# パッケージ
from fastapi import FastAPI, HTTPException, Response
from pydantic import BaseModel
import requests
from docx import Document
from openpyxl import load_workbook
from datetime import datetime
from io import BytesIO
import os
from config import GOOGLE_TRANSLATE_API_KEY, GOOGLE_MAPS_API_KEY
from fastapi.middleware.cors import CORSMiddleware

TEMPLATE_DIR = "/var/data/"

# SSL検証回避
requests.packages.urllib3.disable_warnings()

app = FastAPI()

# CORS 設定を追加
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 必要なら特定のオリジンに制限可能
    allow_credentials=True,
    allow_methods=["*"],  # すべてのHTTPメソッドを許可（GET, POST, OPTIONSなど）
    allow_headers=["*"],  # すべてのヘッダーを許可
)

class FormData(BaseModel):
    companyName: str
    address: str
    presidentName: str
    presidentAddress: str
    year: int
    month: int
    day: int
    birthyear: int
    birthmonth: int
    birthday: int
    purpose1: str
    purpose2: str
    purpose3: str
    purpose4: str
    purpose5: str

# Google Maps API を使用して住所を日本語に変換
def get_japanese_address(address: str) -> str:
    params = {
        "address": address,
        "key": GOOGLE_MAPS_API_KEY,
        "language": "ja"
    }
    response = requests.get("https://maps.googleapis.com/maps/api/geocode/json", params=params, verify=False)
    geocode_result = response.json()

    if geocode_result.get("status") == "OK":
        return geocode_result["results"][0]["formatted_address"].split("〒")[1][8:]
    else:
        raise HTTPException(status_code=500, detail="Address conversion failed")

# 翻訳関数
def translate_text(text: str, target_lang: str = "ja") -> str:
    params = {"q": text, "target": target_lang, "key": GOOGLE_TRANSLATE_API_KEY}
    response = requests.post("https://translation.googleapis.com/language/translate/v2", params=params, verify=False)
    translation_result = response.json()

    if "data" in translation_result:
        return translation_result["data"]["translations"][0]["translatedText"]
    else:
        raise HTTPException(status_code=500, detail="Translation failed")

# カタカナ変換用関数
def to_katakana(text: str) -> str:
    params = {"q": text, "target": "ja-Hira", "key": GOOGLE_TRANSLATE_API_KEY}
    response = requests.post("https://translation.googleapis.com/language/translate/v2", params=params, verify=False)
    translation_result = response.json()

    if "data" in translation_result:
        return translation_result["data"]["translations"][0]["translatedText"].replace(" ", "")
    else:
        raise HTTPException(status_code=500, detail="Katakana conversion failed")

# 法人届出書
@app.post("/generate-word")
def generate_word(data: FormData):
    template_path = "template_word_registration.docx"
    output_path = "created_registration.docx"

    if not os.path.exists(template_path):
        raise HTTPException(status_code=500, detail="Template file not found")
    try:
        doc = Document(template_path)
    except Exception as e:
        doc = Document(os.path.join(TEMPLATE_DIR, template_path))


    # 各入力フィールドを翻訳
    translated_company_name = to_katakana(data.companyName)
    translated_address = get_japanese_address(data.address)
    translated_president_name = to_katakana(data.presidentName)  # 名前をカタカナに変換
    translated_president_address = get_japanese_address(data.presidentAddress)
    translated_purpose = translate_text(data.purpose1)
    translated_purpose2 = translate_text(data.purpose2)
    translated_purpose3 = translate_text(data.purpose3)
    translated_purpose4 = translate_text(data.purpose4)
    translated_purpose5 = translate_text(data.purpose5)

    # テンプレートの ( ) 内の項目を翻訳結果で置換
    for paragraph in doc.paragraphs:
        paragraph.text = paragraph.text.replace("(A商号)", data.companyName)
        paragraph.text = paragraph.text.replace("(A商号のメインパートのフリガナ)", translated_company_name)
        paragraph.text = paragraph.text.replace("(Pending1B・本店住所フル)", translated_address)
        paragraph.text = paragraph.text.replace("(C社員住所)", translated_president_address)
        paragraph.text = paragraph.text.replace("(D社員氏名)", translated_president_name)
        paragraph.text = paragraph.text.replace("(E設立日・和暦)", str(data.year) + "年" + str(data.month) + "月" + str(data.day) + "日")
        paragraph.text = paragraph.text.replace("(G社員生年月日・暦年)", str(data.birthyear) + "年" + str(data.birthmonth) + "月" + str(data.birthday) + "日")
        paragraph.text = paragraph.text.replace("(B目的1)", translated_purpose)
        paragraph.text = paragraph.text.replace("(B目的2)", translated_purpose2)
        paragraph.text = paragraph.text.replace("(B目的3)", translated_purpose3)
        paragraph.text = paragraph.text.replace("(B目的4)", translated_purpose4)
        paragraph.text = paragraph.text.replace("(B目的5)", translated_purpose5)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                cell.text = cell.text.replace("(A商号)", data.companyName)
                cell.text = cell.text.replace("(A商号のメインパートのフリガナ)", translated_company_name)
                cell.text = cell.text.replace("(Pending1B・本店住所フル)", translated_address)
                cell.text = cell.text.replace("(C社員住所)", translated_president_address)
                cell.text = cell.text.replace("(D社員氏名)", translated_president_name)
                cell.text = cell.text.replace("(E設立日・和暦)", str(data.year) + "年" + str(data.month) + "月" + str(data.day) + "日")
                cell.text = cell.text.replace("(G社員生年月日・暦年)", str(data.birthyear) + "年" + str(data.birthmonth) + "月" + str(data.birthday) + "日")
                cell.text = cell.text.replace("(B目的1)", translated_purpose)
                cell.text = cell.text.replace("(B目的2)", translated_purpose2)
                cell.text = cell.text.replace("(B目的3)", translated_purpose3)
                cell.text = cell.text.replace("(B目的4)", translated_purpose4)
                cell.text = cell.text.replace("(B目的5)", translated_purpose5)

    # 生成された Word ファイルを保存
    doc.save(output_path)

    headers = {
        "Content-Disposition": "attachment; filename=created_registration.docx",
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    }

    with open(output_path, "rb") as file:
        return Response(content=file.read(), headers=headers, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

@app.get("/get-created-word")
def get_created_word():
    file_path = "created_registration.docx"
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Word file not found")

    try:
        with open(file_path, "rb") as file:
            file_stream = BytesIO(file.read())
    except Exception as e:
        with open(os.path.join(TEMPLATE_DIR, file_path), "rb") as file:
            file_stream = BytesIO(file.read())

    headers = {
        "Content-Disposition": "attachment; filename=created_registration.docx",
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    }
    return Response(content=file_stream.getvalue(), headers=headers, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# 定款作成
@app.post("/generate-word2")
def generate_word(data: FormData):
    template_path = "template_word_incorparticles.docx"
    output_path = "created_incorparticles.docx"

    if not os.path.exists(template_path):
        raise HTTPException(status_code=500, detail="Template file not found")
    try:
        doc = Document(template_path)
    except Exception as e:
        doc = Document(os.path.join(TEMPLATE_DIR, template_path))

    # 各入力フィールドを翻訳
    translated_address = get_japanese_address(data.address)
    translated_president_name = to_katakana(data.presidentName)  # 名前をカタカナに変換
    translated_president_address = get_japanese_address(data.presidentAddress)
    translated_purpose = translate_text(data.purpose1)
    translated_purpose2 = translate_text(data.purpose2)
    translated_purpose3 = translate_text(data.purpose3)
    translated_purpose4 = translate_text(data.purpose4)
    translated_purpose5 = translate_text(data.purpose5)

    if data.month == 1:
        E_month = 12
        E_day = 31
    elif data.month == 5 | data.month == 7 | data.month == 10 | data.month == 12:
        E_month = data.month - 1
        E_day = 30
    elif data.month == 3:
        E_month = data.month - 1
        E_day = 28
    else:
        E_month = data.month - 1
        E_day = 31

    # テンプレートの ( ) 内の項目を翻訳結果で置換
    for paragraph in doc.paragraphs:
        paragraph.text = paragraph.text.replace("(A商号)", data.companyName)
        paragraph.text = paragraph.text.replace("(本店住所●Pending1A=東京都△△区)", translated_address)
        paragraph.text = paragraph.text.replace("(C社員住所)", translated_president_address)
        paragraph.text = paragraph.text.replace("(D社員氏名)", translated_president_name)
        paragraph.text = paragraph.text.replace("(E設立日がある月の1日)", str(data.year) + "年" + str(data.month) + "月" + "1日")
        paragraph.text = paragraph.text.replace("(E設立日がある月から11ヶ月後の月末)", str(data.year+1) + "年" + str(E_month) + "月" + str(E_day) + "日")
        paragraph.text = paragraph.text.replace("(F定款作成日・暦年)", datetime.now().strftime("%Y年%m月%d日"))
        paragraph.text = paragraph.text.replace("(B目的1)", translated_purpose)
        paragraph.text = paragraph.text.replace("(B目的2)", translated_purpose2)
        paragraph.text = paragraph.text.replace("(B目的3)", translated_purpose3)
        paragraph.text = paragraph.text.replace("(B目的4)", translated_purpose4)
        paragraph.text = paragraph.text.replace("(B目的5)", translated_purpose5)

    # 生成された Word ファイルを保存
    doc.save(output_path)

    headers = {
        "Content-Disposition": "attachment; filename=created_incorparticles.docx",
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    }

    with open(output_path, "rb") as file:
        return Response(content=file.read(), headers=headers, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

@app.get("/get-created-word2")
def get_created_word():
    file_path = "created_incorparticles.docx"
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Word file not found")

    with open(file_path, "rb") as file:
        file_stream = BytesIO(file.read())

    headers = {
        "Content-Disposition": "attachment; filename=created_incorparticles.docx",
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    }
    return Response(content=file_stream.getvalue(), headers=headers, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

# 印鑑届出書
@app.post("/generate-excel")
def generate_excel(data: FormData):
    template_path = "template_excel_corporation_application.xlsx"
    output_path = "created_corporation_application.xlsx"

    if not os.path.exists(template_path):
        raise HTTPException(status_code=500, detail="Excel template file not found")

    wb = load_workbook(template_path)
    ws = wb.active

    # 各入力フィールドを翻訳
    translated_address = get_japanese_address(data.address)  # 住所をGoogle Mapsで変換
    translated_president_name = to_katakana(data.presidentName)  # 名前をカタカナに変換
    translated_president_address = get_japanese_address(data.presidentAddress)

    def set_merged_cell_value(ws, cell_range, value):
        """マージセルの左上セルに値をセットする関数"""
        ws.unmerge_cells(cell_range)  # マージを解除
        start_cell = cell_range.split(":")[0]  # 最初のセルを取得
        ws[start_cell] = value  # 値をセット
        ws.merge_cells(cell_range)  # 再びマージ

    # Excel のセルに値を設定
    set_merged_cell_value(ws, "AH7:BC9", data.companyName)
    set_merged_cell_value(ws, "AH10:BC13", translated_address)
    set_merged_cell_value(ws, "P52:BC52", translated_president_address)
    set_merged_cell_value(ws, "AH18:BC21", translated_president_name)
    set_merged_cell_value(ws, "P53:BC53", translated_president_name)
    set_merged_cell_value(ws, "G51:AC51", str(data.year) + "年" + str(data.month) + "月" + str(data.day) + "日")
    set_merged_cell_value(ws, "AH22:BC24", str(data.birthyear) + "年" + str(data.birthmonth) + "月" + str(data.birthday) + "日")

    #AH8:A,AH11:B,P52:C,AH19/P53:D,G51:E,AH23:G,
    # ws["AH8"] = data.companyName
    # ws["AH11"] = translated_address
    # ws["P52"] = translated_president_address
    # ws["AH19"] = translated_president_name
    # ws["P53"] = translated_president_name
    # ws["G515"] =  str(data.year) + "年" + str(data.month) + "月" + str(data.day) + "日"
    # ws["AH23"] = str(data.birthyear) + "年" + str(data.birthmonth) + "月" + str(data.birthday) + "日"

    # 生成された Excel ファイルを保存
    wb.save(output_path)

    return {"message": "Excel file successfully generated"}

@app.get("/get-created-excel")
def get_created_excel():
    file_path = "created_corporation_application.xlsx"
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Excel file not found")

    with open(file_path, "rb") as file:
        file_stream = BytesIO(file.read())

    headers = {
        "Content-Disposition": "attachment; filename=created_corporation_application.xlsx",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    }
    return Response(content=file_stream.getvalue(), headers=headers, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
