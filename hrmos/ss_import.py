# --- 必要なインポート ---
import streamlit as st
import pandas as pd
import numpy as np
import re
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io
import gspread

# --- 設定 ---
json_key_path = "/Users/poca/hrmos/mineral-liberty-460106-m7-a24c4c78154f.json"
drive_folder_id = "1tQGYGjOmWR0MBWJ6NpSQ9etg3rxTzvuY"
target_filename = "kintai_2025-04.csv"
sheet_url = "https://docs.google.com/spreadsheets/d/1Ymt2OrvY2dKFs9puCX8My7frS_BS1sg3Yev3BLQm9xQ/edit"

# --- サービスアカウント認証 ---
scopes = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets"
]
credentials = service_account.Credentials.from_service_account_file(json_key_path, scopes=scopes)

drive_service = build("drive", "v3", credentials=credentials)
gspread_client = gspread.authorize(credentials)

# --- DriveからファイルIDを取得 ---
results = drive_service.files().list(
    q=f"name='{target_filename}' and '{drive_folder_id}' in parents and trashed=false",
    spaces="drive",
    fields="files(id, name)"
).execute()
items = results.get("files", [])

if not items:
    st.error(f"❌ ファイルが見つかりません: {target_filename}")
else:
    file_id = items[0]["id"]
    st.success(f"✅ ファイルを検出: {target_filename}")

    # --- DriveからCSVをダウンロード ---
    request = drive_service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()

    fh.seek(0)
    try:
        df = pd.read_csv(fh, encoding="cp932")
    except UnicodeDecodeError:
        fh.seek(0)
        df = pd.read_csv(fh, encoding="shift_jis")

    df = df.fillna("")
    st.dataframe(df)

    # --- 時間列の定義 ---
    time_columns = [
        '所定内勤務時間',
        '所定時間外勤務時間',
        '所定外休日勤務時間',
        '法定外休日勤務時間',
        '法定休日勤務時間',
        '深夜勤務時間',
        '勤務時間',
        '実勤務時間',
        '確定_有給なし_残業時間'
    ]

    # --- 時間フォーマット変換処理 ---
    def preprocess_value(val):
        if pd.isna(val):
            return ""
        if isinstance(val, str):
            val = val.replace("'", "")
            time_pattern = re.compile(r'^(\d{1,3}):(\d{2})')
            match = time_pattern.match(val)
            if match:
                hours, minutes = match.groups()
                return f"{hours}:{minutes}:00"
            return val
        if isinstance(val, (int, float)):
            return val
        return str(val)

    # --- スプレッドシートへ書き込み処理 ---
    try:
        spreadsheet = gspread_client.open_by_url(sheet_url)
        worksheet = spreadsheet.worksheet("貼り付け用")

        # データ前処理
        processed_data = []
        for row in df.values.tolist():
            processed_row = [preprocess_value(val) for val in row]
            processed_data.append(processed_row)
        processed_headers = [preprocess_value(col) for col in df.columns.values.tolist()]

        # シートをクリアしてデータ書き込み
        worksheet.clear()
        worksheet.update([processed_headers] + processed_data, value_input_option='USER_ENTERED')

        # 時間列の書式を整える
        for col_name in time_columns:
            if col_name in df.columns:
                col_index = df.columns.get_loc(col_name)
                col_letter = chr(65 + col_index)  # A〜Z対応（列数が多い場合は gspread.utils.toA1推奨）
                time_range = f'{col_letter}2:{col_letter}{len(processed_data) + 1}'
                worksheet.format(time_range, {
                    "numberFormat": {
                        "type": "TIME",
                        "pattern": "[h]:mm:ss"
                    }
                })

        st.success("✅ '貼り付け用' シートへ更新しました")

    except Exception as e:
        st.error(f"❌ スプレッドシートの更新中にエラーが発生しました: {str(e)}")
