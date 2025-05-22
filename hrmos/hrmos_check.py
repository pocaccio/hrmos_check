# --- 必要なインポート ---
import streamlit as st
import pandas as pd
from google.oauth2 import service_account
import gspread
import requests
import jwt
from datetime import datetime, timedelta
import hashlib
import hmac

# --- Google OAuth設定 ---
GOOGLE_CLIENT_ID = st.secrets["GOOGLE_CLIENT_ID"]  # Streamlit Secretsから取得
GOOGLE_CLIENT_SECRET = st.secrets["GOOGLE_CLIENT_SECRET"]
REDIRECT_URI = st.secrets["REDIRECT_URI"]  # https://your-app.streamlit.app/

# --- 認証関数 ---
def get_google_auth_url():
    """Google OAuth認証URLを生成"""
    auth_url = f"https://accounts.google.com/o/oauth2/auth?client_id={GOOGLE_CLIENT_ID}&redirect_uri={REDIRECT_URI}&scope=email%20profile&response_type=code&access_type=offline"
    return auth_url

def get_google_user_info(code):
    """認証コードからユーザー情報を取得"""
    try:
        # アクセストークン取得
        token_url = "https://oauth2.googleapis.com/token"
        token_data = {
            "client_id": GOOGLE_CLIENT_ID,
            "client_secret": GOOGLE_CLIENT_SECRET,
            "code": code,
            "grant_type": "authorization_code",
            "redirect_uri": REDIRECT_URI
        }
        token_response = requests.post(token_url, data=token_data)
        token_json = token_response.json()
        
        if "access_token" not in token_json:
            return None
            
        # ユーザー情報取得
        user_info_url = f"https://www.googleapis.com/oauth2/v2/userinfo?access_token={token_json['access_token']}"
        user_response = requests.get(user_info_url)
        return user_response.json()
    except Exception as e:
        st.error(f"認証エラー: {e}")
        return None

def check_user_permission(email, df_staff):
    """ユーザーの権限チェック"""
    # B列「ログインID」と権限をチェック
    valid_permissions = ["4. 承認者", "3. 利用者・承認者", "2. システム管理者"]
    user_data = df_staff[
        (df_staff["ログインID"] == email) & 
        (df_staff["権限"].isin(valid_permissions))
    ]
    
    if len(user_data) > 0:
        user_info = user_data.iloc[0]
        return True, user_info
    else:
        return False, None

# --- セッション状態の初期化 ---
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "user_email" not in st.session_state:
    st.session_state.user_email = None
if "user_name" not in st.session_state:
    st.session_state.user_name = None
if "user_info" not in st.session_state:
    st.session_state.user_info = None

# --- 認証処理 ---
def handle_authentication():
    """認証処理のメイン関数"""
    # URLパラメータからcodeを取得
    query_params = st.experimental_get_query_params()
    
    if "code" in query_params and not st.session_state.authenticated:
        code = query_params["code"][0]
        user_info = get_google_user_info(code)
        
        if user_info and "email" in user_info:
            # 社員一覧を読み込んで権限チェック
            json_key_path = "/Users/poca/hrmos/mineral-liberty-460106-m7-a24c4c78154f.json"
            scopes = ["https://www.googleapis.com/auth/spreadsheets"]
            credentials = service_account.Credentials.from_service_account_file(json_key_path, scopes=scopes)
            gspread_client = gspread.authorize(credentials)
            
            sheet_url = "https://docs.google.com/spreadsheets/d/1Ymt2OrvY2dKFs9puCX8My7frS_BS1sg3Yev3BLQm9xQ/edit"
            spreadsheet = gspread_client.open_by_url(sheet_url)
            worksheet_staff = spreadsheet.worksheet("社員一覧")
            df_staff = pd.DataFrame(worksheet_staff.get_all_records())
            
            has_permission, staff_info = check_user_permission(user_info["email"], df_staff)
            if has_permission:
                st.session_state.authenticated = True
                st.session_state.user_email = user_info["email"]
                st.session_state.user_name = user_info.get("name", "")
                st.session_state.user_info = staff_info
                st.experimental_rerun()
            else:
                st.error("アクセス権限がありません。権限が「承認者」「利用者・承認者」「システム管理者」に設定されているメールアドレスでログインしてください。")
                st.stop()
        else:
            st.error("認証に失敗しました。")
            st.stop()

# --- ログアウト機能 ---
def logout():
    """ログアウト処理"""
    st.session_state.authenticated = False
    st.session_state.user_email = None
    st.session_state.user_name = None
    st.session_state.user_info = None
    st.experimental_rerun()

# --- メインアプリケーション ---
def main_app():
    """メインアプリケーション"""
    # --- スプレッドシート設定 ---
    json_key_path = "/Users/poca/hrmos/mineral-liberty-460106-m7-a24c4c78154f.json"
    sheet_url = "https://docs.google.com/spreadsheets/d/1Ymt2OrvY2dKFs9puCX8My7frS_BS1sg3Yev3BLQm9xQ/edit"

    # --- サービスアカウント認証 ---
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    credentials = service_account.Credentials.from_service_account_file(json_key_path, scopes=scopes)
    gspread_client = gspread.authorize(credentials)

    # --- スプレッドシート読み込み ---
    spreadsheet = gspread_client.open_by_url(sheet_url)
    worksheet_kintai = spreadsheet.worksheet("勤怠確認シート(打刻管理)")
    worksheet_staff = spreadsheet.worksheet("社員一覧")

    # 社員一覧の列構成:
    # B列: ログインID (Googleアカウント)
    # D列: 社員番号
    # E列: 姓
    # F列: 名
    # BL列: 権限 (2. システム管理者 / 3. 利用者・承認者 / 4. 承認者)

    # --- ヘッダー重複エラー回避 ---
    headers_kintai_raw = worksheet_kintai.row_values(1)
    headers_kintai = []
    seen = {}
    for col in headers_kintai_raw:
        if col in seen:
            seen[col] += 1
            headers_kintai.append(f"{col}_{seen[col]}")
        else:
            seen[col] = 0
            headers_kintai.append(col)

    records_kintai = worksheet_kintai.get_all_values()[1:]
    df_kintai = pd.DataFrame(records_kintai, columns=headers_kintai)

    # 社員一覧読み込み
    df_staff = pd.DataFrame(worksheet_staff.get_all_records())
    df_kintai = df_kintai[df_kintai["社員番号"].str.strip() != ""]

    # --- データ整形 ---
    # 社員一覧のD列「社員番号」を使用してマージ
    merged = pd.merge(df_kintai, df_staff[["社員番号", "第一承認者"]], on="社員番号", how="left")
    merged = merged.rename(columns={"第一承認者": "承認者"})

    # --- 承認者で絞り込み（認証ユーザーの権限に応じた表示） ---
    user_email = st.session_state.user_email
    user_info = st.session_state.user_info
    user_permission = user_info["権限"] if user_info is not None else None
    
    # 権限に応じてデータをフィルタリング
    if user_permission == "2. システム管理者":
        # システム管理者は全データ閲覧可能
        filtered = merged.copy()
    elif user_permission in ["4. 承認者", "3. 利用者・承認者"]:
        # 承認者・利用者承認者は第一承認者として割り当てられたスタッフのみ閲覧可能
        user_login_id = user_info["ログインID"] if user_info is not None else user_email
        filtered = merged[merged["承認者"] == user_login_id]
    else:
        # その他は空のデータ
        filtered = merged.iloc[0:0]

    # --- 表示項目の整理 ---
    display_columns = [
        "社員番号", "名前", "休日出勤", "有休日数",
        "欠勤日数", "出勤時間", "総残業時間",
        "規定残業時間", "規定残業超過分",
        "深夜残業時間", "60時間超過残業",
        "打刻ズレ", "勤怠マイナス分"
    ]
    display_columns = [col for col in display_columns if col in filtered.columns]

    # 数値列をfloat変換
    for col in ["打刻ズレ", "勤怠マイナス分"]:
        if col in filtered.columns:
            filtered[col] = pd.to_numeric(filtered[col].replace("", "0").replace("-", "0"), errors="coerce").fillna(0)

    # ソート
    filtered = filtered.sort_values(by=["勤怠マイナス分", "打刻ズレ"], ascending=[True, True])

    # --- UI表示 ---
    st.markdown("""
    <style>
        .st-emotion-cache-1w723zb{
            max-width:90%;
        }
        .st-am{
            width:40%;
        }
        .block-container {
            padding-top: 2rem;
        }
        .header-box {
            font-size: 20px;
            font-weight: bold;
            padding: 0.5rem;
            border: 0px solid #ccc;
            display: inline-block;
            margin-bottom: 1rem;
        }
        .user-info {
            background-color: #f0f2f6;
            padding: 1rem;
            border-radius: 0.5rem;
            margin-bottom: 1rem;
        }
    </style>
    """, unsafe_allow_html=True)

    # ヘッダー部分
    col1, col2 = st.columns([3, 1])
    with col1:
        st.title("勤怠確認チェックツール")
    with col2:
        if st.button("ログアウト"):
            logout()

    # ユーザー情報表示
    st.markdown(f"""
    <div class='user-info'>
        <strong>ログインユーザー:</strong> {st.session_state.user_name} ({st.session_state.user_email})
    </div>
    """, unsafe_allow_html=True)

    # データ表示
    if len(filtered) > 0:
        permission_label = "全スタッフ" if user_permission == "2. システム管理者" else "承認対象スタッフ"
        st.markdown(f"<div class='header-box'>{permission_label}: {len(filtered)}名</div>", unsafe_allow_html=True)
        display_df = filtered[display_columns].copy()
        st.dataframe(display_df.style.set_properties(**{'text-align': 'center'}))
    else:
        if user_permission in ["4. 承認者", "3. 利用者・承認者"]:
            st.info("承認対象のスタッフがいません。第一承認者として割り当てられているスタッフのデータのみ表示されます。")
        else:
            st.info("表示可能なデータがありません。")

# --- メイン実行部分 ---
if __name__ == "__main__":
    if not st.session_state.authenticated:
        handle_authentication()
        
        # 認証されていない場合のログイン画面
        st.title("勤怠確認チェックツール")
        st.markdown("---")
        
        st.markdown("### ログイン")
        st.markdown("このアプリケーションを使用するには、適切な権限が設定されているGoogleアカウントでログインしてください。")
        
        auth_url = get_google_auth_url()
        st.markdown(f"[Googleアカウントでログイン]({auth_url})")
        
        st.markdown("---")
        st.markdown("**権限について:**")
        st.markdown("- **システム管理者**: 全スタッフの勤怠データを閲覧可能")
        st.markdown("- **承認者・利用者承認者**: 第一承認者として割り当てられたスタッフのみ閲覧可能")
        st.markdown("")
        st.markdown("**注意**: 社員一覧のログインID列にGoogleアカウントが登録され、適切な権限が設定されている必要があります。")
        
    else:
        main_app()