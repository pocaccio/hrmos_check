# --- 必要なインポート ---
import streamlit as st
import pandas as pd
import os

# Google関連のインポートをtry-exceptで囲む
try:
    from google.oauth2 import service_account
    import gspread
    GOOGLE_LIBS_AVAILABLE = True
except ImportError as e:
    st.error(f"Google ライブラリが見つかりません: {e}")
    st.error("requirements.txt に以下が含まれていることを確認してください:")
    st.code("""
google-auth>=2.0.0
google-auth-oauthlib>=0.5.0
gspread>=5.0.0
    """)
    st.stop()

# --- 設定の初期化 ---
@st.cache_data
def get_config():
    """設定情報を取得"""
    config = {
        "development_mode": True,  # デフォルトは開発モード
        "sheet_url": "https://docs.google.com/spreadsheets/d/1Ymt2OrvY2dKFs9puCX8My7frS_BS1sg3Yev3BLQm9xQ/edit",
        "has_secrets": False,
        "has_gcp_account": False
    }
    
    # Streamlit Secretsの確認
    try:
        if hasattr(st, 'secrets') and st.secrets:
            config["has_secrets"] = True
            config["development_mode"] = st.secrets.get("DEVELOPMENT_MODE", True)
            
            # Google Service Accountの確認
            if "gcp_service_account" in st.secrets:
                config["has_gcp_account"] = True
    except Exception:
        pass
    
    return config

# --- 認証情報の取得 ---
@st.cache_resource
def get_credentials():
    """Google Sheets認証情報を取得"""
    config = get_config()
    
    try:
        if config["has_gcp_account"]:
            # Streamlit Secretsからサービスアカウント情報を取得
            credentials = service_account.Credentials.from_service_account_info(
                st.secrets["gcp_service_account"],
                scopes=["https://www.googleapis.com/auth/spreadsheets"]
            )
            return credentials
        else:
            # 代替手段: 環境変数やローカルファイル
            json_paths = [
                "/Users/poca/hrmos/mineral-liberty-460106-m7-a24c4c78154f.json",
                "./service_account.json",
                os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "")
            ]
            
            for path in json_paths:
                if path and os.path.exists(path):
                    credentials = service_account.Credentials.from_service_account_file(
                        path,
                        scopes=["https://www.googleapis.com/auth/spreadsheets"]
                    )
                    return credentials
            
            # どの方法でも認証情報が取得できない場合
            st.error("Google Service Account認証情報が見つかりません。")
            st.info("以下のいずれかの方法で設定してください:")
            st.code("""
1. Streamlit Secrets設定:
   [gcp_service_account]
   type = "service_account"
   project_id = "your-project-id"
   private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
   client_email = "your-service-account@project.iam.gserviceaccount.com"
   # ... 他の設定

2. 環境変数:
   GOOGLE_APPLICATION_CREDENTIALS=/path/to/service-account.json

3. ローカルファイル:
   ./service_account.json
            """)
            return None
            
    except Exception as e:
        st.error(f"認証エラー: {e}")
        return None

# --- データ読み込み ---
@st.cache_data(ttl=300)  # 5分間キャッシュ
def load_spreadsheet_data():
    """スプレッドシートからデータを読み込み"""
    credentials = get_credentials()
    if not credentials:
        return None, None
    
    try:
        gspread_client = gspread.authorize(credentials)
        config = get_config()
        spreadsheet = gspread_client.open_by_url(config["sheet_url"])
        
        # 勤怠データの読み込み
        worksheet_kintai = spreadsheet.worksheet("勤怠確認シート(打刻管理)")
        headers_kintai_raw = worksheet_kintai.row_values(1)
        
        # ヘッダー重複回避
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
        df_kintai = df_kintai[df_kintai["社員番号"].str.strip() != ""]
        
        # 社員一覧の読み込み
        worksheet_staff = spreadsheet.worksheet("社員一覧")
        df_staff = pd.DataFrame(worksheet_staff.get_all_records())
        
        return df_kintai, df_staff
        
    except Exception as e:
        st.error(f"スプレッドシート読み込みエラー: {e}")
        st.info("以下を確認してください:")
        st.info("1. スプレッドシートのURLが正しいか")
        st.info("2. サービスアカウントがスプレッドシートに共有されているか")
        st.info("3. 「勤怠確認シート(打刻管理)」シートと「社員一覧」シートが存在するか")
        return None, None

# --- 認証システム ---
def handle_authentication():
    """認証処理"""
    config = get_config()
    
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    
    if st.session_state.authenticated:
        return True
    
    # ログイン画面
    st.title("勤怠確認チェックツール")
    st.markdown("---")
    
    # 設定状況の表示
    if config["development_mode"]:
        st.info("🔧 開発モードで動作中")
    
    if not config["has_secrets"]:
        st.warning("⚠️ Streamlit Secrets が設定されていません")
    
    if not config["has_gcp_account"]:
        st.warning("⚠️ Google Service Account が設定されていません")
    
    # データ読み込みテスト
    with st.spinner("データ読み込み中..."):
        df_kintai, df_staff = load_spreadsheet_data()
    
    if df_staff is None:
        st.error("データの読み込みに失敗しました。設定を確認してください。")
        st.stop()
    
    # 権限のあるユーザーを取得
    valid_permissions = ["4. 承認者", "3. 利用者・承認者", "2. システム管理者"]
    
    if "権限" not in df_staff.columns:
        st.error("社員一覧に「権限」列が見つかりません。")
        st.info("必要な列: ログインID(B列), 社員番号(D列), 姓(E列), 名(F列), 権限(BL列)")
        st.stop()
    
    authorized_users = df_staff[df_staff["権限"].isin(valid_permissions)]
    
    if len(authorized_users) == 0:
        st.error("権限のあるユーザーが見つかりません。")
        st.info("社員一覧の権限列に以下のいずれかが設定されているユーザーが必要です:")
        for perm in valid_permissions:
            st.info(f"- {perm}")
        st.stop()
    
    # ユーザー選択
    st.markdown("### ログイン")
    
    user_options = ["選択してください"]
    user_data = {}
    
    for _, user in authorized_users.iterrows():
        surname = str(user.get('姓', '')).strip()
        given_name = str(user.get('名', '')).strip()
        name = f"{surname}{given_name}" if surname or given_name else "名前なし"
        login_id = str(user.get('ログインID', '')).strip()
        permission = str(user.get('権限', '')).strip()
        
        display_text = f"{name} ({login_id}) - {permission}"
        user_options.append(display_text)
        user_data[display_text] = user.to_dict()
    
    selected_user = st.selectbox("ログインするユーザーを選択", user_options)
    
    if selected_user != "選択してください":
        if st.button("ログイン", type="primary"):
            user_info = user_data[selected_user]
            
            # セッション状態設定
            st.session_state.authenticated = True
            st.session_state.user_info = user_info
            st.session_state.user_email = user_info.get('ログインID', '')
            surname = str(user_info.get('姓', '')).strip()
            given_name = str(user_info.get('名', '')).strip()
            st.session_state.user_name = f"{surname}{given_name}"
            
            st.success("ログインしました！")
            st.rerun()
    
    return False

# --- メインアプリケーション ---
def main_app():
    """メインアプリケーション"""
    
    # データ読み込み
    df_kintai, df_staff = load_spreadsheet_data()
    if df_kintai is None or df_staff is None:
        st.error("データの読み込みに失敗しました。")
        return
    
    # データ整形
    if "第一承認者" in df_staff.columns:
        merged = pd.merge(df_kintai, df_staff[["社員番号", "第一承認者"]], on="社員番号", how="left")
        merged = merged.rename(columns={"第一承認者": "承認者"})
    else:
        st.warning("社員一覧に「第一承認者」列が見つかりません。")
        merged = df_kintai.copy()
        merged["承認者"] = ""
    
    # 権限に基づくフィルタリング
    user_info = st.session_state.user_info
    user_permission = user_info.get("権限", "")
    
    if user_permission == "2. システム管理者":
        filtered = merged.copy()
    elif user_permission in ["4. 承認者", "3. 利用者・承認者"]:
        user_login_id = user_info.get("ログインID", "")
        filtered = merged[merged["承認者"] == user_login_id]
    else:
        filtered = merged.iloc[0:0]
    
    # UI
    st.markdown("""
    <style>
        .user-info {
            background-color: #f0f2f6; padding: 1rem;
            border-radius: 0.5rem; margin-bottom: 1rem;
        }
        .header-box {
            font-size: 20px; font-weight: bold; padding: 0.5rem;
            display: inline-block; margin-bottom: 1rem;
        }
    </style>
    """, unsafe_allow_html=True)
    
    # ヘッダー
    col1, col2 = st.columns([3, 1])
    with col1:
        st.title("勤怠確認チェックツール")
    with col2:
        if st.button("ログアウト"):
            st.session_state.authenticated = False
            st.rerun()
    
    # ユーザー情報表示
    st.markdown(f"""
    <div class='user-info'>
        <strong>ログインユーザー:</strong> {st.session_state.user_name} ({st.session_state.user_email})<br>
        <strong>権限:</strong> {user_permission}
    </div>
    """, unsafe_allow_html=True)
    
    # データ表示
    display_columns = [
        "社員番号", "名前", "休日出勤", "有休日数", "欠勤日数", "出勤時間",
        "総残業時間", "規定残業時間", "規定残業超過分", "深夜残業時間",
        "60時間超過残業", "打刻ズレ", "勤怠マイナス分"
    ]
    
    # 存在する列のみ表示
    available_columns = [col for col in display_columns if col in filtered.columns]
    
    if len(filtered) > 0:
        permission_label = "全スタッフ" if user_permission == "2. システム管理者" else "承認対象スタッフ"
        st.markdown(f"<div class='header-box'>{permission_label}: {len(filtered)}名</div>", unsafe_allow_html=True)
        
        if available_columns:
            display_df = filtered[available_columns].copy()
            
            # 数値列の変換
            for col in ["打刻ズレ", "勤怠マイナス分"]:
                if col in display_df.columns:
                    display_df[col] = pd.to_numeric(
                        display_df[col].astype(str).str.replace("", "0").replace("-", "0"), 
                        errors="coerce"
                    ).fillna(0)
            
            # ソート
            sort_cols = [col for col in ["勤怠マイナス分", "打刻ズレ"] if col in display_df.columns]
            if sort_cols:
                display_df = display_df.sort_values(by=sort_cols, ascending=True)
            
            st.dataframe(display_df, use_container_width=True)
        else:
            st.warning("表示可能な列が見つかりません。")
    else:
        if user_permission in ["4. 承認者", "3. 利用者・承認者"]:
            st.info("承認対象のスタッフがいません。第一承認者として割り当てられているスタッフのデータのみ表示されます。")
        else:
            st.info("表示可能なデータがありません。")

# --- メイン実行 ---
if __name__ == "__main__":
    if handle_authentication():
        main_app()