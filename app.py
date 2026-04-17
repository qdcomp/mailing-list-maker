import streamlit as st
import pandas as pd
from datetime import datetime
import io

# ページの設定
st.set_page_config(page_title="Blastメール 差分抽出ツール", layout="centered")

# --- 1. ログイン機能の設定 ---
def check_password():
    """パスワードが正しいか確認し、結果をセッションに保存する"""
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    # すでにログイン済みならTrueを返す
    if st.session_state["password_correct"]:
        return True

    # ログイン画面の表示
    st.title("🔒 ログインが必要です")
    password = st.text_input("パスワードを入力してください", type="password")
    
    if st.button("ログイン"):
        if password == "RimanJP2026!":
            st.session_state["password_correct"] = True
            # 古いStreamlitバージョン対策
            if hasattr(st, "rerun"):
                st.rerun()
            else:
                st.experimental_rerun()
        else:
            st.error("パスワードが正しくありません")
    return False

# ログインチェックを実行
if check_password():

    # --- 2. メインツール処理 ---
    st.title("📧 Blastメール用 差分抽出ツール")
    st.write(f"Logged in: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    # 共通のマッピング定義
    mapping = {
        'MemberID': 'MemberID',
        'Sponsor #': 'ReferrerMainFK',
        'First Name': 'fname',
        'Last Name': 'lname',
        'Gender': 'Gender',
        'Type': 'MemberType',
        'Status': 'MemberStatus',
        'Company Name': 'company',
        'City': 'city',
        'State': 'state',
        'Zip': 'zip',
        'Email': 'E-Mail'
    }

    def convert_df_to_csv_bytes(df, is_excluded=False):
        if is_excluded:
            output_df = df[['Email']].rename(columns={'Email': 'E-Mail'})
        else:
            output_df = df[list(mapping.keys())].rename(columns=mapping)
            output_df['Other'] = ""
        
        csv_buffer = io.StringIO()
        try:
            output_df.to_csv(csv_buffer, index=False, encoding='cp932', errors='replace')
        except Exception:
            csv_buffer = io.StringIO()
            output_df.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
        return csv_buffer.getvalue()

    st.subheader("1. ファイルをアップロード")
    col_files = st.columns(2)
    with col_files[0]:
        current_file = st.file_uploader("今回の Member.xlsx (最新)", type=["xlsx"], key="current")
    with col_files[1]:
        previous_file = st.file_uploader("前回の Member.xlsx (比較用)", type=["xlsx"], key="previous")

    if current_file is not None:
        try:
            # サーバー環境に合わせて engine='openpyxl' を明示
            df_curr_raw = pd.read_excel(current_file, header=1, engine='openpyxl')
            
            def filter_eligible(df):
                required_cols = ['Status', 'Email Opt-In', 'Type', 'MemberID', 'Email']
                if not all(c in df.columns for c in required_cols):
                    st.error("必要な列が見つかりません")
                    return pd.DataFrame()
                return df[(df['Status'] == 'Active') & (df['Email Opt-In'] == 'Yes')]

            df_curr_eligible = filter_eligible(df_curr_raw)
            planner_data = df_curr_eligible[df_curr_eligible['Type'] == 'Planner']
            shopping_data = df_curr_eligible[df_curr_eligible['Type'] == 'Customer']
            date_str = datetime.now().strftime("%Y%m%d")

            excluded_data = pd.DataFrame()
            if previous_file is not None:
                df_prev_raw = pd.read_excel(previous_file, header=1, engine='openpyxl')
                df_prev_eligible = filter_eligible(df_prev_raw)
                if not df_prev_eligible.empty:
                    excluded_mask = df_prev_eligible['MemberID'].isin(df_curr_eligible['MemberID'])
                    excluded_data = df_prev_eligible[~excluded_mask]

            st.divider()
            st.subheader("2. 抽出結果")
            res_col1, res_col2, res_col3 = st.columns(3)
            res_col1.metric("Planner", f"{len(planner_data)}件")
            res_col2.metric("Shopping", f"{len(shopping_data)}件")
            if previous_file:
                res_col3.metric("削除対象", f"{len(excluded_data)}件")

            st.divider()
            btn_col1, btn_col2 = st.columns(2)
            with btn_col1:
                st.write("▼ 配信対象CSV")
                if not planner_data.empty:
                    st.download_button("Planner用保存", convert_df_to_csv_bytes(planner_data).encode('cp932', errors='replace'), f"BLASTMAIL_planner_{date_str}.csv")
                if not shopping_data.empty:
                    st.download_button("Shopping用保存", convert_df_to_csv_bytes(shopping_data).encode('cp932', errors='replace'), f"BLASTMAIL_shopping_{date_str}.csv")

            with btn_col2:
                st.write("▼ 配信停止リスト")
                if previous_file and not excluded_data.empty:
                    st.warning("配信停止が必要な人がいます")
                    st.download_button("削除用リスト保存", convert_df_to_csv_bytes(excluded_data, is_excluded=True).encode('cp932', errors='replace'), f"DELETE_LIST_{date_str}.csv")

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")

    # サイドバーにログアウトボタン
    if st.sidebar.button("ログアウト"):
        st.session_state["password_correct"] = False
        if hasattr(st, "rerun"):
            st.rerun()
        else:
            st.experimental_rerun()