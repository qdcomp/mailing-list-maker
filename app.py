import streamlit as st
import pandas as pd
from datetime import datetime
import io

# ページの設定
st.set_page_config(page_title="Blastメール 抽出ツール", layout="centered")

# --- 1. ログイン機能 ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if st.session_state["password_correct"]:
        return True

    st.title("🔒 ログインが必要です")
    password = st.text_input("パスワードを入力してください", type="password")
    
    if st.button("ログイン"):
        if password == "RimanJP2026!":
            st.session_state["password_correct"] = True
            if hasattr(st, "rerun"):
                st.rerun()
            else:
                st.experimental_rerun()
        else:
            st.error("パスワードが正しくありません")
    return False

if check_password():
    st.title("📧 Blastメール 抽出ツール")
    st.write(f"Logged in: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    # マッピング定義
    mapping = {
        'MemberID': 'MemberID', 'Sponsor #': 'ReferrerMainFK', 'First Name': 'fname',
        'Last Name': 'lname', 'Gender': 'Gender', 'Type': 'MemberType',
        'Status': 'MemberStatus', 'Company Name': 'company', 'City': 'city',
        'State': 'state', 'Zip': 'zip', 'Email': 'E-Mail'
    }

    def convert_df_to_csv_bytes(df, is_excluded=False):
        if is_excluded:
            output_df = df[['Email']].rename(columns={'Email': 'E-Mail'})
        else:
            # マッピングに含まれる列のみを抽出し、名称を変更
            output_df = df[list(mapping.keys())].rename(columns=mapping)
            output_df['Other'] = ""
            output_df['Email Opt-In'] = df['Email Opt-In']
        
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
            df_curr_raw = pd.read_excel(current_file, header=1, engine='openpyxl')
            
            # --- ロジック適用 ---
            # 配信対象 (Active & Yes)
            df_curr_eligible = df_curr_raw[(df_curr_raw['Status'] == 'Active') & (df_curr_raw['Email Opt-In'] == 'Yes')]
            planner_data = df_curr_eligible[df_curr_eligible['Type'] == 'Planner']
            shopping_data = df_curr_eligible[df_curr_eligible['Type'] == 'Customer']
            
            # Opt-Outリスト (Active & No)
            df_curr_optout = df_curr_raw[(df_curr_raw['Status'] == 'Active') & (df_curr_raw['Email Opt-In'] == 'No')]
            
            date_str = datetime.now().strftime("%Y%m%d")

            # 削除リスト
            excluded_data = pd.DataFrame()
            if previous_file is not None:
                df_prev_raw = pd.read_excel(previous_file, header=1, engine='openpyxl')
                df_prev_eligible = df_prev_raw[(df_prev_raw['Status'] == 'Active') & (df_prev_raw['Email Opt-In'] == 'Yes')]
                excluded_data = df_prev_eligible[~df_prev_eligible['MemberID'].isin(df_curr_eligible['MemberID'])]

            st.markdown("---")
            st.subheader("2. 抽出結果")
            res_col1, res_col2, res_col3, res_col4, res_col5 = st.columns(5)
            res_col1.metric("Planner", f"{len(planner_data)}件")
            res_col2.metric("Shopping", f"{len(shopping_data)}件")
            res_col3.metric("Opt-Out", f"{len(df_curr_optout)}件")
            res_col4.metric("全件(無加工)", f"{len(df_curr_raw)}件")
            if previous_file:
                res_col5.metric("削除対象", f"{len(excluded_data)}件")

            st.markdown("---")
            # レイアウト調整：ボタンを4列に変更
            btn_col1, btn_col2, btn_col3, btn_col4 = st.columns(4)
            
            with btn_col1:
                st.write("▼ 配信対象 (Yes)")
                if not planner_data.empty:
                    st.download_button("Planner用保存", convert_df_to_csv_bytes(planner_data).encode('cp932', errors='replace'), f"BLAST_planner_{date_str}.csv")
                if not shopping_data.empty:
                    st.download_button("Shopping用保存", convert_df_to_csv_bytes(shopping_data).encode('cp932', errors='replace'), f"BLAST_shopping_{date_str}.csv")

            with btn_col2:
                st.write("▼ Opt-Out (No)")
                if not df_curr_optout.empty:
                    st.download_button("Opt-Out保存", convert_df_to_csv_bytes(df_curr_optout).encode('cp932', errors='replace'), f"ACTIVE_OPTOUT_{date_str}.csv")
                
                st.write("▼ 全データ")
                # 【新機能】フィルタリングなしの全データ出力
                st.download_button("全データ保存", convert_df_to_csv_bytes(df_curr_raw).encode('cp932', errors='replace'), f"ALL_MAPPED_DATA_{date_str}.csv")

            with btn_col3:
                st.write("▼ 削除リスト")
                if previous_file and not excluded_data.empty:
                    st.warning(f"{len(excluded_data)}名停止")
                    st.download_button("削除CSV保存", convert_df_to_csv_bytes(excluded_data, is_excluded=True).encode('cp932', errors='replace'), f"DELETE_LIST_{date_str}.csv")

        except Exception as e:
            st.error(f"エラーが発生しました: {e}")

    if st.sidebar.button("ログアウト"):
        st.session_state["password_correct"] = False
        if hasattr(st, "rerun"):
            st.rerun()
        else:
            st.experimental_rerun()