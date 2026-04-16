import streamlit as st
import pandas as pd
from datetime import datetime
import io

# ページの設定
st.set_page_config(page_title="Blastメール 差分抽出ツール", layout="centered")

st.title("📧 Blastメール用 差分抽出ツール")
st.write("新旧ファイルを比較し、配信対象(Planner/Shopping)と配信停止(E-Mailのみ)を作成します。")

# 共通のマッピング定義（配信対象用）
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

# CSV変換用関数（配信対象用：全項目）
def process_to_csv(target_df):
    output_df = target_df[list(mapping.keys())].rename(columns=mapping)
    output_df['Other'] = ""
    return convert_df_to_csv_bytes(output_df)

# CSV変換用関数（削除リスト用：E-Mailのみ）
def process_excluded_to_csv(target_df):
    # Email列のみを抽出し、ヘッダを E-Mail に変更
    output_df = target_df[['Email']].rename(columns={'Email': 'E-Mail'})
    return convert_df_to_csv_bytes(output_df)

# エンコーディングエラー回避用の共通処理
def convert_df_to_csv_bytes(df):
    csv_buffer = io.StringIO()
    try:
        # Shift-JIS (CP932) で試行、不明な文字は '?' に置換
        df.to_csv(csv_buffer, index=False, encoding='cp932', errors='replace')
    except Exception:
        # 失敗した場合は UTF-8 (BOM付き)
        csv_buffer = io.StringIO()
        df.to_csv(csv_buffer, index=False, encoding='utf-8-sig')
    return csv_buffer.getvalue()

# --- メイン処理 ---
st.subheader("1. ファイルをアップロード")
col_files = st.columns(2)
with col_files[0]:
    current_file = st.file_uploader("今回の Member.xlsx (最新)", type=["xlsx"], key="current")
with col_files[1]:
    previous_file = st.file_uploader("前回の Member.xlsx (比較用)", type=["xlsx"], key="previous")

if current_file is not None:
    try:
        # データの読み込み
        df_curr_raw = pd.read_excel(current_file, header=1)
        
        # 抽出条件: Active 且つ Email Opt-In が Yes
        def filter_eligible(df):
            required_cols = ['Status', 'Email Opt-In', 'Type', 'MemberID', 'Email']
            for col in required_cols:
                if col not in df.columns:
                    st.error(f"ファイルに '{col}' 列が見つかりません。")
                    return pd.DataFrame()
            return df[(df['Status'] == 'Active') & (df['Email Opt-In'] == 'Yes')]

        df_curr_eligible = filter_eligible(df_curr_raw)
        
        # 配信対象の振り分け
        planner_data = df_curr_eligible[df_curr_eligible['Type'] == 'Planner']
        shopping_data = df_curr_eligible[df_curr_eligible['Type'] == 'Customer']

        date_str = datetime.now().strftime("%Y%m%d")

        # 差分比較
        excluded_data = pd.DataFrame()
        if previous_file is not None:
            df_prev_raw = pd.read_excel(previous_file, header=1)
            df_prev_eligible = filter_eligible(df_prev_raw)
            if not df_prev_eligible.empty:
                # 前回は対象だったが、今回対象外になった人を特定
                excluded_mask = df_prev_eligible['MemberID'].isin(df_curr_eligible['MemberID'])
                excluded_data = df_prev_eligible[~excluded_mask]

        st.divider()
        st.subheader("2. 処理結果")
        
        res_col1, res_col2, res_col3 = st.columns(3)
        res_col1.metric("Planner (今回)", f"{len(planner_data)}件")
        res_col2.metric("Shopping (今回)", f"{len(shopping_data)}件")
        if previous_file:
            res_col3.metric("削除対象", f"{len(excluded_data)}件", delta_color="inverse")

        st.divider()
        
        btn_col1, btn_col2 = st.columns(2)
        with btn_col1:
            st.write("▼ 配信対象リスト")
            if not planner_data.empty:
                st.download_button("Planner用CSV", process_to_csv(planner_data).encode('cp932', errors='replace'),
                                 f"BLASTMAIL_planner_{date_str}.csv", "text/csv")
            if not shopping_data.empty:
                st.download_button("Shopping用CSV", process_to_csv(shopping_data).encode('cp932', errors='replace'),
                                 f"BLASTMAIL_shopping_{date_str}.csv", "text/csv")

        with btn_col2:
            st.write("▼ 削除用リスト")
            if previous_file and not excluded_data.empty:
                st.warning("配信停止が必要なメンバーがいます")
                # 削除リスト用関数を呼び出し（E-Mailのみ）
                csv_e = process_excluded_to_csv(excluded_data)
                st.download_button("削除リスト(E-Mailのみ)を保存", csv_e.encode('cp932', errors='replace'),
                                 f"DELETE_LIST_{date_str}.csv", "text/csv")
            elif previous_file:
                st.info("新しく削除が必要な人はいません。")

    except Exception as e:
        st.error(f"エラーが発生しました: {e}")