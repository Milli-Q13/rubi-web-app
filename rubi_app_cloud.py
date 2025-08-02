import streamlit as st
import json
import pandas as pd
from rubi_core import extract_terms

# 🏷️ タイトル
st.title("語句抽出＆TSV出力ツール（Streamlit Cloud対応）")

# 初期化
if "override_dict" not in st.session_state:
    st.session_state.override_dict = {}

# 📚 辞書アップロード
uploaded_dict_file = st.file_uploader("📚 あなたの override.json をアップロード", type=["json"])
if uploaded_dict_file:
    try:
        st.session_state.override_dict = json.load(uploaded_dict_file)
        st.success(f"{len(st.session_state.override_dict)} 件の語句を読み込みました")
    except Exception as e:
        st.error(f"辞書の読み込みに失敗しました: {e}")

# ✏️ 編集UI
df_dict = pd.DataFrame([{"語句": k, "読み": v} for k, v in st.session_state.override_dict.items()])
edited_dict_df = st.data_editor(df_dict, num_rows="dynamic")

# 💾 保存（セッション内）
if st.button("辞書を更新"):
    st.session_state.override_dict = {
        row["語句"]: row["読み"]
        for _, row in edited_dict_df.iterrows()
        if row["語句"] and row["読み"]
    }
    st.success("辞書を更新しました！（セッション内）")


# 📄 Wordファイルのアップロード
uploaded_files = st.file_uploader("📄 処理対象の Word ファイル（.docx）を選択", type=["docx"], accept_multiple_files=True)

if uploaded_files:
    for uploaded_file in uploaded_files:
        st.subheader(f"📄 処理中: {uploaded_file.name}")

        # 語句抽出（バイトデータを直接渡す）
        terms = extract_terms(uploaded_file, override_dict)

        # ✏️ 抽出語句の編集
        st.subheader("📘 抽出語句の編集")
        df_terms = pd.DataFrame(terms)
        edited_terms_df = st.data_editor(df_terms, num_rows="dynamic")
        terms = edited_terms_df.to_dict(orient="records")

        # TSV生成（文字列として保持）
        tsv_content = "\n".join(f"{term.get('word', '')}\t{term.get('reading', '')}" for term in terms)
        tsv_bytes = tsv_content.encode("cp932")

        # ダウンロードボタン
        st.download_button(
            label=f"{uploaded_file.name} のTSVをダウンロード",
            data=tsv_bytes,
            file_name=uploaded_file.name.replace(".docx", ".tsv"),
            mime="text/tab-separated-values"
        )