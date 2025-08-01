import streamlit as st
from pathlib import Path
import os
from RubiGUI3 import extract_terms

st.title("語句抽出＆TSV出力ツール（複数ファイル対応）")

uploaded_files = st.file_uploader("ファイルを選択してください", type=["docx"], accept_multiple_files=True)

if uploaded_files:
    temp_dir = Path("temp_files")
    temp_dir.mkdir(exist_ok=True)

    for uploaded_file in uploaded_files:
        st.subheader(f"📄 処理中: {uploaded_file.name}")

        # 一時保存
        temp_path = temp_dir / uploaded_file.name
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # 語句抽出
        override_dict = {}
        terms = extract_terms(str(temp_path), override_dict)

        # TSV生成
        tsv_path = str(temp_path).replace(".docx", ".tsv")
        with open(tsv_path, "w", encoding="cp932") as f:
            for term in terms:
                f.write(f"{term.get('word', '')}\t{term.get('reading', '')}\n")

        st.success(f"{uploaded_file.name} の語句抽出完了！")
        st.write(f"抽出語句数: {len(terms)}")

        # ダウンロードボタン
        if os.path.exists(tsv_path):
            with open(tsv_path, "rb") as f:
                st.download_button(
                    label=f"{uploaded_file.name} のTSVをダウンロード",
                    data=f.read(),
                    file_name=Path(tsv_path).name,
                    mime="text/tab-separated-values"
                )
        else:
            st.error(f"TSVファイルが見つかりませんでした: {tsv_path}")