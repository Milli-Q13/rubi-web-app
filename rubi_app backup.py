import streamlit as st
from pathlib import Path
import os
import json
import pandas as pd
from rubi_core import extract_terms
import subprocess

st.title("語句抽出＆TSV出力ツール（複数ファイル対応）")
subprocess.run(["sudachipy", "link", "-t", "full"])

# 📚 辞書の読み込み（アップロード or 既存ファイル）
override_dict = {}
default_dict_path = Path(__file__).parent / "override.json"
uploaded_dict_file = st.file_uploader("📚 ルビ辞書（override.json）をアップロード", type=["json"])

if uploaded_dict_file:
    try:
        override_dict = json.load(uploaded_dict_file)
        st.success(f"{len(override_dict)} 件の語句をアップロード辞書から読み込みました")
    except Exception as e:
        st.error(f"アップロード辞書の読み込みに失敗しました: {e}")
elif default_dict_path.exists():
    try:
        with open(default_dict_path, "r", encoding="utf-8") as f:
            override_dict = json.load(f)
        st.info(f"既存の辞書（{default_dict_path}）を読み込みました")
    except Exception as e:
        st.error(f"既存辞書の読み込みに失敗しました: {e}")

# ✏️ 辞書編集UI（読み込み後に表示）
st.subheader("📝 辞書の編集")

# override_dict（辞書）を DataFrame に変換
df = pd.DataFrame([{"語句": k, "読み": v} for k, v in override_dict.items()])

# 編集可能なテーブルとして表示
edited_df = st.data_editor(df, num_rows="dynamic")


# 💾 保存ボタン
if st.button("辞書を保存（override.json に上書き）"):
    try:
        new_dict = {row["語句"]: row["読み"] for _, row in edited_df.iterrows() if row["語句"]}
        with open(default_dict_path, "w", encoding="utf-8") as f:
            json.dump(new_dict, f, ensure_ascii=False, indent=2)
        override_dict = new_dict
        st.success("辞書を保存しました！")
    except Exception as e:
        st.error(f"保存に失敗しました: {e}")

# 📄 ファイルのアップロード
uploaded_files = st.file_uploader("📄 処理対象の Word ファイル（.docx）を選択", type=["docx"], accept_multiple_files=True)

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
        terms = extract_terms(str(temp_path), override_dict)

        # ✏️ 編集UIを追加
        st.write("✏️ 抽出語句の編集")
        df_terms = pd.DataFrame(terms)
        edited_terms = st.data_editor(df_terms, num_rows="dynamic")

        # 編集結果を terms に反映（以降の処理に使う）
        terms = edited_terms.to_dict(orient="records")

        # TSV生成（編集後の terms を使用）
        tsv_path = str(temp_path).replace(".docx", ".tsv")
        with open(tsv_path, "w", encoding="cp932") as f:
            for term in terms:
                f.write(f"{term.get('word', '')}\t{term.get('reading', '')}\n")

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
