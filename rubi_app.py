import streamlit as st
from pathlib import Path
import os
import json
import pandas as pd
from rubi_core import extract_terms
import subprocess

# 📁 ベースディレクトリ（このファイルと同じ階層）
base_dir = Path(__file__).parent.resolve()

# 📁 必要なフォルダの初期化
data_dir = base_dir / "ルビデータ"
output_dir = base_dir / "出力（ルビ付き）"
data_dir.mkdir(exist_ok=True)
output_dir.mkdir(exist_ok=True)

# 📄 override.json の初期化
override_path = base_dir / "override.json"
if not override_path.exists():
    with open(override_path, "w", encoding="utf-8") as f:
        json.dump({}, f, ensure_ascii=False, indent=2)
    st.info("📄 override.json を新規作成しました（空の辞書）")

# 📄 override.json の読み込み
try:
    with open(override_path, "r", encoding="utf-8") as f:
        override_dict = json.load(f)
except Exception as e:
    override_dict = {}
    st.error(f"override.json の読み込みに失敗しました: {e}")

# 🏷️ タイトル
st.title("語句抽出＆TSV出力ツール（ルビ振りフォルダ対応）")

# 📦 SudachiPy 辞書リンク（初回のみ必要）
subprocess.run(["sudachipy", "link", "-t", "full"])

# 📚 辞書のアップロード（任意）
uploaded_dict_file = st.file_uploader("📚 ルビ辞書（override.json）をアップロード", type=["json"])
if uploaded_dict_file:
    try:
        override_dict = json.load(uploaded_dict_file)
        st.success(f"{len(override_dict)} 件の語句をアップロード辞書から読み込みました")
    except Exception as e:
        st.error(f"アップロード辞書の読み込みに失敗しました: {e}")
else:
    st.info("既存の override.json を使用します")

# ✏️ 辞書編集UI
st.subheader("📝 辞書の編集")
df_dict = pd.DataFrame([{"語句": k, "読み": v} for k, v in override_dict.items()])
edited_dict_df = st.data_editor(df_dict, num_rows="dynamic")

# 💾 辞書保存ボタン
if st.button("辞書を保存（override.json に上書き）"):
    try:
        new_dict = {row["語句"]: row["読み"] for _, row in edited_dict_df.iterrows() if row["語句"]}
        with open(override_path, "w", encoding="utf-8") as f:
            json.dump(new_dict, f, ensure_ascii=False, indent=2)
        override_dict = new_dict
        st.success("辞書を保存しました！")
    except Exception as e:
        st.error(f"保存に失敗しました: {e}")

# 📄 Wordファイルのアップロード
uploaded_files = st.file_uploader("📄 処理対象の Word ファイル（.docx）を選択", type=["docx"], accept_multiple_files=True)

if uploaded_files:
    for uploaded_file in uploaded_files:
        st.subheader(f"📄 処理中: {uploaded_file.name}")

        # 一時保存（ルビデータフォルダに保存）
        temp_path = data_dir / uploaded_file.name
        with open(temp_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # 語句抽出
        terms = extract_terms(str(temp_path), override_dict)

        # ✏️ 抽出語句の編集
        st.subheader("📘 抽出語句の編集")
        df_terms = pd.DataFrame(terms)
        edited_terms_df = st.data_editor(df_terms, num_rows="dynamic")
        terms = edited_terms_df.to_dict(orient="records")

        # TSV保存（ルビデータフォルダに保存）
        tsv_path = data_dir / uploaded_file.name.replace(".docx", ".tsv")
        try:
            with open(tsv_path, "w", encoding="cp932") as f:
                for term in terms:
                    f.write(f"{term.get('word', '')}\t{term.get('reading', '')}\n")
        except Exception as e:
            st.error(f"TSVファイルの保存に失敗しました: {e}")

        # ダウンロードボタン
        if tsv_path.exists():
            with open(tsv_path, "rb") as f:
                st.download_button(
                    label=f"{uploaded_file.name} のTSVをダウンロード",
                    data=f.read(),
                    file_name=tsv_path.name,
                    mime="text/tab-separated-values"
                )
        else:
            st.error(f"TSVファイルが見つかりませんでした: {tsv_path}")