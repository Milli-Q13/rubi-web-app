import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
import jaconv
import json
import os
from pathlib import Path
from sudachipy import tokenizer, dictionary

# Sudachi初期化
tokenizer_obj = dictionary.Dictionary(dict_type="core").create()
mode = tokenizer.Tokenizer.SplitMode.C

# override辞書の読み込み
def load_override_dict():
    if os.path.exists("override.json"):
        with open("override.json", "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

# override辞書の保存
def save_override_dict(override_dict):
    with open("override.json", "w", encoding="utf-8") as f:
        json.dump(override_dict, f, ensure_ascii=False, indent=2)

# 語句抽出関数
def extract_terms(file_path, override_dict):
    with zipfile.ZipFile(file_path, "r") as docx:
        with docx.open("word/document.xml") as file:
            tree = ET.parse(file)
            root = tree.getroot()
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            texts = [node.text for node in root.findall(".//w:t", ns) if node.text]

    full_text = "".join(texts)
    words = {}

    for m in tokenizer_obj.tokenize(full_text, mode):
        surface = m.surface()
        if len(surface) <= 1 or surface in words:
            continue
        if all('\u3040' <= ch <= '\u309F' for ch in surface):
            continue
        if surface in override_dict:
            reading = override_dict[surface]
        else:
            reading = jaconv.kata2hira(m.reading_form())
        if surface == reading:
            continue
        words[surface] = reading

    return [{"word": w, "reading": r} for w, r in words.items()]

# TSV保存
def save_tsv(terms, file_name):
    base_name = Path(file_name).stem
    save_dir = Path("ルビデータ")
    save_dir.mkdir(exist_ok=True)
    tsv_path = save_dir / f"{base_name}（ルビ）.tsv"

    with open(tsv_path, "w", encoding="cp932") as f:
        for term in terms:
            f.write(f"{term['word']}\t{term['reading']}\n")
    return tsv_path

# Streamlit UI
st.title("📘 ルビ編集ツール（Streamlit版）")

override_dict = load_override_dict()

uploaded_files = st.file_uploader("Wordファイルをアップロード（複数可）", type=["docx"], accept_multiple_files=True)

if uploaded_files:
    for i, uploaded_file in enumerate(uploaded_files):
        st.markdown(f"### 📄 {uploaded_file.name}")

        file_bytes = uploaded_file.read()
        temp_path = Path(f"temp_{i}.docx")  # 一時ファイル名をユニークに

        with open(temp_path, "wb") as f:
            f.write(file_bytes)

        terms = extract_terms(temp_path, override_dict)
        st.success(f"{len(terms)} 件の語句を抽出しました")

        edited_terms = st.data_editor(
            terms,
            column_config={
                "word": "語句",
                "reading": "読み"
            },
            num_rows="dynamic",
            key=f"editor_{i}"  # 複数エディタにユニークキーを
        )


    if st.button("📄 TSV保存"):
        tsv_path = save_tsv(edited_terms, uploaded_file.name)
        st.info(f"保存完了：{tsv_path}")

    st.markdown("---")
    st.subheader("📚 辞書編集（override.json）")

    dict_items = [{"word": k, "reading": v} for k, v in override_dict.items()]
    edited_dict = st.data_editor(dict_items, num_rows="dynamic")

    if st.button("💾 辞書保存"):
        new_dict = {item["word"]: item["reading"] for item in edited_dict if item["word"] and item["reading"]}
        save_override_dict(new_dict)
        st.success("辞書を保存しました")


