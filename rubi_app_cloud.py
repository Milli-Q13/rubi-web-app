
import streamlit as st
import json
import pandas as pd
from rubi_core import extract_terms
import tempfile

# 🏷️ タイトル
st.title("語句抽出＆TSV出力ツール（Streamlit Cloud対応）")

# 📘 使い方ガイド（簡易表示）
with st.expander("📘 アプリの使い方（簡易ガイド）"):
    st.markdown("""
### 🧑‍🏫 基本の流れ
1. `override.json`（語句と読み仮名の辞書）をアップロード（任意）
2. Wordファイル（.docx）をアップロード
3. 語句抽出結果を編集
4. TSVファイルをダウンロード

### 📂 フォルダ構成（推奨）

```plaintext
デスクトップ/
└── ルビ振り/
    ├── override.json
    ├── ルビデータ/
    └── 出力（ルビ付き）/
```
""")
# 📥 TSVファイルの扱い
with st.expander("📥 TSVファイルの扱い"):
    st.markdown("""
- ダウンロード後、TSVファイルを「ルビデータ」フォルダに手動で移動してください
""")

# 🧩 VBA連携案内
with st.expander("🧩 Word VBAとの連携方法"):
    st.markdown("""
このアプリで生成したTSVファイルを使って、Word文書にルビ（ふりがな）を自動挿入できます。

### 🔧 マクロの準備
- Wordを開き、対象の `.docx` ファイルを開く  
- 「開発」タブ → 「Visual Basic」からVBAエディタを開く  
- 「Normal」テンプレートの標準モジュールにマクロを貼り付け

### 🔄 実行の流れ
- マクロ名：`InsertFuriganaFromTSV_SaveToNewFile_Stable`  
- 処理後のファイルは `出力（ルビ付き）` フォルダに保存されます

👉 詳しい説明はこちら：[GitHubのREADMEを見る](https://github.com/Milli-Q13/rubi-web-app/blob/main/README.md)
""")     
 
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
        # ✅ ここで辞書を安全に参照
        override_dict = st.session_state.override_dict
        
        # ✅ 一時ファイルとして保存
        import tempfile
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(uploaded_file.getbuffer())
            tmp_path = tmp.name

        # 語句抽出（バイトデータを直接渡す）
        terms = extract_terms(tmp_path, override_dict)

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