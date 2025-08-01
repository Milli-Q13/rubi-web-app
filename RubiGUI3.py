from tkinterdnd2 import DND_FILES, TkinterDnD
import tkinter as tk
from tkinter import ttk, filedialog, messagebox as msgbox
import json
import os
import zipfile
import xml.etree.ElementTree as ET
import jaconv
from sudachipy import tokenizer, dictionary
import win32com.client
import glob
from pathlib import Path
import tkinter.filedialog as filedialog


# ✅ Sudachi初期化
tokenizer_obj = dictionary.Dictionary().create()
mode = tokenizer.Tokenizer.SplitMode.C

def to_hiragana(katakana):
    return jaconv.kata2hira(katakana)

class RubyEditorApp:
    def center_main_window(self, width=800, height=600):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def center_window_auto(self, win):
        win.update_idletasks()
        width = win.winfo_reqwidth()
        height = win.winfo_reqheight()
        x = (win.winfo_screenwidth() // 2) - (width // 2)
        y = (win.winfo_screenheight() // 2) - (height // 2)
        win.geometry(f"+{x}+{y}")

    def __init__(self, root):
        self.root = root
        self.root.title("ルビ編集ツール")
        self.root.geometry("800x600")
        self.center_main_window(800,600)
        self.data = []
        self.override_dict = {}
        self.file_path = None
        self.docx_files = []
        self.current_index = 0
        self.setup_ui()
        self.load_override_dict()
        self.target_folder = "元データ"

    def batch_process(self):
        docx_files = [f for f in glob.glob(f"{self.target_folder}/*.docx") if not Path(f).name.startswith("~$")]
        if not docx_files:
            msgbox.showinfo("一括処理", f"{self.target_folder} に .docx ファイルが見つかりません。")
            return

        for file_path in docx_files:
            try:
                print(f"処理中: {file_path}")
                terms = extract_terms(file_path, self.override_dict)
                print(f"抽出語句数: {len(terms)}")
                save_tsv(terms, file_path)
                print("TSV保存完了")
                run_vba_macro(file_path)
                print("VBA実行完了")
                save_ruby_word(file_path)
                print("ルビ付きWord保存完了")
            except Exception as e:
                print(f"エラー: {file_path} → {e}")

        msgbox.showinfo("一括処理", f"{len(docx_files)} 件のファイルを処理しました。")
    def start_batch_review(self):
        self.docx_files = [f for f in glob.glob(f"{self.target_folder}/*.docx") if not Path(f).name.startswith("~$")]
        self.current_index = 0
        if not self.docx_files:
            msgbox.showinfo("一括処理", f"{self.target_folder} に .docx ファイルが見つかりません。")
            return
        self.process_next_file()
    def process_next_file(self):
        if self.current_index >= len(self.docx_files):
            msgbox.showinfo("完了", "すべてのファイルを処理しました")
            return

        file_path = self.docx_files[self.current_index]
        self.file_path = file_path
        terms = extract_terms(file_path, self.override_dict)
        self.data = [(t["word"], t["reading"]) for t in terms]
        self.tree.delete(*self.tree.get_children())
        for word, reading in self.data:
            self.tree.insert("", "end", values=(word, reading))
        msgbox.showinfo("確認", f"{Path(file_path).name} の語句を確認・修正してください")
    def confirm_and_continue(self):
        if not self.file_path or not self.data:
            msgbox.showwarning("エラー", "処理対象がありません")
            return

        terms = [{"word": w, "reading": r} for w, r in self.data]
        save_tsv(terms, self.file_path)
        run_vba_macro(self.file_path)
        save_ruby_word(self.file_path)
        self.current_index += 1
        self.process_next_file()
    def select_folder(self):
        folder = filedialog.askdirectory(title="処理対象フォルダを選択")
        if folder:
            self.target_folder = folder
            msgbox.showinfo("フォルダ選択", f"選択されたフォルダ:\n{folder}")
    def setup_ui(self):
        self.drop_label = tk.Label(self.root, text="ここにWordファイルをドロップ\nまたはクリックして選択", relief="ridge", height=4)
        self.drop_label.pack(fill="x", padx=10, pady=10)
        self.drop_label.bind("<Button-1>", self.select_file)
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind("<<Drop>>", self.on_drop)

        self.tree = ttk.Treeview(self.root, columns=("word", "reading"), show="headings")
        self.tree.heading("word", text="語句")
        self.tree.heading("reading", text="読み")
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)
        self.tree.bind("<Double-1>", self.edit_item)

        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)
    
        tk.Button(btn_frame, text="TSV保存", command=self.save_tsv).pack(side="left", padx=5)
        tk.Button(btn_frame, text="辞書編集", command=self.edit_override_dict).pack(side="left", padx=5)
        tk.Button(btn_frame, text="ルビ再適用", command=self.apply_ruby).pack(side="left", padx=5)
        tk.Button(btn_frame, text="ルビ付きWord出力", command=self.run_vba_macro).pack(side="left", padx=5)
        btn_batch_process = tk.Button(btn_frame, text="📁 一括処理", command=self.batch_process)
        btn_batch_process.pack(side="left", padx=5)
        tk.Button(btn_frame, text="一括処理開始（確認あり）", command=self.start_batch_review).pack(side="left", padx=5)
        tk.Button(btn_frame, text="▶ 次へ", command=self.confirm_and_continue).pack(side="left", padx=5)
        tk.Button(btn_frame, text="フォルダ選択", command=self.select_folder).pack(side="left", padx=5)

    def select_file(self, event=None):
        path = filedialog.askopenfilename(filetypes=[("Wordファイル", "*.docx")])
        if path:
            self.file_path = path
            self.extract_words(path)

    def on_drop(self, event):
        path = event.data.strip("{}")
        if path.lower().endswith(".docx"):
            self.file_path = path
            self.extract_words(path)
        else:
            messagebox.showwarning("形式エラー", "Wordファイル（.docx）をドロップしてください")

    def extract_words(self, path):
        self.data.clear()
        self.tree.delete(*self.tree.get_children())

        with zipfile.ZipFile(path, "r") as docx:
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
            if surface in self.override_dict:
                reading = self.override_dict[surface]
            else:
                reading = to_hiragana(m.reading_form())
            if surface == reading:
                continue
            words[surface] = reading

        for word, reading in sorted(words.items()):
            self.data.append((word, reading))
            self.tree.insert("", "end", values=(word, reading))

    def edit_item(self, event):
        item_id = self.tree.focus()
        if not item_id:
            return
        word, reading = self.tree.item(item_id, "values")

        edit_win = tk.Toplevel(self.root)
        edit_win.title("編集")
        edit_win.geometry("300x200")
        tk.Label(edit_win, text="語句").pack()
        word_entry = tk.Entry(edit_win)
        word_entry.insert(0, word)
        word_entry.pack()

        tk.Label(edit_win, text="読み").pack()
        reading_entry = tk.Entry(edit_win)
        reading_entry.insert(0, reading)
        reading_entry.pack()
    
        def save_edit():
            new_word = word_entry.get()
            new_reading = reading_entry.get()
            self.tree.item(item_id, values=(new_word, new_reading))
            for i, (w, r) in enumerate(self.data):
                if w == word and r == reading:
                    self.data[i] = (new_word, new_reading)
                    break
            edit_win.destroy()

        tk.Button(edit_win, text="保存", command=save_edit).pack(pady=5)
        self.center_window_auto(edit_win)  # ← ここで中央配置！

    def center_window_auto(self, win):  # ← edit_item と同じインデントでOK！
        win.update_idletasks()
        width = win.winfo_reqwidth()
        height = win.winfo_reqheight()
        x = (win.winfo_screenwidth() // 2) - (width // 2)
        y = (win.winfo_screenheight() // 2) - (height // 2)
        win.geometry(f"+{x}+{y}")

    def save_tsv(self):
        if not self.file_path:
            messagebox.showwarning("エラー", "Wordファイルが未選択です")
            return

        base_name = os.path.splitext(os.path.basename(self.file_path))[0]

    # プロジェクトフォルダを基準に保存先を構築
        project_dir = os.path.abspath(os.path.join(os.path.dirname(self.file_path), ".."))
        save_dir = os.path.join(project_dir, "ルビデータ")
        os.makedirs(save_dir, exist_ok=True)

        tsv_path = os.path.join(save_dir, f"{base_name}（ルビ）.tsv")

        with open(tsv_path, "w", encoding="cp932") as f:
            for word, reading in self.data:
                f.write(f"{word}\t{reading}\n")

        messagebox.showinfo("保存完了", f"TSVファイルを保存しました：\n{tsv_path}")

    def load_override_dict(self):
        if os.path.exists("override.json"):
            with open("override.json", "r", encoding="utf-8") as f:
                self.override_dict = json.load(f)

    def edit_override_dict(self):
        edit_win = tk.Toplevel(self.root)
        edit_win.title("辞書編集")
        edit_win.geometry("400x400")
        self.center_window_auto(edit_win)  # ← ここで中央配置！

        tree = ttk.Treeview(edit_win, columns=("word", "reading"), show="headings")
        tree.heading("word", text="語句")
        tree.heading("reading", text="読み")
        tree.pack(fill="both", expand=True)

        for word, reading in self.override_dict.items():
            tree.insert("", "end", values=(word, reading))

        def edit_item(event):
            item_id = tree.focus()
            if not item_id:
                return
            word, reading = tree.item(item_id, "values")

            popup = tk.Toplevel(edit_win)
            popup.title("編集")
            popup.geometry("300x200")
            self.center_window_auto(popup)  # ← ここで中央配置！
            tk.Label(popup, text="語句").pack()
            word_entry = tk.Entry(popup)
            word_entry.insert(0, word)
            word_entry.pack()

            tk.Label(popup, text="読み").pack()
            reading_entry = tk.Entry(popup)
            reading_entry.insert(0, reading)
            reading_entry.pack()

            def save():
                new_word = word_entry.get()
                new_reading = reading_entry.get()
                tree.item(item_id, values=(new_word, new_reading))
                self.override_dict.pop(word, None)
                self.override_dict[new_word] = new_reading
                popup.destroy()

            tk.Button(popup, text="保存", command=save).pack(pady=5)

        tree.bind("<Double-1>", edit_item)

        def save_dict():
            with open("override.json", "w", encoding="utf-8") as f:
                json.dump(self.override_dict, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("保存完了", "辞書を保存しました")
            edit_win.destroy()

        tk.Button(edit_win, text="保存", command=save_dict).pack(pady=5)

    def apply_ruby(self):
        if not self.data:
            messagebox.showwarning("エラー", "語句一覧が空です")
            return

        updated = 0
        for i, (word, reading) in enumerate(self.data):
            if word in self.override_dict:
                new_reading = self.override_dict[word]
                self.data[i] = (word, new_reading)
                updated += 1

        self.tree.delete(*self.tree.get_children())
        for word, reading in self.data:
            self.tree.delete(*self.tree.get_children())
        for word, reading in self.data:
            self.tree.insert("", "end", values=(word, reading))

        messagebox.showinfo("再適用完了", f"辞書の読みを {updated} 件適用しました")
    def run_vba_macro(self):
        if not self.file_path:
            messagebox.showwarning("エラー", "Wordファイルが未選択です")
            return

        macro_name = "InsertFuriganaFromTSV_SaveToNewFile_Stable"
        try:
            import win32com.client
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False

            doc = word.Documents.Open(self.file_path)
            word.Run(macro_name)
            doc.Close(False)
            word.Quit()

            # 出力先をプロジェクトフォルダ基準で構築
            project_dir = os.path.abspath(os.path.join(os.path.dirname(self.file_path), ".."))
            name = os.path.splitext(os.path.basename(self.file_path))[0]
            ext = os.path.splitext(self.file_path)[1]
            out_path = os.path.join(project_dir, "出力（ルビ付き）", f"{name}（ルビ）{ext}")

            messagebox.showinfo("完了", f"ルビ付きWordを出力しました：\n{out_path}")
        except Exception as e:
            messagebox.showerror("実行エラー", f"VBAマクロの実行に失敗しました：\n{e}")

def extract_terms(file_path, override_dict):
    import zipfile
    import xml.etree.ElementTree as ET
    from sudachipy import tokenizer, dictionary
    import jaconv

    tokenizer_obj = dictionary.Dictionary().create()
    mode = tokenizer.Tokenizer.SplitMode.C

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
def extract_terms(file_path, override_dict):
    import zipfile
    import xml.etree.ElementTree as ET
    from sudachipy import tokenizer, dictionary
    import jaconv

    tokenizer_obj = dictionary.Dictionary().create()
    mode = tokenizer.Tokenizer.SplitMode.C

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
def save_tsv(terms, file_path):
    base_name = Path(file_path).stem
    project_dir = Path(file_path).resolve().parent.parent
    save_dir = project_dir / "ルビデータ"
    save_dir.mkdir(exist_ok=True)
    tsv_path = save_dir / f"{base_name}（ルビ）.tsv"

    with open(tsv_path, "w", encoding="cp932") as f:
        for term in terms:
            f.write(f"{term['word']}\t{term['reading']}\n")

def save_ruby_word(file_path):
    base_name = Path(file_path).stem
    ext = Path(file_path).suffix
    project_dir = Path(file_path).resolve().parent.parent
    output_dir = project_dir / "出力（ルビ付き）"
    output_dir.mkdir(exist_ok=True)
    ruby_path = Path(file_path).with_name(f"{base_name}（ルビ）{ext}")
    final_path = output_dir / ruby_path.name

    if ruby_path.exists():
        ruby_path.replace(final_path)            
def run_vba_macro(file_path):
    macro_name = "InsertFuriganaFromTSV_SaveToNewFile_Stable"  # ← マクロ名は必要に応じて変更
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(Path(file_path).resolve()))
        word.Run(macro_name)
        doc.Close(False)
        word.Quit()
    except Exception as e:
        print(f"VBA実行エラー: {e}")
def start_batch_review(self):
    self.docx_files = [f for f in glob.glob("元データ/*.docx") if not Path(f).name.startswith("~$")]
    self.current_index = 0
    if not self.docx_files:
        msgbox.showinfo("一括処理", "元データフォルダにファイルがありません。")
        return
    self.process_next_file()
def process_next_file(self):
    if self.current_index >= len(self.docx_files):
        msgbox.showinfo("完了", "すべてのファイルを処理しました")
        return

    file_path = self.docx_files[self.current_index]
    self.file_path = file_path
    terms = extract_terms(file_path, self.override_dict)
    self.data = [(t["word"], t["reading"]) for t in terms]
    self.tree.delete(*self.tree.get_children())
    for word, reading in self.data:
        self.tree.insert("", "end", values=(word, reading))
    msgbox.showinfo("確認", f"{Path(file_path).name} の語句を確認・修正してください")
def confirm_and_continue(self):
    if not self.file_path or not self.data:
        msgbox.showwarning("エラー", "処理対象がありません")
        return

    terms = [{"word": w, "reading": r} for w, r in self.data]
    save_tsv(terms, self.file_path)
    run_vba_macro(self.file_path)
    save_ruby_word(self.file_path)
    self.current_index += 1
    self.process_next_file()
if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = RubyEditorApp(root)
    root.mainloop()


