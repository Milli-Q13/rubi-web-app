## rubi-web-app
## 📝 語句抽出＆TSV出力ツール（Streamlit Cloud対応）

## 📌 概要
このWebアプリは、Word（.docx）ファイルから語句を抽出し、読み仮名付きのTSVファイルとして出力するツールです。
独自辞書（override.json）をアップロード・編集することで、語句の読みを自由にカスタマイズできます。
教材作成や語彙管理など、教育現場での活用を想定しています。

## 🌐 アクセス方法
以下のURLからアプリにアクセスできます：
[rubi-web-app on Streamlit Cloud](https://rubi-web-app-njraef8ijfsvyqk6qxqybr.streamlit.app)




## 🧑‍🏫 Streamlitアプリの使い方
1. 📚 辞書ファイル（override.json）のアップロード（任意）
語句と読み仮名の辞書ファイル（JSON形式）をアップロードします。
形式例：
{
  "名古屋": "なごや",
  "教育": "きょういく"
}


アップロード後、読み込まれた語句数が表示されます。

2. ✏️ 辞書の編集
表形式で語句と読み仮名を編集できます。
「辞書を更新」ボタンでセッション内に保存されます。

3. 📄 Wordファイル（.docx）のアップロード
複数ファイルの同時アップロードが可能です。
ファイルごとに語句抽出が行われ、編集画面が表示されます。

4. 📥 TSVファイルのダウンロード
- 語句と読み仮名をタブ区切りで出力
- ファイル名は元のWordファイル名に基づいて自動生成されます
- 文字コードは cp932（Windows環境向け）です
- ダウンロード後、TSVファイルを「ルビデータ」フォルダに手動で移動してください（後述）

## 🔍 読み仮名の生成について
- 読み仮名は、アップロードされた辞書（override.json）を優先して使用します
- 辞書に未登録の語句については、SudachiPy によって自動生成されます
- 自動生成された読み仮名はカタカナで取得され、jaconv.kata2hira() によりひらがなに変換されます
- 誤変換の可能性があるため、抽出後の編集画面で確認・修正してください

## 📁 ファイル形式について
以下の表は、使用されるファイル種別とその内容を示します：

| ファイル種別 | 拡張子 | 内容 | 
|:---:|:---:|:---:|
| 辞書ファイル | .json | 語句と読み仮名の対応表 | 
| 処理対象ファイル | .docx | 語句抽出対象のWord文書 | 
| 出力ファイル | .tsv | 語句と読み仮名のタブ区切りテキスト | 

## 🗂️ 作業フォルダ構成（Streamlit → VBA連携）

TSVファイルを使ってWord VBAでルビ振りを行う場合、以下のようなフォルダ構成を推奨します：
```
デスクトップ/
└── ルビ振り/
    ├── override.json              # 読み仮名辞書（Streamlitアプリで編集・保存）
    ├── ルビデータ/                # Streamlitアプリで出力されたTSVファイルを保存
    └── 出力（ルビ付き）/         # Word VBAで処理されたルビ付きWordファイルの保存先
```
🔄 Streamlit → VBA の連携フロー

```
    ├── override.json              # 読み仮名辞書（Streamlitアプリで編集・保存）
    ├── ルビデータ/                # Streamlitアプリで出力されたTSVファイルを保存
    └── 出力（ルビ付き）/         # Word VBAで処理されたルビ付きWordファイルの保存先
```

## 🔄 Streamlit → VBA の連携フロー
 ce35cff (変更内容の説明)
- Streamlitアプリで .docx ファイルをアップロードし、TSVファイルを生成
- ダウンロードしたTSVファイルを「ルビデータ」フォルダに手動で移動
- Wordで元の .docx ファイルを開く
- VBAマクロ InsertFuriganaFromTSV_SaveToNewFile_Stable を実行
- 語句にルビが挿入された新しいWordファイルが 出力（ルビ付き） フォルダに保存される

## 🧩 Word VBAによるルビ挿入処理
このツールで生成されたTSVファイルをもとに、Word VBAマクロを使って文書にルビ（ふりがな）を自動挿入できます。

## 🔧 マクロの準備
- Wordを開き、対象の .docx ファイルを開く
- 「開発」タブ → 「Visual Basic」からVBAエディタを開く
- 「Normal」テンプレートの標準モジュール（例：Module1）に以下のコードを貼り付け

## 💻 VBAマクロコード
```vba
Sub InsertFuriganaFromTSV_SaveToNewFile_Stable()
    Dim docOriginal As Document, docNew As Document
    Dim docName As String, nameOnly As String, extOnly As String
    Dim basePath As String, tsvPath As String, savePath As String
    Dim fso As Object, FileNum As Integer
    Dim LineData As String, WordParts() As String
    Dim TargetWord As String, Furigana As String
    Dim rng As Range
    Set docOriginal = ActiveDocument
    docName = docOriginal.Name
    nameOnly = Left(docName, InStrRev(docName, ".") - 1)
    extOnly = Mid(docName, InStrRev(docName, "."))
    Set fso = CreateObject("Scripting.FileSystemObject")
    basePath = fso.GetParentFolderName(docOriginal.Path)
    Set docNew = Documents.Add
    docNew.Content.FormattedText = docOriginal.Content.FormattedText
    tsvPath = basePath & "\ルビデータ\" & nameOnly & ".tsv"
    If Dir(tsvPath) = "" Then
        MsgBox "TSVファイルが見つかりません：" & vbCrLf & tsvPath, vbCritical
        Exit Sub
    End If
    If Not fso.FolderExists(basePath & "\出力（ルビ付き）") Then
        fso.CreateFolder basePath & "\出力（ルビ付き）"
    End If
    savePath = basePath & "\出力（ルビ付き）\" & nameOnly & "（ルビ）" & extOnly
    FileNum = FreeFile
    Open tsvPath For Input As FileNum
    Do Until EOF(FileNum)
        Line Input #FileNum, LineData
        WordParts = Split(LineData, vbTab)
        If UBound(WordParts) = 1 Then
            TargetWord = WordParts(0)
            Furigana = WordParts(1)
            Set rng = docNew.Range(0, 0)
            With rng.Find
                .Text = TargetWord
                .Forward = True
                .Wrap = wdFindStop
                .MatchWholeWord = False
                .MatchCase = False
            End With
            Do While rng.Find.Execute
                rng.PhoneticGuide Text:=Furigana, Alignment:=wdPhoneticGuideAlignmentCenter, _
                    Raise:=12, FontSize:=6, FontName:="MS Mincho"
                Set rng = docNew.Range(rng.End, docNew.Content.End)
                DoEvents
            Loop
        End If
    Loop
    Close FileNum
    docNew.SaveAs2 FileName:=savePath, FileFormat:=wdFormatXMLDocument
End Sub
```
- Word文書を開いた状態で、マクロ InsertFuriganaFromTSV_SaveToNewFile_Stable を実行
- 語句にルビが挿入された新しいWordファイルが 出力（ルビ付き） フォルダに保存されます

## 📂 ファイルの命名規則

```
| 種類 | ファイル名例 | 保存場所 | 
| 元文書 | 教材.docx | 任意（Wordで開く） | 
| TSVファイル | 教材.tsv | ルビ振り/ルビデータ/ | 
| 出力文書 | 教材（ルビ）.docx | ルビ振り/出力（ルビ付き）/ | 
```

## ⚠️ 注意点
- TSVファイルは必ず「ルビデータ」フォルダに手動で移動してください
- 語句が文書内に複数回登場する場合、すべてにルビが挿入されます
- 誤変換がある場合は、TSVを修正して再実行してください
- マクロは「Normal」テンプレートに登録することで、すべての文書で利用可能になります

## 💡 よくある用途
- 教材作成時の語句リスト抽出
- 読み仮名付き語彙管理
- 生徒向けの語彙確認シート作成
- Word文書へのルビ挿入（VBA連携）

## 📄 ライセンス（必要に応じて記載してください）
※ ご自身の利用方針に応じて、MITやCC BYなどのライセンスを記載してください。
