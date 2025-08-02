from sudachipy import tokenizer, dictionary
import jaconv
from sudachipy import dictionary

def extract_terms(file_path, override_dict):
    import zipfile
    import xml.etree.ElementTree as ET
    from sudachipy import tokenizer, dictionary
    import jaconv

    tokenizer_obj = dictionary.Dictionary(dict_type="full").create()
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