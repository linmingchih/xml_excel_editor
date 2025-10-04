from pathlib import Path

text = Path(r"D:\OneDrive - ANSYS, Inc\a-client-repositories\lmz_xml_excel_editor_202510\xml_excel_editor\stackup_editor.html").read_text(encoding="utf-8")
pattern = "      } else if (collection === \"layers\") {"
start = text.index(pattern)
end = text.index(pattern, start + len(pattern))
print(text[start:end])
