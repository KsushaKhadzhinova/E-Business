from docx import Document
import os

base = r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА"

target_files = [
    r"ЛР1\ОТЧЕТ_ЛР1_NotaCode.docx",
    r"ЛР2-3\ОТЧЕТ_ЛР2-3_NotaCode.docx",
    r"ЛР4\ОТЧЕТ_ЛР4_NotaCode.docx",
    r"ЛР5\ОТЧЕТ_ЛР5_NotaCode.docx",
    r"ЛР6\ОТЧЕТ_ЛР6_NotaCode.docx",
    r"ЛР7\ОТЧЕТ_ЛР7_NotaCode.docx",
]

for rel in target_files:
    path = os.path.join(base, rel)
    if not os.path.exists(path):
        print(f"NOT FOUND: {rel}")
        continue
    print(f"\n=== {rel} ===")
    try:
        doc = Document(path)
        # Show first 30 non-empty paragraphs
        shown = 0
        for p in doc.paragraphs:
            txt = p.text.strip()
            if not txt:
                continue
            # Show headings and title-page-ish text
            style_name = p.style.name
            if any(x in style_name for x in ['Heading', 'Title', 'ЛР', 'Заголовок']):
                print(f"  [{style_name}] {txt[:100]}")
                shown += 1
            elif shown < 30:
                print(f"  [{style_name[:20]}] {txt[:80]}")
                shown += 1
            if shown >= 30:
                break
    except Exception as e:
        print(f"  ERROR: {e}")
