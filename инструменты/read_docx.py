from docx import Document
import sys

path = r"E:\ИИТ\ЭлектронныйБизнес\СДАЧА\ЛР1\ОТЧЕТ_ЛР1_NotaCode.docx"
doc = Document(path)

# Print first 80 paragraphs to understand structure
for i, p in enumerate(doc.paragraphs[:80]):
    if p.text.strip():
        print(f"[{i:3d}] style={p.style.name!r:30s} | {p.text[:120]}")
