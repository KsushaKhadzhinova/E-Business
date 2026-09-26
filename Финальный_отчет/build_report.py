"""Сборка финального отчёта по ЛР1–7 (Word, СТП БГУИР).

Запуск: python Финальный_отчет/build_report.py
Использует общий сборщик _tools/build_lab_docx.py: склеивает 00_Введение.md,
ОТЧЕТ.md всех ЛР, 99_Заключение.md и 98_Список_источников.md и сохраняет
ОТЧЕТ_по_лабораторным_работам_1-7_NotaCode.docx.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "_tools"))

import build_lab_docx

if __name__ == "__main__":
    build_lab_docx.build_final()
