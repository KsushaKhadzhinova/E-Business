# -*- coding: utf-8 -*-
"""ЛР6. Рендер HTML-мокапов экранов NotaCode в PNG для раздела 6.4 (страницы и экраны).

Источник мокапов: E:\\VisualDSL-Platform\\design\\mockups\\*.dc.html (прототип прежней версии продукта,
старое имя DiagramCode). При рендере во временной копии имя бренда заменяется на NotaCode —
это тот же продукт, переименованный (см. E:\\NotaCode\\README.md).

Рендер выполняется локально через Microsoft Edge в headless-режиме (--screenshot).
Запуск:  python lr6_mockups.py
Результат: ../Рисунки/mock_XX_<имя>.png
"""
import os
import re
import shutil
import subprocess
import tempfile

SRC = r"E:\VisualDSL-Platform\design\mockups"
HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.normpath(os.path.join(HERE, "..", "Рисунки"))
EDGE_CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
]

# (файл мокапа, имя PNG, интерфейсная единица реестра 6.4)
MOCKUPS = [
    ("Main.dc.html", "mock_01_main_ide", "U-010 Рабочая область IDE (общий вид)"),
    ("IDE-Light.dc.html", "mock_02_ide_light", "U-010 IDE, светлая тема"),
    ("IDE-Dark.dc.html", "mock_03_ide_dark", "U-010 IDE, тёмная тема"),
    ("Login.dc.html", "mock_04_login", "U-008 Вход"),
    ("Register.dc.html", "mock_05_register", "U-009 Регистрация"),
    ("Validation.dc.html", "mock_06_validation_problems", "U-013 Problems / состояние ошибки"),
    ("AIModalText.dc.html", "mock_07_ai_modal", "U-014 AI-ассистент"),
    ("AIResultBPMN.dc.html", "mock_08_ai_result", "U-014 AI-ассистент, результат"),
    ("SaveDropdown.dc.html", "mock_09_save_dropdown", "U-015 Меню сохранения"),
    ("VersionHistory.dc.html", "mock_10_version_history", "U-017 История версий"),
    ("DiffView.dc.html", "mock_11_diff", "U-018 Diff версий"),
    ("DocumentationModal.dc.html", "mock_12_documentation", "U-020 Документация в IDE"),
    ("SidebarDetail.dc.html", "mock_13_sidebar_templates", "U-012 Панель шаблонов"),
    ("PanelResize.dc.html", "mock_14_panel_resize", "U-010 Изменение размеров панелей"),
    ("TabletSplit.dc.html", "mock_15_tablet", "U-010 Планшетная раскладка"),
    ("AIImageUpload.dc.html", "mock_16_ai_image_upload", "R4: изображение -> DSL (вне MVP)"),
    ("Moodboard.dc.html", "mock_17_moodboard", "Дизайн-токены (мудборд)"),
]


def find_browser():
    for p in EDGE_CANDIDATES:
        if os.path.exists(p):
            return p
    return None


def window_size(html):
    m = re.search(r"width:(\d+)px;height:(\d+)px", html)
    return (int(m.group(1)), int(m.group(2))) if m else (1440, 900)


def main():
    browser = find_browser()
    if not browser:
        print("Браузер Edge/Chrome не найден — мокапы не отрендерены.")
        return
    os.makedirs(OUT, exist_ok=True)
    tmp = tempfile.mkdtemp(prefix="nc_mock_")
    done = []
    for fname, png, unit in MOCKUPS:
        src = os.path.join(SRC, fname)
        if not os.path.exists(src):
            print("нет файла", src)
            continue
        with open(src, encoding="utf-8") as fh:
            html = fh.read().replace("DiagramCode", "NotaCode")
        w, h = window_size(html)
        tmp_html = os.path.join(tmp, fname)
        with open(tmp_html, "w", encoding="utf-8") as fh:
            fh.write(html)
        out_png = os.path.join(OUT, png + ".png")
        cmd = [browser, "--headless=new", "--disable-gpu", "--hide-scrollbars",
               f"--window-size={w},{h}", f"--screenshot={out_png}",
               "file:///" + tmp_html.replace("\\", "/")]
        subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=90)
        ok = os.path.exists(out_png)
        done.append((png, unit, ok))
        print(("OK  " if ok else "ERR ") + png + ".png  ->  " + unit)
    shutil.rmtree(tmp, ignore_errors=True)
    print(f"Готово: {sum(1 for d in done if d[2])} из {len(done)}")


if __name__ == "__main__":
    main()
