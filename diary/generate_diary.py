#!/usr/bin/env python3
"""
Генератор щоденника розробки дистрибутива Devuan Security OS
Форматування відповідно до ДСТУ 3008:2015
Запуск: PYTHONPATH=~/.nix-profile/lib/python3.13/site-packages python3 generate_diary.py
"""

import os
import sys
import glob

nix_site_packages = os.path.expanduser("~/.nix-profile/lib/python3.13/site-packages")
if nix_site_packages not in sys.path:
    sys.path.insert(0, nix_site_packages)

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import datetime

# === КОНФІГУРАЦІЯ ===
SCREENSHOTS_DIR = "/home/zoozienix/projects/Devuan_Project/screenshots"
OUTPUT_FILE = "/home/zoozienix/projects/Devuan_Project/diary/Devuan_Security_OS_Diary.docx"

# ДСТУ 3008:2015 — параметри сторінки
PAGE_LEFT   = Cm(2.5)
PAGE_RIGHT  = Cm(1.5)
PAGE_TOP    = Cm(2.0)
PAGE_BOTTOM = Cm(2.0)

# Шрифт
FONT_MAIN  = "Times New Roman"
FONT_SIZE  = Pt(14)
FONT_CODE  = "Courier New"
INDENT     = Cm(1.25)  # відступ першого рядка абзацу

# ============================================================
# ДОПОМІЖНІ ФУНКЦІЇ
# ============================================================

def apply_dstu_paragraph(paragraph, first_line_indent=True, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    """Застосовує базове ДСТУ-форматування до абзацу"""
    pf = paragraph.paragraph_format
    pf.alignment = align
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    if first_line_indent:
        pf.first_line_indent = INDENT
    else:
        pf.first_line_indent = Pt(0)

def apply_dstu_run(run, bold=False, italic=False, size=None):
    """Застосовує ДСТУ-форматування до run"""
    run.font.name = FONT_MAIN
    run.font.size = size or FONT_SIZE
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor(0, 0, 0)

def add_dstu_paragraph(doc, text="", bold=False, italic=False,
                        first_line_indent=True, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
    """Додає абзац у стилі ДСТУ"""
    p = doc.add_paragraph()
    apply_dstu_paragraph(p, first_line_indent, align)
    if text:
        run = p.add_run(text)
        apply_dstu_run(run, bold=bold, italic=italic)
    return p

def add_dstu_heading(doc, text, level=1):
    """
    Заголовки за ДСТУ 3008:2015:
      level=1 → розділ (14pt, жирний, по центру, велика літера)
      level=2 → підрозділ (14pt, жирний, з відступом)
      level=3 → пункт (14pt, жирний курсив, з відступом)
    """
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.first_line_indent = Pt(0)

    run = p.add_run(text)
    run.font.name = FONT_MAIN
    run.font.size = FONT_SIZE
    run.font.color.rgb = RGBColor(0, 0, 0)

    if level == 1:
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run.bold = True
        run.font.all_caps = True
    elif level == 2:
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.first_line_indent = INDENT
        run.bold = True
    elif level == 3:
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.first_line_indent = INDENT
        run.bold = True
        run.italic = True
    return p

def add_screenshot(doc, filepath, fig_num, caption_text, width_cm=14):
    """Вставляє рисунок з підписом у стилі ДСТУ"""
    add_dstu_paragraph(doc, first_line_indent=False)  # відступ перед рисунком

    # Пошук файлу (на випадок пробілів у назві)
    actual_path = filepath
    if not os.path.exists(filepath):
        matches = glob.glob(filepath.rstrip() + "*")
        if matches:
            actual_path = matches[0]

    if os.path.exists(actual_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(0)
        p_img.paragraph_format.space_after = Pt(0)
        run = p_img.add_run()
        run.add_picture(actual_path, width=Cm(width_cm))
    else:
        p_err = add_dstu_paragraph(
            doc, f"[РИСУНОК ВІДСУТНІЙ: {os.path.basename(filepath)}]",
            align=WD_ALIGN_PARAGRAPH.CENTER, first_line_indent=False
        )
        p_err.runs[0].font.color.rgb = RGBColor(0xFF, 0x00, 0x00)

    # Підпис рисунку: "Рисунок N — Назва" (по центру, без відступу)
    p_cap = doc.add_paragraph()
    p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cap.paragraph_format.space_before = Pt(3)
    p_cap.paragraph_format.space_after = Pt(6)
    p_cap.paragraph_format.first_line_indent = Pt(0)
    p_cap.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    run = p_cap.add_run(f"Рисунок {fig_num} \u2014 {caption_text}")
    run.font.name = FONT_MAIN
    run.font.size = FONT_SIZE
    run.font.color.rgb = RGBColor(0, 0, 0)

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def add_dstu_table(doc, headers, rows):
    """Додає таблицю у стилі ДСТУ"""
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = 'Table Grid'

    # Заголовки
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        set_cell_bg(hdr_cells[i], "C0C0C0")
        hdr_cells[i].text = ""
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        run = p.add_run(h)
        run.font.name = FONT_MAIN
        run.font.size = FONT_SIZE
        run.font.bold = True

    # Рядки
    for row_data in rows:
        cells = table.add_row().cells
        for i, val in enumerate(row_data):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
            p.paragraph_format.first_line_indent = Pt(0)
            run = p.add_run(val)
            run.font.name = FONT_MAIN
            run.font.size = FONT_SIZE
            if i == 0:
                run.font.bold = True

    return table

def add_page_number(doc):
    """Додає нумерацію сторінок знизу по центру"""
    section = doc.sections[0]
    footer = section.footer
    footer_para = footer.paragraphs[0]
    footer_para.clear()
    footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer_para.add_run()
    run.font.name = FONT_MAIN
    run.font.size = FONT_SIZE

    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.text = 'PAGE'
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)


# ============================================================
# ГЕНЕРАЦІЯ ДОКУМЕНТУ
# ============================================================

def generate():
    doc = Document()

    # --- Параметри сторінки (ДСТУ 3008:2015) ---
    section = doc.sections[0]
    section.page_width  = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin   = PAGE_LEFT
    section.right_margin  = PAGE_RIGHT
    section.top_margin    = PAGE_TOP
    section.bottom_margin = PAGE_BOTTOM

    # Нумерація сторінок
    add_page_number(doc)

    # Очищаємо стиль Normal
    normal = doc.styles['Normal']
    normal.font.name = FONT_MAIN
    normal.font.size = FONT_SIZE
    normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

    # ==========================================
    # ТИТУЛЬНА СТОРІНКА
    # ==========================================
    for _ in range(3):
        add_dstu_paragraph(doc, first_line_indent=False)

    # Назва організації
    add_dstu_paragraph(
        doc, "Щоденник розробки дипломного проєкту",
        align=WD_ALIGN_PARAGRAPH.CENTER, first_line_indent=False, bold=True
    )
    add_dstu_paragraph(doc, first_line_indent=False)

    # Назва проєкту
    p_title = add_dstu_paragraph(
        doc, "DEVUAN SECURITY OS",
        align=WD_ALIGN_PARAGRAPH.CENTER, first_line_indent=False, bold=True
    )
    p_title.runs[0].font.size = Pt(16)

    add_dstu_paragraph(doc, first_line_indent=False)
    add_dstu_paragraph(
        doc,
        "Кастомний дистрибутив GNU/Linux на базі Devuan Excalibur 6.x,\n"
        "орієнтований на безпеку, без системи ініціалізації systemd",
        align=WD_ALIGN_PARAGRAPH.CENTER, first_line_indent=False, italic=True
    )

    for _ in range(8):
        add_dstu_paragraph(doc, first_line_indent=False)

    meta = [
        ("Автор:", "zoozienix"),
        ("Репозиторій:", "github.com/zoozieuniver/Devuan_Project"),
        ("Дата початку:", datetime.date.today().strftime("%d.%m.%Y")),
        ("Версія документу:", "0.1"),
    ]
    for label, value in meta:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        p.paragraph_format.first_line_indent = Pt(0)
        r1 = p.add_run(f"{label} ")
        apply_dstu_run(r1, bold=True)
        r2 = p.add_run(value)
        apply_dstu_run(r2)

    doc.add_page_break()

    # ==========================================
    # ЗМІСТ
    # ==========================================
    add_dstu_heading(doc, "ЗМІСТ", level=1)
    add_dstu_paragraph(doc, first_line_indent=False)

    toc_items = [
        ("Вступ", "3"),
        ("1 Технічний стек та обґрунтування вибору", "3"),
        ("2 Фаза 1. Підготовка середовища (Build Box)", "4"),
        ("  2.1 Завантаження ISO-образу Devuan", "4"),
        ("  2.2 Створення віртуальної машини у VMware", "5"),
        ("3 Фаза 2. Встановлення Devuan без графічного середовища", "\u2014"),
        ("4 Фаза 3. Конфігурування live-build", "\u2014"),
        ("5 Фаза 4. Збірка ISO-образу", "\u2014"),
        ("6 Фаза 5. Тестування та оптимізація", "\u2014"),
    ]
    for item, page in toc_items:
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        p.paragraph_format.first_line_indent = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(item)
        apply_dstu_run(r, bold=item[0].isdigit() and "  " not in item)
        dots = "." * max(2, 65 - len(item))
        r2 = p.add_run(f" {dots} {page}")
        apply_dstu_run(r2)

    doc.add_page_break()

    # ==========================================
    # ВСТУП
    # ==========================================
    add_dstu_heading(doc, "ВСТУП", level=1)
    add_dstu_paragraph(doc, first_line_indent=False)

    add_dstu_paragraph(
        doc,
        "Даний щоденник фіксує процес розробки власного дистрибутиву GNU/Linux на базі "
        "Devuan Excalibur 6.x. Devuan є форком дистрибутиву Debian, головною особливістю "
        "якого є відмова від системи ініціалізації systemd на користь класичного SysVinit."
    )
    add_dstu_paragraph(
        doc,
        "Метою проєкту є створення мінімальної, безпечної та відтворюваної (reproducible) "
        "системи, яка може використовуватися як основа для спеціалізованих середовищ безпеки. "
        "Кінцевим результатом є завантажувальний ISO-образ, зібраний інструментом live-build, "
        "з повністю задокументованим та версіонованим процесом збірки через систему контролю "
        "версій Git."
    )
    add_dstu_paragraph(
        doc,
        "Кожен крок розробки фіксується у вигляді Git-комітів із відповідними скріншотами, "
        "що забезпечує повну відтворюваність проєкту та слугує технічним щоденником."
    )

    doc.add_page_break()

    # ==========================================
    # РОЗДІЛ 1: ТЕХНІЧНИЙ СТЕК
    # ==========================================
    add_dstu_heading(doc, "1 ТЕХНІЧНИЙ СТЕК ТА ОБҐРУНТУВАННЯ ВИБОРУ", level=1)
    add_dstu_paragraph(doc, first_line_indent=False)

    add_dstu_paragraph(
        doc,
        "У таблиці 1.1 наведено перелік технологій та інструментів, обраних для реалізації "
        "проєкту, із зазначенням обґрунтування кожного вибору."
    )
    add_dstu_paragraph(doc, "Таблиця 1.1 \u2014 Технічний стек проєкту",
                       align=WD_ALIGN_PARAGRAPH.LEFT, first_line_indent=False)

    add_dstu_table(doc,
        ["Компонент", "Вибір", "Обґрунтування"],
        [
            ("База дистрибутиву", "Devuan Excalibur 6.1.1", "Debian без systemd, стабільна база"),
            ("Ядро", "Linux (стабільний реліз)", "Широка підтримка обладнання"),
            ("Init система", "SysVinit", "Простота, передбачуваність, без залежностей"),
            ("Робочий стіл", "XFCE 4", "Легковаговість, повна функціональність"),
            ("Інструмент збірки", "live-build", "Декларативна конфігурація, стандарт Debian"),
            ("Гіпервізор", "VMware Workstation Pro 25H2", "Build Box та тестове середовище"),
            ("Версіонування", "Git + GitHub", "Відтворюваність збірки, журнал змін"),
        ]
    )

    doc.add_page_break()

    # ==========================================
    # РОЗДІЛ 2: ФАЗА 1
    # ==========================================
    add_dstu_heading(doc, "2 ФАЗА 1. ПІДГОТОВКА СЕРЕДОВИЩА (BUILD BOX)", level=1)
    add_dstu_paragraph(doc, first_line_indent=False)

    add_dstu_paragraph(
        doc,
        "Метою даної фази є підготовка еталонного середовища збірки (Build Box) у вигляді "
        "віртуальної машини під управлінням VMware Workstation Pro 25H2. Саме в цьому "
        "середовищі надалі виконуватиметься збірка дистрибутиву за допомогою live-build."
    )

    # 2.1
    add_dstu_heading(doc, "2.1 Завантаження ISO-образу Devuan", level=2)
    add_dstu_paragraph(doc, first_line_indent=False)

    add_dstu_paragraph(
        doc,
        "Для встановлення Build Box обрано образ типу netinstall дистрибутиву "
        "Devuan Excalibur 6.1.1 (amd64). Даний тип образу є мінімальним встановлювачем, "
        "що завантажує лише базову систему, а всі пакети отримує з мережі під час "
        "встановлення. Це забезпечує актуальність встановлених пакетів."
    )
    add_dstu_paragraph(
        doc,
        "Образ завантажено з офіційного архіву Devuan за адресою files.devuan.org. "
        "Версія Excalibur 6.x базується на Debian 13 (Trixie) та є поточним стабільним "
        "релізом станом на липень 2026 року. Розмір образу складає 594 МБ."
    )
    add_dstu_paragraph(doc, first_line_indent=False)

    add_screenshot(
        doc,
        os.path.join(SCREENSHOTS_DIR, "1.1_devuan_release_archive.png"),
        "2.1", "Офіційний архів ISO-образів Devuan Excalibur на files.devuan.org"
    )

    # 2.2
    add_dstu_heading(doc, "2.2 Створення віртуальної машини у VMware Workstation", level=2)
    add_dstu_paragraph(doc, first_line_indent=False)

    add_dstu_paragraph(
        doc,
        "Для Build Box створено нову віртуальну машину у VMware Workstation Pro 25H2. "
        "Параметри підібрані з урахуванням вимог процесу збірки через live-build, "
        "що потребує достатньо оперативної пам'яті та обчислювальних ресурсів для "
        "завантаження та розпакування пакетів."
    )

    add_dstu_paragraph(doc, "У таблиці 2.1 наведено параметри створеної віртуальної машини.")
    add_dstu_paragraph(doc, "Таблиця 2.1 \u2014 Параметри віртуальної машини Build Box",
                       align=WD_ALIGN_PARAGRAPH.LEFT, first_line_indent=False)

    add_dstu_table(doc,
        ["Параметр", "Значення"],
        [
            ("Назва ВМ", "Devuan_Security"),
            ("Гостьова ОС (VMware)", "Debian 13.x 64-bit"),
            ("Тип конфігурації", "Typical (Recommended)"),
            ("ISO-образ", "devuan_excalibur_6.1.1_amd64_netinstall.iso"),
            ("Оперативна пам'ять", "2048 МБ (2 ГБ)"),
            ("Процесор", "2 ядра"),
            ("Жорсткий диск", "20 ГБ (Split into multiple files)"),
            ("Мережевий адаптер", "NAT"),
        ]
    )

    add_dstu_paragraph(doc, first_line_indent=False)
    add_dstu_paragraph(
        doc,
        "На рисунках 2.2\u20132.8 наведено послідовність екранів майстра створення "
        "нової віртуальної машини у VMware Workstation Pro 25H2."
    )

    screenshots_2 = [
        ("1.2_vmware_main_screen.png",      "2.2", "Головне меню VMware Workstation Pro 25H2"),
        ("1.3_vmware_wizard_type.png",      "2.3", "Вибір типу конфігурації ВМ (Typical / Custom)"),
        ("1.4_vmware_iso_selection.png",    "2.4", "Підключення ISO-образу Devuan Excalibur 6.1.1"),
        ("1.5_vmware_guest_os_debian13.png","2.5", "Вибір гостьової ОС: Debian 13.x 64-bit"),
        ("1.6_vmware_vm_name.png",          "2.6", "Задання імені віртуальної машини: Devuan_Security"),
        ("1.7_vmware_disk_size.png",        "2.7", "Налаштування розміру віртуального диску: 20 ГБ"),
        ("1.8_vmware_summary.png",          "2.8", "Фінальні параметри ВМ перед створенням"),
    ]
    for fname, num, caption in screenshots_2:
        add_screenshot(doc, os.path.join(SCREENSHOTS_DIR, fname), num, caption)

    # ==========================================
    # РОЗДІЛ 3–6: ЗАГЛУШКИ
    # ==========================================
    placeholders = [
        ("3 ФАЗА 2. ВСТАНОВЛЕННЯ DEVUAN БЕЗ ГРАФІЧНОГО СЕРЕДОВИЩА",
         "Цей розділ буде заповнено після завершення встановлення Devuan у VMware Workstation. "
         "Ключовим етапом є відмова від будь-якого графічного середовища на кроці вибору "
         "компонентів у Tasksel та встановлення лише SSH server і Standard system utilities.",
         ["3.1 \u2014 Boot menu Devuan Excalibur",
          "3.2 \u2014 Вибір мови встановлення",
          "3.3 \u2014 Розмітка диска",
          "3.4 \u2014 Екран Tasksel (без графічного середовища)",
          "3.5 \u2014 Перший вхід у систему"]),
        ("4 ФАЗА 3. КОНФІГУРУВАННЯ LIVE-BUILD", "", []),
        ("5 ФАЗА 4. ЗБІРКА ISO-ОБРАЗУ", "", []),
        ("6 ФАЗА 5. ТЕСТУВАННЯ ТА ОПТИМІЗАЦІЯ (QA)", "", []),
    ]

    for title, body, items in placeholders:
        doc.add_page_break()
        add_dstu_heading(doc, title, level=1)
        add_dstu_paragraph(doc, first_line_indent=False)
        if body:
            add_dstu_paragraph(doc, body)
        if items:
            add_dstu_paragraph(doc, "Скріншоти для даного розділу:", bold=True)
            for item in items:
                p = add_dstu_paragraph(doc, f"\u2014 {item} [очікується]", italic=True)
        if not body:
            add_dstu_paragraph(doc,
                "Розділ буде заповнено на відповідному етапі розробки.", italic=True)

    # ==========================================
    # ЗБЕРЕЖЕННЯ
    # ==========================================
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    doc.save(OUTPUT_FILE)
    print(f"✅ Щоденник (ДСТУ) збережено: {OUTPUT_FILE}")


if __name__ == "__main__":
    generate()
