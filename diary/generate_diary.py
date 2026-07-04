#!/usr/bin/env python3
"""
Генератор щоденника розробки дистрибутива Devuan Security OS
Запуск: python3 generate_diary.py
"""

import os
import sys

# Додаємо шлях до nix-profile пакетів
nix_site_packages = os.path.expanduser("~/.nix-profile/lib/python3.13/site-packages")
if nix_site_packages not in sys.path:
    sys.path.insert(0, nix_site_packages)

from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import datetime

# === КОНФІГУРАЦІЯ ===
SCREENSHOTS_DIR = "/home/zoozienix/projects/Devuan_Project/screenshots"
OUTPUT_FILE = "/home/zoozienix/projects/Devuan_Project/diary/Devuan_Security_OS_Diary.docx"

# === ДОПОМІЖНІ ФУНКЦІЇ ===

def add_horizontal_line(paragraph):
    """Додає горизонтальну лінію після параграфу"""
    p = paragraph._p
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '4A90D9')
    pBdr.append(bottom)
    pPr.append(pBdr)

def set_cell_background(cell, color_hex):
    """Встановлює колір фону комірки таблиці"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def add_screenshot(doc, filepath, caption, width_inches=5.5):
    """Додає скріншот із підписом"""
    if os.path.exists(filepath):
        # Вирівнювання по центру для зображення
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(filepath, width=Inches(width_inches))

        # Підпис під скріншотом
        caption_p = doc.add_paragraph(f"Рис. {caption}")
        caption_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        caption_p.runs[0].italic = True
        caption_p.runs[0].font.size = Pt(9)
        caption_p.runs[0].font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    else:
        # Заглушка якщо файл не знайдено
        p = doc.add_paragraph(f"[СКРІНШОТ ВІДСУТНІЙ: {os.path.basename(filepath)}]")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
        p.runs[0].italic = True

def add_phase_header(doc, phase_num, phase_title, status="В процесі"):
    """Додає заголовок фази"""
    doc.add_heading(f"Фаза {phase_num}: {phase_title}", level=1)
    status_colors = {
        "Завершено": "28A745",
        "В процесі": "FFC107",
        "Очікує": "6C757D"
    }
    color = status_colors.get(status, "6C757D")
    p = doc.add_paragraph()
    run = p.add_run(f"▶ Статус: {status}")
    run.bold = True
    run.font.color.rgb = RGBColor(
        int(color[0:2], 16),
        int(color[2:4], 16),
        int(color[4:6], 16)
    )
    doc.add_paragraph()  # відступ

def add_step(doc, step_num, title, description, commands=None):
    """Додає крок із описом та командами"""
    doc.add_heading(f"Крок {step_num}: {title}", level=2)
    p = doc.add_paragraph(description)
    p.style.font.size = Pt(11)

    if commands:
        doc.add_paragraph("Виконані команди:", style='Intense Quote')
        for cmd in commands:
            code_p = doc.add_paragraph(cmd)
            code_p.style = doc.styles['No Spacing']
            run = code_p.runs[0]
            run.font.name = 'Courier New'
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(0x00, 0x80, 0x00)


# === ГЕНЕРАЦІЯ ДОКУМЕНТУ ===

def generate_diary():
    doc = Document()

    # --- Налаштування сторінки ---
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)

    # --- Налаштування стилів ---
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)

    h1_style = doc.styles['Heading 1']
    h1_style.font.color.rgb = RGBColor(0x1A, 0x56, 0x9A)
    h1_style.font.size = Pt(16)
    h1_style.font.bold = True

    h2_style = doc.styles['Heading 2']
    h2_style.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)
    h2_style.font.size = Pt(13)

    # ==========================================
    # ТИТУЛЬНА СТОРІНКА
    # ==========================================
    for _ in range(4):
        doc.add_paragraph()

    title = doc.add_paragraph("Щоденник розробки")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.runs[0].font.size = Pt(28)
    title.runs[0].bold = True
    title.runs[0].font.color.rgb = RGBColor(0x1A, 0x56, 0x9A)

    subtitle = doc.add_paragraph("Devuan Security OS")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.runs[0].font.size = Pt(22)
    subtitle.runs[0].bold = True
    subtitle.runs[0].font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)

    doc.add_paragraph()

    desc = doc.add_paragraph(
        "Кастомний дистрибутив на базі Devuan Excalibur 6.x\n"
        "Орієнтований на безпеку • Без systemd • XFCE інтерфейс"
    )
    desc.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in desc.runs:
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    for _ in range(6):
        doc.add_paragraph()

    meta_info = [
        ("Автор:", "zoozienix"),
        ("Репозиторій:", "github.com/zoozieuniver/Devuan_Project"),
        ("Дата початку:", datetime.date.today().strftime("%d.%m.%Y")),
        ("Версія документу:", "0.1"),
    ]
    for label, value in meta_info:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_label = p.add_run(f"{label} ")
        run_label.bold = True
        run_label.font.size = Pt(11)
        run_value = p.add_run(value)
        run_value.font.size = Pt(11)
        run_value.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)

    doc.add_page_break()

    # ==========================================
    # ЗМІСТ (Ручний)
    # ==========================================
    doc.add_heading("Зміст", level=1)

    toc_items = [
        ("1. Вступ та мета проєкту", "3"),
        ("2. Технічний стек та обґрунтування вибору", "3"),
        ("Фаза 1: Підготовка середовища (Build Box)", "4"),
        ("  1.1. Завантаження ISO-образу Devuan", "4"),
        ("  1.2. Створення віртуальної машини у VMware", "4"),
        ("Фаза 2: Встановлення Devuan без графічного середовища", "—"),
        ("Фаза 3: Конфігурування live-build", "—"),
        ("Фаза 4: Збірка ISO-образу", "—"),
        ("Фаза 5: Тестування та оптимізація", "—"),
    ]

    for item, page in toc_items:
        p = doc.add_paragraph()
        run = p.add_run(item)
        if item.startswith("Фаза") or item.startswith("1.") or item.startswith("2."):
            run.bold = True
        run.font.size = Pt(11)
        p.add_run(f"{'.' * max(1, 60 - len(item))} {page}").font.color.rgb = RGBColor(0x99, 0x99, 0x99)

    doc.add_page_break()

    # ==========================================
    # РОЗДІЛ 1: ВСТУП
    # ==========================================
    doc.add_heading("1. Вступ та мета проєкту", level=1)
    intro_text = (
        "Даний щоденник фіксує процес розробки власного дистрибутиву GNU/Linux на базі "
        "Devuan Excalibur 6.x. Devuan є форком дистрибутиву Debian, головною особливістю "
        "якого є відмова від системи ініціалізації systemd на користь класичного SysVinit "
        "або альтернативних рішень (runit, OpenRC).\n\n"
        "Мета проєкту — створити мінімальну, безпечну та відтворювану (reproducible) систему, "
        "яка може використовуватися як основа для спеціалізованих середовищ безпеки. "
        "Кінцевий результат — завантажувальний ISO-образ, зібраний інструментом live-build, "
        "з повністю задокументованим та версіонованим процесом збірки через Git."
    )
    doc.add_paragraph(intro_text)

    doc.add_heading("2. Технічний стек та обґрунтування вибору", level=1)
    tech_table = doc.add_table(rows=1, cols=3)
    tech_table.style = 'Table Grid'
    hdr_cells = tech_table.rows[0].cells
    headers = ["Компонент", "Вибір", "Обґрунтування"]
    for i, hdr in enumerate(headers):
        hdr_cells[i].text = hdr
        set_cell_background(hdr_cells[i], "1A569A")
        for para in hdr_cells[i].paragraphs:
            for run in para.runs:
                run.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    tech_rows = [
        ("База дистрибутиву", "Devuan Excalibur 6.1.1", "Debian без systemd, стабільна база"),
        ("Ядро", "Linux (стабільний реліз)", "Широка підтримка обладнання"),
        ("Init система", "SysVinit", "Простота, передбачуваність, без залежностей"),
        ("Робочий стіл", "XFCE 4", "Легковаговість, повна функціональність"),
        ("Інструмент збірки", "live-build", "Декларативна конфігурація, стандарт Debian"),
        ("Гіпервізор", "VMware Workstation Pro 25H2", "Build Box та тестове середовище"),
        ("Версіонування", "Git + GitHub", "Відтворюваність збірки, журнал змін"),
    ]
    for row_data in tech_rows:
        row_cells = tech_table.add_row().cells
        for i, val in enumerate(row_data):
            row_cells[i].text = val
            if i == 0:
                for para in row_cells[i].paragraphs:
                    for run in para.runs:
                        run.bold = True

    doc.add_page_break()

    # ==========================================
    # ФАЗА 1
    # ==========================================
    add_phase_header(doc, 1, "Підготовка середовища (Build Box)", "В процесі")

    doc.add_paragraph(
        "Мета цієї фази — підготувати еталонне середовище (Build Box) у вигляді "
        "віртуальної машини, де надалі відбуватиметься збірка дистрибутиву. "
        "Обрана платформа — VMware Workstation Pro 25H2, оскільки вона надає "
        "стабільне та ізольоване середовище для роботи."
    )

    # Крок 1.1
    add_step(
        doc, "1.1", "Завантаження ISO-образу Devuan",
        "Для встановлення Build Box обрано образ типу netinstall дистрибутиву "
        "Devuan Excalibur 6.1.1 (amd64). Цей тип образу є мінімальним встановлювачем, "
        "який завантажує лише базову систему, а всі пакети отримує з мережі під час "
        "встановлення. Це гарантує актуальність встановлених пакетів.\n\n"
        "Образ завантажено з офіційного архіву Devuan: files.devuan.org. "
        "Версія Excalibur 6.x базується на Debian 13 (Trixie) та є поточним "
        "стабільним релізом."
    )
    doc.add_paragraph()
    add_screenshot(
        doc,
        os.path.join(SCREENSHOTS_DIR, "1.1_devuan_release_archive.png"),
        "1.1 — Офіційний архів ISO-образів Devuan Excalibur на files.devuan.org"
    )

    doc.add_paragraph()

    # Крок 1.2
    add_step(
        doc, "1.2", "Створення віртуальної машини у VMware Workstation",
        "Для Build Box створено нову віртуальну машину у VMware Workstation Pro 25H2. "
        "Параметри ВМ підібрані з урахуванням вимог процесу збірки дистрибутиву "
        "через live-build, який потребує достатньо оперативної пам'яті та обчислювальних ресурсів."
    )

    # Параметри ВМ у таблиці
    doc.add_paragraph("Параметри створеної віртуальної машини:").runs[0].bold = True
    vm_table = doc.add_table(rows=1, cols=2)
    vm_table.style = 'Table Grid'
    hdr = vm_table.rows[0].cells
    hdr[0].text = "Параметр"
    hdr[1].text = "Значення"
    set_cell_background(hdr[0], "2E75B6")
    set_cell_background(hdr[1], "2E75B6")
    for cell in hdr:
        for para in cell.paragraphs:
            for run in para.runs:
                run.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    vm_params = [
        ("Назва ВМ", "Devuan_Security"),
        ("Гостьова ОС (VMware)", "Debian 13.x 64-bit"),
        ("Тип конфігурації", "Typical (Recommended)"),
        ("ISO-образ", "devuan_excalibur_6.1.1_amd64_netinstall.iso"),
        ("Оперативна пам'ять", "2048 МБ (2 ГБ)"),
        ("Процесор", "2 ядра"),
        ("Жорсткий диск", "20 ГБ (Split into multiple files)"),
        ("Мережевий адаптер", "NAT"),
    ]
    for param, value in vm_params:
        row = vm_table.add_row().cells
        row[0].text = param
        row[0].paragraphs[0].runs[0].bold = True
        row[1].text = value

    doc.add_paragraph()
    doc.add_paragraph("Процес створення ВМ (покрокові скріншоти):").runs[0].bold = True
    doc.add_paragraph()

    screenshots_phase1 = [
        ("1.2_vmware_main_screen.png", "1.2 — Головне меню VMware Workstation Pro 25H2"),
        ("1.3_vmware_wizard_type.png", "1.3 — Вибір типу конфігурації ВМ (Typical / Custom)"),
        ("1.4_vmware_iso_selection.png", "1.4 — Підключення ISO-образу Devuan Excalibur"),
        ("1.5_vmware_guest_os_debian13.png", "1.5 — Вибір гостьової ОС: Debian 13.x 64-bit"),
        ("1.6_vmware_vm_name.png", "1.6 — Задання імені віртуальної машини: Devuan_Security"),
        ("1.7_vmware_disk_size.png", "1.7 — Налаштування розміру віртуального диску: 20 ГБ"),
        ("1.8_vmware_summary.png", "1.8 — Фінальний підсумок перед створенням ВМ"),
    ]

    for filename, caption in screenshots_phase1:
        add_screenshot(doc, os.path.join(SCREENSHOTS_DIR, filename), caption)
        doc.add_paragraph()

    # ==========================================
    # ФАЗА 2 (заглушка)
    # ==========================================
    doc.add_page_break()
    add_phase_header(doc, 2, "Встановлення Devuan без графічного середовища", "Очікує")
    doc.add_paragraph(
        "Цей розділ буде заповнено після завершення встановлення Devuan у VMware. "
        "Ключовим кроком є відмова від будь-якого графічного середовища (Desktop Environment) "
        "на етапі вибору компонентів у Tasksel, та встановлення лише SSH server та "
        "Standard system utilities."
    )
    p = doc.add_paragraph("Скріншоти для цієї фази:")
    p.runs[0].italic = True
    items = ["2.1 — Boot menu Devuan", "2.2 — Вибір мови встановлення",
             "2.3 — Розмітка диска", "2.4 — Екран Tasksel (без DE)", "2.5 — Перший вхід у систему"]
    for item in items:
        li = doc.add_paragraph(f"• {item} [ОЧІКУЄ]")
        li.runs[0].font.color.rgb = RGBColor(0x99, 0x99, 0x99)
        li.runs[0].italic = True

    # ==========================================
    # ФАЗА 3, 4, 5 (заглушки)
    # ==========================================
    for phase_num, phase_name in [
        (3, "Конфігурування live-build"),
        (4, "Збірка ISO-образу"),
        (5, "Тестування та оптимізація (QA)")
    ]:
        doc.add_page_break()
        add_phase_header(doc, phase_num, phase_name, "Очікує")
        p = doc.add_paragraph("Цей розділ буде заповнено на відповідному етапі розробки.")
        p.runs[0].italic = True
        p.runs[0].font.color.rgb = RGBColor(0x99, 0x99, 0x99)

    # ==========================================
    # ЗБЕРЕЖЕННЯ
    # ==========================================
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    doc.save(OUTPUT_FILE)
    print(f"✅ Щоденник збережено: {OUTPUT_FILE}")


if __name__ == "__main__":
    generate_diary()
