"""
doc3_check_tasks.py
Generates: output/03_проверочные_задания.docx
Audience: Children
Format: A4 portrait, minimum 14pt font
Principle: Scaffolding — from simple to complex (5 levels)
"""

import os
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUTPUT_DIR = "output"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "03_проверочные_задания.docx")

COLOR_NOUN = RGBColor(0x1E, 0x90, 0xFF)
COLOR_ADJ = RGBColor(0x32, 0xCD, 0x32)
COLOR_VERB = RGBColor(0xFF, 0x45, 0x00)


def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def set_table_border(table, color="AAAAAA", size="4"):
    tbl = table._tbl
    tblPr = tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{side}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), size)
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), color)
        tblBorders.append(border)
    tblPr.append(tblBorders)


def add_illustration_placeholder(doc, height_cm, dall_e_prompt, width_cm=None):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    cell = tbl.rows[0].cells[0]
    set_cell_bg(cell, "D3D3D3")
    set_table_border(tbl, "999999", "6")

    tr = tbl.rows[0]._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement("w:trHeight")
    trHeight.set(qn("w:val"), str(int(Cm(height_cm).pt * 20)))
    trHeight.set(qn("w:hRule"), "exact")
    trPr.append(trHeight)

    if width_cm:
        tbl.columns[0].width = Cm(width_cm)

    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("[ ИЛЛЮСТРАЦИЯ ]\n")
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    run2 = p.add_run(f"DALL-E: {dall_e_prompt}")
    run2.italic = True
    run2.font.size = Pt(8)
    run2.font.color.rgb = RGBColor(0x44, 0x44, 0x44)
    doc.add_paragraph()


def add_task_heading(doc, task_num, title, level_label, bg_color="F0F0F0"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    cell = tbl.rows[0].cells[0]
    set_cell_bg(cell, bg_color)
    set_table_border(tbl, "888888", "6")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(f"ЗАДАНИЕ {task_num} — {title}   ({level_label})")
    run.bold = True
    run.font.size = Pt(15)
    doc.add_paragraph()


def add_instruction(doc, text, size=13):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(f"📌 {text}")
    run.bold = True
    run.font.size = Pt(size)
    return p


def add_mini_cheatsheet(doc):
    """Add a small cheatsheet table at the top of each page."""
    tbl = doc.add_table(rows=1, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_border(tbl, "888888", "4")

    data = [
        ("СУЩ: КТО? ЧТО?  _____", "DDEEFF"),
        ("ПРИЛ: КАКОЙ?  ~~~~~", "DDFFDD"),
        ("ГЛ: ЧТО ДЕЛАЕТ?  \u2550\u2550\u2550\u2550\u2550", "FFE8E0"),
    ]

    for i, (text, bg) in enumerate(data):
        cell = tbl.rows[0].cells[i]
        set_cell_bg(cell, bg)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(10)

    doc.add_paragraph()


def add_answer_box(doc, text):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    cell = tbl.rows[0].cells[0]
    set_cell_bg(cell, "E8FFE8")
    set_table_border(tbl, "32CD32", "6")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.font.size = Pt(12)
    run.italic = True
    doc.add_paragraph()


def create_check_tasks():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    doc = Document()

    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)

    # =========================================================
    # PAGE HEADER
    # =========================================================
    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_name = p_name.add_run("Имя: ___________________________   Дата: ___________")
    run_name.font.size = Pt(14)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("ПРОВЕРЯЕМ СЕБЯ — ЧАСТИ РЕЧИ \U0001F31F")
    run_title.bold = True
    run_title.font.size = Pt(22)

    doc.add_paragraph()
    add_mini_cheatsheet(doc)

    # =========================================================
    # TASK 1 — RECOGNIZE NOUN (Level 1)
    # =========================================================
    add_task_heading(doc, 1, "УЗНАЮ СУЩЕСТВИТЕЛЬНОЕ", "Уровень 1", "DDEEFF")
    add_instruction(doc, "Обведи в кружок слова-существительные (КТО? ЧТО?)")

    add_illustration_placeholder(
        doc, 5,
        "8 simple cartoon images in a grid for Russian language exercise: КОТ (cat), БЕЖИТ (running legs), "
        "КРАСНЫЙ (red color swatch), ЯБЛОКО (apple), ПРЫГАЕТ (jumping figure), ДОМ (house), "
        "ВЕСЁЛЫЙ (happy face emoji-style), КНИГА (book). Each image has its Russian word label beneath it. "
        "Clean white background, simple bold outlines, child-friendly, for 7-year-olds.",
        width_cm=15
    )

    # Word cards
    words_task1 = ["КОТ", "БЕЖИТ", "КРАСНЫЙ", "ЯБЛОКО", "ПРЫГАЕТ", "ДОМ", "ВЕСЁЛЫЙ", "КНИГА"]
    p_words = doc.add_paragraph()
    p_words.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_words.paragraph_format.space_after = Pt(10)
    for word in words_task1:
        run = p_words.add_run(f"  {word}  ")
        run.bold = True
        run.font.size = Pt(16)

    add_answer_box(doc, "✅ Существительные: КОТ, ЯБЛОКО, ДОМ, КНИГА")

    # =========================================================
    # TASK 2 — SORT WORDS (Level 2)
    # =========================================================
    add_task_heading(doc, 2, "СОРТИРУЮ СЛОВА", "Уровень 2", "DDFFDD")
    add_instruction(doc, "Напиши каждое слово в нужную корзину")

    add_illustration_placeholder(
        doc, 5,
        "Three cartoon baskets for sorting exercise. First basket labeled СУЩЕСТВИТЕЛЬНОЕ with a small "
        "object icon (toy block). Second basket labeled ПРИЛАГАТЕЛЬНОЕ with a color palette icon. "
        "Third basket labeled ГЛАГОЛ with a running shoe icon. Colorful, cheerful, child-friendly, "
        "white background, Russian labels clearly visible.",
        width_cm=15
    )

    words_task2 = ["СОЛНЦЕ", "ПРЫГАТЬ", "СИНИЙ", "СОБАКА", "ЛЕТИТ", "МЯГКИЙ", "ЦВЕТОК", "РИСУЕТ",
                   "БОЛЬШОЙ", "СТОЛ", "БЕЛЫЙ", "ПОЁТ"]
    p_words2 = doc.add_paragraph()
    p_words2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_words2.paragraph_format.space_after = Pt(10)
    for word in words_task2:
        run = p_words2.add_run(f"  {word}  ")
        run.bold = True
        run.font.size = Pt(14)

    sort_tbl = doc.add_table(rows=2, cols=3)
    sort_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_border(sort_tbl, "888888", "6")

    sort_headers = [
        ("СУЩЕСТВИТЕЛЬНОЕ", "DDEEFF"),
        ("ПРИЛАГАТЕЛЬНОЕ", "DDFFDD"),
        ("ГЛАГОЛ", "FFE8E0"),
    ]
    for i, (hdr, bg) in enumerate(sort_headers):
        cell = sort_tbl.rows[0].cells[i]
        set_cell_bg(cell, bg)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(hdr)
        run.bold = True
        run.font.size = Pt(13)

    for i, bg in enumerate([s[1] for s in sort_headers]):
        cell = sort_tbl.rows[1].cells[i]
        set_cell_bg(cell, bg)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run("_____________\n_____________\n_____________\n_____________")
        run.font.size = Pt(14)

    doc.add_paragraph()

    # =========================================================
    # TASK 3 — UNDERLINE (Level 3) — new page
    # =========================================================
    doc.add_page_break()

    p_name2 = doc.add_paragraph()
    run_n2 = p_name2.add_run("Имя: ___________________________   Дата: ___________")
    run_n2.font.size = Pt(14)
    add_mini_cheatsheet(doc)

    add_task_heading(doc, 3, "ПОДЧЁРКИВАЮ", "Уровень 3", "FFE8E0")
    add_instruction(doc, "Подчеркни части речи правильными линиями")

    # Color reminder
    reminder_tbl = doc.add_table(rows=1, cols=3)
    reminder_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_border(reminder_tbl, "888888", "4")
    reminder_data = [
        ("СУЩ = _____", "DDEEFF", COLOR_NOUN),
        ("ПРИЛ = ~~~~~", "DDFFDD", COLOR_ADJ),
        ("ГЛ = \u2550\u2550\u2550\u2550\u2550", "FFE8E0", COLOR_VERB),
    ]
    for i, (text, bg, color) in enumerate(reminder_data):
        cell = reminder_tbl.rows[0].cells[i]
        set_cell_bg(cell, bg)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(12)
        run.font.color.rgb = color

    doc.add_paragraph()

    sentences_task3 = [
        "1.  Маленький щенок весело бежит домой.",
        "2.  Пушистая кошка пьёт холодное молоко.",
        "3.  Яркое солнце светит над зелёным лесом.",
        "4.  Весёлые дети громко поют красивую песню.",
        "5.  Большая рыба плывёт в синей реке.",
    ]

    for sentence in sentences_task3:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(16)
        run = p.add_run(sentence)
        run.font.size = Pt(15)
        # Space for underlining
        p2 = doc.add_paragraph()
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(8)
        run2 = p2.add_run("   " * 30)
        run2.font.size = Pt(6)

    # =========================================================
    # TASK 4 — FILL IN WORD (Level 4)
    # =========================================================
    add_task_heading(doc, 4, "ВСТАВЛЯЮ СЛОВО", "Уровень 4", "FFF5DD")
    add_instruction(doc, "Вставь подходящее слово. Подсказка в скобках!")

    fill_sentences = [
        ("1. Мохнатый ", "_______ (КТО? = существительное)", " громко мяукает. \U0001F431"),
        ("2. ", "_______ (КАКОЙ? = прилагательное)", " пёс виляет хвостом. \U0001F436"),
        ("3. Птица радостно ", "_______ (ЧТО ДЕЛАЕТ? = глагол)", " на ветке. \U0001F426"),
        ("4. ", "_______ (КАКОЕ? = прилагательное)", " яблоко лежит на столе. \U0001F34E"),
        ("5. Мальчик быстро ", "_______ (ЧТО ДЕЛАЕТ? = глагол)", " к школе. \U0001F3C3"),
    ]

    for before, blank, after in fill_sentences:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(14)

        if before:
            run = p.add_run(before)
            run.font.size = Pt(14)

        run_blank = p.add_run(blank)
        run_blank.bold = True
        run_blank.font.size = Pt(14)
        run_blank.font.color.rgb = RGBColor(0x88, 0x00, 0x88)

        if after:
            run_after = p.add_run(after)
            run_after.font.size = Pt(14)

    # =========================================================
    # TASK 5 — COMPOSE SENTENCE ★ (Level 5) — new page
    # =========================================================
    doc.add_page_break()

    p_name3 = doc.add_paragraph()
    run_n3 = p_name3.add_run("Имя: ___________________________   Дата: ___________")
    run_n3.font.size = Pt(14)
    add_mini_cheatsheet(doc)

    # Star task header
    tbl_star = doc.add_table(rows=1, cols=1)
    tbl_star.alignment = WD_TABLE_ALIGNMENT.LEFT
    cell_star = tbl_star.rows[0].cells[0]
    set_cell_bg(cell_star, "FFF8DC")
    set_table_border(tbl_star, "FFD700", "8")
    p_star = cell_star.paragraphs[0]
    p_star.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_star = p_star.add_run("ЗАДАНИЕ 5 — СОСТАВЛЯЮ ПРЕДЛОЖЕНИЕ  ★   (Уровень 5 — для сильных)")
    run_star.bold = True
    run_star.font.size = Pt(15)
    doc.add_paragraph()

    add_instruction(doc,
                    "Посмотри на картинку. Составь предложение, используя все три части речи!", size=13)

    add_illustration_placeholder(
        doc, 8,
        "Cheerful cartoon scene: a fluffy orange cat sitting on a red chair in a cozy room, suitable "
        "for a Russian language writing prompt for 7-year-old children. Clear, simple, colorful.",
        width_cm=10
    )

    p_sent1 = doc.add_paragraph()
    run_s1 = p_sent1.add_run("Моё предложение: _____________________________________________")
    run_s1.font.size = Pt(14)

    p_check1 = doc.add_paragraph()
    run_c1 = p_check1.add_run(
        "Проверь себя: подчеркни  СУЩ _____   ПРИЛ ~~~~~   ГЛ \u2550\u2550\u2550\u2550\u2550"
    )
    run_c1.font.size = Pt(12)
    run_c1.italic = True

    doc.add_paragraph()

    add_illustration_placeholder(
        doc, 8,
        "Cheerful cartoon scene: a little girl with pigtails running in a sunny green park with a yellow dog, "
        "suitable for a Russian language writing prompt for 7-year-old children. Clear, simple, colorful.",
        width_cm=10
    )

    p_sent2 = doc.add_paragraph()
    run_s2 = p_sent2.add_run("Моё предложение: _____________________________________________")
    run_s2.font.size = Pt(14)

    p_check2 = doc.add_paragraph()
    run_c2 = p_check2.add_run(
        "Проверь себя: подчеркни  СУЩ _____   ПРИЛ ~~~~~   ГЛ \u2550\u2550\u2550\u2550\u2550"
    )
    run_c2.font.size = Pt(12)
    run_c2.italic = True

    doc.save(OUTPUT_FILE)
    print(f"✅  Создан: {OUTPUT_FILE}")


if __name__ == "__main__":
    create_check_tasks()
