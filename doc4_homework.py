"""
doc4_homework.py
Generates: output/04_домашняя_работа.docx
Audience: Children (parents help)
Format: A4 portrait, 14pt minimum
Tone: Friendly, stress-free, motivating
Total time: ~30 minutes
"""

import os
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUTPUT_DIR = "output"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "04_домашняя_работа.docx")

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


def add_illustration_placeholder(doc, height_cm, dall_e_prompt, width_cm=None,
                                  align=WD_TABLE_ALIGNMENT.LEFT):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = align
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


def add_drawing_placeholder(doc, height_cm, label_text, width_cm=None):
    """Child's drawing area placeholder."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.rows[0].cells[0]
    set_cell_bg(cell, "FAFAFA")
    set_table_border(tbl, "CCCCCC", "6")

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
    run = p.add_run(label_text)
    run.italic = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0x99, 0x99, 0x99)
    doc.add_paragraph()


def add_task_heading(doc, task_num, title, time_hint, bg_color="F0F0F0", border_color="888888"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    cell = tbl.rows[0].cells[0]
    set_cell_bg(cell, bg_color)
    set_table_border(tbl, border_color, "6")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(f"ЗАДАНИЕ {task_num} — {title}   ({time_hint})")
    run.bold = True
    run.font.size = Pt(15)
    doc.add_paragraph()


def add_instruction(doc, text, size=14):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(f"📌 {text}")
    run.bold = True
    run.font.size = Pt(size)
    return p


def add_mini_cheatsheet(doc):
    tbl = doc.add_table(rows=1, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_border(tbl, "888888", "4")
    data = [
        ("СУЩ (КТО?ЧТО?) = _____", "DDEEFF"),
        ("ПРИЛ (КАКОЙ?) = ~~~~~", "DDFFDD"),
        ("ГЛ (ЧТО ДЕЛАЕТ?) = \u2550\u2550\u2550\u2550\u2550", "FFE8E0"),
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


def create_homework():
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
    # HEADER
    # =========================================================
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("\U0001F3E0 ДОМАШНЯЯ РАБОТА — ЧАСТИ РЕЧИ")
    run_title.bold = True
    run_title.font.size = Pt(22)

    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_name = p_name.add_run("Имя: _______________________   Дата: ___________")
    run_name.font.size = Pt(14)

    p_star = doc.add_paragraph()
    p_star.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_star = p_star.add_run("\u2b50 Ты справишься! Мы верим в тебя! \u2b50")
    run_star.bold = True
    run_star.font.size = Pt(16)
    run_star.font.color.rgb = RGBColor(0xFF, 0xA5, 0x00)

    doc.add_paragraph()

    # Illustration placeholder (top-right area via a simple centered table)
    add_illustration_placeholder(
        doc, 5,
        "Warm, encouraging cartoon illustration of a happy child sitting at a desk doing homework, "
        "Russian language book open, pencil in hand, cozy lamp, small potted plant, cheerful atmosphere. "
        "For 7-year-old Russian heritage language learner. Soft warm colors.",
        width_cm=7,
        align=WD_TABLE_ALIGNMENT.RIGHT
    )

    add_mini_cheatsheet(doc)

    # =========================================================
    # TASK 1 — MY OBJECTS (~10 min)
    # =========================================================
    add_task_heading(doc, 1, "МОИ ПРЕДМЕТЫ", "~10 минут", "DDEEFF", "1E90FF")
    add_instruction(
        doc,
        "Найди дома 5 предметов. Нарисуй или напиши их. Подчеркни одной линией ___"
    )

    items_tbl = doc.add_table(rows=6, cols=3)
    items_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_border(items_tbl, "888888", "4")

    col_headers = ["#", "Нарисуй или напиши слово", "Подчеркни _____"]
    for i, hdr in enumerate(col_headers):
        cell = items_tbl.rows[0].cells[i]
        set_cell_bg(cell, "DDEEFF")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(hdr)
        run.bold = True
        run.font.size = Pt(12)

    for row_idx in range(1, 6):
        row = items_tbl.rows[row_idx]

        # Number cell
        cell_num = row.cells[0]
        p_num = cell_num.paragraphs[0]
        p_num.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_num = p_num.add_run(str(row_idx))
        run_num.bold = True
        run_num.font.size = Pt(14)

        # Drawing cell
        cell_draw = row.cells[1]
        p_draw = cell_draw.paragraphs[0]
        p_draw.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_draw = p_draw.add_run(" " * 40)
        run_draw.font.size = Pt(20)

        # Underline check cell
        cell_check = row.cells[2]
        p_check = cell_check.paragraphs[0]
        p_check.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run_check = p_check.add_run("✓")
        run_check.font.size = Pt(16)
        run_check.font.color.rgb = COLOR_NOUN

    doc.add_paragraph()

    # =========================================================
    # TASK 2 — DESCRIBE OBJECTS (~10 min)
    # =========================================================
    add_task_heading(doc, 2, "ОПИСЫВАЮ ПРЕДМЕТЫ", "~10 минут", "DDFFDD", "32CD32")
    add_instruction(
        doc,
        "Выбери 3 предмета из задания 1. Опиши каждый! Подчеркни волнистой линией ~~~~~"
    )

    add_illustration_placeholder(
        doc, 4,
        "Small cute illustration of a child thinking with a thought bubble containing colorful adjective "
        "words in Russian: МЯГКИЙ, БОЛЬШОЙ, КРАСНЫЙ, ТЁПЛЫЙ, КРУГЛЫЙ. "
        "Cartoon style, white background, child-friendly.",
        width_cm=6
    )

    adj_lines = [
        "1. ________________________ — он (она/оно) ________________________ (подчеркни ~~~~~)",
        "2. ________________________ — он (она/оно) ________________________ (подчеркни ~~~~~)",
        "3. ________________________ — он (она/оно) ________________________ (подчеркни ~~~~~)",
    ]
    for line in adj_lines:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(14)
        run = p.add_run(line)
        run.font.size = Pt(14)

    # =========================================================
    # TASK 3 — WHAT DO OBJECTS DO? (~10 min) — new page
    # =========================================================
    doc.add_page_break()

    p_name2 = doc.add_paragraph()
    run_n2 = p_name2.add_run("Имя: _______________________   Дата: ___________")
    run_n2.font.size = Pt(14)
    add_mini_cheatsheet(doc)

    add_task_heading(doc, 3, "ЧТО ДЕЛАЮТ МОИ ПРЕДМЕТЫ?", "~10 минут", "FFE8E0", "FF4500")
    add_instruction(
        doc,
        "Напиши, что делает предмет. Подчеркни двойной линией \u2550\u2550\u2550\u2550\u2550"
    )

    add_illustration_placeholder(
        doc, 4,
        "Small cute illustration showing household objects doing actions: a cat sleeping, a clock ticking, "
        "a book being read, a cup steaming. Russian action words shown as labels. "
        "Cartoon style, white background.",
        width_cm=6
    )

    verb_lines = [
        "1. ________________________ (что делает?) ________________________   (подчеркни \u2550\u2550\u2550\u2550\u2550)",
        "2. ________________________ (что делает?) ________________________   (подчеркни \u2550\u2550\u2550\u2550\u2550)",
        "3. ________________________ (что делает?) ________________________   (подчеркни \u2550\u2550\u2550\u2550\u2550)",
    ]
    for line in verb_lines:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(14)
        run = p.add_run(line)
        run.font.size = Pt(14)

    doc.add_paragraph()

    # =========================================================
    # TASK 4 — MY STORY ★ (optional, ~10 min)
    # =========================================================
    tbl_star = doc.add_table(rows=1, cols=1)
    tbl_star.alignment = WD_TABLE_ALIGNMENT.LEFT
    cell_star = tbl_star.rows[0].cells[0]
    set_cell_bg(cell_star, "FFF8DC")
    set_table_border(tbl_star, "FFD700", "8")
    p_star_hdr = cell_star.paragraphs[0]
    p_star_hdr.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_star_hdr = p_star_hdr.add_run("ЗАДАНИЕ 4 — МОЙ РАССКАЗ  ★   (по желанию, ~10 минут)")
    run_star_hdr.bold = True
    run_star_hdr.font.size = Pt(15)
    doc.add_paragraph()

    add_instruction(
        doc,
        "Составь 2–3 предложения про твой любимый предмет. Используй все три части речи!",
        size=14
    )

    add_drawing_placeholder(doc, 6, "Нарисуй свой предмет здесь", width_cm=6)

    # Lined writing area
    for _ in range(6):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run("_" * 70)
        run.font.size = Pt(14)

    doc.add_paragraph()

    # Self-check box
    check_tbl = doc.add_table(rows=1, cols=1)
    check_tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    check_cell = check_tbl.rows[0].cells[0]
    set_cell_bg(check_cell, "E8FFE8")
    set_table_border(check_tbl, "32CD32", "6")
    p_chk = check_cell.paragraphs[0]
    p_chk.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_chk = p_chk.add_run(
        "\u2705 Проверь себя:\n"
        "\u25a1 Я подчеркнул(а) существительное _____\n"
        "\u25a1 Я подчеркнул(а) прилагательное ~~~~~\n"
        "\u25a1 Я подчеркнул(а) глагол \u2550\u2550\u2550\u2550\u2550"
    )
    run_chk.font.size = Pt(13)

    doc.add_paragraph()

    # Motivational ending
    p_end = doc.add_paragraph()
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_end = p_end.add_run("\U0001F31F Отличная работа! До встречи в воскресенье! \U0001F31F")
    run_end.bold = True
    run_end.font.size = Pt(18)
    run_end.font.color.rgb = RGBColor(0xFF, 0xA5, 0x00)

    doc.save(OUTPUT_FILE)
    print(f"✅  Создан: {OUTPUT_FILE}")


if __name__ == "__main__":
    create_homework()
