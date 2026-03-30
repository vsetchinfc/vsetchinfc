"""
doc2_visual_content.py
Generates: output/02_визуальные_плакаты.docx
Audience: Children (and teacher for demonstration)
Format:
  - Posters 1-3: A4 landscape, headings 36pt bold, text 18pt
  - Sheet 4 (cheatsheet): A4 portrait
"""

import os
from docx import Document
from docx.shared import Cm, Pt, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.section import WD_ORIENT

OUTPUT_DIR = "output"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "02_визуальные_плакаты.docx")

COLOR_NOUN = RGBColor(0x1E, 0x90, 0xFF)
COLOR_ADJ = RGBColor(0x32, 0xCD, 0x32)
COLOR_VERB = RGBColor(0xFF, 0x45, 0x00)
COLOR_NOUN_HEX = "1E90FF"
COLOR_ADJ_HEX = "32CD32"
COLOR_VERB_HEX = "FF4500"
COLOR_PLACEHOLDER = "D3D3D3"


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
    """Add a grey rectangle placeholder with DALL-E prompt text inside."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.rows[0].cells[0]
    set_cell_bg(cell, COLOR_PLACEHOLDER)
    set_table_border(tbl, "999999", "6")

    # Set cell height via row height
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
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    run2 = p.add_run(f"DALL-E: {dall_e_prompt}")
    run2.italic = True
    run2.font.size = Pt(8)
    run2.font.color.rgb = RGBColor(0x44, 0x44, 0x44)

    doc.add_paragraph()


def add_colored_heading(doc, text, size, color_rgb, align=WD_ALIGN_PARAGRAPH.CENTER):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(size)
    run.font.color.rgb = color_rgb
    return p


def add_rule_box(doc, text, border_color_hex, size=14):
    """Add a text in a colored single-cell table acting as a box."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.rows[0].cells[0]
    set_table_border(tbl, border_color_hex, "8")
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(size)
    doc.add_paragraph()


def add_marking_example(doc, word, marking_symbol, description, size=14):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{word}  {marking_symbol}")
    run.bold = True
    run.font.size = Pt(size)
    run2 = p.add_run(f"  ← {description}")
    run2.font.size = Pt(12)
    run2.italic = True


def add_example_sentence(doc, before, keyword, after, marking, color_rgb, size=14):
    """Add a sentence with a marked keyword."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.left_indent = Cm(1)
    p.paragraph_format.space_after = Pt(8)

    if before:
        run = p.add_run(before)
        run.font.size = Pt(size)

    run_kw = p.add_run(keyword)
    run_kw.bold = True
    run_kw.font.size = Pt(size)
    run_kw.font.color.rgb = color_rgb
    if marking == "underline":
        run_kw.underline = True
    elif marking == "wave":
        run_kw.underline = True  # Word doesn't support wavy underline via python-docx easily
        # Add wavy symbol beneath
    elif marking == "double":
        run_kw.underline = True

    if after:
        run_after = p.add_run(after)
        run_after.font.size = Pt(size)

    # Add marking line below
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p2.paragraph_format.left_indent = Cm(1)
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(4)

    marking_chars = {
        "underline": "_" * len(keyword),
        "wave": "~" * len(keyword),
        "double": "\u2550" * len(keyword),
    }
    run_mark = p2.add_run(marking_chars.get(marking, ""))
    run_mark.font.size = Pt(size)
    run_mark.font.color.rgb = color_rgb


def set_landscape(section):
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Cm(29.7)
    section.page_height = Cm(21)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)


def set_portrait(section):
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2)
    section.right_margin = Cm(2)


def add_page_break_with_section(doc, landscape=True):
    """Add a new section with the given orientation."""
    new_section = doc.add_section()
    if landscape:
        set_landscape(new_section)
    else:
        set_portrait(new_section)


def create_visual_content():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    doc = Document()

    # First section: landscape for poster 1
    section = doc.sections[0]
    set_landscape(section)

    # =========================================================
    # POSTER 1 — СУЩЕСТВИТЕЛЬНОЕ (blue)
    # =========================================================
    add_colored_heading(doc, "СУЩЕСТВИТЕЛЬНОЕ", 36, COLOR_NOUN)

    add_illustration_placeholder(
        doc, 8,
        "Bright, child-friendly cartoon illustration for Russian language class, showing 6 common objects "
        "with Russian labels: КОТ (orange tabby cat), МЯЧ (red ball), ДОМ (colorful house), "
        "ЯБЛОКО (green apple), КНИГА (open book), СОЛНЦЕ (smiling sun). Each Russian word written below "
        "its object with a single straight underline. Bold, simple font. White background, primary colors, "
        "cheerful style for 7-year-old children."
    )

    add_rule_box(
        doc,
        "« Существительное — это ПРЕДМЕТ или ЖИВОЕ СУЩЕСТВО »",
        COLOR_NOUN_HEX, size=16
    )

    p_q = doc.add_paragraph()
    p_q.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_q = p_q.add_run("КТО?   ЧТО?")
    run_q.bold = True
    run_q.font.size = Pt(24)
    run_q.font.color.rgb = COLOR_NOUN

    doc.add_paragraph()
    add_marking_example(doc, "КОТ", "_____", "подчёркиваем одной линией", size=16)
    doc.add_paragraph()

    examples = [
        ("Рыжий ", "КОТ", " пьёт молоко."),
        ("Круглый ", "МЯЧ", " лежит на полу."),
        ("Маленькая ", "КНИГА", " лежит на столе."),
    ]
    for before, keyword, after in examples:
        add_example_sentence(doc, before, keyword, after, "underline", COLOR_NOUN, size=16)

    # =========================================================
    # POSTER 2 — ПРИЛАГАТЕЛЬНОЕ (green) — new landscape section
    # =========================================================
    add_page_break_with_section(doc, landscape=True)

    add_colored_heading(doc, "ПРИЛАГАТЕЛЬНОЕ", 36, COLOR_ADJ)

    add_illustration_placeholder(
        doc, 8,
        "Bright, child-friendly cartoon illustration showing adjectives in Russian: "
        "РЫЖИЙ КОТ (orange cat labeled РЫЖИЙ with wavy underline), "
        "КРАСНЫЙ МЯЧ (red ball labeled КРАСНЫЙ with wavy underline), "
        "БОЛЬШОЙ ДОМ (big house labeled БОЛЬШОЙ with wavy underline), "
        "МАЛЕНЬКОЕ ЯБЛОКО (tiny apple labeled МАЛЕНЬКОЕ with wavy underline). "
        "White background, cheerful cartoon style, primary colors, for 7-year-old children learning Russian."
    )

    add_rule_box(
        doc,
        "« Прилагательное описывает предмет — говорит нам, КАКОЙ он! »",
        COLOR_ADJ_HEX, size=16
    )

    p_q2 = doc.add_paragraph()
    p_q2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_q2 = p_q2.add_run("КАКОЙ?   КАКАЯ?   КАКОЕ?   КАКИЕ?")
    run_q2.bold = True
    run_q2.font.size = Pt(22)
    run_q2.font.color.rgb = COLOR_ADJ

    doc.add_paragraph()
    add_marking_example(doc, "РЫЖИЙ", "~~~~~", "подчёркиваем волнистой линией", size=16)

    # Mnemonic
    p_mn = doc.add_paragraph()
    p_mn.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_mn = p_mn.add_run(
        "\U0001F4A1 ПРИЛАГАТЕЛЬНОЕ «прилагается» к существительному — как прилипает!"
    )
    run_mn.bold = True
    run_mn.font.size = Pt(14)
    run_mn.font.color.rgb = COLOR_ADJ

    doc.add_paragraph()

    adj_examples = [
        ("", "РЫЖИЙ", " кот спит на диване."),
        ("", "КРУГЛЫЙ", " мяч лежит в траве."),
        ("", "БОЛЬШОЕ", " яблоко упало с дерева."),
    ]
    for before, keyword, after in adj_examples:
        add_example_sentence(doc, before, keyword, after, "wave", COLOR_ADJ, size=16)

    # =========================================================
    # POSTER 3 — ГЛАГОЛ (red) — new landscape section
    # =========================================================
    add_page_break_with_section(doc, landscape=True)

    add_colored_heading(doc, "ГЛАГОЛ", 36, COLOR_VERB)

    add_illustration_placeholder(
        doc, 8,
        "Bright, child-friendly cartoon illustration showing action verbs in Russian: "
        "a cat БЕЖИТ (running, labeled БЕЖИТ with double underline), "
        "a ball ЛЕТИТ (flying through air, labeled ЛЕТИТ with double underline), "
        "children ЧИТАЮТ (reading books, labeled ЧИТАЮТ with double underline), "
        "sun СВЕТИТ (shining, labeled СВЕТИТ with double underline). "
        "Dynamic movement lines, cartoon style, white background, for 7-year-old children learning Russian."
    )

    add_rule_box(
        doc,
        "« Глагол — это ДЕЙСТВИЕ предмета или существа »",
        COLOR_VERB_HEX, size=16
    )

    p_q3 = doc.add_paragraph()
    p_q3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_q3 = p_q3.add_run("ЧТО ДЕЛАЕТ?   ЧТО ДЕЛАЮТ?   ЧТО СДЕЛАЛ?")
    run_q3.bold = True
    run_q3.font.size = Pt(20)
    run_q3.font.color.rgb = COLOR_VERB

    doc.add_paragraph()
    add_marking_example(doc, "БЕЖИТ", "\u2550\u2550\u2550\u2550\u2550", "подчёркиваем двойной линией", size=16)

    p_mn3 = doc.add_paragraph()
    p_mn3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_mn3 = p_mn3.add_run(
        "\U0001F4A1 Глагол можно ПОКАЗАТЬ жестом или изобразить в движении!"
    )
    run_mn3.bold = True
    run_mn3.font.size = Pt(14)
    run_mn3.font.color.rgb = COLOR_VERB

    doc.add_paragraph()

    verb_examples = [
        ("Кот ", "БЕЖИТ", " домой."),
        ("Мяч ", "ЛЕТИТ", " высоко."),
        ("Солнце ", "СВЕТИТ", " ярко."),
    ]
    for before, keyword, after in verb_examples:
        add_example_sentence(doc, before, keyword, after, "double", COLOR_VERB, size=16)

    # =========================================================
    # SHEET 4 — CHEATSHEET (portrait)
    # =========================================================
    add_page_break_with_section(doc, landscape=False)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("МОЯ ШПАРГАЛКА — ЧАСТИ РЕЧИ")
    run_title.bold = True
    run_title.font.size = Pt(24)

    doc.add_paragraph()

    # Cheatsheet table
    cheat_table = doc.add_table(rows=4, cols=4)
    cheat_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_border(cheat_table, "888888", "6")

    cheat_headers = ["Часть речи", "Вопросы", "Линия", "Пример"]
    cheat_hdr_row = cheat_table.rows[0]
    for i, hdr in enumerate(cheat_headers):
        cell = cheat_hdr_row.cells[i]
        set_cell_bg(cell, "2F4F8F")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(hdr)
        run.bold = True
        run.font.size = Pt(13)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    cheat_data = [
        ("СУЩЕСТВИТЕЛЬНОЕ", "КТО? ЧТО?", "_______", "КОТ, МЯЧ, КНИГА", "DDEEFF"),
        ("ПРИЛАГАТЕЛЬНОЕ", "КАКОЙ? КАКАЯ? КАКОЕ?", "~~~~~~~", "РЫЖИЙ, КРУГЛЫЙ", "DDFFDD"),
        ("ГЛАГОЛ", "ЧТО ДЕЛАЕТ?", "\u2550\u2550\u2550\u2550\u2550\u2550\u2550", "БЕЖИТ, ЧИТАЕТ", "FFE8E0"),
    ]

    cheat_colors = [COLOR_NOUN, COLOR_ADJ, COLOR_VERB]

    for row_idx, (pos, questions, line, example, bg) in enumerate(cheat_data, start=1):
        row = cheat_table.rows[row_idx]
        for col_idx, val in enumerate([pos, questions, line, example]):
            cell = row.cells[col_idx]
            set_cell_bg(cell, bg)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(val)
            run.font.size = Pt(14)
            if col_idx == 0:
                run.bold = True
                run.font.color.rgb = cheat_colors[row_idx - 1]
            if col_idx == 2:
                run.font.color.rgb = cheat_colors[row_idx - 1]

    doc.add_paragraph()
    doc.add_paragraph()

    p_name = doc.add_paragraph()
    p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_name = p_name.add_run("Имя: _______________________     Дата: _______________________")
    run_name.font.size = Pt(14)

    doc.save(OUTPUT_FILE)
    print(f"✅  Создан: {OUTPUT_FILE}")


if __name__ == "__main__":
    create_visual_content()
