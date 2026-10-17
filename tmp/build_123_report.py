from pathlib import Path
import re

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


OUTPUT = Path("output") / "รายงานโครงงาน_ระบบจัดการร้าน123พาณิชย์ปลีกส่ง_ฉบับปรับปรุง_วิธีการดำเนินงานเดิม.docx"
FONT = "TH Sarabun PSK"
INK = RGBColor(0, 0, 0)
THAI_WORD_SPACE = re.compile(r"(?<=[\u0E00-\u0E7F]) (?=[\u0E00-\u0E7F])")
PROJECT_NAME_BREAK = re.compile(r"(123พาณิชย์ปลีกส่ง)(?=[\u0E00-\u0E7F])")


def thai_justify_text(text):
    """Keep Thai prose compact while preserving non-expanding technical-term gaps."""
    text = PROJECT_NAME_BREAK.sub(lambda match: f"{match.group(1)}\u2009", text)
    text = THAI_WORD_SPACE.sub("", text)
    return text.replace(" ", "\u2009")


def set_rfonts(rpr, name=FONT):
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for key in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn(f"w:{key}"), name)


def set_style_font(style, size, bold=False, italic=False):
    style.font.name = FONT
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.italic = italic
    style.font.color.rgb = INK
    rpr = style.element.get_or_add_rPr()
    set_rfonts(rpr)
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.append(lang)
    lang.set(qn("w:val"), "th-TH")
    lang.set(qn("w:eastAsia"), "th-TH")
    lang.set(qn("w:bidi"), "th-TH")


def set_run_font(run, size=16, bold=None, italic=None):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.color.rgb = INK
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    rpr = run._element.get_or_add_rPr()
    set_rfonts(rpr)
    lang = rpr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rpr.append(lang)
    lang.set(qn("w:val"), "th-TH")
    lang.set(qn("w:eastAsia"), "th-TH")
    lang.set(qn("w:bidi"), "th-TH")


def set_paragraph_keep(paragraph, keep_with_next=False, keep_together=False):
    ppr = paragraph._p.get_or_add_pPr()
    if keep_with_next:
        tag = qn("w:keepNext")
        if ppr.find(tag) is None:
            ppr.append(OxmlElement("w:keepNext"))
    if keep_together:
        tag = qn("w:keepLines")
        if ppr.find(tag) is None:
            ppr.append(OxmlElement("w:keepLines"))


def set_cell_margins(cell, top=80, start=100, bottom=80, end=100):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for side, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{side}"))
        if node is None:
            node = OxmlElement(f"w:{side}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def shade_cell(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")


def set_table_geometry(table, widths_cm):
    widths_dxa = [round(width / 2.54 * 1440) for width in widths_cm]
    total = sum(widths_dxa)
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.first_child_found_in("w:tblW")
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(total))
    tbl_w.set(qn("w:type"), "dxa")
    tbl_ind = tbl_pr.first_child_found_in("w:tblInd")
    if tbl_ind is None:
        tbl_ind = OxmlElement("w:tblInd")
        tbl_pr.append(tbl_ind)
    tbl_ind.set(qn("w:w"), "0")
    tbl_ind.set(qn("w:type"), "dxa")
    layout = tbl_pr.first_child_found_in("w:tblLayout")
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")

    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths_dxa:
        col = OxmlElement("w:gridCol")
        col.set(qn("w:w"), str(width))
        grid.append(col)

    for row in table.rows:
        for cell, width in zip(row.cells, widths_dxa):
            cell.width = Cm(width / 1440 * 2.54)
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.first_child_found_in("w:tcW")
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(width))
            tc_w.set(qn("w:type"), "dxa")
            set_cell_margins(cell)


def repeat_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    marker = OxmlElement("w:tblHeader")
    marker.set(qn("w:val"), "true")
    tr_pr.append(marker)


def prevent_row_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    marker = OxmlElement("w:cantSplit")
    tr_pr.append(marker)


def set_page_number_start(section, start=1):
    pg_num = section._sectPr.find(qn("w:pgNumType"))
    if pg_num is None:
        pg_num = OxmlElement("w:pgNumType")
        section._sectPr.append(pg_num)
    pg_num.set(qn("w:start"), str(start))


def continue_page_numbers(section):
    """Allow a new chapter section to continue Arabic pagination."""
    pg_num = section._sectPr.find(qn("w:pgNumType"))
    if pg_num is not None:
        section._sectPr.remove(pg_num)


def clear_paragraph(paragraph):
    """Clear runs from a paragraph while retaining its paragraph properties."""
    for child in list(paragraph._p):
        if child.tag != qn("w:pPr"):
            paragraph._p.remove(child)


def add_page_field(paragraph):
    run = paragraph.add_run()
    set_run_font(run, size=16)
    r = run._r
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    r.extend([begin, instr, separate, text, end])


def set_header_page_number(section):
    section.header.is_linked_to_previous = False
    p = section.header.paragraphs[0]
    clear_paragraph(p)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    add_page_field(p)


def configure_body_section(section, page_start=None):
    """Apply the SNRU page layout and hide the number on a chapter's first page."""
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.left_margin = Cm(3.81)
    section.right_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.header_distance = Cm(1.27)
    section.footer_distance = Cm(1.27)
    section.different_first_page_header_footer = True
    section.first_page_header.is_linked_to_previous = False
    clear_paragraph(section.first_page_header.paragraphs[0])
    if page_start is None:
        continue_page_numbers(section)
    else:
        set_page_number_start(section, page_start)
    set_header_page_number(section)


def configure_styles(doc):
    styles = doc.styles

    normal = styles["Normal"]
    set_style_font(normal, 16)
    # Use ordinary justify. Thai word spaces are replaced by zero-width break
    # opportunities in body paragraphs, preventing Word from over-stretching them.
    normal.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    normal.paragraph_format.line_spacing = 1.0
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.first_line_indent = Cm(1.27)

    styles["List Number"].base_style = normal
    set_style_font(styles["List Number"], 16)
    styles["List Number"].paragraph_format.line_spacing = 1.0
    styles["List Number"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    styles["List Number"].paragraph_format.left_indent = Cm(1.6)
    styles["List Number"].paragraph_format.first_line_indent = Cm(-0.55)
    styles["List Number"].paragraph_format.space_after = Pt(0)

    for name, size, bold, before, after in (
        ("Heading 1", 18, True, 12, 4),
        ("Heading 2", 16, True, 8, 2),
        ("Heading 3", 16, True, 4, 0),
    ):
        style = styles[name]
        style.base_style = normal
        set_style_font(style, size, bold=bold)
        style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        style.paragraph_format.line_spacing = 1.0
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.first_line_indent = Cm(0 if name == "Heading 1" else 1.27)

    custom = {
        "Cover Title": (22, True, WD_ALIGN_PARAGRAPH.CENTER, 0, 5, 0),
        "Cover English": (20, True, WD_ALIGN_PARAGRAPH.CENTER, 0, 16, 0),
        "Cover Name": (20, True, WD_ALIGN_PARAGRAPH.CENTER, 0, 0, 0),
        "Cover Detail": (20, True, WD_ALIGN_PARAGRAPH.CENTER, 0, 0, 0),
        "Chapter": (20, True, WD_ALIGN_PARAGRAPH.CENTER, 0, 0, 0),
        "Chapter Title": (20, True, WD_ALIGN_PARAGRAPH.CENTER, 0, 12, 0),
        "TOC Title": (20, True, WD_ALIGN_PARAGRAPH.CENTER, 0, 12, 0),
        "Table Caption": (16, True, WD_ALIGN_PARAGRAPH.CENTER, 6, 2, 0),
        "Table Text": (16, False, WD_ALIGN_PARAGRAPH.LEFT, 0, 0, 0),
        "Table Header": (16, True, WD_ALIGN_PARAGRAPH.CENTER, 0, 0, 0),
        "Note": (16, False, WD_ALIGN_PARAGRAPH.LEFT, 4, 4, 0),
    }
    for name, (size, bold, align, before, after, indent) in custom.items():
        style = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        style.base_style = normal
        set_style_font(style, size, bold=bold)
        style.paragraph_format.alignment = align
        style.paragraph_format.line_spacing = 1.0
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.first_line_indent = Cm(indent)


def add_paragraph(doc, text="", style="Normal", align=None, indent=None, keep_next=False, keep_together=False):
    p = doc.add_paragraph(style=style)
    if align is not None:
        p.alignment = align
    if indent is not None:
        p.paragraph_format.first_line_indent = Cm(indent)
    if text:
        run = p.add_run(text)
        sizes = {
            "Cover Title": 22,
            "Cover English": 20,
            "Cover Name": 20,
            "Cover Detail": 20,
            "Chapter": 20,
            "Chapter Title": 20,
            "Heading 1": 18,
        }
        size = sizes.get(style, 16)
        set_run_font(run, size=size, bold=style in ("Cover Title", "Cover English", "Cover Name", "Cover Detail", "Chapter", "Chapter Title", "Heading 1", "Heading 2", "Heading 3", "Table Caption", "Table Header"))
    set_paragraph_keep(p, keep_next, keep_together)
    return p


def add_heading(doc, text, level=1):
    return add_paragraph(doc, text, style=f"Heading {level}", keep_next=True, keep_together=True)


def add_body(doc, text, no_indent=False):
    return add_paragraph(
        doc,
        thai_justify_text(text),
        style="Normal",
        indent=0 if no_indent else 1.27,
        keep_together=True,
    )


def add_numbered(doc, text):
    p = doc.add_paragraph(style="List Number")
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.keep_together = True
    set_run_font(p.add_run(thai_justify_text(text)), size=16)
    return p


def add_caption(doc, text):
    return add_paragraph(doc, text, style="Table Caption", keep_next=True, keep_together=True)


def add_table(doc, headers, rows, widths_cm, header_align=None):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_geometry(table, widths_cm)
    header = table.rows[0]
    repeat_header(header)
    prevent_row_split(header)
    for index, label in enumerate(headers):
        cell = header.cells[index]
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        shade_cell(cell, "F2F2F2")
        p = cell.paragraphs[0]
        p.style = "Table Header"
        p.alignment = (header_align[index] if header_align else WD_ALIGN_PARAGRAPH.CENTER)
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.keep_together = True
        set_run_font(p.add_run(label), size=16, bold=True)
    for values in rows:
        cells = table.add_row().cells
        for index, value in enumerate(values):
            cell = cells[index]
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.style = "Table Text"
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.keep_together = True
            set_run_font(p.add_run(str(value)), size=16)
        prevent_row_split(table.rows[-1])
    set_table_geometry(table, widths_cm)
    return table


def add_chapter(doc, number, title, page_break=True):
    if page_break:
        configure_body_section(doc.add_section(WD_SECTION_START.NEW_PAGE))
    add_paragraph(doc, f"บทที่ {number}", style="Chapter", keep_next=True, keep_together=True)
    add_paragraph(doc, title, style="Chapter Title", keep_next=True, keep_together=True)


def add_cover(doc):
    for height in (2.0, 0.55):
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(height * 20)
    add_paragraph(
        doc,
        "ระบบจัดการร้าน 123พาณิชย์ปลีกส่ง อำเภอนิคมน้ำอูน\nด้วยเว็บแอปพลิเคชัน",
        style="Cover Title",
        keep_together=True,
    )
    add_paragraph(
        doc,
        "Web Application for 123 Phanit Wholesale Shop Management",
        style="Cover English",
        keep_together=True,
    )
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(72)
    add_paragraph(doc, "ภรัณยู โยธาพล", style="Cover Name", keep_together=True)
    add_paragraph(doc, "66102122113", style="Cover Name", keep_together=True)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(84)
    add_paragraph(doc, "โครงงานคอมพิวเตอร์นี้เป็นส่วนหนึ่งของการศึกษาตามหลักสูตรปริญญาวิทยาศาสตรบัณฑิต", style="Cover Detail", keep_together=True)
    add_paragraph(doc, "สาขาวิชาเทคโนโลยีคอมพิวเตอร์และดิจิทัล", style="Cover Detail", keep_together=True)
    add_paragraph(doc, "มหาวิทยาลัยราชภัฏสกลนคร", style="Cover Detail", keep_together=True)
    add_paragraph(doc, "พ.ศ. 2569", style="Cover Detail", keep_together=True)


def add_toc_page(doc):
    """Add a front-matter page where the Word TOC field is inserted after build."""
    add_paragraph(doc, "สารบัญ", style="TOC Title", keep_next=True, keep_together=True)
    p = add_paragraph(doc, "[[TOC]]", style="Normal", align=WD_ALIGN_PARAGRAPH.LEFT, indent=0)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(0)


def build_document():
    doc = Document()
    configure_styles(doc)
    props = doc.core_properties
    props.title = "ระบบจัดการร้าน 123พาณิชย์ปลีกส่ง ด้วยเว็บแอปพลิเคชัน"
    props.subject = "รายงานโครงงานฉบับปรับปรุงให้สอดคล้องกับระบบเว็บ"
    props.author = "ภรัณยู โยธาพล"
    props.comments = "จัดทำจากเอกสารต้นฉบับและตรวจสอบความสอดคล้องกับซอร์สโค้ดของระบบ"

    cover_section = doc.sections[0]
    cover_section.page_width = Cm(21)
    cover_section.page_height = Cm(29.7)
    cover_section.top_margin = Cm(2.54)
    cover_section.left_margin = Cm(3.81)
    cover_section.right_margin = Cm(2.54)
    cover_section.bottom_margin = Cm(2.54)
    cover_section.header_distance = Cm(1.27)
    cover_section.footer_distance = Cm(1.27)
    add_cover(doc)

    toc_section = doc.add_section(WD_SECTION_START.NEW_PAGE)
    configure_body_section(toc_section)
    add_toc_page(doc)

    body_section = doc.add_section(WD_SECTION_START.NEW_PAGE)
    configure_body_section(body_section, page_start=1)

    # Chapter 1
    add_chapter(doc, 1, "บทนำ", page_break=False)
    add_heading(doc, "1.1 ที่มาและความสำคัญ")
    add_body(doc, "ปัจจุบันเทคโนโลยีดิจิทัลมีบทบาทต่อการดำเนินธุรกิจทุกระดับ โดยเฉพาะร้านค้าปลีกส่งที่ต้องบริหารข้อมูลสินค้า คำสั่งซื้อ และการจัดส่งให้ทันต่อการใช้งานจริง ร้าน 123พาณิชย์ปลีกส่งมีลักษณะการทำงานที่เกี่ยวข้องกับการรับคำสั่งซื้อจากผู้ค้าปลีก การเตรียมสินค้า การนำส่ง และการตรวจสอบสถานะงาน หากบันทึกข้อมูลด้วยกระดาษหรือสื่อที่แยกกัน อาจทำให้ตรวจสอบจำนวนสินค้าและสถานะคำสั่งซื้อได้ไม่สะดวก")
    add_body(doc, "โครงงานนี้จึงพัฒนาระบบจัดการร้านในรูปแบบเว็บแอปพลิเคชัน เพื่อให้ผู้เกี่ยวข้องใช้งานผ่านเว็บเบราว์เซอร์ได้ตามสิทธิ์ของตน ระบบรองรับการจัดการสินค้า การปรับสต็อก การส่งคำสั่งซื้อ และการติดตามการนำส่งโดยอาศัยตำแหน่ง GPS ที่พนักงานอนุญาตให้แชร์ระหว่างทำงาน ทั้งนี้การติดตามมีจุดประสงค์เพื่อสื่อสารสถานะการจัดส่งและช่วยให้เจ้าของร้านหรือผู้ค้าปลีกตรวจสอบงานได้สะดวกขึ้น")
    add_body(doc, "ระบบที่พัฒนามีผู้ใช้ 4 บทบาท ได้แก่ เจ้าของร้าน ผู้ดูแลระบบ พนักงาน และผู้ค้าปลีก ข้อมูลถูกจัดเก็บในฐานข้อมูล Supabase PostgreSQL โดยกำหนดสิทธิ์การใช้งานในฝั่งระบบและฐานข้อมูล เพื่อลดการเข้าถึงข้อมูลเกินความจำเป็น การทำงานดังกล่าวช่วยลดขั้นตอนการสอบถามข้อมูลซ้ำ ๆ และทำให้ภาพรวมของคำสั่งซื้อและการจัดส่งอ่านได้ชัดเจนขึ้น")

    add_heading(doc, "1.2 วัตถุประสงค์")
    add_numbered(doc, "เพื่อพัฒนาระบบจัดการร้าน 123พาณิชย์ปลีกส่งด้วยเว็บแอปพลิเคชัน")
    add_numbered(doc, "เพื่อสนับสนุนการจัดการสินค้า คำสั่งซื้อ สต็อก และสถานะการจัดส่งให้เป็นระบบ")
    add_numbered(doc, "เพื่อใช้เป็นแนวทางประเมินประสิทธิภาพและความพึงพอใจของผู้ใช้งานระบบ")

    add_heading(doc, "1.3 ขอบเขตของโครงงาน")
    add_body(doc, "ขอบเขตด้านเนื้อหากำหนดจากความสามารถที่มีอยู่ในระบบเว็บปัจจุบัน โดยแบ่งสิทธิ์การใช้งานตามบทบาทดังตารางที่ 1.1 ระบบไม่มีโมดูลรับชำระเงิน ออกใบเสร็จ หรือระบบบัญชี จึงใช้คำว่า “ยืนยันคำสั่งซื้อ” แทนการชำระสินค้าในเอกสารฉบับนี้")
    add_caption(doc, "ตารางที่ 1.1 ขอบเขตการใช้งานตามบทบาท")
    add_table(
        doc,
        ["บทบาท", "ขอบเขตการทำงาน"],
        [
            ("OWNER\n(เจ้าของร้าน)", "ดูภาพรวมสินค้าและรายการสต็อกน้อย เพิ่ม ปรับ หรือลบสินค้า ดูคำสั่งซื้อทั้งหมด และติดตามการนำส่ง"),
            ("ADMIN\n(ผู้ดูแลระบบ)", "ตรวจสอบรายการสินค้าและเพิ่มหรือลดจำนวนสต็อกตามสิทธิ์ที่กำหนด"),
            ("EMPLOYEE\n(พนักงาน)", "จัดการคำสั่งซื้อที่รับผิดชอบ เปลี่ยนสถานะตามลำดับ PENDING, PACKED, DELIVERING และ COMPLETED รวมทั้งแชร์พิกัดจริงระหว่างนำส่ง"),
            ("RETAILER\n(ผู้ค้าปลีก)", "สมัครสมาชิกด้วยเบอร์โทรศัพท์ ชื่อร้าน และรหัสผ่าน เลือกสินค้าตามหน่วยกล่องหรือแพ็ค ยืนยันคำสั่งซื้อ และดูเฉพาะคำสั่งซื้อกับพิกัดของร้านตนเอง"),
        ],
        [3.55, 11.10],
    )
    add_paragraph(doc, "หมายเหตุ ตะกร้าสินค้าถูกเก็บไว้ในเว็บเบราว์เซอร์ของผู้ค้าปลีกจนกว่าจะยืนยันคำสั่งซื้อ จากนั้นระบบจึงสร้างข้อมูลคำสั่งซื้อในฐานข้อมูล", style="Note", indent=0)

    add_heading(doc, "1.4 เครื่องมือที่ใช้ในการพัฒนา")
    add_heading(doc, "1.4.1 ภาษาคอมพิวเตอร์และเทคโนโลยีเว็บ", level=2)
    add_body(doc, "ระบบพัฒนาด้วย TypeScript และ TSX บน Next.js 16 ร่วมกับ React 19 สำหรับส่วนติดต่อผู้ใช้ ใช้ HTML และ CSS ในการจัดโครงสร้างและรูปแบบของหน้าเว็บ และใช้ Tailwind CSS 4 สำหรับชุดเครื่องมือด้านการจัดรูปแบบ โดยทำงานบน Node.js ตามโครงสร้างของเว็บแอปพลิเคชัน")
    add_heading(doc, "1.4.2 ซอฟต์แวร์และบริการ", level=2)
    add_body(doc, "ใช้ Visual Studio Code สำหรับพัฒนาและแก้ไขซอร์สโค้ด ใช้ Supabase PostgreSQL เป็นฐานข้อมูลและเรียกใช้ข้อมูลผ่าน REST/RPC จากฝั่งเซิร์ฟเวอร์ของ Next.js ใช้ Git และ GitHub สำหรับควบคุมเวอร์ชันของซอร์สโค้ด และเตรียมเชื่อมต่อ Vercel สำหรับกระบวนการ build และนำเว็บแอปพลิเคชันขึ้นใช้งานจริง")
    add_heading(doc, "1.4.3 ฮาร์ดแวร์", level=2)
    add_body(doc, "เครื่องคอมพิวเตอร์พกพาที่ใช้พัฒนาระบบมีหน่วยประมวลผล AMD Ryzen 5 5600H หน่วยความจำ 16 GB หน่วยเก็บข้อมูล SSD 477 GB และใช้ระบบปฏิบัติการ Windows 11")

    add_heading(doc, "1.5 วิธีการดำเนินงาน")
    add_body(doc, "สำหรับการดำเนินงานระบบจัดการร้าน 123พาณิชย์ปลีกส่ง ใช้เทคโนโลยีโปรแกรมประยุกต์บนเว็บ (Web Application) ในการพัฒนาระบบ ใช้ TypeScript และ TSX ร่วมกับ HTML, CSS, Node.js, Next.js, React และ Tailwind CSS ใช้ Supabase PostgreSQL เป็นฐานข้อมูล และใช้ Visual Studio Code เป็นเครื่องมือหลักในการพัฒนาเว็บแอปพลิเคชัน โดยมีขั้นตอนการพัฒนา ดังต่อไปนี้")
    add_heading(doc, "1.5.1 วิเคราะห์และออกแบบระบบ", level=2)
    add_numbered(doc, "แผนภาพความสัมพันธ์ของข้อมูล (ER Diagram)")
    add_numbered(doc, "แผนภาพกรณีใช้งาน (Use Case Diagram)")
    add_numbered(doc, "การออกแบบส่วนติดต่อผู้ใช้ (User Interface)")
    add_numbered(doc, "การออกแบบส่วนเชื่อมต่อโปรแกรมประยุกต์ (API)")
    add_numbered(doc, "การพัฒนาและทดสอบระบบ")
    add_numbered(doc, "การทดสอบประสิทธิภาพของระบบ")

    add_heading(doc, "1.6 ประโยชน์ที่คาดว่าจะได้รับ")
    add_body(doc, "ผู้จัดทำได้ฝึกทักษะการพัฒนาระบบแบบ Full Stack ตั้งแต่การออกแบบหน้าจอ API ฐานข้อมูล การกำหนดสิทธิ์ผู้ใช้ และการนำระบบขึ้นใช้งาน ร้านค้าสามารถลดเวลาการตรวจสอบข้อมูลสินค้าและคำสั่งซื้อได้ รวมถึงมีข้อมูลสถานะการนำส่งที่ดูได้จากระบบเดียวกัน การทำงานของร้านจึงมีโครงสร้างที่ช่วยให้ติดตามงานได้สะดวกขึ้น")

    add_heading(doc, "1.7 แผนการดำเนินงาน")
    add_caption(doc, "ตารางที่ 1.2 ระยะเวลาในการดำเนินงาน ปี พ.ศ. 2569")
    months = ["กิจกรรม", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค."]
    schedule = [
        ("ศึกษาและรวบรวมข้อมูล", "✓", "✓", "", "", "", "", "", "", ""),
        ("วิเคราะห์และออกแบบระบบ", "", "✓", "✓", "", "", "", "", "", ""),
        ("ออกแบบหน้าจอผู้ใช้", "", "", "✓", "✓", "", "", "", "", ""),
        ("พัฒนาระบบ", "", "", "", "✓", "✓", "✓", "", "", ""),
        ("ทดสอบและแก้ไขข้อผิดพลาด", "", "", "", "", "", "✓", "✓", "", ""),
        ("จัดทำเอกสารเพื่อขึ้นสอบ", "", "", "", "", "", "", "", "✓", ""),
        ("สอบจบขั้นสุดท้าย", "", "", "", "", "", "", "", "", "✓"),
        ("จัดทำเล่มสมบูรณ์", "", "", "", "", "", "", "", "", "✓"),
    ]
    add_table(doc, months, schedule, [4.25] + [1.155] * 9)

    # Chapter 2
    add_chapter(doc, 2, "เอกสารและงานวิจัยที่เกี่ยวข้อง")
    add_body(doc, "ร้าน 123พาณิชย์ปลีกส่งมีการรับสินค้า ตรวจสอบจำนวน และจัดส่งให้ผู้ค้าปลีกตามรอบการจัดส่งของร้าน โดยรอบวันจันทร์และวันพฤหัสบดีจัดส่งในพื้นที่อำเภอวาริชภูมิ ตำบลวาริชภูมิ และตำบลคำบ่อ ส่วนรอบวันอังคารและวันศุกร์จัดส่งในพื้นที่อำเภอวาริชภูมิ ตำบลวาริชภูมิ และตำบลปลาโหล การมีระบบกลางช่วยให้ข้อมูลสินค้าและคำสั่งซื้อที่เกี่ยวข้องกับแต่ละรอบงานตรวจสอบได้ง่ายกว่าการสื่อสารผ่านหลายช่องทาง")

    add_heading(doc, "2.1 แนวคิดการพัฒนาระบบแบบ Full Stack")
    add_body(doc, "การพัฒนาแบบ Full Stack คือการพัฒนาที่ครอบคลุมทั้งส่วนติดต่อผู้ใช้และส่วนทำงานเบื้องหลังของระบบ ในโครงงานนี้ ส่วนหน้าบ้านใช้ Next.js และ React เพื่อแสดงหน้าจอให้เหมาะกับบทบาทของผู้ใช้ ส่วนหลังบ้านใช้ API route ของ Next.js สำหรับตรวจสอบ session รับข้อมูล และเรียกข้อมูลจาก Supabase PostgreSQL การออกแบบลักษณะนี้ช่วยให้ส่วนที่ผู้ใช้มองเห็นและส่วนที่จัดการข้อมูลทำงานเชื่อมกันอย่างเป็นระบบ")
    add_body(doc, "ข้อมูลสำคัญ เช่น สินค้า คำสั่งซื้อ รายการสินค้าในคำสั่งซื้อ และตำแหน่งการจัดส่ง ถูกเก็บในฐานข้อมูลกลาง การดำเนินการที่มีผลต่อข้อมูลสำคัญ เช่น การปรับสต็อก การสร้างคำสั่งซื้อ และการเปลี่ยนสถานะจัดส่ง ถูกเรียกผ่านฟังก์ชัน RPC เพื่อให้การตรวจสอบเงื่อนไขและการบันทึกข้อมูลทำได้ในขั้นตอนที่ควบคุมได้")

    add_heading(doc, "2.2 การวิเคราะห์และออกแบบระบบ")
    add_heading(doc, "2.2.1 แผนภาพความสัมพันธ์ระหว่างเอนทิตี", level=2)
    add_body(doc, "การออกแบบฐานข้อมูลใช้แนวคิด Entity-Relationship Diagram เพื่ออธิบายเอนทิตี ข้อมูลคุณลักษณะ และความสัมพันธ์ในระบบ เอนทิตีหลักของระบบที่พัฒนาจริงประกอบด้วย profiles, products, orders, order_items และ delivery_tracking ส่วน app_credentials ใช้เก็บ salt และ hash ของรหัสผ่านสำหรับระบบยืนยันตัวตนแบบกำหนดเอง")
    add_body(doc, "ความสัมพันธ์สำคัญคือผู้ค้าปลีกหนึ่งรายสามารถมีคำสั่งซื้อได้หลายรายการ คำสั่งซื้อหนึ่งรายการมีรายการสินค้าได้หลายรายการ และแต่ละรายการสินค้าอ้างอิงสินค้าในคลัง การติดตามการจัดส่งเก็บพิกัดล่าสุดหนึ่งชุดต่อคำสั่งซื้อ โดยอ้างอิงพนักงานที่ได้รับมอบหมายงาน จึงไม่มีตาราง category, cart หรือ employee แยกต่างหากเหมือนโครงสร้างเดิม")
    add_heading(doc, "2.2.2 การกำหนดสิทธิ์และการยืนยันตัวตน", level=2)
    add_body(doc, "ระบบใช้ Custom Auth โดยให้เบอร์โทรศัพท์ไทยเป็นชื่อผู้ใช้ร่วมกับรหัสผ่าน รหัสผ่านไม่ถูกจัดเก็บเป็นข้อความปกติ แต่เก็บเป็น salt และ scrypt hash ในตาราง app_credentials หลังเข้าสู่ระบบสำเร็จ ระบบสร้าง session ใน HTTP-only cookie เพื่อช่วยลดความเสี่ยงจากการเข้าถึง token ด้วย JavaScript ฝั่งเบราว์เซอร์ ผู้ค้าปลีกสมัครบัญชีได้จากหน้าเว็บ ส่วนบัญชี OWNER, ADMIN และ EMPLOYEE ถูกเตรียมโดยผู้ดูแลระบบ")

    add_heading(doc, "2.3 ซอฟต์แวร์และบริการที่เกี่ยวข้อง")
    add_heading(doc, "2.3.1 Visual Studio Code", level=2)
    add_body(doc, "Visual Studio Code เป็นโปรแกรมแก้ไขซอร์สโค้ดที่รองรับ TypeScript, JavaScript, Node.js และ Git จึงเหมาะสำหรับใช้พัฒนา ตรวจสอบ และปรับปรุงเว็บแอปพลิเคชันของโครงงาน")
    add_heading(doc, "2.3.2 Node.js, Next.js และ React", level=2)
    add_body(doc, "Node.js ใช้เป็น runtime สำหรับการพัฒนาและ build ระบบ Next.js ขณะที่ Next.js ช่วยจัดโครงสร้างหน้าเว็บและ API route ส่วน React ใช้จัดการองค์ประกอบของหน้าจอและปฏิสัมพันธ์ของผู้ใช้ เทคโนโลยีทั้งสามช่วยให้ระบบใช้ภาษา TypeScript เป็นหลักได้ทั้งส่วนติดต่อและส่วนตรรกะของเว็บ")
    add_heading(doc, "2.3.3 Supabase PostgreSQL", level=2)
    add_body(doc, "Supabase ใช้เป็นบริการฐานข้อมูล PostgreSQL ของระบบ โดย Next.js ฝั่งเซิร์ฟเวอร์เรียกข้อมูลผ่าน REST และ RPC การออกแบบนี้ใช้ service role key เฉพาะฝั่งเซิร์ฟเวอร์ ไม่ส่งค่านี้ไปยังเบราว์เซอร์ และกำหนดสิทธิ์ของข้อมูลตามบทบาทผู้ใช้ในขั้นตอนของระบบ")
    add_heading(doc, "2.3.4 GitHub และ Vercel", level=2)
    add_body(doc, "GitHub ใช้เป็นพื้นที่เก็บซอร์สโค้ดออนไลน์และบันทึกประวัติการพัฒนา ทำให้สามารถตรวจสอบการเปลี่ยนแปลงของโครงการได้ เมื่อเชื่อม repository กับ Vercel ระบบสามารถ build และ deploy จาก GitHub ได้ตาม branch ที่กำหนด โดยต้องตั้งค่า URL ของ Supabase, service role key และ session secret ใน Vercel แยกจากซอร์สโค้ด เพื่อไม่ให้ข้อมูลสำคัญถูกเผยแพร่")

    add_heading(doc, "2.4 งานวิจัยและวรรณกรรมที่เกี่ยวข้อง")
    add_body(doc, "กิตติกูล กรรณส่วนบุญ (2567) ได้พัฒนาระบบซื้อขายเสื้อผ้าออนไลน์ในลักษณะ Full Stack ซึ่งครอบคลุมการวิเคราะห์และออกแบบระบบสำหรับทั้งผู้ใช้และผู้ดูแลระบบ แนวคิดดังกล่าวเป็นประโยชน์ต่อการกำหนดบทบาทผู้ใช้และการแยกส่วนการทำงานของระบบจัดการร้าน")
    add_body(doc, "สุภาพร คำเตจา วรัญญา สองสีใส ศศิชา สุขกาย นคร ไชยวงศ์ศักดา วิลาสินี ศรีสุวรรณ และเสกสรรค์ วินยางค์กูล (2023) อธิบายว่าข้อมูลความเคลื่อนไหวของสินค้าคงคลังช่วยสนับสนุนการวางแผนสั่งซื้อ และเว็บแอปพลิเคชันช่วยลดเวลาและต้นทุนในการจัดการสินค้าคงคลัง แนวคิดนี้สอดคล้องกับการจัดการสินค้าและการปรับสต็อกในโครงงาน")
    add_body(doc, "อนุพงษ์ สุขประเสริฐ ดวงจันทร์ สีหาราช เจษฎาพร ปาคาวัง และจิตรนันท์ ศรีเจริญ (2025) ศึกษาการพัฒนาพาณิชย์อิเล็กทรอนิกส์เพื่อเพิ่มศักยภาพการขายสินค้าของกลุ่มวิสาหกิจชุมชน ขณะที่งานของสุนันท์และคณะ (2565) และกุลนิษฐ์และคณะ (2568) สะท้อนความสำคัญของการจัดการสินค้าคงคลังที่ลดความคลาดเคลื่อนและการตรวจสอบสินค้าก่อนจัดส่ง แนวคิดเหล่านี้ใช้เป็นแนวทางในการออกแบบลำดับสถานะคำสั่งซื้อและการตรวจสอบสต็อกของระบบ")

    # Chapter 3
    add_chapter(doc, 3, "วิธีดำเนินการวิจัย")
    add_body(doc, "การดำเนินการวิจัยเพื่อพัฒนาระบบจัดการร้าน 123พาณิชย์ปลีกส่งด้วยเว็บแอปพลิเคชัน ใช้กระบวนการวิเคราะห์ความต้องการ ออกแบบข้อมูลและส่วนติดต่อผู้ใช้ พัฒนาระบบ ทดสอบการทำงานตามสิทธิ์ และเตรียมขั้นตอนการนำระบบขึ้นใช้งาน โดยรายละเอียดสำคัญมีดังต่อไปนี้")

    add_heading(doc, "3.1 การวิเคราะห์และออกแบบระบบ")
    add_heading(doc, "3.1.1 ขอบเขตการไหลของข้อมูล", level=2)
    add_body(doc, "ระบบทำหน้าที่เป็นศูนย์กลางระหว่างเจ้าของร้าน ผู้ดูแลระบบ พนักงาน และผู้ค้าปลีก ผู้ใช้แต่ละกลุ่มรับส่งข้อมูลกับระบบเฉพาะที่เกี่ยวข้องกับสิทธิ์ของตน ดังตารางที่ 3.1 คำอธิบายนี้ใช้แทนแผนภาพบริบทและแผนภาพการไหลของข้อมูล เพื่อให้เอกสารฉบับนี้ไม่มีรูปภาพตามขอบเขตที่กำหนด")
    add_caption(doc, "ตารางที่ 3.1 การไหลของข้อมูลตามผู้เกี่ยวข้อง")
    add_table(
        doc,
        ["ผู้เกี่ยวข้อง", "ข้อมูลที่ส่งเข้าสู่ระบบ", "ข้อมูลที่ระบบแสดงหรือส่งกลับ"],
        [
            ("OWNER", "ข้อมูลสินค้าและการปรับสต็อก", "ภาพรวมสินค้า สินค้าสต็อกน้อย คำสั่งซื้อ และพิกัดการนำส่ง"),
            ("ADMIN", "รายการเพิ่มหรือลดจำนวนสต็อก", "รายการสินค้าและจำนวนสต็อกปัจจุบัน"),
            ("EMPLOYEE", "การเปลี่ยนสถานะคำสั่งซื้อและพิกัด GPS", "รายการงานที่รับผิดชอบ สถานะ และรายละเอียดคำสั่งซื้อ"),
            ("RETAILER", "ข้อมูลสมัครสมาชิกและรายการสินค้าที่เลือก", "รายการสินค้า ยอดคำสั่งซื้อ สถานะ และพิกัดของคำสั่งซื้อของตน"),
        ],
        [2.45, 5.35, 6.85],
    )

    add_heading(doc, "3.1.2 ลำดับการทำงานของคำสั่งซื้อ", level=2)
    add_body(doc, "ผู้ค้าปลีกลงทะเบียนหรือเข้าสู่ระบบ แล้วเลือกสินค้าในหน่วยกล่องหรือแพ็ค ระบบเก็บตะกร้าไว้ในเบราว์เซอร์และสร้างคำสั่งซื้อเมื่อผู้ใช้ยืนยัน จากนั้นพนักงานสามารถรับงานตามลำดับสถานะ PENDING ไปเป็น PACKED, DELIVERING และ COMPLETED ระบบไม่อนุญาตให้ข้ามขั้นตอนหรือเปลี่ยนสถานะซ้ำ เพื่อช่วยลดความคลาดเคลื่อนของการทำงาน")
    add_body(doc, "เมื่อพนักงานเปลี่ยนสถานะเป็น DELIVERING ระบบตรวจสอบจำนวนสินค้าที่เกี่ยวข้องและตัดสต็อกภายใน transaction เดียว หากสินค้าไม่เพียงพอ ระบบจะไม่เริ่มการนำส่ง สำหรับการติดตามตำแหน่ง พนักงานต้องเริ่มแชร์ Location ด้วยตนเองในช่วง DELIVERING และระบบบันทึกพิกัดล่าสุดของคำสั่งซื้อนั้น")

    add_heading(doc, "3.1.3 โครงสร้างฐานข้อมูล", level=2)
    add_body(doc, "ฐานข้อมูลใช้ UUID เป็นรหัสหลักของข้อมูลส่วนใหญ่ ตาราง profiles เก็บข้อมูลบัญชีและบทบาท ตาราง products เก็บข้อมูลสินค้า ตาราง orders และ order_items เก็บคำสั่งซื้อกับรายละเอียดสินค้า ส่วน delivery_tracking เก็บตำแหน่งล่าสุดของพนักงานต่อคำสั่งซื้อ ตาราง app_credentials เป็นตารางเสริมสำหรับ Custom Auth และไม่เก็บรหัสผ่านจริง")
    add_caption(doc, "ตารางที่ 3.2 พจนานุกรมข้อมูล: profiles และ app_credentials")
    add_table(
        doc,
        ["ตาราง/ฟิลด์", "ชนิดข้อมูล", "คำอธิบาย", "เงื่อนไขสำคัญ"],
        [
            ("profiles.id", "UUID", "รหัสผู้ใช้งาน", "Primary key"),
            ("profiles.phone", "text", "เบอร์โทรศัพท์ของผู้ใช้", "ไม่ซ้ำกัน"),
            ("profiles.shop_name", "text", "ชื่อร้านหรือชื่อแสดงของผู้ใช้", "ต้องระบุ"),
            ("profiles.role", "app_role", "บทบาท OWNER, ADMIN, EMPLOYEE หรือ RETAILER", "ค่าเริ่มต้น RETAILER"),
            ("profiles.auth_version", "integer", "รุ่นของข้อมูลยืนยันตัวตน", "ใช้ยกเลิก session เดิมเมื่อรีเซ็ตรหัสผ่าน"),
            ("app_credentials.profile_id", "UUID", "รหัสผู้ใช้ที่อ้างอิงจาก profiles", "Primary key / Foreign key"),
            ("app_credentials.password_salt", "text", "salt สำหรับสร้าง hash", "ไม่เก็บรหัสผ่านจริง"),
            ("app_credentials.password_hash", "text", "ค่า scrypt hash ของรหัสผ่าน", "ไม่เก็บรหัสผ่านจริง"),
        ],
        [3.5, 2.0, 5.2, 3.95],
    )
    add_caption(doc, "ตารางที่ 3.3 พจนานุกรมข้อมูล: products")
    add_table(
        doc,
        ["ฟิลด์", "ชนิดข้อมูล", "คำอธิบาย", "เงื่อนไขสำคัญ"],
        [
            ("id", "UUID", "รหัสสินค้า", "Primary key"),
            ("name", "text", "ชื่อสินค้า", "ต้องระบุและไม่เป็นค่าว่าง"),
            ("price_box", "numeric(12,2)", "ราคาต่อกล่อง", "ต้องไม่น้อยกว่า 0"),
            ("price_pack", "numeric(12,2)", "ราคาต่อแพ็ค", "ระบุได้หรือเว้นว่างได้"),
            ("stock", "integer", "จำนวนสต็อก", "ค่าเริ่มต้น 0 และไม่ติดลบ"),
            ("created_at", "timestamptz", "วันที่และเวลาที่สร้างข้อมูล", "ระบบกำหนดอัตโนมัติ"),
            ("created_by", "UUID", "ผู้สร้างรายการสินค้า", "อ้างอิง profiles เมื่อมีข้อมูล"),
        ],
        [3.1, 2.45, 5.2, 3.9],
    )
    add_caption(doc, "ตารางที่ 3.4 พจนานุกรมข้อมูล: orders และ order_items")
    add_table(
        doc,
        ["ตาราง/ฟิลด์", "ชนิดข้อมูล", "คำอธิบาย", "เงื่อนไขสำคัญ"],
        [
            ("orders.id", "UUID", "รหัสคำสั่งซื้อ", "Primary key"),
            ("orders.retailer_id", "UUID", "ผู้ค้าปลีกเจ้าของคำสั่งซื้อ", "Foreign key ไป profiles"),
            ("orders.assigned_employee_id", "UUID", "พนักงานที่รับผิดชอบนำส่ง", "ระบุเมื่อรับงาน"),
            ("orders.status", "order_status", "สถานะ PENDING, PACKED, DELIVERING หรือ COMPLETED", "ค่าเริ่มต้น PENDING"),
            ("orders.total_amount", "numeric(12,2)", "ยอดรวมของคำสั่งซื้อ", "ต้องไม่น้อยกว่า 0"),
            ("order_items.product_id", "UUID", "สินค้าที่อยู่ในคำสั่งซื้อ", "อ้างอิง products"),
            ("order_items.quantity_box", "integer", "จำนวนที่สั่งเป็นกล่อง", "ไม่น้อยกว่า 0"),
            ("order_items.quantity_pack", "integer", "จำนวนที่สั่งเป็นแพ็ค", "ไม่น้อยกว่า 0"),
            ("order_items.line_total", "numeric(12,2)", "ยอดรวมของสินค้าแต่ละรายการ", "คำนวณจากจำนวนและราคา"),
        ],
        [3.5, 2.0, 5.2, 3.95],
    )
    add_caption(doc, "ตารางที่ 3.5 พจนานุกรมข้อมูล: delivery_tracking")
    add_table(
        doc,
        ["ฟิลด์", "ชนิดข้อมูล", "คำอธิบาย", "เงื่อนไขสำคัญ"],
        [
            ("order_id", "UUID", "คำสั่งซื้อที่กำลังติดตาม", "Primary key / Foreign key ไป orders"),
            ("employee_id", "UUID", "พนักงานที่ส่งพิกัด", "Foreign key ไป profiles"),
            ("latitude", "double precision", "ละติจูดล่าสุด", "อยู่ระหว่าง -90 ถึง 90"),
            ("longitude", "double precision", "ลองจิจูดล่าสุด", "อยู่ระหว่าง -180 ถึง 180"),
            ("updated_at", "timestamptz", "เวลาที่บันทึกพิกัดล่าสุด", "ระบบกำหนดอัตโนมัติ"),
        ],
        [3.1, 2.45, 5.2, 3.9],
    )

    add_heading(doc, "3.2 การออกแบบส่วนติดต่อผู้ใช้")
    add_body(doc, "เอกสารฉบับนี้อธิบายส่วนติดต่อผู้ใช้เป็นข้อความแทนภาพหน้าจอตามขอบเขตที่กำหนด การออกแบบหน้าจอเน้นให้ผู้ใช้เห็นเฉพาะข้อมูลและปุ่มคำสั่งที่สัมพันธ์กับบทบาทของตน เพื่อลดความซับซ้อนและช่วยให้เส้นทางการทำงานชัดเจน")
    add_heading(doc, "3.2.1 หน้าจอ OWNER", level=2)
    add_body(doc, "หน้าจอเจ้าของร้านแสดงภาพรวมสินค้าและรายการสต็อกน้อย รองรับการค้นหา เพิ่ม ปรับ และลบสินค้า เจ้าของร้านยังสามารถดูคำสั่งซื้อทั้งหมด พร้อมสถานะและข้อมูลตำแหน่งล่าสุดของงานที่กำลังนำส่งได้")
    add_heading(doc, "3.2.2 หน้าจอ ADMIN", level=2)
    add_body(doc, "หน้าจอผู้ดูแลระบบออกแบบสำหรับตรวจสอบรายการสินค้าและปรับจำนวนสต็อก จึงแยกสิทธิ์ออกจากหน้าจอเจ้าของร้านเพื่อให้การเข้าถึงข้อมูลสอดคล้องกับหน้าที่การทำงาน")
    add_heading(doc, "3.2.3 หน้าจอ EMPLOYEE", level=2)
    add_body(doc, "หน้าจอพนักงานแสดงร้านค้าและคำสั่งซื้อที่เกี่ยวข้องกับการจัดส่ง พนักงานเลือกดำเนินการตามลำดับสถานะของคำสั่งซื้อ และในช่วงที่สถานะเป็น DELIVERING สามารถเริ่มหรือหยุดการแชร์ตำแหน่งจริงจาก GPS ได้")
    add_heading(doc, "3.2.4 หน้าจอ RETAILER", level=2)
    add_body(doc, "หน้าจอผู้ค้าปลีกมีส่วนสำหรับสมัครสมาชิก เข้าสู่ระบบ เลือกสินค้า ระบุจำนวนกล่องหรือแพ็ค และยืนยันคำสั่งซื้อ หลังยืนยันแล้ว ผู้ค้าปลีกดูรายละเอียดคำสั่งซื้อของร้านตนเอง รวมถึงสถานะและตำแหน่งล่าสุดของการจัดส่งได้")

    add_heading(doc, "3.3 การนำระบบขึ้นใช้งาน")
    add_body(doc, "ซอร์สโค้ดของระบบจัดเก็บด้วย Git และเผยแพร่ใน GitHub repository เพื่อควบคุมเวอร์ชันและใช้เป็นต้นทางสำหรับการนำขึ้นใช้งาน เมื่อเชื่อม repository กับ Vercel ผู้พัฒนากำหนดค่าการ build ของ Next.js และตั้งค่าตัวแปรสภาพแวดล้อมในหน้า Project Settings ของ Vercel โดยไม่ใส่ค่าดังกล่าวไว้ในซอร์สโค้ดหรือ commit ขึ้น GitHub")
    add_body(doc, "ตัวแปรที่ต้องกำหนดอย่างน้อยประกอบด้วย NEXT_PUBLIC_SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY และ CUSTOM_AUTH_SESSION_SECRET โดย service role key และ session secret เป็นข้อมูลลับที่ต้องใช้เฉพาะฝั่งเซิร์ฟเวอร์ Vercel จะสร้าง deployment ตามการเปลี่ยนแปลงของ branch ที่เชื่อมไว้ และการใช้งานจริงผ่าน HTTPS ทำให้เบราว์เซอร์สามารถขอสิทธิ์ GPS สำหรับฟังก์ชันติดตามการจัดส่งได้")

    add_heading(doc, "3.4 แนวทางการทดสอบระบบ")
    add_body(doc, "การทดสอบระบบกำหนดกรณีทดสอบตามขั้นตอนที่มีผลต่อสิทธิ์และข้อมูลสำคัญ โดยเน้นว่าผู้ใช้ต้องเข้าถึงเฉพาะข้อมูลที่เกี่ยวข้อง การเปลี่ยนสถานะต้องเป็นไปตามลำดับ และการตัดสต็อกต้องไม่ทำให้จำนวนคงเหลือติดลบ")
    add_caption(doc, "ตารางที่ 3.6 ตัวอย่างกรณีทดสอบเชิงหน้าที่")
    add_table(
        doc,
        ["กรณีทดสอบ", "เงื่อนไขที่ตรวจสอบ", "ผลที่คาดหวัง"],
        [
            ("การเข้าสู่ระบบ", "เบอร์โทรศัพท์และรหัสผ่านถูกต้องตามบทบาท", "นำผู้ใช้ไปยัง dashboard ที่ตรงกับสิทธิ์"),
            ("การสร้างคำสั่งซื้อ", "RETAILER เลือกสินค้าอย่างน้อยหนึ่งรายการและระบุจำนวนถูกต้อง", "ระบบสร้าง orders และ order_items พร้อมคำนวณยอดรวม"),
            ("การเปลี่ยนสถานะ", "EMPLOYEE เปลี่ยนจาก PENDING ไป PACKED แล้ว DELIVERING", "ระบบรับเฉพาะลำดับที่ถูกต้องและผูกพนักงานกับงาน"),
            ("การตัดสต็อก", "เริ่ม DELIVERING โดยมีสินค้าคงเหลือเพียงพอ", "ตัดสต็อกใน transaction เดียว; หากไม่พอ ระบบไม่เริ่มนำส่ง"),
            ("การแชร์ตำแหน่ง", "EMPLOYEE แชร์ GPS ของงาน DELIVERING ที่รับผิดชอบ", "บันทึกพิกัดล่าสุด และ OWNER/RETAILER ที่เกี่ยวข้องดูได้"),
        ],
        [3.1, 5.7, 5.85],
    )
    add_body(doc, "ผลการประเมินประสิทธิภาพและความพึงพอใจควรเก็บจากผู้ใช้จริงด้วยแบบประเมินในระยะต่อไป เพื่อให้สอดคล้องกับวัตถุประสงค์ของโครงงาน โดยเอกสารฉบับนี้ไม่ระบุค่าผลประเมินที่ยังไม่มีข้อมูลรองรับ")

    # Chapter 4
    add_chapter(doc, 4, "ผลการดำเนินงาน")
    add_body(doc, "บทนี้นำเสนอผลการพัฒนาที่ตรวจสอบได้จากซอร์สโค้ด โครงสร้างฐานข้อมูล และผลการตรวจสอบคำสั่งของโครงการ โดยเชื่อมโยงกับวิธีดำเนินงานในบทที่ 3 เอกสารฉบับนี้อธิบายผลด้วยข้อความและตารางตามขอบเขตที่ไม่ใช้รูปภาพ ทั้งนี้ผลการใช้งานจริงกับบัญชี Supabase และผู้ใช้ภาคสนามจะแยกให้เห็นชัดเจนจากผลการตรวจสอบซอร์สโค้ด เพื่อไม่ให้สรุปผลเกินหลักฐานที่มีอยู่")

    add_heading(doc, "4.1 ผลการพัฒนาระบบ")
    add_body(doc, "ผลการพัฒนาพบว่าระบบจัดการร้าน 123พาณิชย์ปลีกส่งถูกจัดทำเป็นเว็บแอปพลิเคชันที่แยกหน้าที่และสิทธิ์ผู้ใช้ออกเป็น 4 บทบาท ได้แก่ OWNER, ADMIN, EMPLOYEE และ RETAILER การออกแบบดังกล่าวเชื่อมการจัดการสินค้า คำสั่งซื้อ สถานะงาน และตำแหน่งการนำส่งไว้ในระบบเดียวกัน โดยรายละเอียดสรุปได้ดังตารางที่ 4.1")
    add_caption(doc, "ตารางที่ 4.1 ผลการพัฒนาฟังก์ชันสำคัญของระบบ")
    add_table(
        doc,
        ["ส่วนการทำงาน", "ผลการพัฒนาที่ตรวจสอบได้จากโครงสร้างระบบ"],
        [
            ("การยืนยันตัวตนและสิทธิ์ผู้ใช้", "รองรับการเข้าสู่ระบบด้วยเบอร์โทรศัพท์และรหัสผ่านแบบ Custom Auth รหัสผ่านเก็บเป็น salt และ scrypt hash พร้อม session ใน HTTP-only cookie ผู้ค้าปลีกสมัครบัญชีได้ ส่วนบัญชีบุคลากรจัดเตรียมโดยผู้ดูแล"),
            ("สินค้าและสต็อก", "OWNER เพิ่ม ปรับ และลบสินค้าได้ ขณะที่ ADMIN ปรับจำนวนสต็อกตามสิทธิ์ สินค้ามีราคาแบบกล่องและราคาแบบแพ็คที่ระบุได้ตามข้อมูลสินค้า"),
            ("คำสั่งซื้อและการนำส่ง", "RETAILER เลือกสินค้าในตะกร้าบนเบราว์เซอร์และยืนยันคำสั่งซื้อ จากนั้น EMPLOYEE ดำเนินงานตามสถานะ PENDING, PACKED, DELIVERING และ COMPLETED โดยไม่ข้ามลำดับ"),
            ("การตัดสต็อก", "เมื่อเริ่มสถานะ DELIVERING ระบบตรวจสอบสต็อกและตัดจำนวนผ่าน transaction เดียว หากสินค้าไม่เพียงพอ ระบบจะไม่เริ่มการนำส่ง"),
            ("การติดตามตำแหน่ง", "EMPLOYEE เริ่มแชร์ตำแหน่งจากเบราว์เซอร์ได้เฉพาะระหว่าง DELIVERING ระบบเก็บพิกัดล่าสุดของคำสั่งซื้อ และ OWNER หรือ RETAILER ที่เกี่ยวข้องเรียกดูข้อมูลได้ตามสิทธิ์"),
            ("ฐานข้อมูลและการนำขึ้นใช้งาน", "ข้อมูลหลักจัดเก็บใน Supabase PostgreSQL ผ่าน REST/RPC ของ Next.js ซอร์สโค้ดมี GitHub repository เป็นแหล่งควบคุมเวอร์ชัน และเตรียมขั้นตอนเชื่อมต่อ Vercel โดยตั้งค่าความลับในตัวแปรสภาพแวดล้อมแยกจาก repository"),
        ],
        [3.55, 11.10],
    )
    add_body(doc, "โครงสร้างข้อมูลที่นำมาใช้จริงประกอบด้วย profiles, products, orders, order_items และ delivery_tracking โดยมี app_credentials เป็นตารางเสริมสำหรับข้อมูลยืนยันตัวตนแบบกำหนดเอง ระบบไม่ประกอบด้วยโมดูลรับชำระเงิน ออกใบเสร็จ ระบบบัญชี หรือตารางตะกร้าสินค้าในฐานข้อมูล จึงกำหนดขอบเขตการทำงานเป็นการยืนยันคำสั่งซื้อและการจัดส่งสินค้า")

    add_heading(doc, "4.2 ผลการทดสอบระบบ")
    add_heading(doc, "4.2.1 ผลการตรวจสอบซอร์สโค้ด", level=2)
    add_body(doc, "ผู้จัดทำตรวจสอบโครงการด้วยคำสั่งมาตรฐานของ repository ณ วันที่ 2 กันยายน พ.ศ. 2569 ผลการตรวจสอบแสดงในตารางที่ 4.2 โดยเป็นหลักฐานระดับซอร์สโค้ด ไม่ใช่ผลยืนยันการทำงานกับข้อมูลจริงบนบริการภายนอก")
    add_caption(doc, "ตารางที่ 4.2 ผลการตรวจสอบคุณภาพซอร์สโค้ด")
    add_table(
        doc,
        ["รายการตรวจสอบ", "เกณฑ์ที่คาดหวัง", "ผลที่ตรวจพบ", "สถานะ"],
        [
            ("npm run lint", "คำสั่ง ESLint ทำงานครบและไม่รายงานข้อผิดพลาด", "คำสั่งทำงานเสร็จสมบูรณ์โดยไม่พบ ESLint error", "ผ่าน"),
            ("ตรวจ TypeScript", "ใช้คำสั่ง tsc แบบ no emit และปิด incremental เพื่อตรวจชนิดข้อมูลแบบ strict", "คำสั่งทำงานเสร็จสมบูรณ์และไม่พบ TypeScript error", "ผ่าน"),
            ("npm run build", "โครงการสร้างรุ่น production ได้เมื่อกำหนดค่าระบบครบ", "ยังไม่มีบันทึกผลการทดสอบคำสั่งนี้ในหลักฐานที่ตรวจสอบ", "รอทดสอบ"),
            ("ชุดทดสอบเชิงหน้าที่อัตโนมัติ", "มี script หรือไฟล์ทดสอบที่บันทึกผลการทดสอบได้", "ไม่พบ script หรือไฟล์ทดสอบเชิงหน้าที่ของโครงการใน repository", "ยังสรุปผลไม่ได้"),
        ],
        [3.10, 4.05, 5.00, 2.50],
        [WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.CENTER],
    )
    add_body(doc, "ผลในตารางที่ 4.2 แสดงว่าโค้ดส่วนที่อยู่ใน repository ผ่านการตรวจรูปแบบและการตรวจชนิดข้อมูลได้ แต่ยังไม่มีหลักฐานผลการสร้างรุ่น production อย่างไรก็ตาม การตรวจดังกล่าวไม่ได้แทนการทดสอบการเข้าสู่ระบบ การเรียก RPC ของ Supabase การอนุญาต GPS หรือสิทธิ์ของผู้ใช้บนข้อมูลจริง จึงต้องบันทึกผลทดสอบเชิงหน้าที่แยกต่างหาก")

    add_heading(doc, "4.2.2 กรณีทดสอบเชิงหน้าที่สำหรับสภาพแวดล้อมจริง", level=2)
    add_body(doc, "เพื่อให้การประเมินมีความครบถ้วน ได้กำหนดกรณีทดสอบตามความสามารถของระบบในตารางที่ 4.3 แต่ใน repository ที่ตรวจสอบยังไม่มีบันทึกผลการทดสอบกับบัญชี Supabase ทดสอบหรือผู้ใช้จริง จึงระบุสถานะตามข้อเท็จจริงว่าเป็นรายการรอทดสอบ ไม่กำหนดผลผ่านหรือไม่ผ่านแทนข้อมูลที่ยังไม่ได้เก็บ")
    add_caption(doc, "ตารางที่ 4.3 กรณีทดสอบเชิงหน้าที่สำหรับทดสอบกับ Supabase และผู้ใช้จริง")
    add_table(
        doc,
        ["รหัส", "กรณีทดสอบ", "ผลที่คาดหวัง", "สถานะในเอกสาร"],
        [
            ("TC‑01", "เข้าสู่ระบบด้วยข้อมูลของแต่ละบทบาท", "ผู้ใช้เข้าสู่หน้าใช้งานที่ตรงกับสิทธิ์ และเข้าถึงข้อมูลเท่าที่ได้รับอนุญาต", "รอทดสอบและบันทึกผล"),
            ("TC‑02", "ผู้ค้าปลีกยืนยันคำสั่งซื้อ", "ระบบสร้างรายการสั่งซื้อพร้อมยอดรวม โดยไม่เปิดเผยคำสั่งซื้อของร้านอื่น", "รอทดสอบและบันทึกผล"),
            ("TC‑03", "พนักงานเปลี่ยนสถานะเป็นกำลังนำส่ง", "ระบบรับเฉพาะลำดับสถานะที่ถูกต้อง และตัดสต็อกเมื่อมีสินค้าเพียงพอ", "รอทดสอบและบันทึกผล"),
            ("TC‑04", "พนักงานแชร์ตำแหน่งระหว่างนำส่ง", "ระบบรับพิกัดเฉพาะงานที่รับผิดชอบ และผู้เกี่ยวข้องเห็นพิกัดล่าสุดได้", "รอทดสอบและบันทึกผล"),
            ("TC‑05", "นำระบบขึ้น Vercel ผ่าน GitHub", "ระบบสร้างรุ่นได้เมื่อกำหนดตัวแปรสภาพแวดล้อมครบ และเก็บข้อมูลลับไว้เฉพาะใน Vercel", "รอทดสอบและบันทึกผล"),
        ],
        [1.75, 3.10, 6.15, 3.65],
    )

    add_heading(doc, "4.3 การวิเคราะห์ผล")
    add_body(doc, "เมื่อเทียบกับวัตถุประสงค์ข้อที่ 1 ระบบสามารถพัฒนาเป็นเว็บแอปพลิเคชันตามแนวทางที่กำหนดได้ โดยใช้ Next.js, React, TypeScript, Tailwind CSS และ Node.js ร่วมกับฐานข้อมูล Supabase PostgreSQL โครงสร้างการกำหนดสิทธิ์ช่วยให้บทบาทของเจ้าของร้าน ผู้ดูแลระบบ พนักงาน และผู้ค้าปลีกแยกหน้าที่กันอย่างชัดเจน")
    add_body(doc, "สำหรับวัตถุประสงค์ข้อที่ 2 ระบบมีองค์ประกอบรองรับการจัดการสินค้า คำสั่งซื้อ สต็อก และสถานะการจัดส่งครบตามขอบเขตที่ระบุไว้ การตัดสต็อกเมื่อเริ่มนำส่งช่วยให้การตัดจำนวนเกิดในจุดที่สัมพันธ์กับการปฏิบัติงานจริง แต่การใช้จำนวนสต็อกเพียงค่าเดียวสำหรับสินค้าทั้งหน่วยกล่องและแพ็คยังต้องกำหนดหน่วยนับให้ตรงกับวิธีดำเนินงานของร้านก่อนใช้งานจริง")
    add_body(doc, "สำหรับวัตถุประสงค์ข้อที่ 3 ระบบมีแนวทางและกรณีทดสอบรองรับการประเมิน แต่ยังไม่ปรากฏข้อมูลแบบประเมินความพึงพอใจหรือผลทดสอบแบบ end-to-end ในสภาพแวดล้อม Supabase และ Vercel ดังนั้น ในระยะนี้จึงสรุปได้เฉพาะผล lint และการตรวจสอบชนิดข้อมูล ไม่ควรสรุปประสิทธิภาพ ความพึงพอใจ หรือความเสถียรในการใช้งานจริงจนกว่าจะเก็บข้อมูลตามตารางที่ 4.3")

    # Chapter 5
    add_chapter(doc, 5, "สรุปผลการดำเนินงาน")
    add_heading(doc, "5.1 สรุปผล")
    add_body(doc, "โครงงานนี้พัฒนาระบบจัดการร้าน 123พาณิชย์ปลีกส่งในรูปแบบเว็บแอปพลิเคชัน เพื่อสนับสนุนการจัดการสินค้า สต็อก คำสั่งซื้อ และสถานะการนำส่งของร้าน ระบบพัฒนาด้วย Next.js, React, TypeScript, Tailwind CSS และ Node.js โดยใช้ Supabase PostgreSQL เป็นฐานข้อมูล และใช้ GitHub เพื่อควบคุมเวอร์ชันของซอร์สโค้ดสำหรับเชื่อมต่อกระบวนการนำขึ้นใช้งานบน Vercel")
    add_body(doc, "ระบบกำหนดผู้ใช้ 4 บทบาท ได้แก่ OWNER, ADMIN, EMPLOYEE และ RETAILER พร้อมกำหนดขอบเขตสิทธิ์ตามหน้าที่ ผู้ค้าปลีกเลือกสินค้าตามหน่วยกล่องหรือแพ็คและยืนยันคำสั่งซื้อ พนักงานดำเนินงานตามสถานะและแชร์ตำแหน่งได้ขณะนำส่ง การตรวจสอบซอร์สโค้ดยืนยันว่า npm run lint และการตรวจสอบชนิดข้อมูล TypeScript ผ่าน ส่วนการทดสอบเชิงหน้าที่กับ Supabase และการใช้งานจริงยังต้องบันทึกผลเพิ่มเติม")

    add_heading(doc, "5.2 อภิปรายผล")
    add_body(doc, "ผลการพัฒนาสอดคล้องกับปัญหาการจัดการข้อมูลที่กระจายอยู่หลายขั้นตอนของร้าน เนื่องจากระบบรวมข้อมูลสินค้า คำสั่งซื้อ และสถานะการจัดส่งไว้ในฐานข้อมูลกลาง การแยกบทบาทลดโอกาสที่ผู้ใช้จะเข้าถึงคำสั่งหรือข้อมูลเกินหน้าที่ ขณะที่การใช้ transaction ในช่วงเริ่มนำส่งช่วยป้องกันการตัดสต็อกบางส่วนเมื่อสินค้าไม่เพียงพอ")
    add_body(doc, "อย่างไรก็ดี ความสามารถด้าน GPS ขึ้นอยู่กับ HTTPS การอนุญาตตำแหน่งจากผู้ใช้ สัญญาณเครือข่าย และการทำงานของเบราว์เซอร์บนอุปกรณ์เคลื่อนที่ ระบบเก็บพิกัดล่าสุดต่อคำสั่งซื้อและดึงข้อมูลเป็นระยะ จึงเหมาะสำหรับติดตามสถานะการนำส่งในระดับงาน แต่ไม่ควรอ้างว่าเป็นระบบบันทึกเส้นทางย้อนหลังหรือระบบติดตามพฤติกรรมพนักงาน")
    add_body(doc, "ด้านการนำขึ้นใช้งาน GitHub มีบทบาทเป็นแหล่งซอร์สโค้ดและประวัติการเปลี่ยนแปลง ส่วน Vercel เป็นแนวทาง deployment ที่ตั้งค่าได้จาก repository การใช้งานจริงต้องกำหนดตัวแปรสภาพแวดล้อมใน Vercel ให้ครบและเก็บ service role key กับ session secret ไว้เฉพาะฝั่งเซิร์ฟเวอร์ เอกสารฉบับนี้จึงไม่ระบุว่าเว็บไซต์ถูกเผยแพร่แล้ว เนื่องจากยังไม่มีหลักฐาน URL หรือผล deployment ที่ตรวจสอบได้")

    add_heading(doc, "5.3 ปัญหาและอุปสรรค")
    add_body(doc, "ปัญหาแรกคือการตรวจสอบการทำงานแบบ end-to-end ต้องใช้โครงการ Supabase สำหรับทดสอบ ข้อมูลตัวอย่าง และบัญชีผู้ใช้ของแต่ละบทบาท หากขาดการตั้งค่าดังกล่าว จะยืนยันได้เพียงผล lint และการตรวจสอบชนิดข้อมูลของซอร์สโค้ด ไม่สามารถยืนยันผลการเรียกฐานข้อมูลและ RPC ได้ครบถ้วน")
    add_body(doc, "ปัญหาถัดมาคือการติดตามตำแหน่งจริงต้องได้รับความยินยอมจากพนักงานและอาศัยอุปกรณ์ที่รองรับ Location บน HTTPS หากผู้ใช้ปิดสิทธิ์ตำแหน่ง สัญญาณไม่เสถียร หรือเบราว์เซอร์ลดการทำงานเบื้องหลัง ระบบอาจไม่ได้รับพิกัดตามช่วงเวลาที่ต้องการ")
    add_body(doc, "นอกจากนี้ สินค้ารองรับการสั่งทั้งกล่องและแพ็ค แต่จำนวนสต็อกเป็นค่าจำนวนเต็มเดียว การนำไปใช้กับสินค้าที่มีอัตราแปลงกล่องต่อแพ็คต่างกันจำเป็นต้องกำหนดหน่วยนับและกติกาการตัดสต็อกให้ชัดเจน หากลบสินค้าที่เคยอยู่ในคำสั่งซื้อ ระบบจะเก็บรายการเดิมโดยไม่ผูก product_id ไว้ จึงต้องกำหนดแนวทางดูแลสินค้าที่เลิกขายอย่างรอบคอบ ส่วนการนำระบบขึ้น Vercel ยังต้องตั้งค่าตัวแปรสภาพแวดล้อมด้วยผู้มีสิทธิ์ จึงยังไม่ควรระบุว่า deployment สำเร็จจนกว่าจะมีผลตรวจสอบจริง")

    add_heading(doc, "5.4 ข้อเสนอแนะ")
    add_body(doc, "ควรจัดเตรียม Supabase project สำหรับทดสอบ แยกจากข้อมูลใช้งานจริง พร้อมสร้างบัญชี OWNER, ADMIN, EMPLOYEE และ RETAILER สำหรับบันทึกผลตามตารางที่ 4.3 หลังจากนั้นควรเก็บแบบประเมินความพึงพอใจจากผู้ใช้งานที่เกี่ยวข้อง และสรุปผลด้วยจำนวนกลุ่มตัวอย่าง เกณฑ์การประเมิน และค่าเฉลี่ยที่ตรวจสอบย้อนกลับได้")
    add_body(doc, "ควรพัฒนาการจัดการหน่วยสินค้าให้รองรับอัตราแปลงระหว่างกล่องและแพ็ค รวมทั้งพิจารณาการจองสต็อกก่อนเริ่มนำส่งหากรูปแบบธุรกิจต้องการลดความเสี่ยงจากการสั่งซื้อเกินจำนวนคงเหลือ และควรใช้การปิดการขายหรือ archive แทนการลบสินค้าที่มีคำสั่งซื้อค้างอยู่ การเปลี่ยนแปลงดังกล่าวควรออกแบบร่วมกับผู้ใช้งานร้านเพื่อให้สอดคล้องกับหน่วยซื้อขายจริง")
    add_body(doc, "สำหรับการใช้งานออนไลน์ ควรกำหนดขั้นตอนตรวจสอบ deployment บน Vercel และตั้งค่า secret ผ่าน Project Settings เท่านั้น พร้อมเพิ่มการทดสอบอัตโนมัติใน GitHub เพื่อให้ตรวจสอบ lint, build และกรณีทดสอบสำคัญก่อนเผยแพร่แต่ละครั้ง เอกสาร .env.example และคู่มือติดตั้งควรระบุชื่อตัวแปรที่จำเป็นให้ครบโดยไม่ใส่ค่าความลับจริง ด้านตำแหน่ง GPS ควรกำหนดข้อความขอความยินยอม ระยะเวลาจัดเก็บข้อมูล และผู้มีสิทธิ์เข้าดูให้ชัดเจน รวมทั้งแสดงหน้าสถานะรอพิกัดให้แตกต่างจากข้อมูลตำแหน่งจริง เพื่อให้การใช้งานเคารพความเป็นส่วนตัวของพนักงานและผู้ค้าปลีก")

    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(OUTPUT)
    print("DOCX generated in output directory")


if __name__ == "__main__":
    build_document()
