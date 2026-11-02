# -*- coding: utf-8 -*-
"""
Helper utility module for generating standardized, professionally styled DOCX documents
for authentic English SMA teaching materials (Modul Ajar, LKPD, and Sumatif).
"""

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

COLOR_NAVY = RGBColor(15, 41, 66)      # #0F2942
COLOR_ROYAL = RGBColor(30, 64, 175)    # #1E40AF
COLOR_SLATE = RGBColor(71, 85, 105)    # #475569
COLOR_BODY = RGBColor(30, 41, 59)      # #1E293B
COLOR_MUTED = RGBColor(100, 116, 139)  # #64748B

HEX_NAVY = "0F2942"
HEX_ROYAL = "1E40AF"
HEX_LIGHT_BLUE = "EFF6FF"
HEX_BORDER = "CBD5E1"
HEX_ZEBRA = "F8FAFC"
HEX_AMBER = "FEF3C7"
HEX_GREEN = "DCFCE7"

def create_styled_doc():
    doc = docx.Document()
    
    # Page Setup
    for section in doc.sections:
        section.top_margin = Inches(0.9)
        section.bottom_margin = Inches(0.9)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.different_first_page_header_footer = False
        
        # Footer with page number
        footer = section.footer
        p_foot = footer.paragraphs[0]
        p_foot.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_foot.text = "Bank Materi Otentik Bahasa Inggris SMA | Kurikulum Merdeka 2025"
        p_foot.style.font.name = "Segoe UI"
        p_foot.style.font.size = Pt(8.5)
        p_foot.style.font.color.rgb = COLOR_MUTED

    # Normal Style configuration
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Segoe UI'
    style_normal.font.size = Pt(10)
    style_normal.font.color.rgb = COLOR_BODY
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(4)
    style_normal.paragraph_format.space_before = Pt(0)
    
    return doc

def set_cell_shading(cell, color_hex):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def set_table_borders(table, color="CBD5E1", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def add_header_banner(doc, title, subtitle):
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_t = p_title.add_run(title)
    run_t.font.name = "Segoe UI"
    run_t.font.size = Pt(16)
    run_t.font.bold = True
    run_t.font.color.rgb = COLOR_NAVY

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(12)
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_s = p_sub.add_run(subtitle)
    run_s.font.name = "Segoe UI"
    run_s.font.size = Pt(10)
    run_s.font.italic = True
    run_s.font.color.rgb = COLOR_ROYAL

def add_tp_info_box(doc, info_dict):
    table = doc.add_table(rows=len(info_dict), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table, color="93C5FD", sz="6")

    widths = [Inches(2.0), Inches(4.5)]
    for i, (k, v) in enumerate(info_dict.items()):
        row = table.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = widths[0], widths[1]
        
        c0.text = k
        c1.text = str(v)
        
        set_cell_shading(c0, "DBEAFE")
        set_cell_shading(c1, "F8FAFC" if i % 2 == 1 else "FFFFFF")
        
        set_cell_margins(c0, 80, 80, 120, 120)
        set_cell_margins(c1, 80, 80, 120, 120)
        
        p0 = c0.paragraphs[0]
        p0.runs[0].font.name = "Segoe UI"
        p0.runs[0].font.size = Pt(9.5)
        p0.runs[0].font.bold = True
        p0.runs[0].font.color.rgb = COLOR_NAVY
        p0.paragraph_format.space_after = Pt(0)
        
        p1 = c1.paragraphs[0]
        if p1.runs:
            p1.runs[0].font.name = "Segoe UI"
            p1.runs[0].font.size = Pt(9.5)
            p1.runs[0].font.color.rgb = COLOR_BODY
        p1.paragraph_format.space_after = Pt(0)

    p_spacer = doc.add_paragraph()
    p_spacer.paragraph_format.space_before = Pt(6)
    p_spacer.paragraph_format.space_after = Pt(6)

def add_heading_1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.name = "Segoe UI"
    run.font.size = Pt(12.5)
    run.font.bold = True
    run.font.color.rgb = COLOR_NAVY

def add_heading_2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.font.name = "Segoe UI"
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = COLOR_ROYAL

def add_callout(doc, title, text, bg="EFF6FF", border_color="1E40AF"):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = tbl.cell(0, 0)
    c.width = Inches(6.5)
    set_cell_shading(c, bg)
    set_cell_margins(c, 100, 100, 150, 150)
    
    # Left border only
    tcPr = c._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="none"/>'
        f'<w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/>'
        f'<w:bottom w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)
    
    p = c.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    r_t = p.add_run(title + "\n")
    r_t.font.name = "Segoe UI"
    r_t.font.size = Pt(10)
    r_t.font.bold = True
    r_t.font.color.rgb = COLOR_ROYAL
    
    r_b = p.add_run(text)
    r_b.font.name = "Segoe UI"
    r_b.font.size = Pt(9.5)
    r_b.font.color.rgb = COLOR_BODY

def add_authentic_text_block(doc, text_num, title, function_role, source_info, source_year="", content_paragraphs=None, analysis_dict=None):
    add_heading_2(doc, f"Materi Otentik {text_num}: {title} ({function_role})")
    
    # Check if content_paragraphs was passed into source_year position
    if isinstance(source_year, (list, tuple)):
        analysis_dict = content_paragraphs
        content_paragraphs = source_year
        if isinstance(source_info, dict):
            source_year = source_info.get("year", "2024")
        else:
            source_year = "2024"
    
    if content_paragraphs is None:
        content_paragraphs = []

    # Handle dictionary or string for source_info
    if isinstance(source_info, dict):
        src_portal = source_info.get("portal", "Situs Web Resmi / Publikasi Kredibel")
        src_title = source_info.get("title", title)
        src_url = source_info.get("url", "https://...")
        src_year = source_info.get("year", source_year or "2024")
        src_cefr = source_info.get("cefr", "Sesuai Kurikulum Merdeka SMA")
    else:
        src_portal = str(source_info)
        src_title = title
        if str(source_year).startswith("http"):
            src_url = str(source_year)
            src_year = "2024"
        else:
            src_url = f"https://www.google.com/search?q={src_portal.replace(' ', '+')}"
            src_year = str(source_year) if source_year else "2024"
        src_cefr = "Sesuai Kurikulum Merdeka SMA"

    # Dedicated Source Attribution Box
    tbl_src = doc.add_table(rows=1, cols=1)
    tbl_src.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_src = tbl_src.cell(0, 0)
    c_src.width = Inches(6.5)
    set_cell_shading(c_src, "EFF6FF")
    set_cell_margins(c_src, 70, 70, 120, 120)
    
    tcPr_src = c_src._tc.get_or_add_tcPr()
    borders_src = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="6" w:space="0" w:color="93C5FD"/>'
        f'<w:left w:val="single" w:sz="18" w:space="0" w:color="1E40AF"/>'
        f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="93C5FD"/>'
        f'<w:right w:val="single" w:sz="6" w:space="0" w:color="93C5FD"/>'
        f'</w:tcBorders>'
    )
    tcPr_src.append(borders_src)
    
    p_s = c_src.paragraphs[0]
    p_s.paragraph_format.space_before = Pt(0)
    p_s.paragraph_format.space_after = Pt(2)
    p_s.paragraph_format.line_spacing = 1.15
    
    r_hdr = p_s.add_run("IDENTIFIKASI SUMBER OTENTIK WEB (CREDIBLE WEB SOURCE):\n")
    r_hdr.font.name = "Segoe UI"
    r_hdr.font.size = Pt(8.5)
    r_hdr.font.bold = True
    r_hdr.font.color.rgb = COLOR_ROYAL
    
    r_meta1 = p_s.add_run(f"• Media / Penerbit: {src_portal} (Tahun: {src_year})\n• Judul Asli: \"{src_title}\"\n")
    r_meta1.font.name = "Segoe UI"
    r_meta1.font.size = Pt(8.5)
    r_meta1.font.color.rgb = COLOR_BODY
    
    r_meta2 = p_s.add_run(f"• Tautan URL Resmi: {src_url}\n• Target Tingkat Kemahiran: {src_cefr} | Peruntukan: {function_role}")
    r_meta2.font.name = "Segoe UI"
    r_meta2.font.size = Pt(8.5)
    r_meta2.font.italic = True
    r_meta2.font.color.rgb = COLOR_SLATE

    p_gap = doc.add_paragraph()
    p_gap.paragraph_format.space_before = Pt(0)
    p_gap.paragraph_format.space_after = Pt(2)

    # Text box
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = tbl.cell(0, 0)
    c.width = Inches(6.5)
    set_cell_shading(c, "F8FAFC")
    set_cell_margins(c, 120, 120, 160, 160)
    
    tcPr = c._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>'
        f'<w:left w:val="single" w:sz="18" w:space="0" w:color="1E40AF"/>'
        f'<w:bottom w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>'
        f'<w:right w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)
    
    # First paragraph in cell
    p_first = c.paragraphs[0]
    p_first.paragraph_format.space_after = Pt(5)
    p_first.paragraph_format.line_spacing = 1.15
    r_h = p_first.add_run(f"--- [READING PASSAGE: {title.upper()}] ---\n\n")
    r_h.font.name = "Segoe UI"
    r_h.font.size = Pt(9.5)
    r_h.font.bold = True
    r_h.font.color.rgb = COLOR_NAVY
    
    for idx, p_text in enumerate(content_paragraphs):
        if idx == 0:
            p = p_first
        else:
            p = c.add_paragraph()
            p.paragraph_format.space_after = Pt(5)
            p.paragraph_format.line_spacing = 1.15
        
        r_body = p.add_run(p_text)
        r_body.font.name = "Georgia"
        r_body.font.size = Pt(10)
        r_body.font.color.rgb = COLOR_BODY

    # Structural / Linguistic Analysis Note
    if analysis_dict:
        p_note_hdr = doc.add_paragraph()
        p_note_hdr.paragraph_format.space_before = Pt(4)
        p_note_hdr.paragraph_format.space_after = Pt(2)
        r_nh = p_note_hdr.add_run("Bedah Struktur Teks & Fitur Bahasa Otentik:")
        r_nh.font.name = "Segoe UI"
        r_nh.font.size = Pt(9.5)
        r_nh.font.bold = True
        r_nh.font.color.rgb = COLOR_ROYAL

        for k, v in analysis_dict.items():
            p_n = doc.add_paragraph()
            p_n.paragraph_format.left_indent = Inches(0.2)
            p_n.paragraph_format.space_before = Pt(0)
            p_n.paragraph_format.space_after = Pt(2)
            r_k = p_n.add_run(f"• {k}: ")
            r_k.font.name = "Segoe UI"
            r_k.font.size = Pt(9)
            r_k.font.bold = True
            r_k.font.color.rgb = COLOR_NAVY
            r_v = p_n.add_run(v)
            r_v.font.name = "Segoe UI"
            r_v.font.size = Pt(9)
            r_v.font.color.rgb = COLOR_BODY

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(4)

def add_vocab_table(doc, vocab_items):
    add_heading_2(doc, "Glosarium Kosakata & Frasa Kunci Otentik (Authentic Lexical Bank)")
    
    headers = ["Kosakata / Frasa Otentik", "Kelas Kata", "Makna Kontekstual (Indonesian)", "Contoh Penggunaan dalam Teks Otentik"]
    widths = [Inches(1.8), Inches(0.9), Inches(1.8), Inches(2.0)]
    
    table = doc.add_table(rows=len(vocab_items) + 1, cols=4)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table, color="CBD5E1", sz="4")
    
    # Header Row
    hdr_row = table.rows[0]
    for c_idx, h_text in enumerate(headers):
        c = hdr_row.cells[c_idx]
        c.width = widths[c_idx]
        c.text = h_text
        set_cell_shading(c, HEX_NAVY)
        set_cell_margins(c, 80, 80, 100, 100)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].font.name = "Segoe UI"
        p.runs[0].font.size = Pt(9)
        p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
        
    for r_idx, item in enumerate(vocab_items, 1):
        row = table.rows[r_idx]
        word, pos, meaning, example = item
        row_vals = [word, pos, meaning, example]
        for c_idx, val in enumerate(row_vals):
            c = row.cells[c_idx]
            c.width = widths[c_idx]
            c.text = val
            if r_idx % 2 == 1:
                set_cell_shading(c, HEX_ZEBRA)
            else:
                set_cell_shading(c, "FFFFFF")
            set_cell_margins(c, 60, 60, 100, 100)
            p = c.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if c_idx == 0:
                p.runs[0].font.name = "Segoe UI"
                p.runs[0].font.size = Pt(9)
                p.runs[0].font.bold = True
                p.runs[0].font.color.rgb = COLOR_ROYAL
            elif c_idx == 1:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.runs[0].font.name = "Segoe UI"
                p.runs[0].font.size = Pt(8.5)
                p.runs[0].font.italic = True
                p.runs[0].font.color.rgb = COLOR_SLATE
            else:
                p.runs[0].font.name = "Segoe UI"
                p.runs[0].font.size = Pt(9)
                p.runs[0].font.color.rgb = COLOR_BODY

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(4)

def add_assessment_questions(doc, literal_q, inferential_q, evaluative_q, performance_task):
    add_heading_2(doc, "Bank Soal & Tugas Pemantik Otentik (Bahan LKPD & Asesmen Sumatif)")

    p_intro = doc.add_paragraph()
    p_intro.paragraph_format.space_after = Pt(4)
    r_in = p_intro.add_run("Pertanyaan dan penugasan dirancang berjenjang dari pemahaman literal, analisis inferensial, hingga evaluasi berpikir kritis (HOTS) serta tugas kinerja produktif:")
    r_in.font.name = "Segoe UI"
    r_in.font.size = Pt(9.5)
    r_in.font.italic = True
    r_in.font.color.rgb = COLOR_MUTED

    categories = [
        ("A. Pertanyaan Pemahaman Literal (Literal Reading Comprehension)", literal_q),
        ("B. Pertanyaan Analisis & Makna Tersirat (Inferential Analysis)", inferential_q),
        ("C. Pertanyaan Berpikir Kritis & Evaluasi (Critical Evaluation / HOTS)", evaluative_q)
    ]

    for cat_title, q_list in categories:
        p_c = doc.add_paragraph()
        p_c.paragraph_format.space_before = Pt(4)
        p_c.paragraph_format.space_after = Pt(2)
        r_c = p_c.add_run(cat_title)
        r_c.font.name = "Segoe UI"
        r_c.font.size = Pt(10)
        r_c.font.bold = True
        r_c.font.color.rgb = COLOR_NAVY

        for idx, q in enumerate(q_list, 1):
            p_q = doc.add_paragraph()
            p_q.paragraph_format.left_indent = Inches(0.2)
            p_q.paragraph_format.space_before = Pt(0)
            p_q.paragraph_format.space_after = Pt(2)
            r_num = p_q.add_run(f"{idx}. ")
            r_num.font.bold = True
            r_num.font.color.rgb = COLOR_ROYAL
            r_txt = p_q.add_run(q)
            r_txt.font.name = "Segoe UI"
            r_txt.font.size = Pt(9.5)

    # Performance Task
    add_callout(doc, "D. Tugas Kinerja Otentik (Authentic Performance Task Prompt):",
                performance_task, bg="FEF3C7", border_color="D97706")

def add_styled_table(doc, headers, data_rows, col_widths=None, header_bg="1E40AF", zebra=True):
    table = doc.add_table(rows=len(data_rows) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table, color="CBD5E1", sz="4")
    
    # Header Row
    hdr_row = table.rows[0]
    for c_idx, h_text in enumerate(headers):
        cell = hdr_row.cells[c_idx]
        if col_widths and c_idx < len(col_widths):
            cell.width = col_widths[c_idx]
        cell.text = str(h_text)
        set_cell_shading(cell, header_bg)
        set_cell_margins(cell, 80, 80, 100, 100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if p.runs:
            p.runs[0].font.name = "Segoe UI"
            p.runs[0].font.size = Pt(9.5)
            p.runs[0].font.bold = True
            p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
            
    # Data Rows
    for r_idx, row_data in enumerate(data_rows):
        row = table.rows[r_idx + 1]
        bg = "F8FAFC" if (zebra and r_idx % 2 == 1) else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            if col_widths and c_idx < len(col_widths):
                cell.width = col_widths[c_idx]
            cell.text = str(val)
            set_cell_shading(cell, bg)
            set_cell_margins(cell, 60, 60, 100, 100)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            if p.runs:
                p.runs[0].font.name = "Segoe UI"
                p.runs[0].font.size = Pt(9)
                p.runs[0].font.color.rgb = COLOR_BODY
                
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(4)
    return table

