# -*- coding: utf-8 -*-
"""
Generator Modul Pegangan Siswa (Student Learning Handout) DOCX
Khusus SMA Negeri 1 Bae Kudus - Berbasis Teoretis Siswa:
1. Konsep Teori & Unsur Teks yang Relevan
2. Contoh Teks Model Otentik Teranotasi
3. Grammar (Tata Bahasa: Rumus, Penjelasan, dan Contoh)
4. Pronunciation Guide (Fonetik IPA, Penekanan Suku Kata, Arti, dan Tips Pengucapan)
"""

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Colors
COLOR_NAVY = RGBColor(27, 54, 93)      # #1B365D
COLOR_STEEL = RGBColor(36, 59, 83)     # #243B53
COLOR_ROYAL = RGBColor(30, 64, 175)    # #1E40AF
COLOR_BODY = RGBColor(30, 41, 59)      # #1E293B
COLOR_MUTED = RGBColor(100, 116, 139)  # #64748B
COLOR_WHITE = RGBColor(255, 255, 255)

HEX_NAVY = "1B365D"
HEX_STEEL = "243B53"
HEX_ROYAL = "1E40AF"
HEX_LIGHT_BLUE = "EFF6FF"
HEX_ZEBRA = "F8FAFC"
HEX_AMBER = "FEF3C7"
HEX_BORDER = "CBD5E1"

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

def create_student_doc():
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.9)
        section.right_margin = Inches(0.9)
        
        footer = section.footer
        p_foot = footer.paragraphs[0]
        p_foot.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_foot.text = "Modul Belajar Siswa Bahasa Inggris | SMA Negeri 1 Bae Kudus (Pemetaan MGMP 2025/2026)"
        p_foot.style.font.name = "Segoe UI"
        p_foot.style.font.size = Pt(8.5)
        p_foot.style.font.color.rgb = COLOR_MUTED

    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Segoe UI'
    style_normal.font.size = Pt(10)
    style_normal.font.color.rgb = COLOR_BODY
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(4)
    return doc

def add_student_header(doc, grade_title, subtitle):
    # Main Header
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_after = Pt(2)
    r_t = p_title.add_run("MODUL MATERI BELAJAR SISWA (STUDENT HANDOUT)\nSMA NEGERI 1 BAE KUDUS")
    r_t.font.name = "Segoe UI"
    r_t.font.size = Pt(14)
    r_t.font.bold = True
    r_t.font.color.rgb = COLOR_NAVY

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(8)
    r_s = p_sub.add_run(f"{grade_title} — {subtitle}\nTahun Ajaran 2025/2026 | Berdasarkan Pemetaan MGMP SMA Kabupaten Kudus")
    r_s.font.name = "Segoe UI"
    r_s.font.size = Pt(10)
    r_s.font.italic = True
    r_s.font.color.rgb = COLOR_ROYAL

    # Student Identity Box
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl, color="93C5FD", sz="6")
    
    c0, c1 = tbl.cell(0, 0), tbl.cell(0, 1)
    c0.width, c1.width = Inches(3.5), Inches(3.2)
    set_cell_shading(c0, "EFF6FF")
    set_cell_shading(c1, "EFF6FF")
    set_cell_margins(c0, 80, 80, 120, 120)
    set_cell_margins(c1, 80, 80, 120, 120)

    p0 = c0.paragraphs[0]
    p0.paragraph_format.space_after = Pt(0)
    r0 = p0.add_run("Nama Siswa : ____________________________________\nKelas / No. : ________________ / ________________")
    r0.font.name = "Segoe UI"
    r0.font.size = Pt(9.5)
    r0.font.bold = True
    r0.font.color.rgb = COLOR_NAVY

    p1 = c1.paragraphs[0]
    p1.paragraph_format.space_after = Pt(0)
    r1 = p1.add_run("Petunjuk Belajar Siswa:\nPelajari teori konsep, unsur teks, rumus tata bahasa, dan panduan pengucapan (pronunciation) sebelum latihan asesmen.")
    r1.font.name = "Segoe UI"
    r1.font.size = Pt(8.5)
    r1.font.italic = True
    r1.font.color.rgb = COLOR_BODY

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(6)

def add_heading_1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.name = "Segoe UI"
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = COLOR_NAVY

def add_heading_2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.font.name = "Segoe UI"
    run.font.size = Pt(10.5)
    run.font.bold = True
    run.font.color.rgb = COLOR_ROYAL

def add_concept_box(doc, definition, social_function):
    tbl = doc.add_table(rows=2, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl, color="CBD5E1", sz="4")
    
    widths = [Inches(1.8), Inches(4.9)]
    rows_data = [
        ("Definisi Teoretis", definition),
        ("Fungsi Sosial (Social Function)", social_function)
    ]
    for idx, (label, val) in enumerate(rows_data):
        c0, c1 = tbl.cell(idx, 0), tbl.cell(idx, 1)
        c0.width, c1.width = widths[0], widths[1]
        c0.text = label
        c1.text = val
        set_cell_shading(c0, "DBEAFE")
        set_cell_shading(c1, "F8FAFC" if idx % 2 == 1 else "FFFFFF")
        set_cell_margins(c0, 70, 70, 100, 100)
        set_cell_margins(c1, 70, 70, 100, 100)
        
        p0 = c0.paragraphs[0]
        p0.runs[0].font.name = "Segoe UI"
        p0.runs[0].font.size = Pt(9.5)
        p0.runs[0].font.bold = True
        p0.runs[0].font.color.rgb = COLOR_NAVY
        
        p1 = c1.paragraphs[0]
        p1.runs[0].font.name = "Segoe UI"
        p1.runs[0].font.size = Pt(9.5)
        p1.runs[0].font.color.rgb = COLOR_BODY

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(4)

def add_elements_table(doc, elements):
    add_heading_2(doc, "A. Unsur-Unsur Struktur Teks (Generic Structure & Text Elements)")
    tbl = doc.add_table(rows=len(elements) + 1, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl, color="CBD5E1", sz="4")
    
    headers = ["Bagian Struktur", "Fungsi & Unsur Utama", "Ciri / Cara Mengidentifikasi"]
    widths = [Inches(1.6), Inches(2.6), Inches(2.5)]
    
    hdr = tbl.rows[0]
    for c_idx, h_text in enumerate(headers):
        c = hdr.cells[c_idx]
        c.width = widths[c_idx]
        c.text = h_text
        set_cell_shading(c, HEX_NAVY)
        set_cell_margins(c, 70, 70, 90, 90)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].font.name = "Segoe UI"
        p.runs[0].font.size = Pt(9)
        p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = COLOR_WHITE
        
    for r_idx, (part, desc, tip) in enumerate(elements, 1):
        row = tbl.rows[r_idx]
        for c_idx, val in enumerate([part, desc, tip]):
            c = row.cells[c_idx]
            c.width = widths[c_idx]
            c.text = val
            set_cell_shading(c, HEX_ZEBRA if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(c, 70, 70, 90, 90)
            p = c.paragraphs[0]
            p.runs[0].font.name = "Segoe UI"
            p.runs[0].font.size = Pt(9)
            if c_idx == 0:
                p.runs[0].font.bold = True
                p.runs[0].font.color.rgb = COLOR_ROYAL
            else:
                p.runs[0].font.color.rgb = COLOR_BODY
                
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(4)

def add_annotated_model_text(doc, title, source, paragraphs_with_annotations):
    add_heading_2(doc, f"B. Contoh Teks yang Relevan (Model Text): \"{title}\"")
    
    # Source note
    p_src = doc.add_paragraph()
    p_src.paragraph_format.space_before = Pt(0)
    p_src.paragraph_format.space_after = Pt(4)
    r_src = p_src.add_run(f"Sumber / Konteks Teks: {source}")
    r_src.font.name = "Segoe UI"
    r_src.font.size = Pt(8.5)
    r_src.font.italic = True
    r_src.font.color.rgb = COLOR_MUTED

    tbl = doc.add_table(rows=len(paragraphs_with_annotations), cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl, color="CBD5E1", sz="4")
    
    widths = [Inches(1.5), Inches(5.2)]
    for idx, (annot, p_text) in enumerate(paragraphs_with_annotations):
        c_ann, c_txt = tbl.cell(idx, 0), tbl.cell(idx, 1)
        c_ann.width, c_txt.width = widths[0], widths[1]
        
        c_ann.text = annot
        c_txt.text = p_text
        
        set_cell_shading(c_ann, "EFF6FF")
        set_cell_shading(c_txt, "FFFFFF")
        set_cell_margins(c_ann, 80, 80, 100, 100)
        set_cell_margins(c_txt, 80, 80, 120, 120)
        
        p0 = c_ann.paragraphs[0]
        p0.runs[0].font.name = "Segoe UI"
        p0.runs[0].font.size = Pt(8.5)
        p0.runs[0].font.bold = True
        p0.runs[0].font.color.rgb = COLOR_ROYAL
        
        p1 = c_txt.paragraphs[0]
        p1.runs[0].font.name = "Georgia"
        p1.runs[0].font.size = Pt(9.5)
        p1.runs[0].font.color.rgb = COLOR_BODY

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(4)

def add_grammar_section(doc, grammar_points):
    add_heading_2(doc, "C. Grammar & Tata Bahasa Fokus (Language Features)")
    
    for gp in grammar_points:
        p_t = doc.add_paragraph()
        p_t.paragraph_format.space_before = Pt(6)
        p_t.paragraph_format.space_after = Pt(2)
        r_t = p_t.add_run(f"📌 {gp['rule_name']}")
        r_t.font.name = "Segoe UI"
        r_t.font.size = Pt(9.5)
        r_t.font.bold = True
        r_t.font.color.rgb = COLOR_NAVY
        
        # Formula box
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        c = tbl.cell(0, 0)
        c.width = Inches(6.7)
        set_cell_shading(c, "F1F5F9")
        set_cell_margins(c, 70, 70, 120, 120)
        
        p = c.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        r_form = p.add_run(f"Rumus / Pola: {gp['formula']}\n")
        r_form.font.name = "Segoe UI"
        r_form.font.size = Pt(9)
        r_form.font.bold = True
        r_form.font.color.rgb = COLOR_ROYAL
        
        r_desc = p.add_run(f"Penjelasan: {gp['explanation']}\n\nContoh Kalimat:\n")
        r_desc.font.name = "Segoe UI"
        r_desc.font.size = Pt(8.5)
        r_desc.font.color.rgb = COLOR_BODY
        
        for ex in gp['examples']:
            r_ex = p.add_run(f"• {ex}\n")
            r_ex.font.name = "Segoe UI"
            r_ex.font.size = Pt(8.5)
            r_ex.font.italic = True
            r_ex.font.color.rgb = COLOR_STEEL

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(4)

def add_pronunciation_section(doc, vocab_pron_list, general_pron_rules):
    add_heading_2(doc, "D. Pronunciation Guide & Speaking Practice (Pelafalan Siswa)")
    
    # Pronunciation Table
    headers = ["Kata Kunci", "Ejaan Fonetik (IPA)", "Penekanan Suku Kata (Stress)", "Arti Bahasa Indonesia"]
    widths = [Inches(1.5), Inches(1.8), Inches(1.9), Inches(1.5)]
    
    tbl = doc.add_table(rows=len(vocab_pron_list) + 1, cols=4)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl, color="CBD5E1", sz="4")
    
    hdr = tbl.rows[0]
    for c_idx, h_text in enumerate(headers):
        c = hdr.cells[c_idx]
        c.width = widths[c_idx]
        c.text = h_text
        set_cell_shading(c, HEX_STEEL)
        set_cell_margins(c, 70, 70, 90, 90)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].font.name = "Segoe UI"
        p.runs[0].font.size = Pt(8.5)
        p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = COLOR_WHITE
        
    for r_idx, item in enumerate(vocab_pron_list, 1):
        word, ipa, stress, meaning = item
        row = tbl.rows[r_idx]
        for c_idx, val in enumerate([word, ipa, stress, meaning]):
            c = row.cells[c_idx]
            c.width = widths[c_idx]
            c.text = val
            set_cell_shading(c, HEX_ZEBRA if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(c, 60, 60, 90, 90)
            p = c.paragraphs[0]
            p.runs[0].font.name = "Segoe UI"
            p.runs[0].font.size = Pt(8.5)
            if c_idx == 0:
                p.runs[0].font.bold = True
                p.runs[0].font.color.rgb = COLOR_NAVY
            elif c_idx == 1:
                p.runs[0].font.italic = True
                p.runs[0].font.color.rgb = COLOR_ROYAL
            else:
                p.runs[0].font.color.rgb = COLOR_BODY
                
    # Rules Callout
    p_rule_hdr = doc.add_paragraph()
    p_rule_hdr.paragraph_format.space_before = Pt(6)
    p_rule_hdr.paragraph_format.space_after = Pt(2)
    r_rh = p_rule_hdr.add_run("💡 Tips Pelafalan untuk Siswa (Pronunciation Rules):")
    r_rh.font.name = "Segoe UI"
    r_rh.font.size = Pt(9.5)
    r_rh.font.bold = True
    r_rh.font.color.rgb = COLOR_ROYAL

    tbl_tip = doc.add_table(rows=1, cols=1)
    tbl_tip.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_tip = tbl_tip.cell(0, 0)
    c_tip.width = Inches(6.7)
    set_cell_shading(c_tip, HEX_AMBER)
    set_cell_margins(c_tip, 80, 80, 120, 120)
    
    p_t = c_tip.paragraphs[0]
    p_t.paragraph_format.space_after = Pt(0)
    for rule in general_pron_rules:
        r_rl = p_t.add_run(f"• {rule}\n")
        r_rl.font.name = "Segoe UI"
        r_rl.font.size = Pt(8.5)
        r_rl.font.color.rgb = COLOR_BODY

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(8)

print("Student DOCX builder functions ready.")
