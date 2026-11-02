import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Common Styles
thin_side = Side(border_style="thin", color="CBD5E1")
med_side = Side(border_style="medium", color="1B365D")
double_bottom = Side(border_style="double", color="1B365D")

table_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
header_border = Border(left=thin_side, right=thin_side, top=med_side, bottom=med_side)

font_title = Font(name="Segoe UI", size=14, bold=True, color="1B365D")
font_subtitle = Font(name="Segoe UI", size=11, bold=True, color="334E68")
font_meta = Font(name="Segoe UI", size=9, italic=True, color="64748B")

font_sec_hdr = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
fill_navy = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
fill_steel = PatternFill(start_color="243B53", end_color="243B53", fill_type="solid")
fill_slate = PatternFill(start_color="334E68", end_color="334E68", fill_type="solid")
fill_accent = PatternFill(start_color="0284C7", end_color="0284C7", fill_type="solid")
fill_blue_light = PatternFill(start_color="F0F4F8", end_color="F0F4F8", fill_type="solid")
fill_alt = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
fill_key = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
fill_pts = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")

font_tbl_hdr = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
font_data = Font(name="Segoe UI", size=9, color="1E293B")
font_data_bold = Font(name="Segoe UI", size=9, bold=True, color="1E293B")
font_key = Font(name="Segoe UI", size=9, bold=True, color="166534")
font_score = Font(name="Segoe UI", size=12, bold=True, color="1B365D")

align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
align_right = Alignment(horizontal="right", vertical="center", wrap_text=True)
align_wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)

def style_range(ws, start_row, start_col, end_row, end_col, font=None, fill=None, alignment=None, border=None):
    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            cell = ws.cell(row=r, column=c)
            if font: cell.font = font
            if fill: cell.fill = fill
            if alignment: cell.alignment = alignment
            if border: cell.border = border

def add_header(ws, title, subtitle, grade, semester, school="SMA NEGERI 1 BAE KUDUS"):
    ws.merge_cells("A1:K1")
    ws.cell(1, 1, school).font = font_title
    ws.cell(1, 1).alignment = align_left
    
    ws.merge_cells("A2:K2")
    ws.cell(2, 1, title).font = font_subtitle
    ws.cell(2, 1).alignment = align_left
    
    ws.merge_cells("A3:K3")
    meta_text = f"Mata Pelajaran: Bahasa Inggris | Jenjang/Kelas: {grade} | Semester: {semester} | Kurikulum Merdeka (Kepka BSKAP No. 020/2026)"
    ws.cell(3, 1, meta_text).font = font_meta
    ws.cell(3, 1).alignment = align_left

def add_passage_box(ws, start_row, title, source_info, passage_lines):
    # Title bar
    ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=11)
    ws.cell(start_row, 1, f"AUTHENTIC READING PASSAGE: {title}").font = font_sec_hdr
    ws.cell(start_row, 1).fill = fill_slate
    ws.cell(start_row, 1).alignment = align_left
    
    # Source metadata
    ws.merge_cells(start_row=start_row+1, start_column=1, end_row=start_row+1, end_column=11)
    ws.cell(start_row+1, 1, f"Atribusi Sumber: {source_info}").font = font_meta
    ws.cell(start_row+1, 1).fill = fill_blue_light
    ws.cell(start_row+1, 1).alignment = align_left
    
    # Passage text box
    ws.merge_cells(start_row=start_row+2, start_column=1, end_row=start_row+4, end_column=11)
    ws.cell(start_row+2, 1, "\n".join(passage_lines)).font = font_data
    ws.cell(start_row+2, 1).fill = fill_blue_light
    ws.cell(start_row+2, 1).alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    
    # Borders
    style_range(ws, start_row, 1, start_row+4, 11, border=table_border)
    return start_row + 6

def build_sts_sheet(ws, grade, topic1_name, topic1_passage, topic2_name, topic2_passage, questions):
    add_header(ws, "LATIHAN ASESMEN TENGAH SEMESTER (STS) - 2 TOPIK ESENSIAL", 
               f"Fokus Materi: Topik 1 ({topic1_name}) & Topik 2 ({topic2_name})", grade, "Gasal")
    
    # Instruction box
    ws.merge_cells("A5:K5")
    ws.cell(5, 1, "PETUNJUK SISWA: Bacalah teks otentik dengan cermat, lalu jawab butir soal HOTS pada kolom 'Jawaban Siswa' (Ketik huruf A, B, C, D, atau E). Nilai dan status koreksi akan terhitung otomatis.").font = font_meta
    ws.cell(5, 1).alignment = align_left
    
    # Passage 1
    next_row = add_passage_box(ws, 7, topic1_name, topic1_passage["source"], topic1_passage["text"])
    
    # Passage 2
    next_row = add_passage_box(ws, next_row, topic2_name, topic2_passage["source"], topic2_passage["text"])
    
    # Table header
    q_start = next_row + 1
    headers = [
        ("No", 4), ("Topik / Domain", 15), ("Indikator Kognitif (KKTP)", 22), 
        ("Butir Soal (Stem)", 38), ("Pilihan Jawaban (A, B, C, D, E)", 38),
        ("Kunci", 7), ("Pembahasan Analitis", 32), ("Jawaban Siswa", 14), 
        ("Koreksi", 9), ("Skor Butir", 9), ("Keterangan", 12)
    ]
    
    for c_idx, (h_name, width) in enumerate(headers, 1):
        cell = ws.cell(q_start, c_idx, h_name)
        cell.font = font_tbl_hdr
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = header_border
        col_letter = get_column_letter(c_idx)
        ws.column_dimensions[col_letter].width = width
    
    ws.row_dimensions[q_start].height = 25
    
    # Populate Questions
    row = q_start + 1
    for q in questions:
        ws.row_dimensions[row].height = 55
        is_even = (row % 2 == 0)
        curr_fill = fill_alt if is_even else None
        
        ws.cell(row, 1, q["no"]).alignment = align_center
        ws.cell(row, 2, q["topic"]).alignment = align_left
        ws.cell(row, 3, q["kktp"]).alignment = align_wrap
        ws.cell(row, 4, q["stem"]).alignment = align_wrap
        ws.cell(row, 5, "\n".join(q["options"])).alignment = align_wrap
        
        k_cell = ws.cell(row, 6, q["key"])
        k_cell.alignment = align_center
        k_cell.font = font_key
        k_cell.fill = fill_key
        
        ws.cell(row, 7, q["explanation"]).alignment = align_wrap
        
        # Student Answer (Col H / 8)
        ans_cell = ws.cell(row, 8, "")
        ans_cell.alignment = align_center
        ans_cell.font = font_data_bold
        ans_cell.fill = fill_pts
        
        # Check formula (Col I / 9)
        # =IF(ISBLANK(H12), "-", IF(UPPER(TRIM(H12))=F12, "BENAR", "SALAH"))
        ws.cell(row, 9, f'=IF(ISBLANK(H{row}), "-", IF(UPPER(TRIM(H{row}))=F{row}, "BENAR", "SALAH"))').alignment = align_center
        
        # Score formula (Col J / 10)
        # =IF(I12="BENAR", 10, 0)
        ws.cell(row, 10, f'=IF(I{row}="BENAR", 10, 0)').alignment = align_center
        
        # Feedback (Col K / 11)
        ws.cell(row, 11, f'=IF(I{row}="BENAR", "Tuntas", IF(I{row}="SALAH", "Perlu Reviu", "-"))').alignment = align_center
        
        style_range(ws, row, 1, row, 11, font=font_data, border=table_border)
        if curr_fill and ans_cell.fill != fill_pts:
            pass
        row += 1
        
    # Summary Box
    tot_row = row + 1
    ws.merge_cells(start_row=tot_row, start_column=1, end_row=tot_row, end_column=8)
    ws.cell(tot_row, 1, "TOTAL SKOR ASESMEN TENGAH SEMESTER (STS)").font = font_title
    ws.cell(tot_row, 1).alignment = align_right
    
    ws.merge_cells(start_row=tot_row, start_column=9, end_row=tot_row, end_column=10)
    score_cell = ws.cell(tot_row, 9, f"=SUM(J{q_start+1}:J{row-1})")
    score_cell.font = font_score
    score_cell.alignment = align_center
    score_cell.fill = fill_pts
    score_cell.border = table_border
    
    ws.cell(tot_row, 11, f'=IF(I{tot_row}>=75, "LULUS KKTP", "REMIDIAL")').alignment = align_center
    ws.cell(tot_row, 11).font = font_data_bold
    ws.cell(tot_row, 11).border = table_border

def build_sas_sheet(ws, grade, topics_summary, questions_pg, questions_essay):
    add_header(ws, "LATIHAN ASESMEN SEMESTER (SAS) - 4 TOPIK ESENSIAL LENGKAP", 
               f"Cakupan 4 Topik Esensial: {topics_summary}", grade, "Gasal")
    
    ws.merge_cells("A5:K5")
    ws.cell(5, 1, "BAGIAN I: 20 BUTIR SOAL PILIHAN GANDA HOTS (Bobot 70%) | Ketik jawaban A/B/C/D/E pada kolom H").font = font_sec_hdr
    ws.cell(5, 1).fill = fill_steel
    ws.cell(5, 1).alignment = align_left
    
    headers = [
        ("No", 4), ("Topik Esensial", 15), ("Teks Otentik / Stimulus", 22), 
        ("Butir Soal HOTS", 38), ("Pilihan Jawaban (A - E)", 38),
        ("Kunci", 7), ("Pembahasan Analitis", 32), ("Jawaban Siswa", 14), 
        ("Koreksi", 9), ("Skor (x3.5)", 9), ("Status", 12)
    ]
    
    q_start = 7
    for c_idx, (h_name, width) in enumerate(headers, 1):
        cell = ws.cell(q_start, c_idx, h_name)
        cell.font = font_tbl_hdr
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = header_border
        col_letter = get_column_letter(c_idx)
        ws.column_dimensions[col_letter].width = width
    
    ws.row_dimensions[q_start].height = 25
    
    row = q_start + 1
    for q in questions_pg:
        ws.row_dimensions[row].height = 55
        ws.cell(row, 1, q["no"]).alignment = align_center
        ws.cell(row, 2, q["topic"]).alignment = align_left
        ws.cell(row, 3, q["stimulus"]).alignment = align_wrap
        ws.cell(row, 4, q["stem"]).alignment = align_wrap
        ws.cell(row, 5, "\n".join(q["options"])).alignment = align_wrap
        
        k_cell = ws.cell(row, 6, q["key"])
        k_cell.alignment = align_center
        k_cell.font = font_key
        k_cell.fill = fill_key
        
        ws.cell(row, 7, q["explanation"]).alignment = align_wrap
        
        # Student Answer (Col H / 8)
        ans_cell = ws.cell(row, 8, "")
        ans_cell.alignment = align_center
        ans_cell.font = font_data_bold
        ans_cell.fill = fill_pts
        
        # Check formula (Col I / 9)
        ws.cell(row, 9, f'=IF(ISBLANK(H{row}), "-", IF(UPPER(TRIM(H{row}))=F{row}, "BENAR", "SALAH"))').alignment = align_center
        
        # Score formula (Col J / 10): 20 questions x 3.5 = 70 points
        ws.cell(row, 10, f'=IF(I{row}="BENAR", 3.5, 0)').alignment = align_center
        
        # Feedback (Col K / 11)
        ws.cell(row, 11, f'=IF(I{row}="BENAR", "Tuntas", IF(I{row}="SALAH", "Reviu", "-"))').alignment = align_center
        
        style_range(ws, row, 1, row, 11, font=font_data, border=table_border)
        row += 1
        
    pg_end_row = row - 1
    
    # BAGIAN II: 5 SOAL ESAI STUDI KASUS (Bobot 30%)
    essay_sec_row = row + 1
    ws.merge_cells(start_row=essay_sec_row, start_column=1, end_row=essay_sec_row, end_column=11)
    ws.cell(essay_sec_row, 1, "BAGIAN II: 5 BUTIR SOAL ESAI ANALITIS & STUDI KASUS AUTENTIK (Bobot 30% | 5 Soal x Skor 6 = 30 Poin)").font = font_sec_hdr
    ws.cell(essay_sec_row, 1).fill = fill_slate
    ws.cell(essay_sec_row, 1).alignment = align_left
    
    essay_headers = [
        ("No", 4), ("Topik Esensial", 15), ("Konteks Studi Kasus Otentik", 25), 
        ("Instruksi Tugas Esai Analitis", 38), ("Rubrik Penilaian & Kriteria Skor 0 - 6", 38),
        ("Skor Maks", 9), ("Panduan Jawaban Guru", 32), ("Nilai Esai Guru (0-6)", 14), 
        ("Koreksi", 9), ("Skor Akhir Esai", 9), ("Catatan Guru", 12)
    ]
    
    e_start = essay_sec_row + 1
    for c_idx, (h_name, width) in enumerate(essay_headers, 1):
        cell = ws.cell(e_start, c_idx, h_name)
        cell.font = font_tbl_hdr
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = header_border
        
    ws.row_dimensions[e_start].height = 25
    
    row = e_start + 1
    for e in questions_essay:
        ws.row_dimensions[row].height = 65
        ws.cell(row, 1, e["no"]).alignment = align_center
        ws.cell(row, 2, e["topic"]).alignment = align_left
        ws.cell(row, 3, e["context"]).alignment = align_wrap
        ws.cell(row, 4, e["prompt"]).alignment = align_wrap
        ws.cell(row, 5, e["rubric"]).alignment = align_wrap
        ws.cell(row, 6, 6).alignment = align_center
        ws.cell(row, 7, e["key_model"]).alignment = align_wrap
        
        # Teacher input score (0-6)
        score_in = ws.cell(row, 8, "")
        score_in.alignment = align_center
        score_in.font = font_data_bold
        score_in.fill = fill_pts
        
        ws.cell(row, 9, f'=IF(ISBLANK(H{row}), "-", IF(H{row}>=4, "MEMADAI", "KURANG"))').alignment = align_center
        ws.cell(row, 10, f'=IF(ISBLANK(H{row}), 0, MIN(H{row}, 6))').alignment = align_center
        ws.cell(row, 11, f'=IF(H{row}>=5, "Sangat Baik", IF(H{row}>=3, "Cukup", "Perlu Bimbingan"))').alignment = align_center
        
        style_range(ws, row, 1, row, 11, font=font_data, border=table_border)
        row += 1
        
    essay_end_row = row - 1
    
    # GRAND TOTAL
    tot_row = row + 1
    ws.merge_cells(start_row=tot_row, start_column=1, end_row=tot_row, end_column=7)
    ws.cell(tot_row, 1, "TOTAL NILAI ASESMEN SEMESTER (PG: 70% + ESAI: 30% = 100 POIN)").font = font_title
    ws.cell(tot_row, 1).alignment = align_right
    
    ws.merge_cells(start_row=tot_row, start_column=8, end_row=tot_row, end_column=10)
    score_cell = ws.cell(tot_row, 8, f"=ROUND(SUM(J{q_start+1}:J{pg_end_row}) + SUM(J{e_start+1}:J{essay_end_row}), 1)")
    score_cell.font = font_score
    score_cell.alignment = align_center
    score_cell.fill = fill_pts
    score_cell.border = table_border
    
    ws.cell(tot_row, 11, f'=IF(H{tot_row}>=75, "LULUS KKTP", "REMIDIAL")').alignment = align_center
    ws.cell(tot_row, 11).font = font_data_bold
    ws.cell(tot_row, 11).border = table_border

def build_gradebook_sheet(ws, grade):
    add_header(ws, "REKAPITULASI NILAI LATIHAN ASESMEN SISWA (GRADEBOOK)", 
               "Terhubung Otomatis dengan Simulasi Penilaian Berkelanjutan", grade, "Gasal")
    
    headers = [
        ("No", 4), ("Nomor Induk Siswa (NIS)", 18), ("Nama Lengkap Siswa", 30),
        ("Nilai STS (Bobot 40%)", 18), ("Nilai SAS (Bobot 60%)", 18),
        ("Nilai Akhir Semester", 18), ("Status Ketercapaian KKTP", 25), ("Tindak Lanjut Pembelajaran", 25)
    ]
    
    start_row = 6
    for c_idx, (h_name, width) in enumerate(headers, 1):
        cell = ws.cell(start_row, c_idx, h_name)
        cell.font = font_tbl_hdr
        cell.fill = fill_navy
        cell.alignment = align_center
        cell.border = header_border
        col_letter = get_column_letter(c_idx)
        ws.column_dimensions[col_letter].width = width
        
    sample_students = [
        "Ahmad Fauzi", "Aisyah Nur Rahma", "Anindya Putri Kusuma", "Bagus Dwi Prasetyo",
        "Bayu Setiawan", "Citra Dewi Lestari", "Dimas Arya Pratama", "Dinda Maharani",
        "Eka Kurniawan", "Fadhil Muhammad", "Farah Nadia", "Galih Saputra",
        "Gita Ayu Safitri", "Hafiz Ridwan", "Indah Permatasari", "Irfan Maulana",
        "Kurnia Ramadhani", "Laila Syahrani", "M. Rizky Ramadhan", "Nabila Khairunnisa",
        "Naufal Alamsyah", "Nurul Hidayah", "Panji Wicaksono", "Putri Rahayu",
        "Raditya Eka Wardhana", "Rara Sekar Kinanti", "Reza Firmansyah", "Rina Wulandari",
        "Rizka Amalia", "Satria Bintang", "Septian Nugroho", "Siti Aisyah",
        "Taufik Hidayat", "Tiara Amanda", "Wahyu Kusuma", "Zahra Novitasari"
    ]
    
    row = start_row + 1
    for idx, name in enumerate(sample_students, 1):
        ws.cell(row, 1, idx).alignment = align_center
        ws.cell(row, 2, f"2025{grade.split()[-1]}{idx:03d}").alignment = align_center
        ws.cell(row, 3, name).alignment = align_left
        
        # Sample simulated STS & SAS values
        sts_val = 70 + (idx * 3) % 28
        sas_val = 72 + (idx * 5) % 26
        
        ws.cell(row, 4, sts_val).alignment = align_center
        ws.cell(row, 5, sas_val).alignment = align_center
        
        # Formula Nilai Akhir: 40% STS + 60% SAS
        ws.cell(row, 6, f"=ROUND(D{row}*0.4 + E{row}*0.6, 1)").alignment = align_center
        ws.cell(row, 6).font = font_data_bold
        
        # Status KKTP (KKTP = 75)
        ws.cell(row, 7, f'=IF(F{row}>=75, "Tuntas (Mencapai KKTP)", "Belum Tuntas (Perlu Remidial)")').alignment = align_center
        ws.cell(row, 8, f'=IF(F{row}>=85, "Pengayaan Mandiri", IF(F{row}>=75, "Latihan Lanjutan", "Tutor Sebaya & Remedial"))').alignment = align_center
        
        style_range(ws, row, 1, row, 8, font=font_data, border=table_border)
        if row % 2 == 0:
            for c in range(1, 9):
                if ws.cell(row, c).fill != fill_pts:
                    ws.cell(row, c).fill = fill_alt
        row += 1
        
    last_data_row = row - 1
    
    # Statistical Summary Rows
    stat_start = row + 1
    ws.merge_cells(start_row=stat_start, start_column=1, end_row=stat_start, end_column=5)
    ws.cell(stat_start, 1, "RATA-RATA NILAI KELAS").font = font_data_bold
    ws.cell(stat_start, 1).alignment = align_right
    ws.cell(stat_start, 6, f"=ROUND(AVERAGE(F{start_row+1}:F{last_data_row}), 1)").alignment = align_center
    ws.cell(stat_start, 6).font = font_score
    style_range(ws, stat_start, 1, stat_start, 8, border=table_border)
    
    ws.merge_cells(start_row=stat_start+1, start_column=1, end_row=stat_start+1, end_column=5)
    ws.cell(stat_start+1, 1, "NILAI TERTINGGI (MAX)").font = font_data_bold
    ws.cell(stat_start+1, 1).alignment = align_right
    ws.cell(stat_start+1, 6, f"=MAX(F{start_row+1}:F{last_data_row})").alignment = align_center
    ws.cell(stat_start+1, 6).font = font_score
    style_range(ws, stat_start+1, 1, stat_start+1, 8, border=table_border)
    
    ws.merge_cells(start_row=stat_start+2, start_column=1, end_row=stat_start+2, end_column=5)
    ws.cell(stat_start+2, 1, "NILAI TERENDAH (MIN)").font = font_data_bold
    ws.cell(stat_start+2, 1).alignment = align_right
    ws.cell(stat_start+2, 6, f"=MIN(F{start_row+1}:F{last_data_row})").alignment = align_center
    ws.cell(stat_start+2, 6).font = font_score
    style_range(ws, stat_start+2, 1, stat_start+2, 8, border=table_border)
    
    ws.merge_cells(start_row=stat_start+3, start_column=1, end_row=stat_start+3, end_column=5)
    ws.cell(stat_start+3, 1, "PERSENTASE KETUNTASAN BELAJAR").font = font_data_bold
    ws.cell(stat_start+3, 1).alignment = align_right
    ws.cell(stat_start+3, 6, f'=COUNTIF(G{start_row+1}:G{last_data_row}, "*Tuntas*")/COUNTA(G{start_row+1}:G{last_data_row})').alignment = align_center
    ws.cell(stat_start+3, 6).number_format = '0.0%'
    ws.cell(stat_start+3, 6).font = font_score
    style_range(ws, stat_start+3, 1, stat_start+3, 8, border=table_border)

print("Builder core functions defined.")
