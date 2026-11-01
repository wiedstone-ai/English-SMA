# -*- coding: utf-8 -*-
"""
Generator script for LKPD (Lembar Kerja Peserta Didik / Student Worksheets) in 5-Modul Ajar XII SMA (Wajib).
Generates student-ready, printable .docx worksheets for TP 12.W.1 through TP 12.W.6.
Each LKPD features: Student Identity, Petunjuk Belajar, Aktivitas 1 (BKoF Vocabulary),
Aktivitas 2 (MoT Reading & Deconstruction), Aktivitas 3 (JCoT Group Task),
Aktivitas 4 (ICoT Independent Performance), and Self/Peer-Assessment Checklist.
"""

import os
import sys
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import docx_helper

TARGET_DIR = r"C:\Users\KOMPUTERKU2\Documents\English SMA\5-Modul Ajar XII SMA"
os.makedirs(TARGET_DIR, exist_ok=True)

COLOR_NAVY = RGBColor(15, 41, 66)
COLOR_ROYAL = RGBColor(30, 64, 175)
COLOR_BODY = RGBColor(30, 41, 59)

def add_student_id_box(doc, lkpd_num, title, tp_code):
    p_bnr = doc.add_paragraph()
    p_bnr.paragraph_format.space_before = Pt(0)
    p_bnr.paragraph_format.space_after = Pt(2)
    p_bnr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t = p_bnr.add_run(f"LEMBAR KERJA PESERTA DIDIK (LKPD) - {lkpd_num}\nBAHASA INGGRIS KELAS XII SMA")
    r_t.font.name = "Segoe UI"
    r_t.font.size = Pt(13)
    r_t.font.bold = True
    r_t.font.color.rgb = COLOR_NAVY

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(8)
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_s = p_sub.add_run(f"Materi: {title} | Kode: {tp_code} | Kurikulum Merdeka 2025")
    r_s.font.name = "Segoe UI"
    r_s.font.size = Pt(9.5)
    r_s.font.italic = True
    r_s.font.color.rgb = COLOR_ROYAL

    # Student Info Table
    tbl = doc.add_table(rows=2, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    docx_helper.set_table_borders(tbl, color="CBD5E1", sz="4")
    
    fields = [
        [("Nama Siswa", "......................................................."), ("Kelas / No. Absen", "XII - ...... / ......")] ,
        [("Hari / Tanggal", "......................................................."), ("Kelompok / Nilai", "........................ / ........")]
    ]
    for r_idx, row_fields in enumerate(fields):
        row = tbl.rows[r_idx]
        for c_idx, (k, v) in enumerate(row_fields):
            cell = row.cells[c_idx]
            cell.width = Inches(3.25)
            docx_helper.set_cell_shading(cell, "F8FAFC")
            docx_helper.set_cell_margins(cell, 60, 60, 100, 100)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r_k = p.add_run(f"{k}: ")
            r_k.bold = True
            r_k.font.size = Pt(9)
            r_k.font.color.rgb = COLOR_NAVY
            r_v = p.add_run(v)
            r_v.font.size = Pt(9)

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(4)

def add_activity_heading(doc, act_num, act_title, instruction):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(f"AKTIVITAS {act_num}: {act_title.upper()}")
    r.font.name = "Segoe UI"
    r.font.size = Pt(11)
    r.font.bold = True
    r.font.color.rgb = COLOR_ROYAL

    p_ins = doc.add_paragraph()
    p_ins.paragraph_format.space_before = Pt(0)
    p_ins.paragraph_format.space_after = Pt(4)
    r_in = p_ins.add_run(f"Petunjuk Pengerjaan: {instruction}")
    r_in.font.name = "Segoe UI"
    r_in.font.size = Pt(9)
    r_in.font.italic = True
    r_in.font.color.rgb = RGBColor(71, 85, 105)

def add_writing_box(doc, prompt_text, num_lines=8):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(prompt_text)
    r.font.name = "Segoe UI"
    r.font.size = Pt(9.5)
    r.font.bold = True

    tbl = doc.add_table(rows=num_lines, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx in range(num_lines):
        cell = tbl.rows[r_idx].cells[0]
        cell.width = Inches(6.5)
        docx_helper.set_cell_margins(cell, 40, 40, 80, 80)
        docx_helper.set_cell_shading(cell, "FFFFFF")
        p_c = cell.paragraphs[0]
        p_c.paragraph_format.space_after = Pt(0)
        p_c.text = " "
    docx_helper.set_table_borders(tbl, color="CBD5E1", sz="4")

    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_after = Pt(4)


# ==============================================================================
# LKPD 1: TP 12.W.1 - DISCUSSION TEXT (FUTURE OF WORK & AUTOMATION)
# ==============================================================================
def create_lkpd_1():
    doc = docx_helper.create_styled_doc()
    add_student_id_box(doc, "1", "Discussion Text: The Future of Work & Automation", "TP 12.W.1")

    docx_helper.add_callout(
        doc,
        "Tujuan Pembelajaran:",
        "Setelah menyelesaikan lembar kerja ini, kamu mampu mengidentifikasi argumen pro dan kontra terkait otomasi AI di tempat kerja, menggunakan konjungsi konsesif, serta menyusun teks diskusi berimbang sepanjang 250-300 kata.",
        bg="EFF6FF", border_color="1E40AF"
    )

    # Aktivitas 1
    add_activity_heading(doc, "1", "Building Background & Vocabulary Matching (BKoF)", 
                         "Pasangkanlah istilah ketenagakerjaan di kolom kiri dengan definisi yang tepat di kolom kanan!")
    vocab_headers = ["No", "Istilah Otentik", "Definisi Kontekstual (Tuliskan Huruf Pasangan)"]
    vocab_rows = [
        ["1", "Displacement", "[  ] A. Bimbingan karier profesional dari staf senior kepada staf yunior."],
        ["2", "Reskilling", "[  ] B. Model kerja gabungan antara bekerja di kantor dan bekerja dari rumah."],
        ["3", "Commute", "[  ] C. Pergeseran atau hilangnya posisi kerja manusia karena digantikan mesin."],
        ["4", "Hybrid model", "[  ] D. Pelatihan ulang keterampilan baru agar tenaga kerja relevan dengan teknologi."],
        ["5", "Mentorship", "[  ] E. Perjalanan harian bolak-balik antara tempat tinggal dan kantor."]
    ]
    docx_helper.add_styled_table(doc, vocab_headers, vocab_rows, [Inches(0.6), Inches(2.2), Inches(3.7)])

    # Aktivitas 2
    add_activity_heading(doc, "2", "Guided Text Deconstruction (MoT)", 
                         "Bacalah Teks Otentik 1 (World Economic Forum: 'Workplace Automation: Technological Boom or Job Threat?') pada bahan ajar otentikmu, kemudian lengkapi tabel analisis struktur teks di bawah ini!")
    struct_headers = ["Bagian Struktur Teks", "Pertanyaan Pemandu Analisis", "Catatan Hasil Analisis Siswa"]
    struct_rows = [
        ["1. Issue (Pendahuluan)", "Isu teknologi apa yang sedang diperdebatkan pada paragraf 1?", "......................................................................................."],
        ["2. Arguments For (Pros)", "Sebutkan 2 keuntungan otomasi AI yang dijelaskan pada paragraf 2!", "1. ...................................................................................\n2. ..................................................................................."],
        ["3. Arguments Against (Cons)", "Sebutkan 2 ancaman disrupsi AI yang dijelaskan pada paragraf 3!", "1. ...................................................................................\n2. ..................................................................................."],
        ["4. Recommendation (Simpulan)", "Apa jalan tengah (kompromi) yang disarankan penulis di paragraf 4?", "......................................................................................."]
    ]
    docx_helper.add_styled_table(doc, struct_headers, struct_rows, [Inches(1.8), Inches(2.3), Inches(2.4)])

    # Aktivitas 3
    add_activity_heading(doc, "3", "Collaborative Group Matrix (JCoT)", 
                         "Diskusikan bersama kelompokmu (4 orang) mengenai studi kasus 'Sistem Kerja Jarak Jauh (Remote Work)' pada Teks Otentik 2. Rumuskan matriks pro-kontra berikut!")
    matrix_headers = ["Aspek Kehidupan Kerja", "Sudut Pandang Mendukung (PRO)", "Sudut Pandang Menolak (CONTRA)"]
    matrix_rows = [
        ["Waktu & Efisiensi", "Hemat waktu perjalanan dan bebas macet.", "Jam kerja dan istirahat menjadi kabur."],
        ["Kesehatan Mental", "...........................................................", "..........................................................."],
        ["Kolaborasi Tim", "...........................................................", "..........................................................."]
    ]
    docx_helper.add_styled_table(doc, matrix_headers, matrix_rows, [Inches(1.8), Inches(2.35), Inches(2.35)])

    # Aktivitas 4
    add_activity_heading(doc, "4", "Independent Discussion Essay Writing (ICoT)", 
                         "Susunlah sebuah esai teks diskusi berimbang (250-300 kata) dengan memilih salah satu mosi di bawah ini!")
    prompt_desc = (
        "Mosi A: Should companies replace customer service representatives with automated AI chatbots?\n"
        "Mosi B: Should all high school students be required to complete an internship in digital technology?\n\n"
        "Tuliskan draf esai lengkap 4 paragraf (Issue -> Arguments For -> Arguments Against -> Recommendation):"
    )
    add_writing_box(doc, prompt_desc, num_lines=10)

    # Self-Assessment
    add_activity_heading(doc, "5", "Self-Assessment & Peer-Review Checklist", 
                         "Berikan tanda centang (V) pada pernyataan berikut sebelum mengumpulkan tugasmu!")
    chk_headers = ["Kriteria Penilaian Mandiri", "Ya", "Belum", "Catatan Koreksi Teman"]
    chk_rows = [
        ["Teks memuat pengenalan isu yang netral pada paragraf pembuka.", "", "", ""],
        ["Argumen pro dan kontra disajikan secara seimbang didukung alasan logis.", "", "", ""],
        ["Menggunakan konjungsi kontras (conversely, on the other hand, although).", "", "", ""],
        ["Paragraf penutup memberikan rekomendasi jalan tengah yang solutif.", "", "", ""],
        ["Bebas dari kesalahan ejaan (spelling) dan tanda baca (punctuation).", "", "", ""]
    ]
    docx_helper.add_styled_table(doc, chk_headers, chk_rows, [Inches(3.5), Inches(0.6), Inches(0.6), Inches(1.8)])

    fpath = os.path.join(TARGET_DIR, "LKPD_TP_12.W.1_Discussion_Text.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# LKPD 2: TP 12.W.2 - NEWS ITEM TEXT & NEWS CASTING
# ==============================================================================
def create_lkpd_2():
    doc = docx_helper.create_styled_doc()
    add_student_id_box(doc, "2", "News Item Text & Media Literacy", "TP 12.W.2")

    docx_helper.add_callout(
        doc,
        "Tujuan Pembelajaran:",
        "Kamu mampu mengidentifikasi unsur 5W+1H pada berita aktual, menganalisis fungsi kutipan narasumber (quotes), menyusun naskah berita siaran 120-150 kata, dan memperagakannya sebagai News Anchor.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_activity_heading(doc, "1", "Vocabulary & Reporting Verbs Drill (BKoF)", 
                         "Lengkapilah kalimat rumpang berikut dengan kata kerja pelapor (reporting verbs) yang tepat: [stated, commended, announced, praised, confirmed]!")
    p_v = doc.add_paragraph()
    p_v.paragraph_format.left_indent = Inches(0.2)
    p_v.add_run("1. The president officially .................................... that the high-speed railway was open to the public.\n")
    p_v.add_run("2. The international judges .................................... the students for their innovative robotics design.\n")
    p_v.add_run("3. Commuters .................................... the smooth and quiet ride of the new electric bullet train.\n")
    p_v.add_run("4. The energy director .................................... that renewable solar power had surpassed coal.")

    add_activity_heading(doc, "2", "News Element Deconstruction (MoT)", 
                         "Berdasarkan Teks Berita Otentik 1 (Whoosh Bullet Train Launch), analisislah unsur 5W+1H pada tabel berikut!")
    news_headers = ["Unsur Berita", "Pertanyaan Pemantik", "Jawaban Fakta dari Teks"]
    news_rows = [
        ["What (Peristiwa)", "Peristiwa penting apa yang terjadi?", "Peluncuran resmi kereta cepat pertama Asia Tenggara ('Whoosh')."],
        ["Who (Tokoh)", "Siapa tokoh utama yang meresmikan?", "......................................................................................."],
        ["Where (Tempat)", "Di mana lokasi peresmian stasiun?", "......................................................................................."],
        ["When (Waktu)", "Kapan peristiwa tersebut berlangsung?", "......................................................................................."],
        ["Why (Tujuan/Sebab)", "Mengapa kereta cepat ini dibangun?", "......................................................................................."],
        ["Sources (Kutipan)", "Tuliskan 1 kutipan langsung resmi dari narasumber!", "'...................................................................................'"]
    ]
    docx_helper.add_styled_table(doc, news_headers, news_rows, [Inches(1.5), Inches(2.3), Inches(2.7)])

    add_activity_heading(doc, "3", "Collaborative News Script Writing (JCoT)", 
                         "Bersama teman sebangkumu, susunlah draf naskah berita siaran televisi (120-150 kata) mengenai prestasi sekolah atau peristiwa daerah!")
    add_writing_box(doc, "Draf Naskah Berita TV (Headline -> Lead -> Background Details -> Official Quotes):", num_lines=8)

    add_activity_heading(doc, "4", "News Anchor Oral Practice Checklist (ICoT)", 
                         "Berlatihlah membacakan naskah beritamu di depan cermin atau kamera ponsel. Ceklis performamu:")
    chk_rows = [
        ["Eye Contact: Menatap kamera/penonton dengan percaya diri dan stabil.", "", "", ""],
        ["Voice Modulation: Mengatur jeda intonasi berita secara tegas dan lugas.", "", "", ""],
        ["Pronunciation: Melafalkan kosakata bahasa Inggris secara akurat.", "", "", ""],
        ["Posture: Duduk atau berdiri tegak dengan bahasa tubuh profesional.", "", "", ""]
    ]
    docx_helper.add_styled_table(doc, ["Aspek Penampilan News Anchor", "Sangat Baik", "Cukup", "Perlu Latihan"], chk_rows, [Inches(3.5), Inches(1.0), Inches(1.0), Inches(1.0)])

    fpath = os.path.join(TARGET_DIR, "LKPD_TP_12.W.2_News_Item_Text.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# LKPD 3: TP 12.W.3 - REVIEW TEXT (CREATIVE PRODUCTS & HERITAGE)
# ==============================================================================
def create_lkpd_3():
    doc = docx_helper.create_styled_doc()
    add_student_id_box(doc, "3", "Review Text: Indonesian Creative Heritage", "TP 12.W.3")

    docx_helper.add_callout(
        doc,
        "Tujuan Pembelajaran:",
        "Kamu mampu menganalisis struktur teks resensi produk (Orientation, Evaluation, Summary), menggunakan kosakata apresiasi estetika, dan menyusun artikel ulasan produk kreatif lokal nusantara.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_activity_heading(doc, "1", "Aesthetic Vocabulary Matching (BKoF)", 
                         "Cocokkanlah kosakata evaluasi estetika berikut dengan makna yang sesuai!")
    v_rows = [
        ["1", "Painstaking", "[  ] A. Sangat indah, rumit, dan mendetail pembuatannya."],
        ["2", "Intricate", "[  ] B. Pusaka bernilai luhur yang diwariskan turun-temurun."],
        ["3", "Heirloom", "[  ] C. Membutuhkan kerja keras, ketelitian, dan kesabaran tinggi."],
        ["4", "Resilience", "[  ] D. Dapat terurai kembali ke tanah secara alami."],
        ["5", "Biodegradable", "[  ] E. Daya tahan lentur yang kokoh menghadapi tekanan."]
    ]
    docx_helper.add_styled_table(doc, ["No", "Kosakata Estetika", "Makna Kontekstual"], v_rows, [Inches(0.6), Inches(2.2), Inches(3.7)])

    add_activity_heading(doc, "2", "Product Review Deconstruction (MoT)", 
                         "Berdasarkan ulasan otentik 'Indonesian Batik Tulis', jawablah pertanyaan analisis berikut!")
    p_q = doc.add_paragraph()
    p_q.paragraph_format.left_indent = Inches(0.2)
    p_q.add_run("1. What makes genuine Batik Tulis fundamentally different from factory printed cloth?\n")
    p_q.add_run("   Jawab: ........................................................................................................................\n")
    p_q.add_run("2. Why does the reviewer consider Batik Tulis an 'heirloom investment' despite its premium price?\n")
    p_q.add_run("   Jawab: ........................................................................................................................\n")
    p_q.add_run("3. What rating and recommendation does the reviewer assign at the end of the text?\n")
    p_q.add_run("   Jawab: ........................................................................................................................")

    add_activity_heading(doc, "3", "Independent Creative Product Review Writing (ICoT)", 
                         "Tulislah sebuah artikel ulasan produk kreatif lokal (200-250 kata) untuk mempromosikan UMKM nusantara (misal: tenun, ukiran kayu, keramik, atau kuliner khas daerahmu)!")
    add_writing_box(doc, "Draf Review: Title -> Orientation -> Evaluation (Kelebihan & Saran) -> Star Rating (1-5):", num_lines=9)

    fpath = os.path.join(TARGET_DIR, "LKPD_TP_12.W.3_Review_Text.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# LKPD 4: TP 12.W.4 - JOB APPLICATION LETTER & CV
# ==============================================================================
def create_lkpd_4():
    doc = docx_helper.create_styled_doc()
    add_student_id_box(doc, "4", "Job Application Letter & Modern ATS CV", "TP 12.W.4")

    docx_helper.add_callout(
        doc,
        "Tujuan Pembelajaran:",
        "Kamu mampu menganalisis iklan lowongan kerja otentik, menyusun surat lamaran pekerjaan formal (Cover Letter) berstandar bisnis, serta merancang Curriculum Vitae (CV) ATS-compliant 1 halaman siap pakai.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_activity_heading(doc, "1", "Analyzing Job Vacancies (BKoF)", 
                         "Berdasarkan Iklan Lowongan Kerja Teks Otentik 3 (Junior Digital Content Creator), identifikasi komponen lowongan berikut!")
    ad_rows = [
        ["Nama Perusahaan & Posisi", "......................................................................................."],
        ["Tipe Pekerjaan & Lokasi", "......................................................................................."],
        ["Tiga Keterampilan Utama yang Dibutuhkan", "1. ...................................................................................\n2. ...................................................................................\n3. ..................................................................................."],
        ["Alamat Email Pengiriman Lamaran", "......................................................................................."]
    ]
    docx_helper.add_styled_table(doc, ["Komponen Iklan Lowongan", "Catatan Hasil Analisis Siswa"], ad_rows, [Inches(2.5), Inches(4.0)])

    add_activity_heading(doc, "2", "Action Verbs for Professional CVs (MoT)", 
                         "Ubahlah kalimat tugas pasif di bawah menjadi kalimat aktif terukur menggunakan Action Verbs (misal: Directed, Engineered, Coordinated, Boosted)!")
    p_av = doc.add_paragraph()
    p_av.paragraph_format.left_indent = Inches(0.2)
    p_av.add_run("1. I was responsible for school social media posts.\n")
    p_av.add_run("   -> Revised: Managed and published 60+ engaging promotional posts on official school channels.\n")
    p_av.add_run("2. I helped organize the annual youth sports event.\n")
    p_av.add_run("   -> Revised: ...................................................................................................................\n")
    p_av.add_run("3. I was in charge of designing event flyers on Canva.\n")
    p_av.add_run("   -> Revised: ...................................................................................................................")

    add_activity_heading(doc, "3", "Independent Cover Letter Drafting (ICoT)", 
                         "Tuliskan surat lamaran pekerjaan formal (200-250 kata) yang ditujukan kepada BrightSpark Media Group sesuai format surat bisnis baku!")
    add_writing_box(doc, "Draf Surat Lamaran: Tanggal -> Alamat -> Salutation -> 3 Body Paragraphs -> Formal Sign-off:", num_lines=10)

    fpath = os.path.join(TARGET_DIR, "LKPD_TP_12.W.4_Job_Application_and_CV.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# LKPD 5: TP 12.W.5 - JOB INTERVIEW SIMULATION (STAR METHOD)
# ==============================================================================
def create_lkpd_5():
    doc = docx_helper.create_styled_doc()
    add_student_id_box(doc, "5", "Job Interview Simulation (STAR Method)", "TP 12.W.5")

    docx_helper.add_callout(
        doc,
        "Tujuan Pembelajaran:",
        "Kamu mampu menyusun Elevator Pitch 60 detik, menerapkan metode STAR untuk menjawab pertanyaan situasional, dan mempraktikkan simulasi wawancara kerja formal secara percaya diri.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_activity_heading(doc, "1", "Crafting Your 60-Second Elevator Pitch (BKoF)", 
                         "Lengkapi draf perkenalan diri profesional 60 detikmu menggunakan formula: Who I am -> What I achieved -> Where I am going!")
    p_ep = doc.add_paragraph()
    p_ep.paragraph_format.left_indent = Inches(0.2)
    p_ep.add_run("Hello! My name is .........................................................., and I am a graduating senior from .................................................... with a strong interest in ...........................................................................\n")
    p_ep.add_run("During high school, I successfully ......................................................................................................................................\n")
    p_ep.add_run("I am excited to apply for this role because I want to contribute my skills in ......................................................................")

    add_activity_heading(doc, "2", "Structuring Answers with the STAR Method (MoT)", 
                         "Pilihlah salah satu pengalaman suksesmu dalam mengatasi masalah (di sekolah/organisasi), lalu susunlah menggunakan kerangka STAR!")
    star_rows = [
        ["S – Situation (Latar Konteks)", "Kapan dan di mana masalah tersebut terjadi?", "......................................................................................."],
        ["T – Task (Tanggung Jawab)", "Apa tugas spesifikmu dalam situasi tersebut?", "......................................................................................."],
        ["A – Action (Tindakan Terukur)", "Langkah nyata apa yang kamu ambil?", "......................................................................................."],
        ["R – Result (Hasil & Pelajaran)", "Apa hasil terukur dan nilai moral yang didapat?", "......................................................................................."]
    ]
    docx_helper.add_styled_table(doc, ["Unsur STAR", "Pertanyaan Pemandu", "Draf Jawaban Siswa"], star_rows, [Inches(1.8), Inches(2.2), Inches(2.5)])

    add_activity_heading(doc, "3", "Pair Mock Interview Observation Sheet (JCoT & ICoT)", 
                         "Praktikkan wawancara kerja berpasangan. Sebagai observer/interviewer, berikan skor 1-4 untuk performa rekanmu!")
    obs_rows = [
        ["1. Professional Posture & Eye Contact", "Menatap mata pewawancara, duduk tegak dan tersenyum ramah.", "[    ] / 4"],
        ["2. Delivery of 60-Second Elevator Pitch", "Menjelaskan identitas dan potensi diri secara lancar tanpa gugup.", "[    ] / 4"],
        ["3. Effective Use of STAR Method", "Menjawab pertanyaan pengalaman dengan urutan S-T-A-R yang jelas.", "[    ] / 4"],
        ["4. Inquiring Question", "Mengajukan minimal 1 pertanyaan cerdas kepada pewawancara.", "[    ] / 4"]
    ]
    docx_helper.add_styled_table(doc, ["Kriteria Performa Wawancara", "Deskripsi Perilaku Teramati", "Skor (1-4)"], obs_rows, [Inches(2.5), Inches(3.0), Inches(1.0)])

    fpath = os.path.join(TARGET_DIR, "LKPD_TP_12.W.5_Job_Interview_Simulation.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# LKPD 6: TP 12.W.6 - CAPTIONS, DATA GRAPHS & INFOGRAPHICS
# ==============================================================================
def create_lkpd_6():
    doc = docx_helper.create_styled_doc()
    add_student_id_box(doc, "6", "Captions, Data Graphs & Multimodal Infographics", "TP 12.W.6")

    docx_helper.add_callout(
        doc,
        "Tujuan Pembelajaran:",
        "Kamu mampu menyusun takarir foto jurnalistik 2 kalimat berstandar AP/World Press Photo, mendeskripsikan tren data grafik statistik, dan menafsirkan infografis multimodal.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_activity_heading(doc, "1", "Data Trend Vocabulary Drill (BKoF)", 
                         "Kelompokkanlah kosakata tren berikut ke dalam kolom yang tepat: [surged, dropped gradually, leveled off, rose steadily, plunged, remained stable]!")
    t_rows = [
        ["Tren Naik (Upward)", "...........................................................", "..........................................................."],
        ["Tren Turun (Downward)", "...........................................................", "..........................................................."],
        ["Kondisi Stabil (Plateau)", "...........................................................", "..........................................................."]
    ]
    docx_helper.add_styled_table(doc, ["Kategori Tren", "Kosakata 1", "Kosakata 2"], t_rows, [Inches(2.0), Inches(2.25), Inches(2.25)])

    add_activity_heading(doc, "2", "Writing Professional Photo Captions (MoT & ICoT)", 
                         "Tuliskan sebuah takarir foto jurnalistik (2 kalimat) untuk foto kegiatan pelestarian lingkungan atau pameran karya siswa!")
    p_cap = doc.add_paragraph()
    p_cap.paragraph_format.left_indent = Inches(0.2)
    p_cap.add_run("• Sentence 1 (Present Tense: Who - What - Where - When):\n")
    p_cap.add_run("  ...................................................................................................................................................................\n")
    p_cap.add_run("• Sentence 2 (Past / Present Perfect: Context & Background Significance):\n")
    p_cap.add_run("  ...................................................................................................................................................................\n")
    p_cap.add_run("• Photo Credit Line: (Photo: ............................................................................................................................)\n")

    add_activity_heading(doc, "3", "Interpreting Multimodal Infographics (ICoT)", 
                         "Berdasarkan Infografis WHO (5 Pillars for Teen Mental Well-Being), pilihlah 2 pilar yang paling relevan bagi siswa kelas XII dan jelaskan alasannya!")
    add_writing_box(doc, "Penjelasan Analisis Infografis Multimodal:", num_lines=7)

    fpath = os.path.join(TARGET_DIR, "LKPD_TP_12.W.6_Captions_and_Infographics.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


if __name__ == "__main__":
    print("Generating complete student LKPD worksheets for 5-Modul Ajar XII SMA (Wajib)...")
    create_lkpd_1()
    create_lkpd_2()
    create_lkpd_3()
    create_lkpd_4()
    create_lkpd_5()
    create_lkpd_6()
    print("All 6 student LKPD (.docx) worksheets for Kelas XII SMA Wajib successfully created!")
