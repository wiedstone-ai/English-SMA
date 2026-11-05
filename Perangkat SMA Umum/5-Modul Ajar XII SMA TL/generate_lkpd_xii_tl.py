# -*- coding: utf-8 -*-
"""
Generator script for LKPD (Lembar Kerja Peserta Didik / Student Worksheets) in 5-Modul Ajar XII SMA TL.
Generates student-ready, printable .docx worksheets for TP 12.TL.1 through TP 12.TL.8.
Each LKPD features: Student Identity, Petunjuk Belajar, Aktivitas 1 (BKoF Vocabulary & Concepts),
Aktivitas 2 (MoT Guided Reading & Deconstruction), Aktivitas 3 (JCoT Collaborative Matrix / Simulation),
Aktivitas 4 (ICoT Independent Authentic Performance Task with guided writing box),
and Aktivitas 5 (Self & Peer Assessment Rubric Checklist).
"""

import os
import sys
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import docx_helper

TARGET_DIR = os.path.dirname(os.path.abspath(__file__))
os.makedirs(TARGET_DIR, exist_ok=True)

COLOR_NAVY = RGBColor(15, 41, 66)
COLOR_ROYAL = RGBColor(30, 64, 175)
COLOR_BODY = RGBColor(30, 41, 59)

def add_student_id_box(doc, lkpd_num, title, tp_code):
    p_bnr = doc.add_paragraph()
    p_bnr.paragraph_format.space_before = Pt(0)
    p_bnr.paragraph_format.space_after = Pt(2)
    p_bnr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t = p_bnr.add_run(f"LEMBAR KERJA PESERTA DIDIK (LKPD) - {lkpd_num}\nBAHASA INGGRIS TINGKAT LANJUT KELAS XII SMA")
    r_t.font.name = "Segoe UI"
    r_t.font.size = Pt(13)
    r_t.font.bold = True
    r_t.font.color.rgb = COLOR_NAVY

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(8)
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_s = p_sub.add_run(f"Materi: {title} | Kode: {tp_code} | Kurikulum Merdeka 2025/2026")
    r_s.font.name = "Segoe UI"
    r_s.font.size = Pt(9.5)
    r_s.font.italic = True
    r_s.font.color.rgb = COLOR_ROYAL

    # Student Info Table
    tbl = doc.add_table(rows=2, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    docx_helper.set_table_borders(tbl, color="CBD5E1", sz="4")
    
    fields = [
        [("Nama Siswa", "......................................................."), ("Kelas / No. Absen", "XII TL - ...... / ......")] ,
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
# LKPD 1: TP 12.TL.1 - GLOBAL BIOETHICS DISCUSSION
# ==============================================================================
def create_lkpd_1():
    doc = docx_helper.create_styled_doc()
    add_student_id_box(doc, "1", "Global Academic Discussion on Bioethics", "TP 12.TL.1")

    docx_helper.add_callout(
        doc,
        "Tujuan Pembelajaran:",
        "Setelah menyelesaikan lembar kerja ini, kamu mampu menganalisis wacana debat akademik tingkat tinggi mengenai rekayasa genetika CRISPR dan bioetika medis, menggunakan stance markers dan klausa konsesif, serta menyusun Position Paper bioetika sepanjang 400-500 kata.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_activity_heading(doc, "1", "Building Background & Academic Vocabulary (BKoF)", 
                         "Pasangkanlah istilah bioetika di kolom kiri dengan definisi kontekstual yang tepat di kolom kanan!")
    v_headers = ["No", "Istilah Bioetika", "Definisi Kontekstual (Tuliskan Huruf Pasangan)"]
    v_rows = [
        ["1", "Germline editing", "[  ] A. Kondisi ekonomi kelimpahan di mana sumber daya alam tidak lagi terbatas."],
        ["2", "Beneficence", "[  ] B. Penundaan atau penghentian sementara secara resmi terhadap aktivitas ilmiah."],
        ["3", "Stratification", "[  ] C. Modifikasi genetik pada sel sperma/sel telur yang diwariskan ke keturunan."],
        ["4", "Moratorium", "[  ] D. Kewajiban moral tenaga medis/ilmuwan untuk bertindak demi kemaslahatan pasien."],
        ["5", "Post-scarcity", "[  ] E. Terjadinya pengkastaan hierarki sosial dan ketimpangan biologis buatan."]
    ]
    docx_helper.add_styled_table(doc, v_headers, v_rows, [Inches(0.6), Inches(2.2), Inches(3.7)])

    add_activity_heading(doc, "2", "Guided Text Deconstruction (MoT)", 
                         "Bacalah Teks Otentik 1 ('CRISPR and Human Germline Modification: Navigating the Ethical Precipice') pada bahan ajar otentikmu, kemudian lengkapi bagan dekonstruksi teks berikut!")
    s_headers = ["Elemen Struktur Teks", "Pertanyaan Pemandu Analisis", "Catatan Hasil Analisis Kritis Siswa"]
    s_rows = [
        ["1. The Ethical Dilemma", "Apa perbedaan konsekuensi biologis antara somatic gene therapy dan germline modification?", "......................................................................................."],
        ["2. Utilitarian Arguments (Pros)", "Sebutkan 2 argumen moral mengapa pendukung menuntut pelegalan modifikasi genetik!", "1. ...................................................................................\n2. ..................................................................................."],
        ["3. Egalitarian Arguments (Cons)", "Sebutkan 2 ancaman bahaya kemanusiaan menurut penentang modifikasi germline!", "1. ...................................................................................\n2. ..................................................................................."],
        ["4. Balanced Policy Recommendation", "Apa kompromi kebijakan regulasi internasional yang disimpulkan pada paragraf akhir?", "......................................................................................."]
    ]
    docx_helper.add_styled_table(doc, s_headers, s_rows, [Inches(1.8), Inches(2.3), Inches(2.4)])

    add_activity_heading(doc, "3", "Collaborative Bioethics Deliberation Matrix (JCoT)", 
                         "Bersama kelompok meja bundarmu (4 orang), analisislah studi kasus AI Triage (Teks 2) atau Kolonisasi Antariksa (Teks 3) dan rumuskan matriks pertimbangan kebijakan berikut!")
    m_headers = ["Dimensi Kebijakan Sains", "Perspektif Inovasi & Manfaat (PRO)", "Perspektif Etika, Moral & Risiko (CON)"]
    m_rows = [
        ["Martabat Manusia", "Mencegah kesalahan diagnosis akibat kelelahan.", "Algoritma tidak memiliki empati dan nurani."],
        ["Keadilan & Akses Sosial", "...........................................................", "..........................................................."],
        ["Regulasi Masa Depan", "...........................................................", "..........................................................."]
    ]
    docx_helper.add_styled_table(doc, m_headers, m_rows, [Inches(1.8), Inches(2.35), Inches(2.35)])

    add_activity_heading(doc, "4", "Independent Bioethics Position Paper Writing (ICoT)", 
                         "Susunlah sebuah Position Paper akademik 400-500 kata dengan memilih salah satu mosi etika sains di bawah ini!")
    prompt_desc = (
        "Mosi A: Regulating Human Germline Gene Editing: Moral Imperative or Technological Peril?\n"
        "Mosi B: Algorithmic Triage in Critical Healthcare: Objective Efficiency vs Dehumanization.\n\n"
        "Tuliskan naskah esai posisi dengan 4 paragraf (Issue -> Arguments For -> Arguments Against -> Ethical Synthesis):"
    )
    add_writing_box(doc, prompt_desc, num_lines=12)

    add_activity_heading(doc, "5", "Self-Assessment & Peer-Review Checklist", 
                         "Berikan tanda centang (V) pada pernyataan berikut sebelum naskah posisi dikumpulkan!")
    c_headers = ["Kriteria Penilaian Mandiri & Kemitraan", "Ya", "Belum", "Catatan Koreksi Rekan"]
    c_rows = [
        ["Isu bioetika diuraikan secara mendalam dan netral pada paragraf pembuka.", "", "", ""],
        ["Mengintegrasikan argumen filosofis (utilitarian vs hak asasi) secara berimbang.", "", "", ""],
        ["Menggunakan minimal 4 penanda sikap akademik (stance markers: arguably, contend).", "", "", ""],
        ["Menggunakan minimal 3 klausa konsesif (notwithstanding, whereas, while).", "", "", ""],
        ["Menawarkan rekomendasi jalan tengah regulasi yang etis dan aplikatif.", "", "", ""]
    ]
    docx_helper.add_styled_table(doc, c_headers, c_rows, [Inches(3.5), Inches(0.6), Inches(0.6), Inches(1.8)])

    fpath = os.path.join(TARGET_DIR, "LKPD_TP_12.TL.1_Global_Bioethics_Discussion.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# LKPD 2: TP 12.TL.2 - ACADEMIC SYNTHESIS ESSAY
# ==============================================================================
def create_lkpd_2():
    doc = docx_helper.create_styled_doc()
    add_student_id_box(doc, "2", "Academic Synthesis Essay & Citation Rigor", "TP 12.TL.2")

    docx_helper.add_callout(
        doc,
        "Tujuan Pembelajaran:",
        "Setelah menyelesaikan lembar kerja ini, kamu mampu merumuskan working thesis statement, menyintesis minimal 3 sumber jurnal ilmiah terakreditasi menggunakan kaidah sitasi APA 7th Edition bebas plagiarisme, serta menyusun draf esai akademik 1.000 kata.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_activity_heading(doc, "1", "Scholarly Lexicon & Citation Rules (BKoF)", 
                         "Lengkapilah kalimat rumpang akademik berikut dengan istilah leksikal yang tepat!")
    v_headers = ["No", "Kalimat Kontekstual Akademik", "Pilihan Kata Tepat"]
    v_rows = [
        ["1", "Synthetic polymers have become __________ throughout aquatic and terrestrial biomes.", "A. Intermittency"],
        ["2", "Failure to acknowledge the original author's ideas constitutes academic __________.", "B. Ubiquitous"],
        ["3", "Researchers noted significant __________ between clinical trials and real-world trials.", "C. Plagiarism"],
        ["4", "Solar and wind energy require battery storage systems to balance meteorological __________.", "D. Discrepancies"],
        ["5", "A literature review must __________ multiple findings rather than summarize single papers.", "E. Synthesize"]
    ]
    docx_helper.add_styled_table(doc, v_headers, v_rows, [Inches(0.6), Inches(4.3), Inches(1.6)])

    add_activity_heading(doc, "2", "Literature Synthesis Matrix (MoT)", 
                         "Telaahlah Teks Otentik 1 ('Microplastic Infiltration: Synthesized Environmental Evidence'). Petakan bagaimana penulis menyintesis 4 studi ke dalam matriks berikut!")
    s_headers = ["Penulis & Tahun", "Metode / Subjek Penelitian", "Temuan Kunci", "Peran dalam Sintesis Esai"]
    s_rows = [
        ["Geyer et al. (2017)", "Estimasi produksi plastik global", "8.3 miliar metrik ton sampah plastik", "Menetapkan skala krisis biosfer"],
        ["Leslie et al. (2022)", "Sampel darah manusia seutuhnya", "...........................................", "..........................................."],
        ["Prata et al. (2023)", "Endotelium vaskular manusia", "...........................................", "..........................................."],
        ["Vethaak & Legler (2021)", "Uji toksikologi in vitro", "...........................................", "..........................................."]
    ]
    docx_helper.add_styled_table(doc, s_headers, s_rows, [Inches(1.8), Inches(1.8), Inches(1.6), Inches(1.3)])

    add_activity_heading(doc, "3", "Collaborative Paraphrasing & APA Citation (JCoT)", 
                         "Bersama rekan sebangku, ubahlah kutipan mentah berikut menjadi kalimat parafrasa akademik yang elegan dengan sitasi in-text APA 7th Edition!")
    p_headers = ["Kutipan Teks Asli (Direct Raw Quote)", "Hasil Parafrasa Akademik Kolaboratif Bebas Plagiarisme"]
    p_rows = [
        ["'Solar panels stop producing power when clouds block sunlight, which destabilizes power frequencies.' (Kroposki, 2020)", 
         "..................................................................................................................."],
        ["'Cash stipends gave recipients financial resilience to launch new small businesses rather than quit working.' (West et al., 2021)", 
         "..................................................................................................................."]
    ]
    docx_helper.add_styled_table(doc, p_headers, p_rows, [Inches(3.25), Inches(3.25)])

    add_activity_heading(doc, "4", "Working Thesis & Outline Formulation (ICoT)", 
                         "Rumuskan working thesis statement dan kerangka esai sintesis akademik 1.000 katamu!")
    prompt_desc = (
        "Working Thesis Statement: (Pastikan memuat klaim terukur yang diperdebatkan)\n"
        "Tuliskan draf outline 5 paragraf (Introduction, Thematic Synthesis 1, Thematic Synthesis 2, Counter-Arguments, Conclusion & References):"
    )
    add_writing_box(doc, prompt_desc, num_lines=10)

    add_activity_heading(doc, "5", "Academic Integrity & Peer-Review Checklist", 
                         "Periksalah draf esai rekanmu dengan indikator kepatuhan standar internasional berikut!")
    c_headers = ["Indikator Kepatuhan Akademik", "Ya", "Belum", "Catatan Kritis Rekan Reviewer"]
    c_rows = [
        ["Thesis statement terumuskan secara spesifik dan eksplisit di akhir paragraf 1.", "", "", ""],
        ["Setiap paragraf tubuh menyintesis minimal 2 sumber literatur yang relevan.", "", "", ""],
        ["Seluruh sitasi dalam teks (in-text citation) mematuhi format resmi APA 7th Edition.", "", "", ""],
        ["Tidak ada kalimat hasil copy-paste langsung tanpa tanda petik yang sah.", "", "", ""],
        ["Daftar Pustaka (References) memuat informasi bibliografis lengkap beserta DOI.", "", "", ""]
    ]
    docx_helper.add_styled_table(doc, c_headers, c_rows, [Inches(3.5), Inches(0.6), Inches(0.6), Inches(1.8)])

    fpath = os.path.join(TARGET_DIR, "LKPD_TP_12.TL.2_Academic_Synthesis_Essay.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# LKPD 3: TP 12.TL.3 - DRAMATIC SCRIPT & PERFORMANCE
# ==============================================================================
def create_lkpd_3():
    doc = docx_helper.create_styled_doc()
    add_student_id_box(doc, "3", "Dramatic Literature, Subtext & Theatrical Performance", "TP 12.TL.3")

    docx_helper.add_callout(
        doc,
        "Tujuan Pembelajaran:",
        "Setelah menyelesaikan lembar kerja ini, kamu mampu menganalisis petunjuk panggung (stage directions) dan subteks psikologis naskah drama sastra dunia (The Crucible, A Raisin in the Sun, Hamlet), mempraktikkan modulasi vokal, serta mementaskan adegan monolog dramatik 3-4 menit.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_activity_heading(doc, "1", "Theatrical Terminology & Vocal Concepts (BKoF)", 
                         "Jodohkanlah istilah seni drama di kolom kiri dengan konsep yang tepat di kolom kanan!")
    v_headers = ["No", "Istilah Teater", "Konsep Seni Peran (Tuliskan Huruf Pasangan)"]
    v_rows = [
        ["1", "Subtext", "[  ] A. Monolog saat tokoh berbicara kepada dirinya sendiri mengungkapkan pergulatan batin."],
        ["2", "Soliloquy", "[  ] B. Makna tersembunyi, emosi, atau motivasi di balik kata-kata yang diucapkan tokoh."],
        ["3", "Stage direction", "[  ] C. Perubahan nada, dinamika volume, dan kecepatan suara demi ekspresi emosi."],
        ["4", "Voice modulation", "[  ] D. Petunjuk posisi aktor, aksi fisik, dan pencahayaan panggung dalam kurung siku."],
        ["5", "Character arc", "[  ] E. Perjalanan evolusi mental atau moral yang dialami tokoh sepanjang cerita."]
    ]
    docx_helper.add_styled_table(doc, v_headers, v_rows, [Inches(0.6), Inches(2.2), Inches(3.7)])

    add_activity_heading(doc, "2", "Subtext & Conflict Analysis (MoT)", 
                         "Telaahlah dialog John Proctor dalam Teks Otentik 1 ('The Crucible: The Climax of Name and Integrity'). Uraikan dialog dan subteks tersembunyinya!")
    s_headers = ["Kutipan Dialog Tokoh", "Emosi Tampak (Surface Emotion)", "Subteks Mendalam (Inner Conflict)"]
    s_rows = [
        ["PROCTOR: 'Because it is my name! Because I cannot have another in my life!'", 
         "Kemarahan yang meluap-luap.", "Menyadari bahwa mengorbankan kehormatan nurani lebih buruk daripada kematian fisik."],
        ["PROCTOR: 'I have given you my soul; leave me my name!'", 
         "Keputusasaan mendalam.", "......................................................................................."],
        ["DANFORTH: 'Surrender the paper so that the village may see your submission...'", 
         "Ketegasan wibawa hakim.", "......................................................................................."]
    ]
    docx_helper.add_styled_table(doc, s_headers, s_rows, [Inches(2.5), Inches(1.8), Inches(2.2)])

    add_activity_heading(doc, "3", "Collaborative Theatrical Script Annotation (JCoT)", 
                         "Bersama mitra pementasanmu, analisislah cuplikan naskah Walter Lee (Teks 2) atau Hamlet (Teks 3). Berikan tanda anotasi olah vokal pada teks berikut!")
    annot_desc = (
        "Gunakan simbol anotasi: // = jeda hening 2 detik | TEBAL = penekanan vokal keras | (bisik) = intonasi lirih\n\n"
        "[LINDNER holds checkbook open. WALTER stands straight, looking down at his son TRAVIS.]\n"
        "WALTER: Well, Mr. Lindner... my father... my father was a laborer all his life. He worked sixty hours a week until his back was crooked and his knuckles bled. (Tambahkan anotasi vokalmu di sini)\n"
        "..................................................................................................................................................................................................................."
    )
    add_writing_box(doc, annot_desc, num_lines=6)

    add_activity_heading(doc, "4", "Monologue Performance Planning & Rehearsal (ICoT)", 
                         "Rencanakan rencana pementasan monolog mandiri 3-4 menit yang akan kamu bawakan di panggung kelas!")
    prompt_desc = (
        "Judul Naskah Pilihan: ..................................................... Tokoh yang Diperankan: .............................\n"
        "Rancangan Kostum/Properti Sederhana: ........................................................................................\n"
        "Deskripsi Blocking Gerak Panggung (Kapan melangkah, berlutut, atau berhenti menatap penonton):"
    )
    add_writing_box(doc, prompt_desc, num_lines=8)

    add_activity_heading(doc, "5", "Theatrical Performance Rubric Checklist", 
                         "Gunakan rubrik seni peran analitik ini untuk mengevaluasi penampilan rekan panggungmu!")
    c_headers = ["Aspek Kinerja Peran", "Skor 4 (Sangat Baik)", "Skor 3 (Baik)", "Skor 2 (Cukup)", "Skor 1 (Kurang)"]
    c_rows = [
        ["Modulasi & Proyeksi Suara", "Vokal sangat lantang, intonasi dinamis, artikulasi sangat jelas.", "Vokal jelas, intonasi cukup variatif.", "Vokal agak datar, artikulasi kurang tegas.", "Vokal tidak terdengar jelas di baris belakang."],
        ["Penghayatan Subteks", "Sangat meresapi emosi tokoh, tatapan mata hidup.", "Menghayati emosi tokoh dengan meyakinkan.", "Kurang mendalam menghayati batin tokoh.", "Sekadar menghafal dan membaca baris dialog."],
        ["Blocking & Bahasa Tubuh", "Gerak tubuh selaras dengan teks dan petunjuk panggung.", "Gerak tubuh luwes dan sesuai adegan.", "Gerak tubuh kaku dan ragu-ragu di panggung.", "Hanya berdiri diam tanpa ekspresi fisik."]
    ]
    docx_helper.add_styled_table(doc, c_headers, c_rows, [Inches(1.8), Inches(1.2), Inches(1.2), Inches(1.2), Inches(1.1)])

    fpath = os.path.join(TARGET_DIR, "LKPD_TP_12.TL.3_Dramatic_Script_and_Performance.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# LKPD 4: TP 12.TL.4 - PERSUASIVE ORATORY & PARLIAMENTARY DEBATE
# ==============================================================================
def create_lkpd_4():
    doc = docx_helper.create_styled_doc()
    add_student_id_box(doc, "4", "Persuasive Oratory & Parliamentary Debate", "TP 12.TL.4")

    docx_helper.add_callout(
        doc,
        "Tujuan Pembelajaran:",
        "Setelah menyelesaikan lembar kerja ini, kamu mampu mengaplikasikan retorika klasik (Ethos, Pathos, Logos), menggunakan perangkat gaya bahasa (parallelism, antithesis, tricolon), merumuskan kasus debat formal, dan menyampaikan pidato oratori persuasif 3-4 menit.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_activity_heading(doc, "1", "Rhetorical Triangle & Oratorical Vocabulary (BKoF)", 
                         "Tentukan pilar retorika Aristoteles (Ethos, Pathos, atau Logos) yang paling dominan pada pernyataan berikut!")
    v_headers = ["No", "Pernyataan Oratorik Cuplikan Pidato", "Pilar Retorika Dominan"]
    v_rows = [
        ["1", "'Subsidies for fossil fuels reached an astronomical seven trillion dollars last year.'", "[  ] LOGOS"],
        ["2", "'Consider the child in Jakarta wading through monsoon waters to reach a flooded classroom.'", "[  ] PATHOS"],
        ["3", "'As elected heads of state, your legitimacy rests on your moral fidelity to protect the vulnerable.'", "[  ] ETHOS"],
        ["4", "'The Intergovernmental Panel on Climate Change established that 1.5°C is an irreversible tipping point.'", "[  ] LOGOS"],
        ["5", "'Posterity will look back at this podium fifty years hence and judge whether you had moral courage.'", "[  ] ETHOS / PATHOS"]
    ]
    docx_helper.add_styled_table(doc, v_headers, v_rows, [Inches(0.6), Inches(4.5), Inches(1.4)])

    add_activity_heading(doc, "2", "Rhetorical Device Deconstruction (MoT)", 
                         "Bacalah Teks Otentik 1 ('The Arithmetic of Survival') dan Teks 3. Temukan contoh perangkat retoris berikut pada naskah!")
    s_headers = ["Nama Perangkat Retorika", "Definisi Singkat", "Contoh Kutipan Langsung dari Naskah Pidato"]
    s_rows = [
        ["Tricolon (Pola Tiga)", "Pengulangan tiga frasa berturut-turut untuk ritme penekanan.", 
         "'The laws of physics do not compromise... permafrost does not wait... tides do not retreat...'"],
        ["Antithesis (Penyandingan Kontras)", "Penyandingan dua ide yang berlawanan dalam struktur seimbang.", 
         "..................................................................................................."],
        ["Parallelism (Keselarasan Struktur)", "Pengulangan struktur gramatikal yang identik secara beruntun.", 
         "..................................................................................................."]
    ]
    docx_helper.add_styled_table(doc, s_headers, s_rows, [Inches(1.8), Inches(2.2), Inches(2.5)])

    add_activity_heading(doc, "3", "Collaborative Debate Case Building & Rebuttal (JCoT)", 
                         "Bersama tim debatmu (3 orang), susunlah kasus mosi: 'This House Would Ban Autonomous Weapon Systems'. Siapkan poin sanggahan taktis terhadap lawan!")
    m_headers = ["Pilar Kasus Tim Afirmasi", "Bukti & Alasan Logis", "Antisipasi Argumen Oposisi & Sanggahan (Rebuttal)"]
    m_rows = [
        ["1. Moral Accountability", "Algoritma tidak memiliki nurani atau belas kasih.", "Oposisi: AI lebih presisi -> Rebuttal: Kecepatan tanpa etika adalah bencana."],
        ["2. Technical Fallibility", "Sensor kamera rentan terhadap kabut/spoofing.", "..................................................................................."],
        ["3. Barrier to War", "Negara lebih mudah memulai agresi tanpa korban tentara.", "..................................................................................."]
    ]
    docx_helper.add_styled_table(doc, m_headers, m_rows, [Inches(1.8), Inches(2.35), Inches(2.35)])

    add_activity_heading(doc, "4", "Persuasive Oratorical Speech Drafting (ICoT)", 
                         "Susunlah naskah pidato oratori persuasif 300-400 kata dengan memilih salah satu mosi global!")
    prompt_desc = (
        "Topik Pilihan: (A) Youth Climate Action, (B) Ethical Regulation of AI, atau (C) Equitable Space Resources.\n"
        "Tuliskan naskah pidato lengkap dengan pembuka (Exordium), pembuktian (Logos), gugahan moral (Ethos/Pathos), dan penutup (Peroratio):"
    )
    add_writing_box(doc, prompt_desc, num_lines=12)

    add_activity_heading(doc, "5", "Public Speaking Adjudication Checklist", 
                         "Evaluasi pidato rekan oratormu menggunakan standar penilaian debat dan oratori internasional!")
    c_headers = ["Kriteria Penilaian Oratori Lisan", "Ya", "Belum", "Catatan Konstruktif Juri"]
    c_rows = [
        ["Membuka pidato dengan hook yang memikat perhatian seluruh audiens.", "", "", ""],
        ["Argumen didukung bukti data konkret (Logos) dan prinsip moral (Ethos).", "", "", ""],
        ["Menggunakan minimal 1 tricolon dan 1 antithesis dengan pengucapan berirama.", "", "", ""],
        ["Kontak mata merata ke seluruh ruangan tanpa terus-menerus menatap teks.", "", "", ""],
        ["Mengakhiri pidato dengan seruan tindakan (call to action) yang membakar semangat.", "", "", ""]
    ]
    docx_helper.add_styled_table(doc, c_headers, c_rows, [Inches(3.5), Inches(0.6), Inches(0.6), Inches(1.8)])

    fpath = os.path.join(TARGET_DIR, "LKPD_TP_12.TL.4_Persuasive_Oratory_and_Debate.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# LKPD 5: TP 12.TL.5 - ADMISSIONS & CAREER DOSSIER
# ==============================================================================
def create_lkpd_5():
    doc = docx_helper.create_styled_doc()
    add_student_id_box(doc, "5", "Professional Admissions & Career Dossier", "TP 12.TL.5")

    docx_helper.add_callout(
        doc,
        "Tujuan Pembelajaran:",
        "Setelah menyelesaikan lembar kerja ini, kamu mampu merancang Executive Cover Letter bisnis formal, menyusun Curriculum Vitae ATS-compliant berbasis Google X-Y-Z formula, serta menulis Personal Statement beasiswa dunia sepanjang 500-600 kata.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_activity_heading(doc, "1", "High-Impact Action Verbs & ATS Concepts (BKoF)", 
                         "Ubahlah kalimat tugas pasif di kolom kiri menjadi peluru capaian (bullet) bertenaga dengan kata kerja aksi!")
    v_headers = ["No", "Kalimat Deskripsi Pasif / Klise", "Revisi dengan Dynamic Action Verb"]
    v_rows = [
        ["1", "I was responsible for making the club Instagram posts.", "Engineered social media visual campaigns..."],
        ["2", "I helped my team build a wave energy generator.", "Co-authored and designed a piezo-electric prototype..."],
        ["3", "I had to organize the student council webinars.", "Orchestrated 5 international webinars engaging 400+ attendees..."],
        ["4", "I worked on cleaning up wastewater data.", "Cleaned, standardized, and analyzed econometric emissions data..."],
        ["5", "I was the leader of the science debate team.", "Spearheaded a 6-member competitive varsity debate squad..."]
    ]
    docx_helper.add_styled_table(doc, v_headers, v_rows, [Inches(0.6), Inches(2.8), Inches(3.1)])

    add_activity_heading(doc, "2", "Personal Statement Narrative Arc Analysis (MoT)", 
                         "Bedahlah Teks Otentik 1 ('Model Personal Statement: Sustainable Energy Access') dan analisislah struktur narasinya!")
    s_headers = ["Bagian Narasi Esai", "Strategi Penulisan yang Digunakan", "Contoh Fakta dari Teks Model"]
    s_rows = [
        ["The Narrative Hook", "Menghubungkan observasi personal masa kecil dengan sains.", "Pemadaman listrik pesisir dan deburan ombak pemecah ombak."],
        ["Academic Trajectory", "Membuktikan rasa ingin tahu melalui proyek ilmiah nyata.", "......................................................................................."],
        ["Institutional Fit", "Menjelaskan mengapa memilih universitas target tersebut.", "......................................................................................."],
        ["Future Impact / Vision", "Menyatakan komitmen kontribusi nyata bagi negara asal.", "......................................................................................."]
    ]
    docx_helper.add_styled_table(doc, s_headers, s_rows, [Inches(1.8), Inches(2.2), Inches(2.5)])

    add_activity_heading(doc, "3", "Collaborative Google X-Y-Z Resume Workshop (JCoT)", 
                         "Bersama rekanmu, terapkan rumus Google: 'Accomplished [X] measured by [Y] by doing [Z]' pada 2 pengalaman prestasimu!")
    m_headers = ["Komponen Rumus X-Y-Z", "Rancangan Prestasi Siswa A", "Rancangan Prestasi Siswa B"]
    m_rows = [
        ["Accomplished [X] (Hasil Kunci)", "Juara 2 Lomba Robotika Regional", "..........................................................."],
        ["Measured by [Y] (Metrik Ukuran)", "Menyelesaikan lintasan dalam 3.4 menit", "..........................................................."],
        ["By doing [Z] (Tindakan Spesifik)", "Memprogram ulang script kamera smartphone", "..........................................................."]
    ]
    docx_helper.add_styled_table(doc, m_headers, m_rows, [Inches(2.1), Inches(2.2), Inches(2.2)])

    add_activity_heading(doc, "4", "Executive Cover Letter Drafting (ICoT)", 
                         "Susunlah sebuah Executive Cover Letter profesional (250-300 kata) ditujukan kepada program magang PBB atau universitas dunia pilihanmu!")
    prompt_desc = (
        "Penerima: Selection Committee for Global Youth Ambassador / University Admissions Office\n"
        "Tuliskan naskah surat lengkap (Salutation, Purpose, Quantitative Strengths, Cultural Fit, Call to Action, Sign-off):"
    )
    add_writing_box(doc, prompt_desc, num_lines=12)

    add_activity_heading(doc, "5", "ATS Compatibility & Dossier Review Checklist", 
                         "Periksalah kelayakan berkas dossier rekanmu sebelum diunggah ke portal seleksi resmi!")
    c_headers = ["Indikator Kelayakan Dossier", "Ya", "Belum", "Catatan Reviewer Profesional"]
    c_rows = [
        ["Cover letter menggunakan register bahasa Inggris bisnis formal yang persuasif.", "", "", ""],
        ["Resume satu kolom polos bebas tabel rumit, ikon, atau gambar yang membingungkan ATS.", "", "", ""],
        ["Setiap poin rekam jejak pada resume memuat metrik kuantitatif terukur (angka/persen).", "", "", ""],
        ["Personal statement memiliki benang merah yang logis antara masa lalu dan cita-cita studi.", "", "", ""],
        ["Bebas dari kesalahan penulisan nama institusi, ejaan kata, dan tanda baca.", "", "", ""]
    ]
    docx_helper.add_styled_table(doc, c_headers, c_rows, [Inches(3.5), Inches(0.6), Inches(0.6), Inches(1.8)])

    fpath = os.path.join(TARGET_DIR, "LKPD_TP_12.TL.5_Admissions_and_Career_Dossier.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# LKPD 6: TP 12.TL.6 - JOB INTERVIEW & ORAL DEFENSE
# ==============================================================================
def create_lkpd_6():
    doc = docx_helper.create_styled_doc()
    add_student_id_box(doc, "6", "Executive Interview & Academic Oral Defense", "TP 12.TL.6")

    docx_helper.add_callout(
        doc,
        "Tujuan Pembelajaran:",
        "Setelah menyelesaikan lembar kerja ini, kamu mampu merespons pertanyaan perilaku (behavioral queries) menggunakan metode STAR, mempertahankan metodologi karya ilmiah secara santun di hadapan dewan penguji, serta mempraktikkan simulasi ujian lisan 4-5 menit.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_activity_heading(doc, "1", "STAR Architecture & Tactical Delay Formulae (BKoF)", 
                         "Pasangkanlah komponen STAR di kolom kiri dengan fungsi penjelasannya di kolom kanan!")
    v_headers = ["No", "Komponen STAR", "Fungsi Penjelasan dalam Wawancara (Tuliskan Huruf Pasangan)"]
    v_rows = [
        ["1", "Situation (S)", "[  ] A. Tindakan inisiatif spesifik yang dilakukan secara mandiri (60% waktu)."],
        ["2", "Task (T)", "[  ] B. Hasil kuantitatif nyata dan refleksi pendewasaan diri (25% waktu)."],
        ["3", "Action (A)", "[  ] C. Konteks latar belakang krisis atau proyek yang dihadapi (10% waktu)."],
        ["4", "Result (R)", "[  ] D. Tanggung jawab atau target yang harus diselesaikan (5% waktu)."],
        ["5", "Tactical Filler", "[  ] E. Kalimat santun untuk membeli waktu berpikir tanpa berkata 'uhm'."]
    ]
    docx_helper.add_styled_table(doc, v_headers, v_rows, [Inches(0.6), Inches(2.2), Inches(3.7)])

    add_activity_heading(doc, "2", "Oral Defense Transcript Deconstruction (MoT)", 
                         "Telaahlah Teks Otentik 2 ('Academic Oral Defense Protocol: Responding to Methodological Interrogation') dan analisislah respon kandidat!")
    s_headers = ["Pertanyaan Kritis Penguji", "Strategi Validasi Santun Kandidat", "Argumen Metodologi Penyelamat"]
    s_rows = [
        ["'Given your small sample size of thirty, isn't it invalid to generalize to all secondary schools?'", 
         "'That is an incisive critique, and I acknowledge that demographic generalizability is a limitation...'", 
         "Menjelaskan bahwa tujuan riset adalah fenomenologis kualitatif, diperkuat triangulasi video screen-capture 4 jam."]
    ]
    docx_helper.add_styled_table(doc, s_headers, s_rows, [Inches(2.5), Inches(2.0), Inches(2.0)])

    add_activity_heading(doc, "3", "Collaborative Mock Panel Practice (JCoT)", 
                         "Bersama kelompokmu (4 orang), simulasikan sesi wawancara STAR untuk pertanyaan: 'Describe a project where your initial plan failed, and how you recovered!'")
    m_headers = ["Elemen Jawaban STAR", "Draf Catatan Respon Siswa (Persiapan Lisan)"]
    m_rows = [
        ["Situation & Task", "Sensor robotik utama rusak 48 jam sebelum kompetisi regional."],
        ["Action (Gunakan 'I')", "..................................................................................................................."],
        ["Result (Data & Pelajaran)", "..................................................................................................................."]
    ]
    docx_helper.add_styled_table(doc, m_headers, m_rows, [Inches(2.5), Inches(4.0)])

    add_activity_heading(doc, "4", "Viva Voce Defense Scripting (ICoT)", 
                         "Tuliskan naskah jawaban pertahanan lisanmu terhadap kemungkinan pertanyaan tersulit yang akan diajukan dewan penguji terhadap makalah ilmiahmu!")
    prompt_desc = (
        "Kemungkinan Pertanyaan Kritis Penguji: ...................................................................................\n"
        "Tuliskan naskah respon lisanmu dengan menggunakan formula: (1) Validasi santun, (2) Klarifikasi batasan lingkup, (3) Bukti triangulasi data:"
    )
    add_writing_box(doc, prompt_desc, num_lines=10)

    add_activity_heading(doc, "5", "Oral Defense & Interview Evaluation Rubric", 
                         "Evaluasi performa lisan rekanmu saat simulasi di hadapan panel meja penguji!")
    c_headers = ["Aspek Pengamatan Ujian Lisan", "Skor 4 (Sangat Baik)", "Skor 3 (Baik)", "Skor 2 (Cukup)", "Skor 1 (Kurang)"]
    c_rows = [
        ["Penerapan Metode STAR", "Runtut sempurna, fokus aksi pribadi, hasil terukur.", "Runtut, aksi jelas, ada hasil.", "Cenderung bertele-tele di latar belakang.", "Tidak terstruktur, melompat-lompat."],
        ["Ketenangan & Wibawa", "Kontak mata tenang, jeda taktis mantap tanpa 'uhm'.", "Cukup tenang, sedikit keraguan.", "Gugup, sering menggunakan filler kata 'uhm'.", "Panik dan terdiam lama."],
        ["Register Bahasa Inggris", "Bahasa Inggris formal presisi, istilah akademik tepat.", "Bahasa formal, sesekali agak santai.", "Banyak kalimat informal/slang sehari-hari.", "Tata bahasa sering rancu."]
    ]
    docx_helper.add_styled_table(doc, c_headers, c_rows, [Inches(1.8), Inches(1.2), Inches(1.2), Inches(1.2), Inches(1.1)])

    fpath = os.path.join(TARGET_DIR, "LKPD_TP_12.TL.6_Job_Interview_and_Oral_Defense.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# LKPD 7: TP 12.TL.7 - RESEARCH ABSTRACT & INFOGRAPHIC
# ==============================================================================
def create_lkpd_7():
    doc = docx_helper.create_styled_doc()
    add_student_id_box(doc, "7", "Scientific Abstract & Infographic Symposium", "TP 12.TL.7")

    docx_helper.add_callout(
        doc,
        "Tujuan Pembelajaran:",
        "Setelah menyelesaikan lembar kerja ini, kamu mampu merangkum karya tulis penelitian ke dalam naskah abstrak terstruktur IMRaD (maksimal 250 kata), menguasai nominalisasi ilmiah, merancang poster infografis tiga kolom berprinsip 'Ten-Second Rule', dan mempresentasikannya dalam simposium ilmiah.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_activity_heading(doc, "1", "Scientific Nominalization Transformation (BKoF)", 
                         "Ubahlah kalimat proses aktif di kolom kiri menjadi klausa nomina padat (nominalization)!")
    v_headers = ["No", "Kalimat Proses Aktif (Wordy)", "Hasil Nominalisasi Saintifik yang Padat (Dense)"]
    v_rows = [
        ["1", "When nanoparticles absorbed heavy metals, they reacted quickly.", "Rapid heavy metal adsorption occurred..."],
        ["2", "Because the heat was absorbed during the reaction, the beaker cooled.", "An endothermic reaction induced rapid cooling..."],
        ["3", "Factories contaminated the river when they discharged untreated dye.", "Untreated dye effluent discharge contaminated the river..."],
        ["4", "Scientists characterized the crystals using X-ray diffraction.", "Crystalline characterization via X-ray diffraction confirmed..."],
        ["5", "The biosorbent removed toxic chromium by 91% across 5 cycles.", "Chromium sequestration reached 91% across 5 regeneration cycles..."]
    ]
    docx_helper.add_styled_table(doc, v_headers, v_rows, [Inches(0.6), Inches(2.8), Inches(3.1)])

    add_activity_heading(doc, "2", "IMRaD Abstract Structure Mapping (MoT)", 
                         "Bedahlah Teks Otentik 1 ('Model Structured Scientific Abstract 245 kata') dan hitung proporsi alokasi katanya!")
    s_headers = ["Elemen Struktur IMRaD", "Fungsi Utama Bagian", "Jumlah Kata pada Teks 1", "Target Kata Siswa"]
    s_rows = [
        ["[Introduction / Background]", "Latar belakang masalah & tujuan penelitian", "45 kata (18%)", "35 - 45 kata"],
        ["[Methodology / Methods]", "Alat karakterisasi & parameter eksperimen", "55 kata (22%)", "50 - 60 kata"],
        ["[Results / Key Findings]", "Data kuantitatif empiris & model kinetika", "110 kata (45%)", "100 - 110 kata"],
        ["[Conclusion & Significance]", "Dampak terapan bagi industri/masyarakat", "35 kata (15%)", "30 - 35 kata"]
    ]
    docx_helper.add_styled_table(doc, s_headers, s_rows, [Inches(1.8), Inches(2.2), Inches(1.3), Inches(1.2)])

    add_activity_heading(doc, "3", "Collaborative Poster Layout Design (JCoT)", 
                         "Rancanglah sketsa tata letak tiga kolom poster infografismu bersama rekan sebangkumu!")
    m_headers = ["Kolom 1: Problem & Purpose", "Kolom 2: Methods & Workflow", "Kolom 3: Key Results & Impact"]
    m_rows = [
        ["• Active Headline\n• Background Context\n• Research Hypothesis", 
         "• Flowchart Sintesis\n• Diagram Alat Uji\n• Variabel Kontrol", 
         "• Grafik Batang/Garis\n• Metrik Persentase\n• Rekomendasi Terapan"]
    ]
    docx_helper.add_styled_table(doc, m_headers, m_rows, [Inches(2.15), Inches(2.15), Inches(2.2)])

    add_activity_heading(doc, "4", "250-Word Structured Abstract Drafting (ICoT)", 
                         "Tuliskan draf abstrak penelitian terstrukturmu di bawah ini (Wajib mematuhi batas ketat 200-250 kata):")
    prompt_desc = (
        "Judul Penelitian: ........................................................................................................\n"
        "Tuliskan naskah abstrak lengkap [Background], [Methods], [Results], dan [Conclusion]:"
    )
    add_writing_box(doc, prompt_desc, num_lines=12)

    add_activity_heading(doc, "5", "Symposium Poster & Abstract Peer-Review", 
                         "Gunakan ceklis ini saat meninjau stan poster rekanmu di ajang simposium ilmiah siswa!")
    c_headers = ["Kriteria Mutu Simposium Ilmiah", "Ya", "Belum", "Catatan Evaluasi Pengunjung"]
    c_rows = [
        ["Abstrak mematuhi batas maksimal 250 kata dan format terstruktur IMRaD.", "", "", ""],
        ["Headline poster berbasis kesimpulan aktif dan memuat angka kuantitatif.", "", "", ""],
        ["Poster mematuhi prinsip 'Ten-Second Rule' dengan hierarki visual tiga kolom yang bersih.", "", "", ""],
        ["Jumlah teks di dalam poster tidak melebihi 400 kata agar tidak padat membingungkan.", "", "", ""],
        ["Presenter mampu menyampaikan flash talk 2 menit secara lancar dan memikat.", "", "", ""]
    ]
    docx_helper.add_styled_table(doc, c_headers, c_rows, [Inches(3.5), Inches(0.6), Inches(0.6), Inches(1.8)])

    fpath = os.path.join(TARGET_DIR, "LKPD_TP_12.TL.7_Research_Abstract_and_Infographic.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# LKPD 8: TP 12.TL.8 - STANDARDIZED READINESS (B2+)
# ==============================================================================
def create_lkpd_8():
    doc = docx_helper.create_styled_doc()
    add_student_id_box(doc, "8", "Standardized English Proficiency Benchmark", "TP 12.TL.8")

    docx_helper.add_callout(
        doc,
        "Tujuan Pembelajaran:",
        "Setelah menyelesaikan lembar kerja ini, kamu mampu menguasai strategi membaca cepat teks akademik panjang (skimming & scanning), mengidentifikasi inferensi saintifik tersirat, mendeskripsikan data grafik analitis (IELTS Writing Task 1) dalam 20 menit, dan mengukur kesiapan kemahiran bahasa Inggris internasional (CEFR B2+).",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_activity_heading(doc, "1", "Standardized Trend Lexicon & Prepositions (BKoF)", 
                         "Lengkapilah kalimat deskripsi data grafik berikut dengan kosakata tren dan preposisi yang tepat!")
    v_headers = ["No", "Kalimat Tren Data Grafik", "Pilihan Kata Tepat"]
    v_rows = [
        ["1", "Clean energy investments surged __________ $650 billion in 2015 to $1,800 billion in 2025.", "A. By"],
        ["2", "Fossil fuel expenditures plummeted __________ nearly 25% due to oil price corrections.", "B. At"],
        ["3", "In 2020, renewable spending reached parity with fossil fuels __________ $900 billion each.", "C. From"],
        ["4", "Over the decade, renewable investments witnessed an exponential upward __________.", "D. Plateaued"],
        ["5", "From 2023 onward, fossil fuel capital allocations __________ at roughly $1,050 billion.", "E. Trajectory"]
    ]
    docx_helper.add_styled_table(doc, v_headers, v_rows, [Inches(0.6), Inches(4.3), Inches(1.6)])

    add_activity_heading(doc, "2", "Academic Reading Strategies & Inferences (MoT)", 
                         "Bacalah Teks Otentik 1 ('Paleoclimatology and Ice Core Chronicles') dengan teknik skimming cepat (3 menit), lalu tentukan status pernyataan berikut (True / False / Not Given)!")
    s_headers = ["Pernyataan Evaluasi Inferensi", "Status (T/F/NG)", "Bukti Kalimat dari Teks"]
    s_rows = [
        ["1. Ancient air bubbles trapped in Antarctic ice provide records dating back 800,000 years.", "TRUE", "Paragraf 1, kalimat 1"],
        ["2. Heavier oxygen isotopes evaporate more rapidly in colder global temperatures.", "..................", "..................................................................."],
        ["3. Current atmospheric CO2 levels have exceeded 420 ppm due to human fossil fuel combustion.", "..................", "..................................................................."],
        ["4. Scientists expect global temperatures to drop naturally over the next 50 years.", "..................", "..................................................................."]
    ]
    docx_helper.add_styled_table(doc, s_headers, s_rows, [Inches(3.2), Inches(1.1), Inches(2.2)])

    add_activity_heading(doc, "3", "Collaborative Writing Task 1 Overview Formulation (JCoT)", 
                         "Telaahlah grafik Teks Otentik 2 bersama rekanmu. Susunlah kalimat Overview utama (paragraf terpenting penentu Band 7.0+)! ")
    m_headers = ["Elemen Paragraf Overview", "Rancangan Kalimat Bersama (JCoT)"]
    m_rows = [
        ["Tren Utama Energi Bersih", "Overall, clean technology investments experienced an exponential upward surge over the decade..."],
        ["Tren Pembanding Bahan Bakar Fosil", "Conversely, fossil fuel expenditures exhibited substantial volatility followed by stagnation..."],
        ["Titik Temu Kunci (Key Milestones)", "..................................................................................................................."]
    ]
    docx_helper.add_styled_table(doc, m_headers, m_rows, [Inches(2.5), Inches(4.0)])

    add_activity_heading(doc, "4", "Timed Academic Writing Task 1 (ICoT - 20 Minutes)", 
                         "Tuliskan laporan analisis data grafik 150-180 kata berdasarkan data investasi energi bersih global (Teks 2). Waktu pengerjaan: 20 menit tepat!")
    prompt_desc = (
        "The chart illustrates global annual capital expenditures across clean energy technologies compared to fossil fuels (2015-2025).\n"
        "Tuliskan laporan analitis lengkap (Introduction Paraphrase -> Overall Overview -> Body Paragraph 1 -> Body Paragraph 2):"
    )
    add_writing_box(doc, prompt_desc, num_lines=12)

    add_activity_heading(doc, "5", "Official IELTS Band Descriptors Checklist", 
                         "Gunakan kriteria penilaian resmi IELTS Band 7.5+ untuk mengevaluasi laporan rekanmu!")
    c_headers = ["Kriteria Band Descriptors", "Ya", "Belum", "Catatan Koreksi Penilai"]
    c_rows = [
        ["Task Achievement: Menyajikan overview yang jelas dan memilih data pembanding yang akurat.", "", "", ""],
        ["Coherence & Cohesion: Paragraf tertata logis dengan transisi bervariasi tanpa repetisi.", "", "", ""],
        ["Lexical Resource: Menggunakan kosakata tren tingkat lanjut dan rasio matematis yang presisi.", "", "", ""],
        ["Grammatical Range: Menggunakan variasi kalimat majemuk, participial clause, dan preposisi akurat.", "", "", ""],
        ["Panjang naskah memenuhi batas minimal 150 kata dalam rentang waktu 20 menit.", "", "", ""]
    ]
    docx_helper.add_styled_table(doc, c_headers, c_rows, [Inches(3.5), Inches(0.6), Inches(0.6), Inches(1.8)])

    fpath = os.path.join(TARGET_DIR, "LKPD_TP_12.TL.8_Standardized_English_Proficiency_Readiness.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


def main():
    print("Generating student worksheets (LKPD) for 5-Modul Ajar XII SMA TL...")
    create_lkpd_1()
    create_lkpd_2()
    create_lkpd_3()
    create_lkpd_4()
    create_lkpd_5()
    create_lkpd_6()
    create_lkpd_7()
    create_lkpd_8()
    print("All 8 LKPD documents successfully generated!")

if __name__ == "__main__":
    main()

