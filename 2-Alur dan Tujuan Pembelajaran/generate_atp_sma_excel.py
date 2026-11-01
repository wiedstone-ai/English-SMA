# -*- coding: utf-8 -*-
"""
Script to generate official ALUR DAN TUJUAN PEMBELAJARAN (ATP) BAHASA INGGRIS SMA
in a single, beautifully formatted Excel workbook with Diagram Alur and complete TP tables.
Conforms to:
- Permendikdasmen No. 13 Tahun 2025
- Capaian Pembelajaran BSKAP Kemendikdasmen No. 046/H/KR/2025 & Kepka 020/2026
- Panduan Mata Pelajaran Bahasa Inggris Puskurjar Kemendikdasmen Revisi 3 (Sept 2025)
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Typography & Palette
FONT_FAMILY = "Segoe UI"
FONT_TITLE = Font(name=FONT_FAMILY, size=14, bold=True, color="FFFFFF")
FONT_SUBTITLE = Font(name=FONT_FAMILY, size=10, italic=True, color="E2E8F0")
FONT_SEC_HEADER = Font(name=FONT_FAMILY, size=11, bold=True, color="FFFFFF")
FONT_CARD_TITLE = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")
FONT_TBL_HEADER = Font(name=FONT_FAMILY, size=9, bold=True, color="FFFFFF")
FONT_TBL_SUBHEADER = Font(name=FONT_FAMILY, size=8, bold=True, color="0F2942")
FONT_DATA = Font(name=FONT_FAMILY, size=9, color="1E293B")
FONT_DATA_BOLD = Font(name=FONT_FAMILY, size=9, bold=True, color="1E293B")
FONT_DATA_CENTER = Font(name=FONT_FAMILY, size=9, color="1E293B")
FONT_DATA_NUMBER = Font(name=FONT_FAMILY, size=9, bold=True, color="0F2942")
FONT_ITALIC_DESC = Font(name=FONT_FAMILY, size=8, italic=True, color="475569")
FONT_TOTAL = Font(name=FONT_FAMILY, size=9, bold=True, color="0F2942")

FONT_DIAGRAM_NODE = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")
FONT_DIAGRAM_DESC = Font(name=FONT_FAMILY, size=8, color="1E293B")
FONT_DIAGRAM_ARROW = Font(name=FONT_FAMILY, size=12, bold=True, color="2563EB")

FILL_NAVY = PatternFill(start_color="0F2942", end_color="0F2942", fill_type="solid")
FILL_ROYAL = PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid")
FILL_BLUE_HEADER = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
FILL_LIGHT_BLUE = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
FILL_VERY_LIGHT = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
FILL_ZEBRA = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
FILL_TOTAL = PatternFill(start_color="BFDBFE", end_color="BFDBFE", fill_type="solid")
FILL_ACCENT_AMBER = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
FILL_ACCENT_GREEN = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
FILL_NODE_GREEN = PatternFill(start_color="059669", end_color="059669", fill_type="solid")
FILL_NODE_BLUE = PatternFill(start_color="1D4ED8", end_color="1D4ED8", fill_type="solid")
FILL_NODE_PURPLE = PatternFill(start_color="6D28D9", end_color="6D28D9", fill_type="solid")
FILL_NODE_AMBER = PatternFill(start_color="D97706", end_color="D97706", fill_type="solid")

BORDER_THIN = Side(style='thin', color="CBD5E1")
BORDER_MEDIUM = Side(style='medium', color="1E40AF")
BORDER_DOUBLE = Side(style='double', color="0F2942")
CELL_BORDER = Border(left=BORDER_THIN, right=BORDER_THIN, top=BORDER_THIN, bottom=BORDER_THIN)
BOX_BORDER = Border(left=BORDER_MEDIUM, right=BORDER_MEDIUM, top=BORDER_MEDIUM, bottom=BORDER_MEDIUM)
TOTAL_BORDER = Border(left=BORDER_THIN, right=BORDER_THIN, top=BORDER_THIN, bottom=BORDER_DOUBLE)

ALIGN_CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
ALIGN_LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal='right', vertical='center')
ALIGN_TOP_LEFT = Alignment(horizontal='left', vertical='top', wrap_text=True)
ALIGN_TOP_CENTER = Alignment(horizontal='center', vertical='top', wrap_text=True)

def apply_title(ws, title, subtitle, max_col):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
    
    c1 = ws.cell(row=1, column=1, value=title)
    c1.font = FONT_TITLE
    c1.fill = FILL_NAVY
    c1.alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 28
    
    c2 = ws.cell(row=2, column=1, value=subtitle)
    c2.font = FONT_SUBTITLE
    c2.fill = FILL_ROYAL
    c2.alignment = ALIGN_CENTER
    ws.row_dimensions[2].height = 20

    for col in range(1, max_col + 1):
        ws.cell(row=1, column=col).fill = FILL_NAVY
        ws.cell(row=2, column=col).fill = FILL_ROYAL

def add_school_identity(ws, school_data, start_row=4):
    for i, (label, val) in enumerate(school_data):
        r = start_row + i
        ws.cell(row=r, column=1, value=label).font = FONT_DATA_BOLD
        ws.cell(row=r, column=1).alignment = ALIGN_LEFT
        ws.cell(row=r, column=2, value=": " + str(val)).font = FONT_DATA
        ws.cell(row=r, column=2).alignment = ALIGN_LEFT
        ws.row_dimensions[r].height = 18
    return start_row + len(school_data) + 1

def add_signatures(ws, start_row, max_col, city="Jakarta"):
    r = start_row + 1
    ws.merge_cells(start_row=r, start_column=max_col-3, end_row=r, end_column=max_col)
    ws.cell(row=r, column=max_col-3, value=f"{city}, Juli 2025").font = FONT_DATA
    ws.cell(row=r, column=max_col-3).alignment = ALIGN_CENTER
    
    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    ws.cell(row=r, column=1, value="Mengetahui,").font = FONT_DATA
    ws.cell(row=r, column=1).alignment = ALIGN_CENTER
    
    ws.merge_cells(start_row=r, start_column=max_col-3, end_row=r, end_column=max_col)
    ws.cell(row=r, column=max_col-3, value="Guru Pengampu Mata Pelajaran,").font = FONT_DATA
    ws.cell(row=r, column=max_col-3).alignment = ALIGN_CENTER

    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    ws.cell(row=r, column=1, value="Kepala SMA .................................").font = FONT_DATA
    ws.cell(row=r, column=1).alignment = ALIGN_CENTER

    r += 4
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    ws.cell(row=r, column=1, value="( .................................................... )").font = FONT_DATA_BOLD
    ws.cell(row=r, column=1).alignment = ALIGN_CENTER

    ws.merge_cells(start_row=r, start_column=max_col-3, end_row=r, end_column=max_col)
    ws.cell(row=r, column=max_col-3, value="( .................................................... )").font = FONT_DATA_BOLD
    ws.cell(row=r, column=max_col-3).alignment = ALIGN_CENTER

    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
    ws.cell(row=r, column=1, value="NIP. .............................................").font = FONT_DATA
    ws.cell(row=r, column=1).alignment = ALIGN_CENTER

    ws.merge_cells(start_row=r, start_column=max_col-3, end_row=r, end_column=max_col)
    ws.cell(row=r, column=max_col-3, value="NIP. .............................................").font = FONT_DATA
    ws.cell(row=r, column=max_col-3).alignment = ALIGN_CENTER

def auto_fit_cols(ws, min_w=10, max_w=50, skip_rows=(1, 2)):
    ws.views.sheetView[0].showGridLines = True
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            if cell.row in skip_rows:
                continue
            v = str(cell.value or '')
            if '\n' in v:
                l_len = max(len(line) for line in v.split('\n'))
            else:
                l_len = len(v)
            if l_len > max_len:
                max_len = l_len
        ws.column_dimensions[col_letter].width = max(min(max_len + 3, max_w), min_w)


# ==============================================================================
# DATA ATP (TUJUAN PEMBELAJARAN, KKTP, GENRE, ELEMEN, PROFIL PANCASILA, JP)
# ==============================================================================

ATP_X_WAJIB = [
    # (No, Kode, Elemen, Rumusan_TP, KKTP_Indikator, Materi_Esensial, Genre, Alur_Tahap, JP, Profil_Pancasila, Asesmen)
    (1, "TP 10.1", "Menyimak - Berbicara & Membaca - Memirsa",
     "Menganalisis fungsi sosial, struktur teks, dan unsur kebahasaan teks deskriptif lisan dan tulis tentang atlet berprestasi dan destinasi wisata dunia secara kritis dan mandiri.",
     "1. Mengidentifikasi informasi spesifik dan ide pokok dari teks deskripsi lisan/audio.\n2. Menganalisis sifat, penampilan fisik, dan prestasi tokoh olahraga.\n3. Mengidentifikasi kosakata adjektiva dan tata bahasa Simple Present Tense.",
     "Describing Great Athletes & Tourism Destinations:\n• Identification & Descriptions\n• Adjectives of physical/personality traits\n• Simple Present Tense & sensory language",
     "Descriptive Text", "Tahap 1 (Awal Sem 1)", 10,
     "Mandiri, Bernalar Kritis, Berkebinekaan Global", "Asesmen Formatif (LKPD) & Tes Menyimak"),

    (2, "TP 10.2", "Membaca - Memirsa & Menulis - Mempresentasikan",
     "Memproduksi teks deskriptif multimodal (brosur digital/poster interaktif) mengenai tempat wisata nusantara atau cagar budaya dengan struktur teks dan ejaan yang tepat.",
     "1. Menyusun deskripsi terperinci mengenai destinasi wisata lokal.\n2. Mengintegrasikan elemen visual (foto/infografis) dan teks deskriptif.\n3. Mempresentasikan brosur wisata secara lisan di depan kelas.",
     "Promoting Cultural & Natural Heritage:\n• Descriptive organization & layout\n• Cohesive devices & sensory adjectives\n• Multimodal presentation skills",
     "Descriptive Text (Multimodal)", "Tahap 2 (Pertengahan Sem 1)", 10,
     "Kreatif, Gotong Royong, Berkebinekaan Global", "Penilaian Proyek Poster & Presentasi"),

    (3, "TP 10.3", "Menyimak - Berbicara & Menulis - Mempresentasikan",
     "Mengidentifikasi konteks, gagasan utama, dan urutan kronologis peristiwa bersejarah serta pengalaman pribadi masa lalu (Recount Text), serta menyusun teks recount faktual.",
     "1. Mengurutkan rangkaian kejadian masa lalu secara kronologis (orientation, events, re-orientation).\n2. Menggunakan action verbs dalam Simple Past dan Past Continuous.\n3. Menulis jurnal refleksi pribadi atau memoar sejarah singkat.",
     "Historical Recount & Inspiring Life Stories:\n• Chronological sequencing & time linkers\n• Simple Past & Past Continuous tenses\n• Biographical recount writing",
     "Recount Text", "Tahap 3 (Akhir Sem 1)", 10,
     "Bernalar Kritis, Mandiri", "Asesmen Portofolio Tulisan & Jurnal Refleksi"),

    (4, "TP 10.4", "Menyimak - Berbicara & Membaca - Memirsa",
     "Menganalisis nilai-nilai moral, alur naratif, dan konflik dalam cerita rakyat Indonesia dan legenda dunia (Narrative Text), serta mendemonstrasikan pembacaan cerita atau bermain peran (roleplay).",
     "1. Mengidentifikasi orientasi, komplikasi, resolusi, dan koda dalam cerita rakyat.\n2. Menganalisis karakteristik tokoh protagonis dan antagonis.\n3. Mempraktikkan teknik membaca cerita dramatik (storytelling) dengan intonasi tepat.",
     "Indonesian Folklore & World Legends with Moral Values:\n• Narrative structure & plot climax\n• Direct and indirect speech\n• Characterization & ethical dilemmas",
     "Narrative Text", "Tahap 4 (Awal Sem 2)", 10,
     "Berkebinekaan Global, Beriman & Bertakwa", "Unjuk Kerja Storytelling & Tes Analisis Teks"),

    (5, "TP 10.5", "Membaca - Memirsa & Menulis - Mempresentasikan",
     "Memahami dan menyusun teks prosedur lisan dan tulis terkait pola hidup sehat (healthy living), resep kuliner nusantara, serta panduan keselamatan berinternet/media sosial secara runut.",
     "1. Mengidentifikasi tujuan (goal), alat/bahan (materials), dan langkah (steps) teks prosedur.\n2. Menggunakan kata kerja imperatif, kalimat perintah, dan kata penghubung sekuensial.\n3. Membuat video tutorial mini atau infografis resep hidup sehat.",
     "Procedures for Healthy Living & Cyber Safety:\n• Goal, Materials, Sequenced Steps\n• Imperative sentences, action verbs, sequencers\n• Adverbials of manner, time, and sequence",
     "Procedure Text", "Tahap 5 (Pertengahan Sem 2)", 10,
     "Mandiri, Bernalar Kritis, Kreatif", "Penilaian Produk Video Tutorial / Infografis"),

    (6, "TP 10.6", "Menyimak - Berbicara, Membaca - Memirsa & Menulis",
     "Menganalisis laporan hasil observasi (Report Text) tentang fenomena alam, fauna endemik, dan inovasi sains, serta mengevaluasi argumen awal dalam teks eksposisi analitis isu lingkungan.",
     "1. Mengklasifikasikan informasi umum dan deskripsi bagian pada Report Text.\n2. Mengidentifikasi tesis dan argumen dalam teks eksposisi lingkungan hidup.\n3. Menyajikan laporan ilmiah ringkas atau esai opini konservasi alam.",
     "Reporting Fauna Conservation & Environmental Issues:\n• General Classification & Detailed Descriptions\n• Passive voice, technical nouns, timeless present\n• Formulating basic thesis and supporting reasons",
     "Report & Intro Exposition", "Tahap 6 (Akhir Sem 2 / Capstone)", 10,
     "Bernalar Kritis, Peduli Lingkungan (Akhlak Alam)", "Laporan Ilmiah Mini & Tes Sumatif Fase E")
]

ATP_XI_WAJIB = [
    (1, "TP 11.W.1", "Membaca - Memirsa & Menulis - Mempresentasikan",
     "Menganalisis fungsi sosial, struktur teks, dan unsur kebahasaan teks eksposisi analitis tentang isu sosial dan etika digital masyarakat modern, serta menyusun esai argumentatif tertulis.",
     "1. Menemukan kalimat tesis (thesis statement) dan argumen pendukung teks eksposisi.\n2. Menggunakan kata kerja mental (feel, believe, conclude) dan konjungsi kausal (because, therefore).\n3. Menulis esai analitis 3 argumen secara logis dan kohesif.",
     "Analytical Exposition on Digital Ethics & Social Issues:\n• Thesis, series of arguments, reiteration\n• Mental verbs, causal conjunctions, formal passive\n• Distinguishing facts from opinions",
     "Analytical Exposition", "Tahap 1 (Awal Sem 3)", 10,
     "Bernalar Kritis, Mandiri", "Esai Argumentatif Analitis & Rubrik Analitik"),

    (2, "TP 11.W.2", "Menyimak - Berbicara & Menulis - Mempresentasikan",
     "Mengidentifikasi argumen dan rekomendasi persuasif dalam teks eksposisi hortatori terkait kesehatan mental remaja dan keselamatan publik, serta mempresentasikan pidato ajakan bertindak.",
     "1. Mengidentifikasi unsur rekomendasi (recommendation) dengan modal auxiliaries (should, ought to).\n2. Menganalisis bahasa emotif dan ajakan persuasif dalam teks wacana.\n3. Menyampaikan pidato persuasif kampanye kesehatan remaja.",
     "Hortatory Exposition for Teen Well-being & Safety:\n• Thesis, arguments, persuasive recommendation\n• Modals of obligation & advice (should, must, ought to)\n• Advocacy presentation & persuasive delivery",
     "Hortatory Exposition", "Tahap 2 (Pertengahan Sem 3)", 10,
     "Mandiri, Bernalar Kritis, Berakhlak Mulia", "Unjuk Kerja Pidato Persuasif & Lembar Observasi"),

    (3, "TP 11.W.3", "Menyimak - Berbicara & Membaca - Memirsa",
     "Menganalisis perkembangan karakter, resolusi konflik, dan pesan moral dalam teks naratif kontemporer dan fabel modern (Fractured Fairy Tales), serta menyusun adaptasi narasi kreatif.",
     "1. Mengidentifikasi variasi sudut pandang penceritaan (point of view).\n2. Menemukan ironi, humor, dan pesan tersirat dalam dongeng yang dimodifikasi.\n3. Menulis ulang akhir cerita dongeng dengan sudut pandang tokoh yang berbeda.",
     "Contemporary Narratives & Modern Fairy Tales:\n• Orientation, complication, evaluation, resolution\n• Past tenses, direct/indirect quotes, descriptive adjectives\n• Creative storytelling from alternative viewpoints",
     "Narrative Text", "Tahap 3 (Akhir Sem 3)", 10,
     "Kreatif, Berkebinekaan Global", "Portofolio Cerita Adaptasi & Penilaian Narasi"),

    (4, "TP 11.W.4", "Membaca - Memirsa & Menulis - Mempresentasikan",
     "Menganalisis proses kausal dan kronologis terjadinya fenomena alam dan sosial budaya (Explanation Text), serta mempresentasikan diagram alir penjelasan ilmiah populer.",
     "1. Menjelaskan hubungan sebab-akibat (cause and effect) dalam fenomena alam/sosial.\n2. Mengidentifikasi kata transisi urutan proses (first, subsequently, resulting in).\n3. Menyajikan eksplanasi ilmiah disertai diagram skematis.",
     "Explanation of Natural & Socio-Cultural Phenomena:\n• General statement, sequenced explanation, concluding remark\n• Cause-effect linkers (due to, as a result, consequently)\n• Passive voice in technical and natural process registers",
     "Explanation Text", "Tahap 4 (Awal Sem 4)", 10,
     "Bernalar Kritis, Kreatif", "Diagram Eksplanasi Ilmiah & Tes Tulis"),

    (5, "TP 11.W.5", "Menyimak - Berbicara & Membaca - Memirsa",
     "Menimbang dua sudut pandang berimbang pro dan kontra terkait isu keseharian remaja dan lingkungan sekolah (Discussion Text), serta berpartisipasi dalam diskusi kelompok terarah.",
     "1. Mengidentifikasi argumen yang mendukung (pros) dan menentang (cons) suatu isu kontroversial.\n2. Menggunakan konjungsi pertentangan (however, on the contrary, nevertheless).\n3. Berpartisipasi aktif dalam diskusi kelas dengan bahasa santun dan objektif.",
     "Discussion on School Realities & Balanced Perspectives:\n• Issue statement, arguments for, arguments against, recommendation\n• Contrastive conjunctions and hedging/neutral expressions\n• Structured classroom round-table discussion",
     "Discussion Text", "Tahap 5 (Pertengahan Sem 4)", 10,
     "Bernalar Kritis, Gotong Royong, Demokratis", "Penilaian Kinerja Diskusi Kelas & Esai Berimbang"),

    (6, "TP 11.W.6", "Membaca - Memirsa & Menulis - Mempresentasikan",
     "Mengevaluasi dan menyusun teks resensi kritis (Review Text) terhadap film edukatif, buku bacaan, atau karya seni budaya lokal nusantara dengan rubrik evaluasi yang tepat.",
     "1. Mengidentifikasi bagian interpretative recount dan evaluasi kritis karya.\n2. Menggunakan kosakata apresiatif dan adjektiva pertimbangan mutu.\n3. Menulis resensi film/buku dan mempublikasikannya di mading/media daring sekolah.",
     "Reviewing Inspiring Films, Books, & Cultural Masterpieces:\n• Orientation, interpretative recount, evaluation, evaluative punch\n• Aesthetic and evaluative vocabulary\n• Published student art & literature review",
     "Review Text", "Tahap 6 (Akhir Sem 4 / Capstone)", 10,
     "Berkebinekaan Global, Kreatif, Bernalar Kritis", "Resensi Karya Tulis & Pameran Resensi")
]

ATP_XI_TL = [
    (1, "TP 11.TL.1", "Membaca - Memirsa & Menulis - Mempresentasikan",
     "Mengevaluasi wacana argumentatif kritis tingkat lanjut dalam teks eksposisi analitis bertema disrupsi kecerdasan buatan (AI) dan transformasi masyarakat digital, serta menyusun makalah posisi akademik.",
     "1. Mengidentifikasi argumen multi-lapis dan counter-arguments dalam teks jurnal/artikel ilmiah populer.\n2. Menggunakan nominalisasi abstrak dan modalitas kepastian tingkat tinggi.\n3. Menyusun academic analytical paper dengan referensi terpercaya.",
     "Advanced Analytical Exposition on AI & Digital Society:\n• Complex thesis formulation & multi-perspective arguments\n• Abstract nominalization, modal adverbs of certainty\n• Logical connectors of condition, concession & result",
     "Analytical Exposition (Advanced)", "Tahap 1 (Sem 3 / 4 mgg)", 20,
     "Bernalar Kritis, Mandiri, Berwawasan Global", "Academic Position Paper & Podcast Analitis"),

    (2, "TP 11.TL.2", "Menyimak - Berbicara & Menulis - Mempresentasikan",
     "Menganalisis dan menyusun teks eksposisi hortatori tingkat lanjut yang mendesak tindakan solutif terkait transisi energi bersih dan keadilan iklim global, serta mempresentasikan policy pitch.",
     "1. Menganalisis strategi retorika persuasif (Ethos, Pathos, Logos) pada pidato pembuat kebijakan.\n2. Menggunakan kata-kata persuasif tingkat lanjut dan permohonan rekomendasi mendesak.\n3. Mempresentasikan student policy brief di hadapan audiens.",
     "Hortatory Exposition for Global Climate Action:\n• Thesis, evaluative arguments, urgent recommendations\n• Emotive vocabulary, strong modal auxiliaries\n• Rhetorical questions & persuasive policy pitch",
     "Hortatory Exposition", "Tahap 2 (Sem 3 / 5 mgg)", 25,
     "Bernalar Kritis, Berakhlak Mulia, Peduli Bumi", "Policy Brief & Presentasi Advokasi Lisan"),

    (3, "TP 11.TL.3", "Menyimak - Berbicara & Membaca - Memirsa",
     "Mengapresiasi, menganalisis kedalaman psikologis karakter, dan merekonstruksi narasi sastra fiksi/cerpen klasik dan kontemporer dengan alur bertingkat dan dilema moral.",
     "1. Mengidentifikasi teknik narasi canggih (foreshadowing, symbolism, stream of consciousness).\n2. Menganalisis motivasi psikologis dan dilema etika tokoh utama.\n3. Menulis analisis kritis sastra atau adaptasi naskah pembacaan naratif.",
     "Literary Analysis of Advanced Narratives & Short Stories:\n• Narrative arc, foreshadowing, symbolism, subtext\n• Literary devices: metaphor, irony, character introspection\n• Comparative character analysis essay",
     "Advanced Narrative & Short Story", "Tahap 3 (Sem 3 / 5 mgg)", 25,
     "Kreatif, Mandiri, Berkebinekaan Global", "Esai Kritik Sastra & Dramatic Reader's Theatre"),

    (4, "TP 11.TL.4", "Menyimak - Berbicara & Membaca - Memirsa",
     "Menganalisis gaya bahasa stilistika, makna simbolik, dan kritik sosial dalam puisi bahasa Inggris dan lirik lagu dunia berpengaruh, serta mendeklamasikannya.",
     "1. Menguraikan rima, meter, bait, citraan (imagery), dan tone puisi.\n2. Menafsirkan pesan kritik sosial terhadap ketidakadilan, kemanusiaan, dan perdamaian.\n3. Membawakan pembacaan puisi ekspresif atau ulasan stilistika lagu.",
     "Stylistic & Societal Reading of Poetry & Protest Songs:\n• Meter, stanza structure, imagery, tone, and subtext\n• Figuring social inequality, peace, and human dignity\n• Poetry recitation and stylistic commentary",
     "Poetry & Song Critique", "Tahap 4 (Sem 3 / 2 mgg)", 10,
     "Berkebinekaan Global, Kreatif", "Resitasi Puisi Ekspresif & Ulasan Stilistika"),

    (5, "TP 11.TL.5", "Menyimak - Berbicara & Menulis - Mempresentasikan",
     "Mengonstruksi dan menimbang dua sudut pandang berimbang dalam teks diskusi mengenai kontroversi e-sports/gaming dan kecerdasan buatan (Model Resmi Bab D Panduan Puskurjar 2025), serta berdebat gaya Oxford.",
     "1. Merumuskan isu kontroversial secara objektif tanpa bias awal.\n2. Menyusun argumen pro dan kontra berbasis data empiris dan bukti ilmiah.\n3. Melaksanakan simulasi debat formal gaya parlemen/Oxford dan menyusun esai sintesis berimbang.",
     "Official Puskurjar Model: Discussion Text on Online Games & Social Issues:\n• Issue statement, Arguments Pro, Arguments Con, Balanced Recommendation\n• Contrasting connectives, Hedging & cautious academic language\n• Oxford-style parliamentary debate & balanced synthesis essay",
     "Discussion Text (Official Puskurjar)", "Tahap 5 (Sem 4 / 5 mgg)", 25,
     "Bernalar Kritis, Demokratis, Gotong Royong", "Simulasi Debat Oxford & Esai Diskusi Sintesis"),

    (6, "TP 11.TL.6", "Membaca - Memirsa & Menulis - Mempresentasikan",
     "Menganalisis mekanisme kausalitas ilmiah kompleks dan urutan terjadinya fenomena rekayasa teknologi mutakhir dan astrofisika dalam teks eksplanasi ilmiah tingkat tinggi.",
     "1. Memetakan rantai kausal multi-cabang (multiple cause-and-effect pathways).\n2. Menggunakan register bahasa ilmiah, nominalisasi kompleks, dan kalimat pasif beruntun.\n3. Membuat artikel ilmiah populer multimodal disertai infografis teknis.",
     "Causal & Sequential Explanation of High-Tech & Astrophysical Phenomena:\n• Sequenced explanation of complex mechanism (why & how)\n• Scientific register, technical terminology, nominal phrases\n• Multimodal science explanation article",
     "Explanation Text (Complex Causal)", "Tahap 6 (Sem 4 / 5 mgg)", 25,
     "Bernalar Kritis, Mandiri", "Artikel Eksplanasi Ilmiah & Infografis Teknis"),

    (7, "TP 11.TL.7", "Membaca - Memirsa & Menulis - Mempresentasikan",
     "Mengevaluasi rekam jejak perjuangan, kepemimpinan transformasional, dan kontribusi historis tokoh-tokoh visioner dunia dalam teks biografi mendalam dan memoar sejarah.",
     "1. Mengidentifikasi tonggak pencapaian sejarah (historical milestones) tokoh dunia.\n2. Menganalisis nilai kepemimpinan, integritas, dan keteladanan karakter.\n3. Menulis mini-monograf biografi analitis atau memproduksi podcast biografi tokoh.",
     "Comprehensive Biography & Historical Memoirs of Visionary Leaders:\n• Orientation, chronological milestones, evaluative contribution\n• Past perfect, complex time clauses, descriptive appraisal\n• Biographical mini-monograph / Biographical podcast",
     "Biography & Historical Memoir", "Tahap 7 (Sem 4 / 4 mgg)", 20,
     "Mandiri, Berakhlak Mulia, Berkebinekaan Global", "Monograf Biografi Tokoh & Rekaman Podcast"),

    (8, "TP 11.TL.8", "Membaca - Memirsa & Menulis - Mempresentasikan",
     "Menyusun resensi dan kritik estetika mendalam (Review Text) terhadap novel sastra dunia atau karya film berkualitas dengan rubrik evaluasi artistik objektif.",
     "1. Menilai elemen intrinsik dan ekstrinsik karya sastra/sinema secara mendalam.\n2. Menggunakan kosakata penilaian estetika tingkat tinggi dan pembandingan karya sejenis.\n3. Mempublikasikan ulasan karya seni dalam portofolio ulasan budaya siswa.",
     "Critical Review of International Literature & Cinematic Works:\n• Orientation, interpretative recount, multi-criteria evaluation, summative punch\n• Aesthetic vocabulary, nuanced judgmental adjectives\n• Critical review portfolio publication",
     "Review Text (Critical Arts Review)", "Tahap 8 (Sem 4 / 2 mgg)", 10,
     "Kreatif, Berkebinekaan Global, Bernalar Kritis", "Portofolio Resensi Seni & Gelar Wicara Budaya")
]

ATP_XII_WAJIB = [
    (1, "TP 12.W.1", "Menyimak - Berbicara & Membaca - Memirsa",
     "Menganalisis teks diskusi mengenai disrupsi dunia kerja masa depan, etika kecerdasan buatan, dan keberlanjutan lingkungan, serta menyusun teks diskusi berimbang.",
     "1. Mengidentifikasi berbagai sudut pandang pemangku kepentingan dalam isu ketenagakerjaan masa depan.\n2. Menilai kekuatan argumen pendukung dan kontra secara objektif.\n3. Menulis teks diskusi berimbang dengan bahasa yang netral dan konstruktif.",
     "Discussion on the Future of Work & Automation:\n• Issue, pro/con viewpoints, balanced conclusion\n• Neutral discourse markers, hedging expressions\n• Collaborative classroom deliberation",
     "Discussion Text", "Tahap 1 (Awal Sem 5)", 10,
     "Bernalar Kritis, Mandiri, Berwawasan Global", "Teks Diskusi Tertulis & Rubrik Penilaian Argumen"),

    (2, "TP 12.W.2", "Membaca - Memirsa & Menulis - Mempresentasikan",
     "Mengidentifikasi fakta utama, latar belakang peristiwa, dan sumber berita dalam teks berita aktual (News Item Text), serta memproduksi siaran berita (news broadcast) informatif.",
     "1. Mengidentifikasi newsworthy event, background events, dan quotes dari narasumber.\n2. Menggunakan reporting verbs (stated, reported, confirmed) dan headline yang menarik.\n3. Mempraktikkan siaran berita lisan (news anchor) secara berpasangan.",
     "News Item Text & Media Literacy:\n• Newsworthy event, background events, sources\n• Headlines, reporting verbs, direct/indirect quotes\n• Video simulasi pembaca berita (news casting)",
     "News Item Text", "Tahap 2 (Pertengahan Sem 5)", 10,
     "Bernalar Kritis, Mandiri, Komunikatif", "Simulasi Penyiaran Berita & Teks Berita Cetak"),

    (3, "TP 12.W.3", "Membaca - Memirsa & Menulis - Mempresentasikan",
     "Mengevaluasi dan menyusun ulasan kritis (Review Text) terhadap produk kreatif nusantara, buku, atau film lokal untuk promosi kekayaan budaya Indonesia secara multimodal.",
     "1. Mengapresiasi karya kreatif lokal dengan kriteria penilaian yang jelas.\n2. Menggunakan bahasa evaluatif yang santun dan membangun.\n3. Menyajikan ulasan dalam bentuk artikel blog atau konten media sosial edukatif.",
     "Reviewing Indonesian Creative Products & Cultural Heritage:\n• Evaluation, interpretative recount, evaluative summation\n• Evaluative adjectives and cultural appreciation terms\n• Creative digital review blog post",
     "Review Text", "Tahap 3 (Akhir Sem 5)", 10,
     "Berkebinekaan Global, Kreatif", "Blog Ulasan Budaya & Penilaian Portofolio"),

    (4, "TP 12.W.4", "Membaca - Memirsa & Menulis - Mempresentasikan",
     "Menyusun surat lamaran pekerjaan formal (Job Application Letter) dan daftar riwayat hidup modern (Curriculum Vitae / Resume) sesuai standar profesional dunia kerja modern.",
     "1. Menganalisis struktur surat lamaran resmi (salutation, opening, qualifications, closing).\n2. Menyusun Curriculum Vitae ringkas, informatif, dan sesuai format pelamar pemula.\n3. Memeriksa ketepatan ejaan, tata bahasa profesional, dan kesantunan register bahasa.",
     "Job Application Letter & Modern Curriculum Vitae:\n• Formal business letter conventions, layout, salutation\n• ATS-friendly CV design & highlighting skills/achievements\n• Professional action verbs and polite formal closing",
     "Application Letter & CV", "Tahap 4 (Awal Sem 6)", 10,
     "Mandiri, Bernalar Kritis, Berintegritas", "Dokumen Berkas Lamaran Kerja & Portofolio CV"),

    (5, "TP 12.W.5", "Menyimak - Berbicara",
     "Mempraktikkan percakapan wawancara kerja formal (Job Interview Simulation) dan keterampilan presentasi diri secara percaya diri, santun, dan runtut.",
     "1. Menjawab pertanyaan wawancara mengenai kualifikasi diri dan rencana karir.\n2. Mengajukan pertanyaan klarifikasi yang tepat dan menunjukkan antusiasme profesional.\n3. Menunjukkan bahasa tubuh yang percaya diri dan kesantunan berbahasa Inggris.",
     "Job Interview Simulation & Professional Presentation:\n• Answering interview queries (strengths, goals, experience)\n• Asking clarifying questions, professional body language\n• Mock job interview simulation with peer/teacher feedback",
     "Job Interview Skills", "Tahap 5 (Pertengahan Sem 6)", 8,
     "Mandiri, Komunikatif, Berakhlak Mulia", "Unjuk Kerja Wawancara Kerja & Lembar Observasi"),

    (6, "TP 12.W.6", "Membaca - Memirsa & Menulis - Mempresentasikan",
     "Menafsirkan makna dan pesan tersirat dalam takarir (Caption), grafik, dan infografis multimodal, serta menyajikan ringkasan visual ke dalam bahasa Inggris yang komunikatif.",
     "1. Mengidentifikasi pesan utama dan konteks foto dari lead dan caption.\n2. Membaca dan menafsirkan data grafik batang, garis, dan diagram lingkaran.\n3. Membuat takarir informatif untuk mendampingi infografis publikasi.",
     "Captions, Graphs, & Multimodal Infographics Literacy:\n• Structure of captions (lead, explanation, source)\n• Interpreting statistical graphs and data charts\n• Writing concise, catchy, and informative photo captions",
     "Caption & Infographic", "Tahap 6 (Akhir Sem 6 / Capstone)", 6,
     "Bernalar Kritis, Kreatif, Literasi Digital", "Proyek Infografis Bertakarir & Ujian Praktik")
]

ATP_XII_TL = [
    (1, "TP 12.TL.1", "Menyimak - Berbicara & Membaca - Memirsa",
     "Menganalisis wacana debat akademik tingkat tinggi (Advanced Discussion) mengenai bioetika medis, rekayasa genetika, eksplorasi antariksa, dan merumuskan sikap etis ilmiah secara komprehensif.",
     "1. Mengidentifikasi argumen filosofis dan implikasi etis dari inovasi sains frontier.\n2. Menggunakan penanda sikap akademik (stance markers, concessive clauses, academic collocations).\n3. Memimpin dialog meja bundar mengenai regulasi etika bioteknologi masa depan.",
     "Global Academic Discussion on Bioethics & Frontier Tech:\n• Multi-layered argumentation, philosophical nuances, ethical stances\n• Advanced stance markers, concessive clauses, academic collocations\n• Scholarly dialogue / Policy deliberation roundtable",
     "Advanced Academic Discussion", "Tahap 1 (Sem 5 / 5 mgg)", 25,
     "Bernalar Kritis, Beriman & Bertakwa, Berwawasan Global", "Meja Bundar Akademik & Esai Posisi Bioetika"),

    (2, "TP 12.TL.2", "Membaca - Memirsa & Menulis - Mempresentasikan",
     "Menyusun esai penelitian akademik (Academic Synthesis Essay) 1.000 kata dengan memadukan multi-referensi, kutipan langsung/tidak langsung, dan daftar pustaka standar internasional (APA/MLA).",
     "1. Memadukan minimal 3 sumber literatur ilmiah untuk mendukung tesis mandiri.\n2. Menggunakan kaidah sitasi dalam teks (in-text citation) dan menyusun daftar pustaka tanpa plagiarisme.\n3. Menyusun esai akademik berstruktur 5 paragraf dengan transisi antarparagraf yang koheren.",
     "Academic Research Essay Writing & Source Synthesis:\n• Formulating working thesis, literature synthesis, topic sentences\n• In-text citation protocols (APA/MLA style), paraphrasing, avoiding plagiarism\n• Cohesive devices, academic hedging, formal objective register",
     "Academic Synthesis Essay", "Tahap 2 (Sem 5 / 5 mgg)", 25,
     "Mandiri, Bernalar Kritis, Berintegritas Akademik", "Makalah Sintesis Akademik 1.000 Kata"),

    (3, "TP 12.TL.3", "Menyimak - Berbicara & Membaca - Memirsa",
     "Menganalisis teks naskah drama klasik/modern, menafsirkan subteks psikologis, serta memperagakan adegan monolog/dialog dramatik dengan intonasi dan pelafalan standar profesional.",
     "1. Menganalisis petunjuk panggung (stage directions), ketegangan dramatik, dan subteks emosional.\n2. Menginterpretasikan pesan moral dan kritik sosial dalam naskah sandiwara.\n3. Mementaskan monolog dramatik atau adegan drama singkat dengan intonasi dan penghayatan kuat.",
     "Dramatic Literature & Theatrical Script Performance:\n• Dramatic structure, conflict escalation, stage directions, subtext\n• Voice modulation, stress, intonation, emotional projection\n• Script adaptation & theatrical live performance / Dramatic monologue recording",
     "Dramatic Script Analysis & Play", "Tahap 3 (Sem 5 / 4 mgg)", 20,
     "Kreatif, Mandiri, Berkebinekaan Global", "Pementasan Drama Singkat & Penilaian Vokal/Akting"),

    (4, "TP 12.TL.4", "Menyimak - Berbicara & Menulis - Mempresentasikan",
     "Merancang dan membawakan pidato persuasif tingkat tinggi (Persuasive Oratory) dan mempraktikkan debat formal gaya parlemen dengan retorika persuasif yang meyakinkan.",
     "1. Mengaplikasikan retorika klasik (Ethos, Pathos, Logos) untuk memengaruhi opini audiens.\n2. Menggunakan perangkat retoris (parallelism, rhetorical questions, antithesis) secara efektif.\n3. Menyampaikan pidato formal dalam simulasi sidang umum internasional.",
     "Persuasive Public Speaking & Parliamentary Debating:\n• Classical rhetoric: Ethos (credibility), Pathos (emotion), Logos (logic)\n• Parallelism, antithesis, persuasive prosody, rebuttal strategies\n• Formal speech delivery in international forum simulation",
     "Persuasive Oratory & Debate", "Tahap 4 (Sem 5 / 2 mgg)", 10,
     "Mandiri, Bernalar Kritis, Komunikatif", "Pidato Oratori Formal & Unjuk Kerja Debat"),

    (5, "TP 12.TL.5", "Membaca - Memirsa & Menulis - Mempresentasikan",
     "Menyusun korespondensi kerja profesional dan akademik internasional: Surat Lamaran Kerja, Curriculum Vitae (CV) ATS-compliant, dan Pernyataan Motivasi Beasiswa/Kuliah (Personal Statement).",
     "1. Menyusun surat lamaran kerja eksekutif dengan register bahasa Inggris bisnis formal.\n2. Merancang Curriculum Vitae digital yang terstruktur dan menonjolkan pencapaian kunci.\n3. Menulis personal statement yang meyakinkan untuk seleksi universitas bergengsi atau beasiswa.",
     "Professional Career & Academic University Admissions Portfolio:\n• Formal business letter format, salutation, professional tone\n• Modern ATS-friendly Curriculum Vitae / Resume design\n• Compelling personal statement for scholarships / college admissions",
     "Workplace & Academic Admissions", "Tahap 5 (Sem 6 / 4 mgg)", 20,
     "Mandiri, Bernalar Kritis, Berwawasan Global", "Berkas Dossier Lamaran & Personal Statement"),

    (6, "TP 12.TL.6", "Menyimak - Berbicara",
     "Mempraktikkan wawancara kerja profesional (Job Interview Simulation) dan pertahanan karya ilmiah lisan (Oral Defense) dengan respons spontan, taktis, dan register bahasa formal.",
     "1. Menerapkan teknik STAR (Situation, Task, Action, Result) dalam menjawab pertanyaan perilaku.\n2. Mempertahankan kesimpulan karya ilmiah di hadapan dewan penguji dengan argumen kuat.\n3. Menggunakan strategi kesantunan, jeda berpikir taktis, dan artikulasi bahasa Inggris presisi.",
     "Job Interview & Academic Oral Defense Simulation:\n• STAR Method (Situation, Task, Action, Result) in behavioral questions\n• Defending research methodology and conclusions under questioning\n• Professional executive presence, composure, and tone",
     "Interview & Oral Defense Simulation", "Tahap 6 (Sem 6 / 3 mgg)", 15,
     "Mandiri, Bernalar Kritis, Berintegritas", "Simulasi Wawancara Eksekutif & Sidang Oral"),

    (7, "TP 12.TL.7", "Membaca - Memirsa & Menulis - Mempresentasikan",
     "Menyusun abstrak laporan penelitian ilmiah (Research Abstract) 250 kata dan mendesain presentasi infografis saintifik untuk diseminasi publik dalam simposium ilmiah siswa.",
     "1. Merangkum latar belakang, metode, temuan, dan simpulan penelitian secara padat dan akurat.\n2. Menggunakan reporting verbs dan frasa nominal saintifik.\n3. Mendesain poster infografis penelitian dan mempresentasikannya di ajang simposium ilmiah.",
     "Mini Academic Paper Abstract & Scientific Infographic Poster:\n• 250-word structured abstract (Background, Method, Results, Conclusion)\n• Scientific reporting verbs, nominalization, data visualization synthesis\n• Academic poster symposium session",
     "Scientific Abstract & Infographic", "Tahap 7 (Sem 6 / 3 mgg)", 15,
     "Bernalar Kritis, Kreatif, Literasi Saintifik", "Abstrak Ilmiah 250 Kata & Simposium Poster"),

    (8, "TP 12.TL.8", "Menyimak, Membaca, Menulis & Berbicara",
     "Mengembangkan kesiapan literasi bahasa Inggris standar internasional (International Academic Literacy: IELTS/TOEFL/CEFR B2+ reading & critical analytical writing).",
     "1. Menerapkan teknik skimming, scanning, dan inferencing pada teks akademik panjang berbobot.\n2. Mengembangkan esai argumentatif dan analisis data grafik di bawah batasan waktu ujian.\n3. Mengukur capaian kompetensi kemahiran bahasa Inggris standar menuju jenjang perguruan tinggi.",
     "Standardized English Proficiency Readiness for Higher Education:\n• Advanced skimming, scanning, synthesizing inferential questions\n• Cohesion, coherence, and lexical precision under timed test conditions\n• Benchmark diagnostic proficiency assessment",
     "Standardized Academic Readiness", "Tahap 8 (Sem 6 / 2 mgg)", 10,
     "Mandiri, Bernalar Kritis, Berdaya Saing Global", "Diagnostic Standardized Benchmark Test")
]


# ==============================================================================
# BUILD SHEET: PRINSIP & KERANGKA ATP
# ==============================================================================

def build_sheet_prinsip_atp(ws):
    ws.title = "Panduan_&_Prinsip_ATP"
    apply_title(ws, "PANDUAN & KERANGKA ALUR TUJUAN PEMBELAJARAN (ATP) BAHASA INGGRIS SMA",
                "Landasan Hukum: Permendikdasmen No. 13 Tahun 2025, CP 2025/2026, dan Panduan Mapel Bahasa Inggris Revisi 3 (Sept 2025)", 6)

    school_data = [
        ("Satuan Pendidikan", "Sekolah Menengah Atas (SMA)"),
        ("Mata Pelajaran", "Bahasa Inggris (Umum/Wajib & Tingkat Lanjut)"),
        ("Fokus Jenjang", "Fase E (Kelas X) & Fase F (Kelas XI dan XII)"),
        ("Tahun Pelajaran", "2025/2026"),
        ("Penyusun Dokumen", "Musyawarah Guru Mata Pelajaran (MGMP) Bahasa Inggris SMA")
    ]
    r_start = add_school_identity(ws, school_data, start_row=4)

    # Section 1
    ws.cell(row=r_start, column=1, value="I. HAKIKAT & PENGERTIAN ALUR TUJUAN PEMBELAJARAN (ATP)").font = FONT_SEC_HEADER
    ws.merge_cells(start_row=r_start, start_column=1, end_row=r_start, end_column=6)
    for c in range(1, 7):
        ws.cell(row=r_start, column=c).fill = FILL_NAVY
    ws.row_dimensions[r_start].height = 24

    desc_points = [
        ("Definisi ATP", "Alur Tujuan Pembelajaran (ATP) adalah rangkaian Tujuan Pembelajaran (TP) yang disusun secara sistematis dan logis di dalam fase pembelajaran secara utuh, mulai dari awal fase hingga akhir fase, sebagai panduan guru dalam merancang Modul Ajar dan melaksanakan pembelajaran mendalam (Deep Learning)."),
        ("Fungsi Pedagogis", "ATP berfungsi sebagai kompas operasional pembelajaran yang memastikan tidak ada kompetensi atau genre teks esensial yang terlewat, serta memastikan terjadinya penahapan belajar (scaffolding) yang terencana dari tingkat pemahaman reseptif hingga produksi mandiri produktif."),
        ("Keterkaitan dengan CP", "Capaian Pembelajaran (CP) diturunkan ke dalam Tujuan Pembelajaran (TP) yang memuat tiga komponen inti: Kompetensi (kata kerja operasional performa), Konten (lingkup materi teks esensial), dan Konteks/Variasi (isu nyata lokal, nasional, dan global yang relevan)."),
        ("Pendekatan Berbasis Teks (GBA)", "ATP mengintegrasikan 4 tahapan Teaching Learning Cycle (TLC) dalam pengajaran teks: Building Knowledge of the Field (BKoF), Modelling of the Text (MoT), Joint Construction of the Text (JCoT), dan Independent Construction of the Text (ICoT)."),
        ("Standar Penjenjangan CEFR", "Fase E (Kelas X) berorientasi pada pencapaian level kemandirian CEFR B1; Fase F Umum (Kelas XI-XII Wajib) memperkuat literasi komunikatif menuju level B1-B2; sedangkan Fase F Tingkat Lanjut (Kelas XI-XII Pilihan) memperdalam penguasaan teks akademik, wacana kritis, dan literasi ilmiah hingga level CEFR B2.")
    ]

    r = r_start + 1
    for label, text in desc_points:
        ws.row_dimensions[r].height = 42
        c_lbl = ws.cell(row=r, column=1, value=label)
        c_lbl.font = FONT_DATA_BOLD
        c_lbl.fill = FILL_LIGHT_BLUE
        c_lbl.alignment = ALIGN_LEFT
        c_lbl.border = CELL_BORDER

        c_txt = ws.cell(row=r, column=2, value=text)
        c_txt.font = FONT_DATA
        c_txt.alignment = ALIGN_TOP_LEFT
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        for c in range(2, 7):
            ws.cell(row=r, column=c).border = CELL_BORDER
            if r % 2 == 1:
                ws.cell(row=r, column=c).fill = FILL_ZEBRA
        r += 1

    r += 1
    ws.cell(row=r, column=1, value="II. PRINSIP PENYUSUNAN ALUR TUJUAN PEMBELAJARAN (ATP)").font = FONT_SEC_HEADER
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws.cell(row=r, column=c).fill = FILL_ROYAL
    ws.row_dimensions[r].height = 24

    prinsip_headers = ["No", "Prinsip Utama", "Penjelasan Operasional", "Penerapan dalam Bahasa Inggris SMA", "Implikasi bagi Siswa", "Kriteria Keberhasilan"]
    r += 1
    for i, h in enumerate(prinsip_headers, 1):
        cell = ws.cell(row=r, column=i, value=h)
        cell.font = FONT_TBL_HEADER
        cell.fill = FILL_NAVY
        cell.alignment = ALIGN_CENTER
        cell.border = CELL_BORDER
    ws.row_dimensions[r].height = 22

    prinsip_data = [
        (1, "Sederhana & Informatif", "Rumusan ATP harus mudah dipahami guru dan memuat kompetensi, materi, serta konteks yang jelas tanpa ambiguitas istilah.", "Penggunaan kode TP yang konsisten (TP 10.1, TP 11.W.1, TP 11.TL.1) dan alokasi waktu yang pasti.", "Siswa memahami peta target capaian belajar mereka.", "Guru mudah menyusun RPP/Modul Ajar."),
        (2, "Esensial & Kontekstual", "Memuat kompetensi kunci (essential competencies) dan teks yang bersentuhan dengan realitas kehidupan murid.", "Tema materi kontekstual: kearifan lokal, teknologi digital & AI, isu iklim, dan etika profesional.", "Pembelajaran terasa bermakna (meaningful) dan aplikatif.", "Siswa mampu berkomunikasi dalam situasi otentik."),
        (3, "Berkesinambungan (Linear & Spiral)", "Penyusunan alur bergerak dari teks konkret/familiar menuju teks abstrak/kompleks, dengan pengulangan kosakata bergradasi.", "Mulai dari Deskriptif & Recount di Fase E, berlanjut ke Eksposisi & Eksplanasi, hingga Diskusi Akademik & Abstrak Ilmiah.", "Terjadinya penumpukan skemata pengetahuan yang kokoh.", "Siswa tidak mengalami lompatan kesulitan mendadak."),
        (4, "Pengintegrasian Deep Learning", "Menyelaraskan pengalaman belajar dengan 3 pilar: Mindful (berkesadaran), Meaningful (bermakna), dan Joyful (menggembirakan).", "Tahap MoT membimbing dekonstruksi teks secara sadar; JCoT memupuk kolaborasi tanpa rasa cemas berbahasa.", "Mereduksi kecemasan berbahasa (language anxiety).", "Kemahiran produktif meningkat secara mandiri."),
        (5, "Kemandirian Belajar & Evaluasi Diri", "Mengalokasikan waktu refleksi metakognitif (learning logs / 3-minute writing) di setiap akhir capaian modul.", "Penilaian formatif berkelanjutan dengan rubrik analitik 4 skala (Above, Meets, Approaching, Below Standards).", "Siswa mampu mengevaluasi kemajuan berbahasanya sendiri.", "Ketercapaian profil kemandirian dan integritas.")
    ]

    r += 1
    for row_val in prinsip_data:
        ws.row_dimensions[r].height = 34
        for col_idx, val in enumerate(row_val, 1):
            cell = ws.cell(row=r, column=col_idx, value=val)
            cell.font = FONT_DATA_BOLD if col_idx in [1, 2] else FONT_DATA
            cell.alignment = ALIGN_CENTER if col_idx == 1 else ALIGN_TOP_LEFT
            cell.border = CELL_BORDER
            if r % 2 == 1:
                cell.fill = FILL_ZEBRA
        r += 1

    auto_fit_cols(ws, min_w=12, max_w=46)


# ==============================================================================
# BUILD SHEET: DIAGRAM ALUR PEMBELAJARAN (VISUAL ROADMAP FLOWCHART)
# ==============================================================================

def build_sheet_diagram_alur(ws):
    ws.title = "Diagram_Alur_Pembelajaran"
    apply_title(ws, "DIAGRAM ALUR TUJUAN PEMBELAJARAN (ROADMAP FLOWCHART) BAHASA INGGRIS SMA",
                "Progresi Belajar Terstruktur: Fase E (Kelas X) -> Fase F Wajib (Kelas XI-XII) & Fase F Tingkat Lanjut (Kelas XI-XII)", 8)

    ws.views.sheetView[0].showGridLines = True

    # 1. HEADER ALUR FASE E
    r = 4
    ws.cell(row=r, column=1, value="ALUR 1: FASE E - KELAS X SMA (BAHASA INGGRIS UMUM / WAJIB - 72 JP) [TARGET CEFR: LEVEL B1]").font = FONT_SEC_HEADER
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    for c in range(1, 9):
        ws.cell(row=r, column=c).fill = FILL_NAVY
    ws.row_dimensions[r].height = 24

    # Flowchart blocks for Fase E
    flow_e = [
        ("TAHAP 1 (SEMESTER 1)", "TP 10.1: Descriptive Text", "Menganalisis & mendeskripsikan Great Athletes & Tourism Destinations (Adjectives, Simple Present)", FILL_NODE_BLUE),
        ("TAHAP 2 (SEMESTER 1)", "TP 10.2: Multimodal Descriptive", "Memproduksi brosur promosi cagar budaya dan destinasi wisata nusantara secara multimodal", FILL_NODE_BLUE),
        ("TAHAP 3 (SEMESTER 1)", "TP 10.3: Historical Recount Text", "Menyusun teks recount peristiwa sejarah nasional dan jurnal pengalaman pribadi bermakna", FILL_NODE_BLUE),
        ("MILESTONE SEM 1", "Sumatif Tengah & Akhir Semester 1", "Evaluasi Portofolio Reseptif-Produktif dan Tes Menyimak Lisan (Capaian Tengah Fase E)", FILL_NODE_AMBER),
        ("TAHAP 4 (SEMESTER 2)", "TP 10.4: Narrative Folklore Text", "Menganalisis nilai moral dan mendemonstrasikan storytelling dongeng nusantara & dunia", FILL_NODE_GREEN),
        ("TAHAP 5 (SEMESTER 2)", "TP 10.5: Procedure Text", "Menyusun teks prosedur pola hidup sehat dan panduan keamanan digital/media sosial", FILL_NODE_GREEN),
        ("TAHAP 6 (SEMESTER 2)", "TP 10.6: Report & Intro Exposition", "Menyajikan laporan observasi sains fauna dan dasar esai eksposisi analitis isu lingkungan", FILL_NODE_GREEN),
        ("CAPSTONE FASE E", "Asesmen Akhir Fase E (Kenaikan)", "Pameran Karya Proyek Multimodal & Uji Kemandirian Berbahasa Inggris (Tuntas Level CEFR B1)", FILL_NODE_PURPLE)
    ]

    r += 1
    for i, (stage, code, desc, fill_node) in enumerate(flow_e, 1):
        ws.row_dimensions[r].height = 26
        c_stg = ws.cell(row=r, column=1, value=stage)
        c_stg.font = FONT_DATA_BOLD
        c_stg.fill = FILL_LIGHT_BLUE if "SEMESTER" in stage else FILL_ACCENT_AMBER
        c_stg.alignment = ALIGN_CENTER
        c_stg.border = CELL_BORDER

        c_box = ws.cell(row=r, column=2, value=code)
        c_box.font = FONT_DIAGRAM_NODE
        c_box.fill = fill_node
        c_box.alignment = ALIGN_CENTER
        c_box.border = BOX_BORDER
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)

        c_arrow = ws.cell(row=r, column=4, value="--->" if i < len(flow_e) else "[SELESAI FASE E]")
        c_arrow.font = FONT_DIAGRAM_ARROW
        c_arrow.alignment = ALIGN_CENTER

        c_desc = ws.cell(row=r, column=5, value=desc)
        c_desc.font = FONT_DATA
        c_desc.alignment = ALIGN_LEFT
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=8)
        for c in range(5, 9):
            ws.cell(row=r, column=c).border = CELL_BORDER
            if r % 2 == 1:
                ws.cell(row=r, column=c).fill = FILL_ZEBRA
        r += 1

    r += 1
    # 2. HEADER ALUR FASE F UMUM / WAJIB
    ws.cell(row=r, column=1, value="ALUR 2: FASE F UMUM - KELAS XI & XII SMA (BAHASA INGGRIS WAJIB - 136 JP) [TARGET CEFR: B1 KE B2]").font = FONT_SEC_HEADER
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    for c in range(1, 9):
        ws.cell(row=r, column=c).fill = FILL_ROYAL
    ws.row_dimensions[r].height = 24

    flow_f_wajib = [
        ("KELAS XI - SEM 3", "TP 11.W.1: Analytical Exposition", "Menganalisis & menyusun esai eksposisi analitis isu sosial & etika digital masyarakat", FILL_NODE_BLUE),
        ("KELAS XI - SEM 3", "TP 11.W.2: Hortatory Exposition", "Menyusun pidato kampanye persuasif keselamatan publik dan kesehatan mental remaja", FILL_NODE_BLUE),
        ("KELAS XI - SEM 3", "TP 11.W.3: Contemporary Narrative", "Menganalisis konflik & menulis adaptasi cerita rakyat modern (Fractured Fairy Tales)", FILL_NODE_BLUE),
        ("KELAS XI - SEM 4", "TP 11.W.4: Explanation Text", "Menjelaskan mekanisme sebab-akibat fenomena alam dan sosial budaya dengan diagram alir", FILL_NODE_GREEN),
        ("KELAS XI - SEM 4", "TP 11.W.5: Discussion Text (Remaja)", "Menimbang sudut pandang pro-kontra isu keseharian siswa dalam diskusi meja bundar", FILL_NODE_GREEN),
        ("KELAS XI - SEM 4", "TP 11.W.6: Review Text (Seni/Film)", "Mengevaluasi & mempublikasikan resensi kritis terhadap novel inspiratif, film, dan karya seni", FILL_NODE_GREEN),
        ("KELAS XII - SEM 5", "TP 12.W.1: Discussion (Masa Depan)", "Menganalisis tren masa depan dunia kerja, otomasi AI, dan keberlanjutan lingkungan", FILL_NODE_PURPLE),
        ("KELAS XII - SEM 5", "TP 12.W.2: News Item Text", "Mengidentifikasi fakta berita & memproduksi siaran berita lisan (news broadcast video)", FILL_NODE_PURPLE),
        ("KELAS XII - SEM 5", "TP 12.W.3: Review Produk Budaya", "Menulis ulasan kritis produk kreatif nusantara untuk promosi warisan budaya lokal", FILL_NODE_PURPLE),
        ("KELAS XII - SEM 6", "TP 12.W.4: Job Application & CV", "Menyusun berkas surat lamaran kerja resmi dan Curriculum Vitae modern ATS-compliant", FILL_NODE_AMBER),
        ("KELAS XII - SEM 6", "TP 12.W.5: Job Interview Skills", "Simulasi wawancara kerja formal dan presentasi diri profesional di hadapan panelis", FILL_NODE_AMBER),
        ("KELAS XII - SEM 6", "TP 12.W.6: Captions & Infographics", "Menafsirkan data infografis multimodal & menyajikan ringkasan komunikatif (Ujian Akhir)", FILL_NODE_AMBER)
    ]

    r += 1
    for i, (stage, code, desc, fill_node) in enumerate(flow_f_wajib, 1):
        ws.row_dimensions[r].height = 26
        c_stg = ws.cell(row=r, column=1, value=stage)
        c_stg.font = FONT_DATA_BOLD
        c_stg.fill = FILL_LIGHT_BLUE if "XI" in stage else FILL_VERY_LIGHT
        c_stg.alignment = ALIGN_CENTER
        c_stg.border = CELL_BORDER

        c_box = ws.cell(row=r, column=2, value=code)
        c_box.font = FONT_DIAGRAM_NODE
        c_box.fill = fill_node
        c_box.alignment = ALIGN_CENTER
        c_box.border = BOX_BORDER
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)

        c_arrow = ws.cell(row=r, column=4, value="--->" if i < len(flow_f_wajib) else "[LULUS SMA]")
        c_arrow.font = FONT_DIAGRAM_ARROW
        c_arrow.alignment = ALIGN_CENTER

        c_desc = ws.cell(row=r, column=5, value=desc)
        c_desc.font = FONT_DATA
        c_desc.alignment = ALIGN_LEFT
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=8)
        for c in range(5, 9):
            ws.cell(row=r, column=c).border = CELL_BORDER
            if r % 2 == 1:
                ws.cell(row=r, column=c).fill = FILL_ZEBRA
        r += 1

    r += 1
    # 3. HEADER ALUR FASE F TINGKAT LANJUT
    ws.cell(row=r, column=1, value="ALUR 3: FASE F TINGKAT LANJUT - KELAS XI & XII SMA (MATA PELAJARAN PILIHAN - 340 JP) [TARGET CEFR: LEVEL B2]").font = FONT_SEC_HEADER
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    for c in range(1, 9):
        ws.cell(row=r, column=c).fill = FILL_BLUE_HEADER
    ws.row_dimensions[r].height = 24

    flow_f_tl = [
        ("KELAS XI - SEM 3", "TP 11.TL.1: Advanced Analytical Exposition", "Evaluasi wacana kritis disrupsi AI, etika siber, dan nominalisasi abstrak kompleks (20 JP)", FILL_NODE_BLUE),
        ("KELAS XI - SEM 3", "TP 11.TL.2: Advanced Hortatory Exposition", "Penyusunan advokasi transisi energi terbarukan & policy pitch keadilan iklim (25 JP)", FILL_NODE_BLUE),
        ("KELAS XI - SEM 3", "TP 11.TL.3: Advanced Literary Narrative", "Kajian psikoanalisis karakter, simbolisme, dan dilema etika cerpen sastra fiksi (25 JP)", FILL_NODE_BLUE),
        ("KELAS XI - SEM 3", "TP 11.TL.4: Poetry & Protest Songs", "Stilistika puisi bahasa Inggris dan interpretasi lirik lagu kritik sosial dunia (10 JP)", FILL_NODE_BLUE),
        ("KELAS XI - SEM 4", "TP 11.TL.5: Discussion Text (Puskurjar)", "Model Resmi Bab D: Perdebatan seimbang pro-kontra online gaming & Oxford debate (25 JP)", FILL_NODE_GREEN),
        ("KELAS XI - SEM 4", "TP 11.TL.6: Complex Causal Explanation", "Mekanisme ilmiah multi-cabang teknologi tinggi, kuantum, dan astrofisika (25 JP)", FILL_NODE_GREEN),
        ("KELAS XI - SEM 4", "TP 11.TL.7: Biography & Historical Memoirs", "Kajian rekam jejak tokoh visioner dunia, kepemimpinan, dan monograf biografi (20 JP)", FILL_NODE_GREEN),
        ("KELAS XI - SEM 4", "TP 11.TL.8: Critical Arts & Film Review", "Kritik estetika sastra internasional dan evaluasi artistik film berbobot (10 JP)", FILL_NODE_GREEN),
        ("KELAS XII - SEM 5", "TP 12.TL.1: Global Bioethics Discussion", "Wacana etika rekayasa genetika, bioetika medis, dan kebijakan luar angkasa (25 JP)", FILL_NODE_PURPLE),
        ("KELAS XII - SEM 5", "TP 12.TL.2: Academic Synthesis Essay", "Penyusunan makalah penelitian akademik 1.000 kata dengan sitasi standar APA/MLA (25 JP)", FILL_NODE_PURPLE),
        ("KELAS XII - SEM 5", "TP 12.TL.3: Dramatic Script & Performance", "Analisis naskah teater, subteks dialog, dan pementasan monolog/sandiwara panggung (20 JP)", FILL_NODE_PURPLE),
        ("KELAS XII - SEM 5", "TP 12.TL.4: Persuasive Oratory & Debate", "Retorika oratori tingkat tinggi (Ethos, Pathos, Logos) dan debat parlemen resmi (10 JP)", FILL_NODE_PURPLE),
        ("KELAS XII - SEM 6", "TP 12.TL.5: Admissions & Career Dossier", "Cover letter eksekutif, ATS CV canggih, dan personal statement beasiswa universitas (20 JP)", FILL_NODE_AMBER),
        ("KELAS XII - SEM 6", "TP 12.TL.6: Job Interview & Oral Defense", "Simulasi wawancara kerja STAR method dan sidang pertahanan karya ilmiah lisan (15 JP)", FILL_NODE_AMBER),
        ("KELAS XII - SEM 6", "TP 12.TL.7: Research Abstract & Infographic", "Penulisan abstrak ilmiah 250 kata dan simposium presentasi infografis saintifik (15 JP)", FILL_NODE_AMBER),
        ("KELAS XII - SEM 6", "TP 12.TL.8: Standardized Readiness (B2+)", "Kesiapan tes kemahiran standar internasional (IELTS/TOEFL format reading & writing) (10 JP)", FILL_NODE_AMBER)
    ]

    r += 1
    for i, (stage, code, desc, fill_node) in enumerate(flow_f_tl, 1):
        ws.row_dimensions[r].height = 26
        c_stg = ws.cell(row=r, column=1, value=stage)
        c_stg.font = FONT_DATA_BOLD
        c_stg.fill = FILL_LIGHT_BLUE if "XI" in stage else FILL_VERY_LIGHT
        c_stg.alignment = ALIGN_CENTER
        c_stg.border = CELL_BORDER

        c_box = ws.cell(row=r, column=2, value=code)
        c_box.font = FONT_DIAGRAM_NODE
        c_box.fill = fill_node
        c_box.alignment = ALIGN_CENTER
        c_box.border = BOX_BORDER
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)

        c_arrow = ws.cell(row=r, column=4, value="--->" if i < len(flow_f_tl) else "[KEMAHIRAN B2]")
        c_arrow.font = FONT_DIAGRAM_ARROW
        c_arrow.alignment = ALIGN_CENTER

        c_desc = ws.cell(row=r, column=5, value=desc)
        c_desc.font = FONT_DATA
        c_desc.alignment = ALIGN_LEFT
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=8)
        for c in range(5, 9):
            ws.cell(row=r, column=c).border = CELL_BORDER
            if r % 2 == 1:
                ws.cell(row=r, column=c).fill = FILL_ZEBRA
        r += 1

    r += 2
    # 4. SIKLUS PEDAGOGIS GBA & DEEP LEARNING DALAM SETIAP TP
    ws.cell(row=r, column=1, value="IV. INTEGRASI SIKLUS PENGAJARAN BERBASIS TEKS (GBA) & PEMBELAJARAN MENDALAM DI SETIAP TP").font = FONT_SEC_HEADER
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    for c in range(1, 9):
        ws.cell(row=r, column=c).fill = FILL_NAVY
    ws.row_dimensions[r].height = 24

    gba_cycle = [
        ("Tahap 1", "Building Knowledge of the Field (BKoF)", "Membangun Pengetahuan Konteks", "Eksplorasi isu nyata, aktivasi skemata awal, dan penguasaan kosakata kunci melalui pemirsaan multimodal.", "Bermakna (Meaningful)"),
        ("Tahap 2", "Modelling of the Text (MoT)", "Dekonstruksi Model Teks", "Analisis sadar terhadap fungsi sosial, struktur organisasi gagasan, dan tata bahasa teks model otentik.", "Berkesadaran (Mindful)"),
        ("Tahap 3", "Joint Construction of the Text (JCoT)", "Konstruksi Bersama Teks", "Siswa bekerja kolaboratif dalam kelompok untuk memproduksi teks dengan bimbingan dan umpan balik guru.", "Menggembirakan (Joyful)"),
        ("Tahap 4", "Independent Construction (ICoT)", "Konstruksi Mandiri Teks", "Siswa secara independen menyusun teks lisan atau tulisan otentik untuk memecahkan masalah nyata.", "Kemandirian (Independent)"),
        ("Tahap 5", "Metacognitive Reflection", "Refleksi & Evaluasi Diri", "Siswa melakukan evaluasi diri (3-minute writing, rubrik analitik, peer-review) untuk perbaikan berkelanjutan.", "Reflektif (Meta-learning)")
    ]

    r += 1
    for step_num, gba_name, ind_name, op_desc, dl_pillar in gba_cycle:
        ws.row_dimensions[r].height = 30
        c1 = ws.cell(row=r, column=1, value=step_num)
        c1.font = FONT_DATA_BOLD
        c1.fill = FILL_LIGHT_BLUE
        c1.alignment = ALIGN_CENTER
        c1.border = CELL_BORDER

        c2 = ws.cell(row=r, column=2, value=gba_name)
        c2.font = FONT_DATA_BOLD
        c2.alignment = ALIGN_LEFT
        c2.border = CELL_BORDER

        c3 = ws.cell(row=r, column=3, value=ind_name)
        c3.font = FONT_DATA
        c3.alignment = ALIGN_LEFT
        c3.border = CELL_BORDER

        c4 = ws.cell(row=r, column=4, value=op_desc)
        c4.font = FONT_DATA
        c4.alignment = ALIGN_TOP_LEFT
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
        for c in range(4, 8):
            ws.cell(row=r, column=c).border = CELL_BORDER

        c5 = ws.cell(row=r, column=8, value=dl_pillar)
        c5.font = FONT_DATA_BOLD
        c5.fill = FILL_ACCENT_GREEN
        c5.alignment = ALIGN_CENTER
        c5.border = CELL_BORDER

        r += 1

    auto_fit_cols(ws, min_w=12, max_w=48)


# ==============================================================================
# BUILD SHEET: DETAIL TABEL ATP PER KELAS
# ==============================================================================

def build_sheet_atp_table(ws, sheet_title, page_title, class_name, fase_name, subject_name, alokasi_jp, atp_data):
    ws.title = sheet_title
    apply_title(ws, f"ALUR DAN TUJUAN PEMBELAJARAN (ATP) - {class_name.upper()} ({fase_name.upper()})",
                f"Mata Pelajaran: {subject_name} | Jenjang SMA | Total Alokasi: {alokasi_jp} JP / Tahun", 10)

    school_data = [
        ("Satuan Pendidikan", "Sekolah Menengah Atas (SMA)"),
        ("Mata Pelajaran", subject_name),
        ("Fase / Kelas", f"{fase_name} / {class_name}"),
        ("Tahun Pelajaran", "2025/2026"),
        ("Alokasi Waktu", f"{alokasi_jp} Jam Pelajaran (JP)"),
        ("Rujukan Regulasi", "Permendikdasmen No. 13 Tahun 2025 & Panduan Mapel Bahasa Inggris Puskurjar 2025")
    ]
    r_start = add_school_identity(ws, school_data, start_row=4)

    headers = [
        "No", "Kode TP", "Elemen Capaian Pembelajaran",
        "Rumusan Tujuan Pembelajaran (TP)",
        "Indikator Ketercapaian TP (KKTP)",
        "Materi Pokok & Ruang Lingkup",
        "Genre Teks", "Alur Penahapan", "JP",
        "Profil Pelajar Pancasila", "Rencana Asesmen Autentik"
    ]

    ws.row_dimensions[r_start].height = 28
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=r_start, column=col_idx, value=h)
        cell.font = FONT_TBL_HEADER
        cell.fill = FILL_NAVY
        cell.alignment = ALIGN_CENTER
        cell.border = CELL_BORDER

    curr_row = r_start + 1
    for row_val in atp_data:
        no, kode, elemen, tp, kktp, materi, genre, alur, jp, profil, asesmen = row_val
        ws.row_dimensions[curr_row].height = 70

        ws.cell(row=curr_row, column=1, value=no).font = FONT_DATA_NUMBER
        ws.cell(row=curr_row, column=1).alignment = ALIGN_CENTER
        ws.cell(row=curr_row, column=1).border = CELL_BORDER

        ws.cell(row=curr_row, column=2, value=kode).font = FONT_DATA_BOLD
        ws.cell(row=curr_row, column=2).alignment = ALIGN_CENTER
        ws.cell(row=curr_row, column=2).border = CELL_BORDER

        ws.cell(row=curr_row, column=3, value=elemen).font = FONT_DATA_BOLD
        ws.cell(row=curr_row, column=3).alignment = ALIGN_CENTER
        ws.cell(row=curr_row, column=3).fill = FILL_LIGHT_BLUE
        ws.cell(row=curr_row, column=3).border = CELL_BORDER

        ws.cell(row=curr_row, column=4, value=tp).font = FONT_DATA
        ws.cell(row=curr_row, column=4).alignment = ALIGN_TOP_LEFT
        ws.cell(row=curr_row, column=4).border = CELL_BORDER

        ws.cell(row=curr_row, column=5, value=kktp).font = FONT_DATA
        ws.cell(row=curr_row, column=5).alignment = ALIGN_TOP_LEFT
        ws.cell(row=curr_row, column=5).border = CELL_BORDER

        ws.cell(row=curr_row, column=6, value=materi).font = FONT_DATA
        ws.cell(row=curr_row, column=6).alignment = ALIGN_TOP_LEFT
        ws.cell(row=curr_row, column=6).border = CELL_BORDER

        ws.cell(row=curr_row, column=7, value=genre).font = FONT_DATA_BOLD
        ws.cell(row=curr_row, column=7).alignment = ALIGN_CENTER
        ws.cell(row=curr_row, column=7).border = CELL_BORDER

        ws.cell(row=curr_row, column=8, value=alur).font = FONT_ITALIC_DESC
        ws.cell(row=curr_row, column=8).alignment = ALIGN_CENTER
        ws.cell(row=curr_row, column=8).border = CELL_BORDER

        ws.cell(row=curr_row, column=9, value=jp).font = FONT_DATA_NUMBER
        ws.cell(row=curr_row, column=9).alignment = ALIGN_CENTER
        ws.cell(row=curr_row, column=9).border = CELL_BORDER

        ws.cell(row=curr_row, column=10, value=profil).font = FONT_ITALIC_DESC
        ws.cell(row=curr_row, column=10).alignment = ALIGN_TOP_LEFT
        ws.cell(row=curr_row, column=10).border = CELL_BORDER

        ws.cell(row=curr_row, column=11, value=asesmen).font = FONT_DATA
        ws.cell(row=curr_row, column=11).alignment = ALIGN_TOP_LEFT
        ws.cell(row=curr_row, column=11).border = CELL_BORDER

        curr_row += 1

    # Total Row
    ws.row_dimensions[curr_row].height = 24
    ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=8)
    c_tot = ws.cell(row=curr_row, column=1, value=f"TOTAL ALOKASI WAKTU TUJUAN PEMBELAJARAN ({fase_name})")
    c_tot.font = FONT_TOTAL
    c_tot.alignment = ALIGN_CENTER
    c_tot.fill = FILL_TOTAL

    c_tot_jp = ws.cell(row=curr_row, column=9, value=f"=SUM(I{r_start+1}:I{curr_row-1})")
    c_tot_jp.font = FONT_TOTAL
    c_tot_jp.alignment = ALIGN_CENTER
    c_tot_jp.fill = FILL_TOTAL

    ws.merge_cells(start_row=curr_row, start_column=10, end_row=curr_row, end_column=11)
    c_tot_ket = ws.cell(row=curr_row, column=10, value="Target Jam Intrakurikuler")
    c_tot_ket.font = FONT_TOTAL
    c_tot_ket.alignment = ALIGN_CENTER
    c_tot_ket.fill = FILL_TOTAL

    for c in range(1, 12):
        ws.cell(row=curr_row, column=c).border = TOTAL_BORDER

    add_signatures(ws, start_row=curr_row + 2, max_col=11)

    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 24
    ws.column_dimensions['D'].width = 38
    ws.column_dimensions['E'].width = 38
    ws.column_dimensions['F'].width = 36
    ws.column_dimensions['G'].width = 22
    ws.column_dimensions['H'].width = 18
    ws.column_dimensions['I'].width = 8
    ws.column_dimensions['J'].width = 24
    ws.column_dimensions['K'].width = 28
    ws.views.sheetView[0].showGridLines = True


# ==============================================================================
# MAIN GENERATOR FUNCTION FOR ATP
# ==============================================================================

def generate_atp_workbook(output_file):
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove default sheet

    # 1. Sheet Panduan & Prinsip ATP
    ws1 = wb.create_sheet(title="Panduan_&_Prinsip_ATP")
    build_sheet_prinsip_atp(ws1)

    # 2. Sheet Diagram Alur Pembelajaran (Visual Roadmap)
    ws2 = wb.create_sheet(title="Diagram_Alur_Pembelajaran")
    build_sheet_diagram_alur(ws2)

    # 3. Sheet ATP Kelas X Fase E Wajib
    ws3 = wb.create_sheet(title="ATP_Kelas_X_Wajib")
    build_sheet_atp_table(ws3, "ATP_Kelas_X_Wajib", "ATP Kelas X", "Kelas X", "Fase E", "Bahasa Inggris Umum (Wajib)", 60, ATP_X_WAJIB)

    # 4. Sheet ATP Kelas XI Fase F Wajib
    ws4 = wb.create_sheet(title="ATP_Kelas_XI_Wajib")
    build_sheet_atp_table(ws4, "ATP_Kelas_XI_Wajib", "ATP Kelas XI Wajib", "Kelas XI", "Fase F Umum", "Bahasa Inggris Umum (Wajib)", 60, ATP_XI_WAJIB)

    # 5. Sheet ATP Kelas XI Fase F Tingkat Lanjut
    ws5 = wb.create_sheet(title="ATP_Kelas_XI_TL")
    build_sheet_atp_table(ws5, "ATP_Kelas_XI_TL", "ATP Kelas XI Tingkat Lanjut", "Kelas XI", "Fase F Tingkat Lanjut", "Bahasa Inggris Tingkat Lanjut (Pilihan)", 160, ATP_XI_TL)

    # 6. Sheet ATP Kelas XII Fase F Wajib
    ws6 = wb.create_sheet(title="ATP_Kelas_XII_Wajib")
    build_sheet_atp_table(ws6, "ATP_Kelas_XII_Wajib", "ATP Kelas XII Wajib", "Kelas XII", "Fase F Umum", "Bahasa Inggris Umum (Wajib)", 54, ATP_XII_WAJIB)

    # 7. Sheet ATP Kelas XII Fase F Tingkat Lanjut
    ws7 = wb.create_sheet(title="ATP_Kelas_XII_TL")
    build_sheet_atp_table(ws7, "ATP_Kelas_XII_TL", "ATP Kelas XII Tingkat Lanjut", "Kelas XII", "Fase F Tingkat Lanjut", "Bahasa Inggris Tingkat Lanjut (Pilihan)", 140, ATP_XII_TL)

    wb.save(output_file)
    print(f"ATP Workbook successfully created at: {output_file}")

if __name__ == "__main__":
    out_dir = r"C:\Users\KOMPUTERKU2\Documents\English SMA\2-Alur dan Tujuan Pembelajaran"
    os.makedirs(out_dir, exist_ok=True)
    out_file = os.path.join(out_dir, "Alur_dan_Tujuan_Pembelajaran_Bahasa_Inggris_SMA.xlsx")
    generate_atp_workbook(out_file)
