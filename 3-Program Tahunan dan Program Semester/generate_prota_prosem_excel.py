# -*- coding: utf-8 -*-
"""
Script to generate a SINGLE unified Excel workbook for:
PROGRAM TAHUNAN (PROTA) DAN PROGRAM SEMESTER (PROSEM) BAHASA INGGRIS SMA
(Kelas X Fase E, Kelas XI & XII Fase F Wajib, Kelas XI & XII Fase F Tingkat Lanjut)
Conforms to:
- Permendikdasmen No. 13 Tahun 2025
- Capaian Pembelajaran 2025/2026
- Panduan Mapel Bahasa Inggris Puskurjar Kemendikdasmen Revisi 3 (Sept 2025)
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
FONT_TBL_HEADER = Font(name=FONT_FAMILY, size=9, bold=True, color="FFFFFF")
FONT_TBL_SUBHEADER = Font(name=FONT_FAMILY, size=8, bold=True, color="0F2942")
FONT_DATA = Font(name=FONT_FAMILY, size=9, color="1E293B")
FONT_DATA_BOLD = Font(name=FONT_FAMILY, size=9, bold=True, color="1E293B")
FONT_DATA_CENTER = Font(name=FONT_FAMILY, size=9, color="1E293B")
FONT_DATA_NUMBER = Font(name=FONT_FAMILY, size=9, bold=True, color="0F2942")
FONT_ITALIC_DESC = Font(name=FONT_FAMILY, size=8, italic=True, color="475569")
FONT_TOTAL = Font(name=FONT_FAMILY, size=9, bold=True, color="0F2942")

FILL_NAVY = PatternFill(start_color="0F2942", end_color="0F2942", fill_type="solid")
FILL_ROYAL = PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid")
FILL_BLUE_HEADER = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")
FILL_LIGHT_BLUE = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
FILL_VERY_LIGHT = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
FILL_ZEBRA = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
FILL_TOTAL = PatternFill(start_color="BFDBFE", end_color="BFDBFE", fill_type="solid")
FILL_ACCENT_AMBER = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
FILL_ACCENT_GREEN = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
FILL_STS_SAS = PatternFill(start_color="FDE68A", end_color="FDE68A", fill_type="solid")
FILL_LIBUR = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

BORDER_THIN = Side(style='thin', color="CBD5E1")
BORDER_DOUBLE = Side(style='double', color="0F2942")
CELL_BORDER = Border(left=BORDER_THIN, right=BORDER_THIN, top=BORDER_THIN, bottom=BORDER_THIN)
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
# DATA PROTA (KELAS X WAJIB, XI WAJIB, XI TL, XII WAJIB, XII TL)
# ==============================================================================

DATA_PROTA_X_WAJIB = [
    ("SEMESTER 1 (GANJIL)", 1, "TP 10.1", 
     "Menganalisis fungsi sosial, struktur teks, dan unsur kebahasaan teks deskriptif lisan dan tulis tentang atlet berprestasi dan destinasi wisata dunia, serta memproduksi teks deskriptif multimodal.",
     "Describing Great Athletes & World-Class Tourism Destinations:\n• Identification & Descriptions\n• Adjectives of personality & physical appearance\n• Sensory language & Simple Present Tense\n• Multimodal promotional brochure/poster",
     "Descriptive Text", 10, "Modul Ajar 1 & LKPD"),
    
    ("SEMESTER 1 (GANJIL)", 2, "TP 10.2", 
     "Mengidentifikasi konteks, gagasan utama, dan detail spesifik peristiwa sejarah dan pengalaman pribadi masa lalu (Recount Text), serta menyusun teks recount personal dan faktual dengan runtut dan akurat.",
     "Exploring Historical Events & Inspiring Personal Recounts:\n• Orientation, Series of Events, Re-orientation\n• Action verbs in Simple Past & Past Continuous\n• Temporal connectives & time linkers\n• Short biographical video / journal entry",
     "Recount Text", 10, "Modul Ajar 2 & LKPD"),
    
    ("SEMESTER 1 (GANJIL)", 3, "TP 10.3", 
     "Menganalisis nilai-nilai moral, pesan tersirat, dan alur naratif cerita rakyat Indonesia dan legenda dunia (Narrative Text), serta mendemonstrasikan pembacaan cerita atau menulis ulang narasi dengan variasi akhir cerita.",
     "Indonesian Folklore & World Legends with Timeless Moral Lessons:\n• Orientation, Complication, Resolution, Coda\n• Past Tense, Direct/Indirect Speech\n• Conflict resolution and ethical values\n• Digital storytelling / Roleplay",
     "Narrative Text", 10, "Modul Ajar 3 & LKPD"),
    
    ("SEMESTER 1 (GANJIL)", 4, "SUMATIF", 
     "Melaksanakan Asesmen Sumatif Lingkup Materi, Sumatif Tengah Semester (STS), dan Sumatif Akhir Semester (SAS) Ganjil.",
     "Evaluasi capaian TP 10.1 - 10.3 (Speaking & Writing portfolio test)",
     "Asesmen & Cadangan", 6, "Asesmen Sumatif"),
    
    ("SEMESTER 2 (GENAP)", 5, "TP 10.4", 
     "Memahami dan menyusun teks prosedur lisan dan tulis terkait pola hidup sehat (healthy living), resep kuliner nusantara, serta panduan keselamatan digital/media sosial.",
     "Procedures for Healthy Living & Safe Digital Practices:\n• Goal, Materials/Equipment, Steps\n• Imperative sentences, action verbs, sequencers\n• Adverbials of time, manner, and place\n• Tutorial video / Infographic presentation",
     "Procedure Text", 10, "Modul Ajar 4 & LKPD"),
    
    ("SEMESTER 2 (GENAP)", 6, "TP 10.5", 
     "Menganalisis teks laporan hasil observasi (Report Text) mengenai fenomena alam, fauna terancam punah Indonesia, dan inovasi sains, serta mempresentasikan laporan ilmiah ringkas secara multimodal.",
     "Reporting Natural Phenomena & Biodiversity Conservation:\n• General Classification & Detailed Descriptions\n• General nouns, technical vocabulary, passive voice\n• Fact-based observation & critical viewing\n• Mini science report / Poster exhibition",
     "Report Text", 10, "Modul Ajar 5 & LKPD"),
    
    ("SEMESTER 2 (GENAP)", 7, "TP 10.6", 
     "Mengevaluasi argumen, bukti pendukung, dan tesis dalam teks eksposisi analitis (Analytical Exposition) bertema pelestarian lingkungan dan gaya hidup berkelanjutan, serta menulis esai argumentatif.",
     "Advocating Environmental Sustainability & Waste Reduction:\n• Thesis, Arguments, Reiteration\n• Causal conjunctions, mental verbs, evaluative language\n• Formulating logical stance and counter-arguments\n• Opinion essay / Student campaign speech",
     "Analytical Exposition", 10, "Modul Ajar 6 & LKPD"),

    ("SEMESTER 2 (GENAP)", 8, "SUMATIF", 
     "Melaksanakan Asesmen Sumatif Lingkup Materi, Sumatif Tengah Semester (STS), dan Sumatif Akhir Semester (SAS/SAT) Genap.",
     "Evaluasi capaian TP 10.4 - 10.6 (Sumatif tulis & pameran proyek berbahasa Inggris)",
     "Asesmen & Cadangan", 6, "Asesmen Sumatif")
]

DATA_PROTA_XI_WAJIB = [
    ("SEMESTER 3 (GANJIL)", 1, "TP 11.W.1", 
     "Menganalisis fungsi sosial, struktur teks, dan unsur kebahasaan teks eksposisi analitis (Analytical Exposition) tentang isu-isu sosial dan teknologi terkini di masyarakat, serta memproduksi teks eksposisi lisan dan tulis yang meyakinkan.",
     "Analytical Exposition on Contemporary Social & Digital Issues:\n• Thesis, series of arguments, reiteration\n• Mental verbs, causal conjunctions, passive constructions\n• Fact vs opinion, logical reasoning\n• Short argumentative essay / Mini presentation",
     "Analytical Exposition", 10, "Modul Ajar 1 & LKPD"),

    ("SEMESTER 3 (GANJIL)", 2, "TP 11.W.2", 
     "Mengidentifikasi gagasan utama, argumen pendukung, dan rekomendasi persuasif dalam teks eksposisi hortatori (Hortatory Exposition) tentang keselamatan publik, kesehatan mental remaja, dan konsumsi etis, serta menyusun teks anjuran bertindak.",
     "Hortatory Exposition for Teen Well-being & Public Safety:\n• Thesis, arguments, recommendation\n• Modals of recommendation (should, ought to, must)\n• Persuasive and emotive language\n• Advocacy poster / Campaign speech",
     "Hortatory Exposition", 10, "Modul Ajar 2 & LKPD"),

    ("SEMESTER 3 (GANJIL)", 3, "TP 11.W.3", 
     "Menganalisis struktur narasi, pengembangan konflik karakter, dan pesan moral dalam teks naratif kontemporer dan fabel modern (Modern/Fractured Fairy Tales), serta menyusun kembali cerita dengan sudut pandang unik.",
     "Contemporary Narratives & Fractured Tales with Meaningful Messages:\n• Orientation, complication, evaluation, resolution\n• Past tenses, descriptive adjectives, dialogue formatting\n• Moral judgment and alternative character perspectives\n• Creative short story writing / Roleplay",
     "Narrative Text", 10, "Modul Ajar 3 & LKPD"),

    ("SEMESTER 3 (GANJIL)", 4, "SUMATIF", 
     "Melaksanakan Asesmen Sumatif Lingkup Materi, Sumatif Tengah Semester (STS), dan Sumatif Akhir Semester (SAS) Ganjil.",
     "Evaluasi capaian TP 11.W.1 - 11.W.3 (Tes tulis dan portofolio teks eksposisi)",
     "Asesmen & Cadangan", 6, "Asesmen Sumatif"),

    ("SEMESTER 4 (GENAP)", 5, "TP 11.W.4", 
     "Menganalisis proses kausal dan urutan kejadian fenomena alam dan sosial budaya (Explanation Text), serta menyusun teks eksplanasi ilmiah populer dengan bahasa yang akurat dan runtut.",
     "Explanation of Natural & Socio-Cultural Phenomena:\n• General statement, sequenced explanation, concluding remark\n• Cause-and-effect connectives (because of, due to, consequently)\n• Passive voice in formal register, technical terms\n• Illustrated explanation diagram / Scientific video voiceover",
     "Explanation Text", 10, "Modul Ajar 4 & LKPD"),

    ("SEMESTER 4 (GENAP)", 6, "TP 11.W.5", 
     "Menimbang dua sudut pandang berimbang pro dan kontra terkait isu remaja dan lingkungan sekolah (Discussion Text), serta berpartisipasi aktif dalam diskusi kelompok terarah.",
     "Discussion on School & Teen Realities (Balanced Perspectives):\n• Issue, arguments for (pro), arguments against (con), conclusion/recommendation\n• Contrasting conjunctions (however, on the other hand, in contrast)\n• Expressing agreement, disagreement, and balanced opinion\n• Structured classroom round-table discussion",
     "Discussion Text", 10, "Modul Ajar 5 & LKPD"),

    ("SEMESTER 4 (GENAP)", 7, "TP 11.W.6", 
     "Mengevaluasi dan menyusun teks resensi kritis (Review Text) terhadap film edukatif, novel remaja, atau karya seni budaya lokal nusantara.",
     "Reviewing Inspiring Films, Books, & Cultural Masterpieces:\n• Orientation, interpretative recount, evaluation, evaluative summation\n• Evaluative adjectives, metaphoric expressions, constructive critique\n• Published student review in school bulletin / blog",
     "Review Text", 10, "Modul Ajar 6 & LKPD"),

    ("SEMESTER 4 (GENAP)", 8, "SUMATIF", 
     "Melaksanakan Asesmen Sumatif Lingkup Materi, Sumatif Tengah Semester (STS), dan Sumatif Akhir Semester (SAT / Kenaikan Kelas).",
     "Evaluasi capaian TP 11.W.4 - 11.W.6 (Ujian tulis, pameran resensi & portofolio)",
     "Asesmen & Cadangan", 6, "Asesmen Sumatif")
]

DATA_PROTA_XI_TL = [
    ("SEMESTER 3 (GANJIL)", 1, "TP 11.TL.1", 
     "Mengevaluasi wacana argumentatif kritis tingkat lanjut dalam teks eksposisi analitis bertema disrupsi kecerdasan buatan (Artificial Intelligence), etika siber, dan transformasi digital secara objektif dan mendalam.",
     "Advanced Analytical Exposition on AI Ethics & Digital Society:\n• Complex thesis formulation & multi-perspective arguments\n• Abstract nominalization, modal adverbs of certainty\n• Logical connectors of condition, concession & result\n• Formal academic argument paper / Analytical podcast",
     "Analytical Exposition (Advanced)", 20, "Modul Ajar 1 (5 JP x 4 mgg)"),

    ("SEMESTER 3 (GANJIL)", 2, "TP 11.TL.2", 
     "Menganalisis dan menyusun teks eksposisi hortatori tingkat lanjut yang mendesak tindakan solutif terkait transisi energi terbarukan, keadilan iklim, dan konsumsi bertanggung jawab.",
     "Hortatory Exposition for Global Climate Action & Sustainable Future:\n• Thesis, Evaluative Arguments, Urgent Recommendations\n• Emotive vocabulary, strong modal auxiliaries (must, ought to)\n• Rhetorical questions & persuasive appeals (Ethos, Logos, Pathos)\n• Persuasive policy brief / Student advocacy pitch",
     "Hortatory Exposition", 25, "Modul Ajar 2 (5 JP x 5 mgg)"),

    ("SEMESTER 3 (GANJIL)", 3, "TP 11.TL.3", 
     "Mengapresiasi, menganalisis kedalaman psikologis karakter, dan merekonstruksi narasi sastra fiksi/cerpen klasik dan kontemporer dengan alur bertingkat dan dilema etika.",
     "Literary Analysis of Advanced Narratives & Short Stories:\n• Narrative arc, foreshadowing, symbolism, stream of consciousness\n• Literary devices: metaphor, irony, character motivation\n• Creative adaptation & alternative endings\n• Literary review essay / Dramatic reader's theatre",
     "Advanced Narrative & Short Story", 25, "Modul Ajar 3 (5 JP x 5 mgg)"),

    ("SEMESTER 3 (GANJIL)", 4, "TP 11.TL.4", 
     "Menganalisis gaya bahasa stilistika, makna simbolik, dan pesan kritik sosial dalam puisi bahasa Inggris dan lirik lagu dunia yang berpengaruh.",
     "Stylistic & Societal Reading of English Poetry & Protest Songs:\n• Meter, stanza structure, imagery, tone, and subtext\n• Figuring social inequality, peace, and human dignity\n• Comparative poetry analysis\n• Poetry recitation / Lyric critique presentation",
     "Poetry & Song Critique", 10, "Modul Ajar 4 (5 JP x 2 mgg)"),

    ("SEMESTER 3 (GANJIL)", 5, "SUMATIF", 
     "Melaksanakan Asesmen Sumatif Lingkup Materi, Sumatif Tengah Semester (STS), dan Sumatif Akhir Semester (SAS) Ganjil.",
     "Ujian analitik lisan (viva/presentasi) dan portofolio esai akademik",
     "Asesmen & Cadangan", 10, "Asesmen Sumatif (5 JP x 2 mgg)"),

    ("SEMESTER 4 (GENAP)", 6, "TP 11.TL.5", 
     "Mengonstruksi dan menimbang dua sudut pandang berimbang dalam teks diskusi (Discussion Text) mengenai kontroversi e-sports/gaming kompetitif, media sosial, dan kecerdasan buatan di dunia pendidikan.",
     "Official Puskurjar Model: Discussion Text on Online Games & Social Issues:\n• Issue statement, Arguments For (Pro), Arguments Against (Con), Balanced Recommendation\n• Contrasting connectives (On the contrary, Whereas, Conversely)\n• Hedging/cautious language (It is widely argued, Evidence suggests)\n• Live Oxford-style classroom debate & balanced synthesis essay",
     "Discussion Text (Official Puskurjar)", 25, "Modul Ajar 5 (5 JP x 5 mgg)"),

    ("SEMESTER 4 (GENAP)", 7, "TP 11.TL.6", 
     "Menganalisis mekanisme kausalitas ilmiah kompleks dan urutan proses terjadinya fenomena alam, astrofisika, dan inovasi rekayasa teknologi tingkat tinggi.",
     "Causal & Sequential Explanation of High-Tech & Astrophysical Phenomena:\n• General statement, sequenced explanation of how & why, concluding summary\n• Complex cause-effect transitions, passive voice in scientific registers\n• Technical terms, nominal phrases, schematic flowcharts\n• Explanatory documentary voiceover / Multimodal scientific article",
     "Explanation Text (Complex Causal)", 25, "Modul Ajar 6 (5 JP x 5 mgg)"),

    ("SEMESTER 4 (GENAP)", 8, "TP 11.TL.7", 
     "Mengevaluasi rekam jejak perjuangan, kepemimpinan transformasional, dan kontribusi historis tokoh-tokoh dunia dalam teks biografi mendalam dan memoar sejarah.",
     "Comprehensive Biography & Historical Memoirs of Visionary Leaders:\n• Orientation, chronological milestone recount, evaluative contribution\n• Past perfect, complex time clauses, descriptive appraisal\n• Historical research & character introspection\n• Biographical mini-monograph / Biographical podcast",
     "Biography & Historical Memoir", 20, "Modul Ajar 7 (5 JP x 4 mgg)"),

    ("SEMESTER 4 (GENAP)", 9, "TP 11.TL.8", 
     "Menyusun resensi dan kritik mendalam (Review Text) terhadap novel sastra, karya seni visual, atau film berbobot dengan rubrik penilaian artistik yang objektif.",
     "Critical Review of International Literature & Cinematic Works:\n• Orientation, interpretative recount, multi-criteria evaluation, summative evaluative punch\n• Aesthetic vocabulary, nuanced judgmental adjectives, cultural commentary\n• Critical review portfolio publication",
     "Review Text (Critical Arts Review)", 10, "Modul Ajar 8 (5 JP x 2 mgg)"),

    ("SEMESTER 4 (GENAP)", 10, "SUMATIF", 
     "Melaksanakan Asesmen Sumatif Lingkup Materi, Sumatif Tengah Semester (STS), dan Sumatif Akhir Semester (SAS/SAT) Genap.",
     "Gelar wicara akademik, pameran publikasi esai, dan ujian sumatif komprehensif",
     "Asesmen & Cadangan", 10, "Asesmen Sumatif (5 JP x 2 mgg)")
]

DATA_PROTA_XII_WAJIB = [
    ("SEMESTER 5 (GANJIL)", 1, "TP 12.W.1", 
     "Menganalisis fungsi sosial, struktur teks, dan unsur kebahasaan teks diskusi (Discussion Text) mengenai tren masa depan dunia kerja, kecerdasan buatan, dan pelestarian lingkungan, serta menyusun teks diskusi berimbang.",
     "Discussion Text on Future of Work & Environmental Sustainability:\n• Issue statement, arguments pro and con, balanced recommendation\n• Contrastive discourse markers, hedging, neutral register\n• Collaborative debate and written balanced discussion article",
     "Discussion Text", 10, "Modul Ajar 1 & LKPD"),

    ("SEMESTER 5 (GANJIL)", 2, "TP 12.W.2", 
     "Mengidentifikasi konteks faktual, struktur berita, dan sumber informasi dalam teks berita aktual (News Item Text), serta memproduksi naskah siaran berita (broadcast news) dan artikel berita cetak/daring.",
     "News Item Text & Media Literacy:\n• Newsworthy event, background events, sources\n• Headlines, reporting verbs, direct/indirect quotes, concise style\n• News broadcast simulation video / Mini journalism project",
     "News Item Text", 10, "Modul Ajar 2 & LKPD"),

    ("SEMESTER 5 (GANJIL)", 3, "TP 12.W.3", 
     "Mengevaluasi dan menyusun teks resensi kritis (Review Text) mengenai produk kreatif, buku, atau film lokal bernilai budaya untuk promosi warisan nusantara secara multimodal.",
     "Cultural Arts & Indonesian Creative Product Review:\n• Orientation, evaluation, interpretative recount, evaluative summary\n• Evaluative adjectives, cultural appreciation terms\n• Creative review blog / Multimodal social media post",
     "Review Text", 10, "Modul Ajar 3 & LKPD"),

    ("SEMESTER 5 (GANJIL)", 4, "SUMATIF", 
     "Melaksanakan Asesmen Sumatif Lingkup Materi, Sumatif Tengah Semester (STS), dan Sumatif Akhir Semester (SAS) Ganjil.",
     "Evaluasi capaian TP 12.W.1 - 12.W.3 (Tes tulis dan presentasi berita)",
     "Asesmen & Cadangan", 6, "Asesmen Sumatif"),

    ("SEMESTER 6 (GENAP)", 5, "TP 12.W.4", 
     "Menyusun surat lamaran pekerjaan formal (Job Application Letter) dan daftar riwayat hidup (Curriculum Vitae / Resume) yang rapi, profesional, dan sesuai standar dunia kerja modern.",
     "Job Application Letter & Modern Curriculum Vitae:\n• Sender/receiver address, salutation, body paragraphs, closing\n• Highlighting skills, education, and relevant experiences\n• Action verbs, formal professional register, polite tone\n• Complete job application dossier",
     "Application Letter & CV", 10, "Modul Ajar 4 & LKPD"),

    ("SEMESTER 6 (GENAP)", 6, "TP 12.W.5", 
     "Mempraktikkan percakapan wawancara kerja formal (Job Interview) dan keterampilan presentasi diri secara percaya diri, santun, dan runtut.",
     "Job Interview Simulation & Professional Self-Presentation:\n• Answering interview questions (strengths, goals, experience)\n• Asking clarifying questions, professional body language\n• Mock job interview simulation with peer/teacher feedback",
     "Job Interview Skills", 8, "Modul Ajar 5 & LKPD"),

    ("SEMESTER 6 (GENAP)", 7, "TP 12.W.6", 
     "Menafsirkan makna dan pesan tersirat dalam takarir (Caption), grafik, dan infografis multimodal, serta menyajikan takarir informatif dan komunikatif.",
     "Captions, Graphs, & Multimodal Infographics:\n• Structure of captions (lead, explanation, source)\n• Interpreting bar charts, line graphs, and pie charts\n• Writing concise, catchy, and accurate photo captions",
     "Caption & Infographic", 6, "Modul Ajar 6 & LKPD"),

    ("SEMESTER 6 (GENAP)", 8, "SUMATIF", 
     "Melaksanakan Ujian Praktik Sekolah Bahasa Inggris dan Asesmen Sumatif Akhir Jenjang (Kelulusan SMA).",
     "Ujian Praktik (Interview & Writing) dan Ujian Akhir Sekolah",
     "Ujian Sekolah / Akhir Jenjang", 4, "Kelulusan SMA")
]

DATA_PROTA_XII_TL = [
    ("SEMESTER 5 (GANJIL)", 1, "TP 12.TL.1", 
     "Menganalisis wacana debat akademik tingkat tinggi (Advanced Discussion) mengenai bioetika medis, rekayasa genetika, eksplorasi antariksa, dan masa depan kemanusiaan.",
     "Global Academic Discussion on Bioethics & Frontier Technologies:\n• Multi-layered argumentation, philosophical nuances, ethical stances\n• Advanced stance markers, concessive clauses, academic collocations\n• Scholarly dialogue / Policy deliberation roundtable",
     "Advanced Academic Discussion", 25, "Modul Ajar 1 (5 JP x 5 mgg)"),

    ("SEMESTER 5 (GANJIL)", 2, "TP 12.TL.2", 
     "Menyusun esai penelitian akademik (Academic Synthesis Essay) 5 paragraf dengan memadukan multi-referensi, kutipan langsung/tidak langsung, dan daftar pustaka standar internasional.",
     "Academic Research Essay Writing & Source Synthesis:\n• Formulating working thesis, literature synthesis, topic sentences\n• In-text citation protocols (APA/MLA style), paraphrasing, avoiding plagiarism\n• Cohesive devices, academic hedging, formal objective register\n• Complete 1,000-word academic synthesis paper",
     "Academic Synthesis Essay", 25, "Modul Ajar 2 (5 JP x 5 mgg)"),

    ("SEMESTER 5 (GANJIL)", 3, "TP 12.TL.3", 
     "Menganalisis teks naskah drama klasik/modern, menafsirkan subteks psikologis, serta memperagakan adegan monolog/dialog dramatik dengan intonasi dan pelafalan standar profesional.",
     "Dramatic Literature & Theatrical Script Performance:\n• Dramatic structure, conflict escalation, stage directions, subtext\n• Voice modulation, stress, intonation, emotional projection\n• Script adaptation & theatrical live performance / Dramatic monologue recording",
     "Dramatic Script Analysis & Play", 20, "Modul Ajar 3 (5 JP x 4 mgg)"),

    ("SEMESTER 5 (GANJIL)", 4, "TP 12.TL.4", 
     "Merancang dan membawakan pidato persuasif tingkat tinggi (Persuasive Oratory) dan mempraktikkan debat formal gaya parlemen dengan retorika yang meyakinkan.",
     "Persuasive Public Speaking & Parliamentary Debating:\n• Classical rhetoric: Ethos (credibility), Pathos (emotion), Logos (logic)\n• Parallelism, antithesis, persuasive prosody, rebuttal strategies\n• Formal speech delivery in international forum simulation",
     "Persuasive Oratory & Debate", 10, "Modul Ajar 4 (5 JP x 2 mgg)"),

    ("SEMESTER 5 (GANJIL)", 5, "SUMATIF", 
     "Melaksanakan Asesmen Sumatif Lingkup Materi, Sumatif Tengah Semester (STS), dan Sumatif Akhir Semester (SAS) Ganjil.",
     "Evaluasi performa pidato, asesmen esai sintesis, dan portofolio akademik",
     "Asesmen & Cadangan", 10, "Asesmen Sumatif (5 JP x 2 mgg)"),

    ("SEMESTER 6 (GENAP)", 6, "TP 12.TL.5", 
     "Menyusun korespondensi kerja profesional dan akademik tingkat internasional: Surat Lamaran Kerja (Job Application Letter), Curriculum Vitae (CV), dan Pernyataan Motivasi Beasiswa/Kuliah (Personal Statement).",
     "Professional Career & Academic University Admissions Portfolio:\n• Formal business letter format, salutation, professional tone\n• Modern ATS-friendly Curriculum Vitae / Resume design\n• Highlighting achievements, transferable skills, career aspirations\n• Comprehensive university/internship application package",
     "Professional Workplace & Academic Correspondence", 20, "Modul Ajar 5 (5 JP x 4 mgg)"),

    ("SEMESTER 6 (GENAP)", 7, "TP 12.TL.6", 
     "Mempraktikkan wawancara kerja profesional (Job Interview Simulation) dan pertahanan karya ilmiah lisan (Oral Defense) dengan respons spontan, taktis, dan register bahasa formal.",
     "Job Interview & Academic Oral Defense Simulation:\n• STAR Method (Situation, Task, Action, Result) in behavioral questions\n• Confident body language, active listening, asking strategic questions\n• Handling unexpected queries & defending academic conclusions\n• Mock professional interview with industry/alumni panelists",
     "Interview & Oral Defense Simulation", 15, "Modul Ajar 6 (5 JP x 3 mgg)"),

    ("SEMESTER 6 (GENAP)", 8, "TP 12.TL.7", 
     "Menyusun abstrak laporan penelitian ilmiah (Research Abstract) dan mendesain presentasi infografis saintifik untuk deseminasi publik.",
     "Mini Academic Paper Abstract & Scientific Infographic Poster:\n• 250-word concise structured abstract (Background, Method, Findings, Conclusion)\n• Concise reporting verbs, nominalization, data visualization synthesis\n• Academic poster symposium session",
     "Scientific Abstract & Infographic Presentation", 15, "Modul Ajar 7 (5 JP x 3 mgg)"),

    ("SEMESTER 6 (GENAP)", 9, "TP 12.TL.8", 
     "Mengembangkan kesiapan literasi bahasa Inggris standar internasional (International Academic Literacy Readiness: IELTS/TOEFL/CEFR B2+ reading & critical analytical writing).",
     "Standardized English Proficiency Readiness for Higher Education:\n• Advanced skimming, scanning, synthesizing inferential questions\n• Cohesion, coherence, and precision under timed conditions\n• Diagnostic proficiency benchmark test",
     "Standardized Academic Readiness", 10, "Modul Ajar 8 (5 JP x 2 mgg)"),

    ("SEMESTER 6 (GENAP)", 10, "SUMATIF", 
     "Melaksanakan Ujian Praktik Sekolah, Ujian Sumatif Akhir Jenjang (Kelulusan), dan Verifikasi Portofolio Kelulusan SMA.",
     "Ujian Praktik Bahasa Inggris Tingkat Lanjut & Asesmen Akhir Jenjang",
     "Asesmen Akhir Jenjang", 10, "Ujian Sekolah (5 JP x 2 mgg)")
]


# ==============================================================================
# DATA PROSEM (SEMESTER 1 - 6 WAJIB & TINGKAT LANJUT)
# ==============================================================================

# PROSEM 1: KELAS X WAJIB (GANJIL - 2 JP/MGG - 36 JP)
DATA_PROSEM_1_WAJIB = [
    (1, "TP 10.1", "Teks Deskriptif: Atlet Hebat & Destinasi Wisata Dunia", 
     "Describing Great Athletes & Tourism Destinations (Adjectives, Simple Present, Multimodal Brochure)", 10,
     ["MPLS", "MPLS", 2, 2, 2, 2, 2, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 2 JP"),
    (2, "TP 10.2", "Teks Recount: Peristiwa Bersejarah & Pengalaman Bermakna", 
     "Historical Recount & Inspiring Life Events (Simple Past, Past Continuous, Time linkers)", 10,
     ["", "", "", "", "", "", "", 2, 2, 2, 2, 2, "", "", "", "", "", "", "", "", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 2 JP"),
    (3, "STS", "Sumatif Tengah Semester (STS) Ganjil", 
     "Asesmen tertulis dan performa lisan lingkup TP 10.1 & 10.2", 2,
     ["", "", "", "", "", "", "", "", "", "", "", "", "STS", "", "", "", "", "", "", "", "", "RPR", "RPR", "LS", "LS"],
     "Tes Tengah Semester"),
    (4, "TP 10.3", "Teks Naratif: Cerita Rakyat Indonesia & Legenda Dunia", 
     "Indonesian Folklore & World Legends with Moral Values (Direct/Indirect Speech, Plot climax)", 10,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", 2, 2, 2, 2, 2, "", "", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 2 JP"),
    (5, "PENGUATAN", "Penguatan Portofolio & Cadangan", 
     "Remedial, pengayaan, dan penuntasan tugas proyek multimodal", 2,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 2, "", "", "RPR", "RPR", "LS", "LS"],
     "Remedial & Pengayaan"),
    (6, "SAS", "Sumatif Akhir Semester (SAS) Ganjil", 
     "Asesmen Sumatif Akhir Semester komprehensif fase E", 2,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 2, "SAS", "RPR", "RPR", "LS", "LS"],
     "Tes Akhir Semester")
]

# PROSEM 2: KELAS X WAJIB (GENAP - 2 JP/MGG - 36 JP)
DATA_PROSEM_2_WAJIB = [
    (1, "TP 10.4", "Teks Prosedur: Gaya Hidup Sehat & Keamanan Digital", 
     "Healthy Living Recipes & Cyber Safety Guides (Imperatives, Sequencers, Action Verbs)", 10,
     ["LU", 2, 2, 2, 2, 2, "", "", "", "", "", "", "", "", "", "", "", "", "", "LU", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 2 JP"),
    (2, "TP 10.5", "Teks Laporan Hasil Observasi (Report Text)", 
     "Natural Phenomena, Endangered Fauna & Science Inventions (Passive voice, General nouns)", 10,
     ["", "", "", "", "", "", 2, 2, 2, 2, 2, "", "", "", "", "", "", "", "", "LU", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 2 JP"),
    (3, "STS", "Sumatif Tengah Semester (STS) Genap", 
     "Asesmen lingkup materi prosedur dan report", 2,
     ["", "", "", "", "", "", "", "", "", "", "", "STS", "", "", "", "", "", "", "", "LU", "", "RPR", "RPR", "LS", "LS"],
     "Tes Tengah Semester"),
    (4, "TP 10.6", "Teks Eksposisi Analitis: Isu Lingkungan & Sampah Plastik", 
     "Analytical Exposition on Environmental Sustainability (Thesis, Causal conjunctions, Arguments)", 10,
     ["", "", "", "", "", "", "", "", "", "", "", "", 2, 2, 2, 2, 2, "", "", "LU", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 2 JP"),
    (5, "PENGUATAN", "Penguatan Karakter & Pameran Karya", 
     "Pameran poster dan gelar wicara mini bahasa Inggris siswa", 2,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 2, "", "LU", "", "RPR", "RPR", "LS", "LS"],
     "Gelar Karya Proyek"),
    (6, "SAT", "Sumatif Akhir Tahun (SAT) / Kenaikan Kelas", 
     "Asesmen Sumatif Akhir Tahun penentuan ketuntasan Fase E", 2,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 2, "LU", "SAT", "RPR", "RPR", "LS", "LS"],
     "Ujian Kenaikan Kelas")
]

# PROSEM 3: KELAS XI WAJIB (GANJIL - 2 JP/MGG - 36 JP)
DATA_PROSEM_3_WAJIB = [
    (1, "TP 11.W.1", "Teks Eksposisi Analitis: Isu Sosial & Literasi Digital", 
     "Analytical Exposition on Contemporary Social & Digital Issues (Thesis, Mental verbs, Logical reasoning)", 10,
     ["MPLS", "MPLS", 2, 2, 2, 2, 2, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 2 JP"),
    (2, "TP 11.W.2", "Teks Eksposisi Hortatori: Kesehatan Mental & Keselamatan", 
     "Hortatory Exposition for Teen Well-being & Public Safety (Should, Ought to, Recommendations)", 10,
     ["", "", "", "", "", "", "", 2, 2, 2, 2, 2, "", "", "", "", "", "", "", "", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 2 JP"),
    (3, "STS", "Sumatif Tengah Semester (STS) Ganjil", 
     "Asesmen tulis dan lisan lingkup TP 11.W.1 & 11.W.2", 2,
     ["", "", "", "", "", "", "", "", "", "", "", "", "STS", "", "", "", "", "", "", "", "", "RPR", "RPR", "LS", "LS"],
     "Tes Tengah Semester"),
    (4, "TP 11.W.3", "Teks Naratif Kontemporer: Fractured Fairy Tales", 
     "Contemporary Narratives & Fractured Tales with Moral Lessons (Dialogue formatting, Perspective shift)", 10,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", 2, 2, 2, 2, 2, "", "", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 2 JP"),
    (5, "PENGUATAN", "Penguatan Portofolio & Cadangan", 
     "Remedial, pengayaan, dan gelar karya esai argumentatif siswa", 2,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 2, "", "", "RPR", "RPR", "LS", "LS"],
     "Remedial & Pengayaan"),
    (6, "SAS", "Sumatif Akhir Semester (SAS) Ganjil", 
     "Asesmen Sumatif Akhir Semester Bahasa Inggris Wajib Kelas XI", 2,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 2, "SAS", "RPR", "RPR", "LS", "LS"],
     "Tes Akhir Semester")
]

# PROSEM 3: KELAS XI TINGKAT LANJUT (GANJIL - 5 JP/MGG - 90 JP)
DATA_PROSEM_3_TL = [
    (1, "TP 11.TL.1", "Advanced Analytical Exposition: AI Ethics & Digital Society", 
     "Critical Analysis on AI & Cyberethics (Complex thesis, Nominalization, Modal adverbs)", 20,
     ["LU", "LU", 5, 5, 5, 5, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "RPR", "RPR", "LS", "LS"],
     "4 Pertemuan @ 5 JP"),
    (2, "TP 11.TL.2", "Hortatory Exposition: Climate Action & Sustainable Future", 
     "Urgent Policy Recommendations & Eco-action (Emotive words, Modals, Rhetorical devices)", 25,
     ["", "", "", "", "", "", 5, 5, 5, 5, 5, "", "", "", "", "", "", "", "", "", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 5 JP"),
    (3, "STS (TL)", "Sumatif Tengah Semester (STS) Tingkat Lanjut", 
     "Asesmen esai analitis dan presentasi advokasi kebijakan", 5,
     ["", "", "", "", "", "", "", "", "", "", "", "STS", "", "", "", "", "", "", "", "", "", "RPR", "RPR", "LS", "LS"],
     "Tes Tengah Semester"),
    (4, "TP 11.TL.3", "Literary Analysis of Advanced Narratives & Short Stories", 
     "Psychological depth, Symbolism, Ethical dilemma in classic & contemporary short stories", 25,
     ["", "", "", "", "", "", "", "", "", "", "", "", 5, 5, 5, 5, 5, "", "", "", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 5 JP"),
    (5, "TP 11.TL.4", "Stylistic & Societal Reading of Poetry & Protest Songs", 
     "Stanza, imagery, social critique analysis in influential English poems and songs", 10,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 5, 5, "", "", "RPR", "RPR", "LS", "LS"],
     "2 Pertemuan @ 5 JP"),
    (6, "SAS (TL)", "Sumatif Akhir Semester (SAS) Tingkat Lanjut", 
     "Ujian viva lisan komprehensif dan uji portofolio esai sastra", 5,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 5, "SAS", "RPR", "RPR", "LS", "LS"],
     "Tes Akhir Semester")
]

# PROSEM 4: KELAS XI WAJIB (GENAP - 2 JP/MGG - 36 JP)
DATA_PROSEM_4_WAJIB = [
    (1, "TP 11.W.4", "Teks Eksplanasi: Fenomena Alam & Sosial Budaya Populer", 
     "Explanation of Natural & Socio-Cultural Phenomena (Cause-effect linkers, Passive voice)", 10,
     ["LU", 2, 2, 2, 2, 2, "", "", "", "", "", "", "", "", "", "", "", "", "", "LU", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 2 JP"),
    (2, "TP 11.W.5", "Teks Diskusi: Isu Remaja & Lingkungan Sekolah", 
     "Discussion on School & Teen Realities (Contrasting conjunctions, Balanced arguments)", 10,
     ["", "", "", "", "", "", 2, 2, 2, 2, 2, "", "", "", "", "", "", "", "", "LU", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 2 JP"),
    (3, "STS", "Sumatif Tengah Semester (STS) Genap", 
     "Asesmen lingkup materi eksplanasi dan diskusi", 2,
     ["", "", "", "", "", "", "", "", "", "", "", "STS", "", "", "", "", "", "", "", "LU", "", "RPR", "RPR", "LS", "LS"],
     "Tes Tengah Semester"),
    (4, "TP 11.W.6", "Teks Resensi Kritis: Apresiasi Film, Buku, & Seni Lokal", 
     "Reviewing Inspiring Films, Books, & Cultural Masterpieces (Evaluative adjectives, Summary)", 10,
     ["", "", "", "", "", "", "", "", "", "", "", "", 2, 2, 2, 2, 2, "", "", "LU", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 2 JP"),
    (5, "PENGUATAN", "Penguatan Karakter & Pameran Resensi", 
     "Gelar wicara apresiasi seni dan publikasi resensi mading/blog", 2,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 2, "", "LU", "", "RPR", "RPR", "LS", "LS"],
     "Gelar Karya Proyek"),
    (6, "SAT", "Sumatif Akhir Tahun (SAT) / Kenaikan Kelas", 
     "Asesmen Sumatif Akhir Tahun Bahasa Inggris Wajib Kelas XI", 2,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 2, "LU", "SAT", "RPR", "RPR", "LS", "LS"],
     "Ujian Kenaikan Kelas")
]

# PROSEM 4: KELAS XI TINGKAT LANJUT (GENAP - 5 JP/MGG - 90 JP)
DATA_PROSEM_4_TL = [
    (1, "TP 11.TL.5", "Discussion Text (Model Resmi Puskurjar: Online Games & AI)", 
     "Pro & Con Arguments, Contrasting connectives, Hedging language, Oxford-style debate", 25,
     ["LU", 5, 5, 5, 5, 5, "", "", "", "", "", "", "", "", "", "", "", "", "", "LU", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 5 JP (Modul Bab D)"),
    (2, "TP 11.TL.6", "Causal & Sequential Explanation of High-Tech & Space Science", 
     "Complex Cause-Effect, Passive Voice, Nominalization, Scientific flowcharts", 25,
     ["", "", "", "", "", "", 5, 5, 5, 5, 5, "", "", "", "", "", "", "", "", "LU", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 5 JP"),
    (3, "STS (TL)", "Sumatif Tengah Semester (STS) Tingkat Lanjut", 
     "Penilaian debat kelas berimbang dan penulisan teks eksplanasi sains", 5,
     ["", "", "", "", "", "", "", "", "", "", "", "STS", "", "", "", "", "", "", "", "LU", "", "RPR", "RPR", "LS", "LS"],
     "Tes Tengah Semester"),
    (4, "TP 11.TL.7", "Comprehensive Biography & Historical Memoirs of Leaders", 
     "Chronological milestones, Descriptive appraisal, Character introspection podcast", 20,
     ["", "", "", "", "", "", "", "", "", "", "", "", 5, 5, 5, 5, "", "", "", "LU", "", "RPR", "RPR", "LS", "LS"],
     "4 Pertemuan @ 5 JP"),
    (5, "TP 11.TL.8", "Critical Review of International Literature & Cinematic Works", 
     "Multi-criteria evaluation, Aesthetic vocabulary, Critical review portfolio", 10,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 5, 5, "", "LU", "", "RPR", "RPR", "LS", "LS"],
     "2 Pertemuan @ 5 JP"),
    (6, "SAT (TL)", "Sumatif Akhir Tahun (SAT) Tingkat Lanjut", 
     "Ujian Sumatif Akhir Tahun dan penentuan kelanjutan Fase F Tingkat Lanjut", 5,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 5, "LU", "SAT", "RPR", "RPR", "LS", "LS"],
     "Ujian Kenaikan Kelas")
]

# PROSEM 5: KELAS XII WAJIB (GANJIL - 2 JP/MGG - 36 JP)
DATA_PROSEM_5_WAJIB = [
    (1, "TP 12.W.1", "Teks Diskusi: Masa Depan Dunia Kerja & Etika Lingkungan", 
     "Discussion Text on Future of Work & Environmental Sustainability (Contrastive markers, Hedging)", 10,
     ["MPLS", "MPLS", 2, 2, 2, 2, 2, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 2 JP"),
    (2, "TP 12.W.2", "Teks Berita Aktual (News Item Text) & Literasi Media", 
     "News Item Text & Media Literacy (Headlines, Reporting verbs, Broadcast simulation)", 10,
     ["", "", "", "", "", "", "", 2, 2, 2, 2, 2, "", "", "", "", "", "", "", "", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 2 JP"),
    (3, "STS", "Sumatif Tengah Semester (STS) Ganjil", 
     "Asesmen tulis dan praktik penyiaran berita singkat", 2,
     ["", "", "", "", "", "", "", "", "", "", "", "", "STS", "", "", "", "", "", "", "", "", "RPR", "RPR", "LS", "LS"],
     "Tes Tengah Semester"),
    (4, "TP 12.W.3", "Teks Resensi Seni Budaya Nusantara & Produk Kreatif", 
     "Cultural Arts & Indonesian Creative Product Review (Evaluative adjectives, Multimodal blog)", 10,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", 2, 2, 2, 2, 2, "", "", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 2 JP"),
    (5, "PENGUATAN", "Penguatan Portofolio Akademik & Cadangan", 
     "Remedial, pengayaan, dan penataan portofolio karya tulis kelas XII", 2,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 2, "", "", "RPR", "RPR", "LS", "LS"],
     "Remedial & Pengayaan"),
    (6, "SAS", "Sumatif Akhir Semester (SAS) Ganjil", 
     "Asesmen Sumatif Akhir Semester Bahasa Inggris Wajib Kelas XII", 2,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 2, "SAS", "RPR", "RPR", "LS", "LS"],
     "Tes Akhir Semester")
]

# PROSEM 5: KELAS XII TINGKAT LANJUT (GANJIL - 5 JP/MGG - 90 JP)
DATA_PROSEM_5_TL = [
    (1, "TP 12.TL.1", "Global Academic Discussion on Bioethics & Frontier Tech", 
     "Complex bioethical stances, Stance markers, Concessions, Scholarly deliberation", 25,
     ["LU", "LU", 5, 5, 5, 5, 5, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 5 JP"),
    (2, "TP 12.TL.2", "Academic Research Essay Writing & Multi-Source Synthesis", 
     "Formulating thesis, In-text citations (APA/MLA), Paraphrasing, 1,000-word academic paper", 25,
     ["", "", "", "", "", "", "", 5, 5, 5, 5, 5, "", "", "", "", "", "", "", "", "", "RPR", "RPR", "LS", "LS"],
     "5 Pertemuan @ 5 JP"),
    (3, "STS (TL)", "Sumatif Tengah Semester (STS) Tingkat Lanjut", 
     "Ujian sintesis esai akademik dan penilaian wacana etika global", 5,
     ["", "", "", "", "", "", "", "", "", "", "", "", "STS", "", "", "", "", "", "", "", "", "RPR", "RPR", "LS", "LS"],
     "Tes Tengah Semester"),
    (4, "TP 12.TL.3", "Dramatic Literature & Theatrical Script Performance", 
     "Subtext interpretation, Voice modulation, Intonation, Theatrical live / video staging", 20,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", 5, 5, 5, 5, "", "", "", "", "RPR", "RPR", "LS", "LS"],
     "4 Pertemuan @ 5 JP"),
    (5, "TP 12.TL.4", "Persuasive Public Speaking & Parliamentary Debating", 
     "Ethos, Pathos, Logos rhetoric, Antithesis, Rebuttal, International forum speech", 10,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 5, 5, "", "", "RPR", "RPR", "LS", "LS"],
     "2 Pertemuan @ 5 JP"),
    (6, "SAS (TL)", "Sumatif Akhir Semester (SAS) Tingkat Lanjut", 
     "Sidang kolokium akademik siswa dan ujian sumatif semester 5", 5,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 5, "SAS", "RPR", "RPR", "LS", "LS"],
     "Tes Akhir Semester")
]

# PROSEM 6: KELAS XII WAJIB (GENAP - 2 JP/MGG - 28 JP / 14 MGG EFEKTIF)
DATA_PROSEM_6_WAJIB = [
    (1, "TP 12.W.4", "Surat Lamaran Kerja (Job Application) & Curriculum Vitae (CV)", 
     "Job Application Letter & Modern Resume (Professional register, Action verbs, ATS layout)", 10,
     ["LU", 2, 2, 2, 2, 2, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
     "5 Pertemuan @ 2 JP"),
    (2, "TP 12.W.5", "Simulasi Wawancara Kerja (Job Interview) & Presentasi Diri", 
     "Job Interview Simulation & Professional Presentation (Confidence, STAR response technique)", 8,
     ["", "", "", "", "", "", 2, 2, 2, 2, "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
     "4 Pertemuan @ 2 JP"),
    (3, "TP 12.W.6", "Takarir (Captions) & Infografis Multimodal", 
     "Interpreting Graphs, Charts, & Modern Photo Captions (Visual literacy, Concise English)", 6,
     ["", "", "", "", "", "", "", "", "", "", 2, 2, 2, "", "", "", "", "", "", "", "", "", "", "", ""],
     "3 Pertemuan @ 2 JP"),
    (4, "UP", "Ujian Praktik Sekolah Bahasa Inggris Wajib", 
     "Ujian praktik kemampuan lisan (wawancara kerja) dan penulisan dokumen profesional", 2,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", "UP", "", "", "", "", "", "", "", "", "", "", ""],
     "Ujian Praktik Kelulusan"),
    (5, "ASAJ", "Asesmen Sumatif Akhir Jenjang (Kelulusan SMA)", 
     "Ujian Teori Asesmen Akhir Jenjang Bahasa Inggris Wajib SMA", 2,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "AJ", "", "", "", "", "", "", "", "", "", ""],
     "Ujian Akhir Kelulusan")
]

# PROSEM 6: KELAS XII TINGKAT LANJUT (GENAP - 5 JP/MGG - 70 JP / 14 MGG EFEKTIF)
DATA_PROSEM_6_TL = [
    (1, "TP 12.TL.5", "Professional Workplace & College Application Portfolio", 
     "Cover Letter, ATS-compliant CV/Resume, Scholarship Personal Statement", 20,
     ["LU", 5, 5, 5, 5, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
     "4 Pertemuan @ 5 JP"),
    (2, "TP 12.TL.6", "Job Interview & Academic Oral Defense Simulation", 
     "STAR method responses, Professional interview register, Defending research conclusions", 15,
     ["", "", "", "", "", 5, 5, 5, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
     "3 Pertemuan @ 5 JP"),
    (3, "TP 12.TL.7", "Scientific Abstract & Infographic Presentation", 
     "250-word structured abstract, Reporting verbs, Scientific infographic symposium", 15,
     ["", "", "", "", "", "", "", "", 5, 5, 5, "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
     "3 Pertemuan @ 5 JP"),
    (4, "TP 12.TL.8", "Standardized Academic English Proficiency Readiness", 
     "IELTS/TOEFL/CEFR B2+ critical reading, Analytical synthesis under timed conditions", 10,
     ["", "", "", "", "", "", "", "", "", "", "", 5, 5, "", "", "", "", "", "", "", "", "", "", "", ""],
     "2 Pertemuan @ 5 JP"),
    (5, "UP (TL)", "Ujian Praktik Sekolah Bahasa Inggris Tingkat Lanjut", 
     "Uji kemahiran berbicara profesional dan presentasi ilmiah bahasa Inggris", 5,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", "UP", "", "", "", "", "", "", "", "", "", "", ""],
     "Ujian Praktik Kelulusan"),
    (6, "ASAJ", "Asesmen Sumatif Akhir Jenjang (Kelulusan SMA)", 
     "Ujian Teori Asesmen Akhir Jenjang Tingkat Lanjut SMA", 5,
     ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "AJ", "", "", "", "", "", "", "", "", "", ""],
     "Ujian Akhir Kelulusan")
]


# ==============================================================================
# HELPER BUILDERS FOR WORKSHEETS
# ==============================================================================

def add_prota_sheet(wb, sheet_name, c_name, f_name, s_name, jp_tot, p_data):
    ws = wb.create_sheet(title=sheet_name)
    apply_title(ws, f"PROGRAM TAHUNAN (PROTA) - {c_name.upper()} ({f_name.upper()})",
                f"Mata Pelajaran: {s_name} | Alokasi: {jp_tot} JP / Tahun | Jenjang: SMA", 8)

    school_data = [
        ("Satuan Pendidikan", "Sekolah Menengah Atas (SMA)"),
        ("Mata Pelajaran", s_name),
        ("Fase / Kelas", f"{f_name} / {c_name}"),
        ("Tahun Pelajaran", "2025/2026"),
        ("Alokasi Waktu", f"{jp_tot} Jam Pelajaran (JP) / Tahun Ajaran"),
        ("Landasan Regulasi", "Permendikdasmen No. 13 Tahun 2025 & Panduan Mapel Bahasa Inggris Puskurjar 2025")
    ]
    r_start = add_school_identity(ws, school_data, start_row=4)

    headers = ["Semester", "No", "Kode TP", "Tujuan Pembelajaran (TP)", "Materi Esensial & Ruang Lingkup", "Genre Teks", "Alokasi (JP)", "Keterangan / Sumber Belajar"]
    ws.row_dimensions[r_start].height = 26
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=r_start, column=col_idx, value=h)
        cell.font = FONT_TBL_HEADER
        cell.fill = FILL_NAVY
        cell.alignment = ALIGN_CENTER
        cell.border = CELL_BORDER

    curr_row = r_start + 1
    for item in p_data:
        sem, no, kode, tp, materi, genre, jp, ket = item
        ws.row_dimensions[curr_row].height = 62

        ws.cell(row=curr_row, column=1, value=sem).font = FONT_DATA_BOLD
        ws.cell(row=curr_row, column=1).alignment = ALIGN_CENTER
        ws.cell(row=curr_row, column=1).fill = FILL_LIGHT_BLUE if "GANJIL" in sem else FILL_VERY_LIGHT
        ws.cell(row=curr_row, column=1).border = CELL_BORDER

        ws.cell(row=curr_row, column=2, value=no).font = FONT_DATA_NUMBER
        ws.cell(row=curr_row, column=2).alignment = ALIGN_CENTER
        ws.cell(row=curr_row, column=2).border = CELL_BORDER

        ws.cell(row=curr_row, column=3, value=kode).font = FONT_DATA_BOLD
        ws.cell(row=curr_row, column=3).alignment = ALIGN_CENTER
        ws.cell(row=curr_row, column=3).border = CELL_BORDER

        ws.cell(row=curr_row, column=4, value=tp).font = FONT_DATA
        ws.cell(row=curr_row, column=4).alignment = ALIGN_TOP_LEFT
        ws.cell(row=curr_row, column=4).border = CELL_BORDER

        ws.cell(row=curr_row, column=5, value=materi).font = FONT_DATA
        ws.cell(row=curr_row, column=5).alignment = ALIGN_TOP_LEFT
        ws.cell(row=curr_row, column=5).border = CELL_BORDER

        ws.cell(row=curr_row, column=6, value=genre).font = FONT_DATA_BOLD
        ws.cell(row=curr_row, column=6).alignment = ALIGN_CENTER
        ws.cell(row=curr_row, column=6).border = CELL_BORDER
        if "Sumatif" in genre or "Asesmen" in genre or "Ujian" in genre:
            ws.cell(row=curr_row, column=6).fill = FILL_STS_SAS

        ws.cell(row=curr_row, column=7, value=jp).font = FONT_DATA_NUMBER
        ws.cell(row=curr_row, column=7).alignment = ALIGN_CENTER
        ws.cell(row=curr_row, column=7).border = CELL_BORDER

        ws.cell(row=curr_row, column=8, value=ket).font = FONT_ITALIC_DESC
        ws.cell(row=curr_row, column=8).alignment = ALIGN_TOP_LEFT
        ws.cell(row=curr_row, column=8).border = CELL_BORDER

        curr_row += 1

    # Total Row
    ws.row_dimensions[curr_row].height = 24
    ws.merge_cells(start_row=curr_row, start_column=1, end_row=curr_row, end_column=6)
    c_tot = ws.cell(row=curr_row, column=1, value="TOTAL ALOKASI WAKTU PROGRAM TAHUNAN")
    c_tot.font = FONT_TOTAL
    c_tot.alignment = ALIGN_CENTER
    c_tot.fill = FILL_TOTAL

    c_tot_jp = ws.cell(row=curr_row, column=7, value=f"=SUM(G{r_start+1}:G{curr_row-1})")
    c_tot_jp.font = FONT_TOTAL
    c_tot_jp.alignment = ALIGN_CENTER
    c_tot_jp.fill = FILL_TOTAL

    c_tot_ket = ws.cell(row=curr_row, column=8, value=f"{jp_tot} JP / Tahun")
    c_tot_ket.font = FONT_TOTAL
    c_tot_ket.alignment = ALIGN_CENTER
    c_tot_ket.fill = FILL_TOTAL

    for c in range(1, 9):
        ws.cell(row=curr_row, column=c).border = TOTAL_BORDER

    add_signatures(ws, start_row=curr_row + 2, max_col=8)

    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 6
    ws.column_dimensions['C'].width = 12
    ws.column_dimensions['D'].width = 38
    ws.column_dimensions['E'].width = 44
    ws.column_dimensions['F'].width = 24
    ws.column_dimensions['G'].width = 14
    ws.column_dimensions['H'].width = 25
    ws.views.sheetView[0].showGridLines = True


def add_prosem_sheet(wb, sheet_name, sem_num, c_name, f_name, s_name, jp_w, s_data):
    is_ganjil = (sem_num % 2 == 1)
    sem_title = f"SEMESTER {sem_num} ({'GANJIL' if is_ganjil else 'GENAP'})"
    ws = wb.create_sheet(title=sheet_name)

    if is_ganjil:
        months = [("JULI", 4), ("AGUSTUS", 5), ("SEPTEMBER", 4), ("OKTOBER", 4), ("NOVEMBER", 4), ("DESEMBER", 4)]
    else:
        months = [("JANUARI", 4), ("FEBRUARI", 4), ("MARET", 5), ("APRIL", 4), ("MEI", 4), ("JUNI", 4)]

    total_weeks = sum(m[1] for m in months)
    max_col = 5 + total_weeks + 1

    title = f"PROGRAM SEMESTER (PROSEM) {sem_title}"
    subtitle = f"Mata Pelajaran: {s_name} | {c_name} ({f_name}) | Beban: {jp_w} JP / Minggu"
    apply_title(ws, title, subtitle, max_col)

    school_data = [
        ("Satuan Pendidikan", "Sekolah Menengah Atas (SMA)"),
        ("Mata Pelajaran", s_name),
        ("Fase / Kelas / Semester", f"{f_name} / {c_name} / {sem_title}"),
        ("Tahun Pelajaran", "2025/2026"),
        ("Beban Belajar", f"{jp_w} JP / Minggu Tatap Muka")
    ]
    r_start = add_school_identity(ws, school_data, start_row=4)

    r_h1 = r_start
    r_h2 = r_start + 1
    ws.row_dimensions[r_h1].height = 20
    ws.row_dimensions[r_h2].height = 18

    fixed_headers = [(1, "No"), (2, "Kode"), (3, "Tujuan Pembelajaran (TP)"), (4, "Materi Esensial & Genre Teks"), (5, "JP")]
    for col_idx, h_text in fixed_headers:
        ws.merge_cells(start_row=r_h1, start_column=col_idx, end_row=r_h2, end_column=col_idx)
        c = ws.cell(row=r_h1, column=col_idx, value=h_text)
        c.font = FONT_TBL_HEADER
        c.fill = FILL_NAVY
        c.alignment = ALIGN_CENTER
        for r_sub in range(r_h1, r_h2 + 1):
            ws.cell(row=r_sub, column=col_idx).border = CELL_BORDER
            ws.cell(row=r_sub, column=col_idx).fill = FILL_NAVY

    curr_c = 6
    week_global_idx = 1
    for m_name, m_weeks in months:
        ws.merge_cells(start_row=r_h1, start_column=curr_c, end_row=r_h1, end_column=curr_c + m_weeks - 1)
        c_m = ws.cell(row=r_h1, column=curr_c, value=m_name)
        c_m.font = FONT_TBL_HEADER
        c_m.fill = FILL_ROYAL
        c_m.alignment = ALIGN_CENTER
        for c_sub in range(curr_c, curr_c + m_weeks):
            ws.cell(row=r_h1, column=c_sub).border = CELL_BORDER
            ws.cell(row=r_h1, column=c_sub).fill = FILL_ROYAL
            c_w = ws.cell(row=r_h2, column=c_sub, value=week_global_idx)
            c_w.font = FONT_TBL_SUBHEADER
            c_w.fill = FILL_LIGHT_BLUE
            c_w.alignment = ALIGN_CENTER
            c_w.border = CELL_BORDER
            week_global_idx += 1
        curr_c += m_weeks

    ket_col = curr_c
    ws.merge_cells(start_row=r_h1, start_column=ket_col, end_row=r_h2, end_column=ket_col)
    c_ket = ws.cell(row=r_h1, column=ket_col, value="Keterangan")
    c_ket.font = FONT_TBL_HEADER
    c_ket.fill = FILL_NAVY
    c_ket.alignment = ALIGN_CENTER
    for r_sub in range(r_h1, r_h2 + 1):
        ws.cell(row=r_sub, column=ket_col).border = CELL_BORDER
        ws.cell(row=r_sub, column=ket_col).fill = FILL_NAVY

    curr_r = r_h2 + 1
    for row_data in s_data:
        no, kode, tp, materi, jp, week_dist, ket = row_data
        ws.row_dimensions[curr_r].height = 55

        ws.cell(row=curr_r, column=1, value=no).font = FONT_DATA_NUMBER
        ws.cell(row=curr_r, column=1).alignment = ALIGN_CENTER
        ws.cell(row=curr_r, column=1).border = CELL_BORDER

        ws.cell(row=curr_r, column=2, value=kode).font = FONT_DATA_BOLD
        ws.cell(row=curr_r, column=2).alignment = ALIGN_CENTER
        ws.cell(row=curr_r, column=2).border = CELL_BORDER

        ws.cell(row=curr_r, column=3, value=tp).font = FONT_DATA
        ws.cell(row=curr_r, column=3).alignment = ALIGN_TOP_LEFT
        ws.cell(row=curr_r, column=3).border = CELL_BORDER

        ws.cell(row=curr_r, column=4, value=materi).font = FONT_DATA
        ws.cell(row=curr_r, column=4).alignment = ALIGN_TOP_LEFT
        ws.cell(row=curr_r, column=4).border = CELL_BORDER

        ws.cell(row=curr_r, column=5, value=jp).font = FONT_DATA_NUMBER
        ws.cell(row=curr_r, column=5).alignment = ALIGN_CENTER
        ws.cell(row=curr_r, column=5).border = CELL_BORDER

        for w_idx in range(total_weeks):
            col_t = 6 + w_idx
            val = week_dist[w_idx] if w_idx < len(week_dist) else ""
            cell_w = ws.cell(row=curr_r, column=col_t, value=val)
            cell_w.border = CELL_BORDER
            cell_w.alignment = ALIGN_CENTER
            if val == "":
                pass
            elif str(val).isdigit():
                cell_w.font = FONT_DATA_BOLD
                cell_w.fill = FILL_ACCENT_GREEN
            elif val in ["STS", "SAS", "SAT", "UP", "AJ", "US"]:
                cell_w.font = FONT_DATA_BOLD
                cell_w.fill = FILL_STS_SAS
            elif val in ["LS", "LU", "MPLS", "RPR"]:
                cell_w.font = FONT_ITALIC_DESC
                cell_w.fill = FILL_LIBUR

        c_k = ws.cell(row=curr_r, column=ket_col, value=ket)
        c_k.font = FONT_ITALIC_DESC
        c_k.alignment = ALIGN_TOP_LEFT
        c_k.border = CELL_BORDER

        curr_r += 1

    # Total Row
    ws.row_dimensions[curr_r].height = 24
    ws.merge_cells(start_row=curr_r, start_column=1, end_row=curr_r, end_column=4)
    c_tot = ws.cell(row=curr_r, column=1, value=f"TOTAL JAM PELAJARAN {sem_title}")
    c_tot.font = FONT_TOTAL
    c_tot.alignment = ALIGN_CENTER
    c_tot.fill = FILL_TOTAL

    c_tot_val = ws.cell(row=curr_r, column=5, value=f"=SUM(E{r_h2+1}:E{curr_r-1})")
    c_tot_val.font = FONT_TOTAL
    c_tot_val.alignment = ALIGN_CENTER
    c_tot_val.fill = FILL_TOTAL

    for w_idx in range(total_weeks):
        col_t = 6 + w_idx
        col_letter = get_column_letter(col_t)
        c_w_tot = ws.cell(row=curr_r, column=col_t, value=f"=SUM({col_letter}{r_h2+1}:{col_letter}{curr_r-1})")
        c_w_tot.font = FONT_DATA_BOLD
        c_w_tot.alignment = ALIGN_CENTER
        c_w_tot.fill = FILL_TOTAL

    c_tot_ket = ws.cell(row=curr_r, column=ket_col, value="Target Semester")
    c_tot_ket.font = FONT_TOTAL
    c_tot_ket.alignment = ALIGN_CENTER
    c_tot_ket.fill = FILL_TOTAL

    for c in range(1, max_col + 1):
        ws.cell(row=curr_r, column=c).border = TOTAL_BORDER

    add_signatures(ws, start_row=curr_r + 2, max_col=max_col)

    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 11
    ws.column_dimensions['C'].width = 36
    ws.column_dimensions['D'].width = 40
    ws.column_dimensions['E'].width = 10
    for w_idx in range(total_weeks):
        ws.column_dimensions[get_column_letter(6 + w_idx)].width = 4.2
    ws.column_dimensions[get_column_letter(ket_col)].width = 25
    ws.views.sheetView[0].showGridLines = True


# ==============================================================================
# MAIN SINGLE CONSOLIDATED WORKBOOK CREATION
# ==============================================================================

def create_single_unified_prota_prosem_workbook(output_file):
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove default sheet

    # 1. SHEET 1: ROADMAP STRUKTUR KURIKULUM & WAKTU
    ws1 = wb.create_sheet(title="Roadmap_Struktur_Kurikulum")
    apply_title(ws1, "MASTER ROADMAP PROGRAM TAHUNAN & PROGRAM SEMESTER BAHASA INGGRIS SMA",
                "Kurikulum Merdeka - Permendikdasmen No. 13 Tahun 2025 & Panduan Mapel Bahasa Inggris Puskurjar 2025", 7)

    ws1.cell(row=4, column=1, value="STRUKTUR ALOKASI WAKTU TAHUNAN & SEMESTERAN (WAJIB & TINGKAT LANJUT)").font = FONT_SEC_HEADER
    ws1.merge_cells("A4:G4")
    for c in range(1, 8):
        ws1.cell(row=4, column=c).fill = FILL_NAVY
    ws1.row_dimensions[4].height = 24

    headers_m = ["Tingkat / Kelas", "Fase", "Kategori Mata Pelajaran", "Beban Belajar", "Total JP / Thn", "Cakupan Semester", "Fokus Capaian Materi Esensial"]
    for i, h in enumerate(headers_m, 1):
        cell = ws1.cell(row=5, column=i, value=h)
        cell.font = FONT_TBL_HEADER
        cell.fill = FILL_ROYAL
        cell.alignment = ALIGN_CENTER
        cell.border = CELL_BORDER
    ws1.row_dimensions[5].height = 24

    roadmap_data = [
        ("Kelas X", "Fase E", "Bahasa Inggris Umum (Wajib)", "2 JP / Minggu", "72 JP", "Semester 1 (36 JP) & Sem 2 (36 JP)", "Teks Deskriptif, Recount, Naratif, Prosedur, Report, Eksposisi Analitis Dasar (CEFR B1)"),
        ("Kelas XI", "Fase F Umum", "Bahasa Inggris Umum (Wajib)", "2 JP / Minggu", "72 JP", "Semester 3 (36 JP) & Sem 4 (36 JP)", "Analytical Exposition, Hortatory Exposition, Contemporary Narrative, Explanation, Discussion, Review (CEFR B1-B2)"),
        ("Kelas XI", "Fase F TL", "Bahasa Inggris Tingkat Lanjut (Pilihan)", "5 JP / Minggu", "180 JP", "Semester 3 (90 JP) & Sem 4 (90 JP)", "Advanced Analytical Exposition, Hortatory, Literary Narrative, Poetry, Discussion Text (Puskurjar Model), Complex Explanation, Biography, Review (CEFR B2)"),
        ("Kelas XII", "Fase F Umum", "Bahasa Inggris Umum (Wajib)", "2 JP / Minggu", "64 JP", "Semester 5 (36 JP) & Sem 6 (28 JP)", "Discussion Text, News Item Text, Creative Product Review, Job Application & CV, Job Interview Skills, Caption & Infographic (CEFR B1-B2)"),
        ("Kelas XII", "Fase F TL", "Bahasa Inggris Tingkat Lanjut (Pilihan)", "5 JP / Minggu", "160 JP", "Semester 5 (90 JP) & Sem 6 (70 JP)", "Advanced Bioethics Discussion, Academic Synthesis Essay (APA/MLA), Drama Script, Persuasive Oratory, Career & Admissions Portfolio, Interview & Defense Simulation, Scientific Abstract (CEFR B2)")
    ]

    for r_i, r_val in enumerate(roadmap_data, 6):
        ws1.row_dimensions[r_i].height = 36
        for c_i, v in enumerate(r_val, 1):
            cell = ws1.cell(row=r_i, column=c_i, value=v)
            cell.font = FONT_DATA_BOLD if c_i in [1, 4, 5] else FONT_DATA
            cell.alignment = ALIGN_CENTER if c_i in [1, 2, 4, 5, 6] else ALIGN_LEFT
            cell.border = CELL_BORDER
            if r_i % 2 == 1:
                cell.fill = FILL_ZEBRA

    auto_fit_cols(ws1, min_w=12, max_w=45)

    # 2. PROTA SHEETS (5 SHEETS)
    add_prota_sheet(wb, "PROTA_Kelas_X_Wajib", "Kelas X", "Fase E", "Bahasa Inggris Umum (Wajib)", 72, DATA_PROTA_X_WAJIB)
    add_prota_sheet(wb, "PROTA_Kelas_XI_Wajib", "Kelas XI", "Fase F Umum", "Bahasa Inggris Umum (Wajib)", 72, DATA_PROTA_XI_WAJIB)
    add_prota_sheet(wb, "PROTA_Kelas_XI_TL", "Kelas XI", "Fase F Tingkat Lanjut", "Bahasa Inggris Tingkat Lanjut (Pilihan)", 180, DATA_PROTA_XI_TL)
    add_prota_sheet(wb, "PROTA_Kelas_XII_Wajib", "Kelas XII", "Fase F Umum", "Bahasa Inggris Umum (Wajib)", 64, DATA_PROTA_XII_WAJIB)
    add_prota_sheet(wb, "PROTA_Kelas_XII_TL", "Kelas XII", "Fase F Tingkat Lanjut", "Bahasa Inggris Tingkat Lanjut (Pilihan)", 160, DATA_PROTA_XII_TL)

    # 3. PROSEM SHEETS (10 SHEETS)
    add_prosem_sheet(wb, "PROSEM_Sem_1_Kelas_X_Wajib", 1, "Kelas X", "Fase E", "Bahasa Inggris Umum (Wajib)", 2, DATA_PROSEM_1_WAJIB)
    add_prosem_sheet(wb, "PROSEM_Sem_2_Kelas_X_Wajib", 2, "Kelas X", "Fase E", "Bahasa Inggris Umum (Wajib)", 2, DATA_PROSEM_2_WAJIB)
    add_prosem_sheet(wb, "PROSEM_Sem_3_Kelas_XI_Wajib", 3, "Kelas XI", "Fase F Umum", "Bahasa Inggris Umum (Wajib)", 2, DATA_PROSEM_3_WAJIB)
    add_prosem_sheet(wb, "PROSEM_Sem_3_Kelas_XI_TL", 3, "Kelas XI", "Fase F Tingkat Lanjut", "Bahasa Inggris Tingkat Lanjut (Pilihan)", 5, DATA_PROSEM_3_TL)
    add_prosem_sheet(wb, "PROSEM_Sem_4_Kelas_XI_Wajib", 4, "Kelas XI", "Fase F Umum", "Bahasa Inggris Umum (Wajib)", 2, DATA_PROSEM_4_WAJIB)
    add_prosem_sheet(wb, "PROSEM_Sem_4_Kelas_XI_TL", 4, "Kelas XI", "Fase F Tingkat Lanjut", "Bahasa Inggris Tingkat Lanjut (Pilihan)", 5, DATA_PROSEM_4_TL)
    add_prosem_sheet(wb, "PROSEM_Sem_5_Kelas_XII_Wajib", 5, "Kelas XII", "Fase F Umum", "Bahasa Inggris Umum (Wajib)", 2, DATA_PROSEM_5_WAJIB)
    add_prosem_sheet(wb, "PROSEM_Sem_5_Kelas_XII_TL", 5, "Kelas XII", "Fase F Tingkat Lanjut", "Bahasa Inggris Tingkat Lanjut (Pilihan)", 5, DATA_PROSEM_5_TL)
    add_prosem_sheet(wb, "PROSEM_Sem_6_Kelas_XII_Wajib", 6, "Kelas XII", "Fase F Umum", "Bahasa Inggris Umum (Wajib)", 2, DATA_PROSEM_6_WAJIB)
    add_prosem_sheet(wb, "PROSEM_Sem_6_Kelas_XII_TL", 6, "Kelas XII", "Fase F Tingkat Lanjut", "Bahasa Inggris Tingkat Lanjut (Pilihan)", 5, DATA_PROSEM_6_TL)

    wb.save(output_file)
    print(f"SINGLE Unified PROTA & PROSEM Workbook successfully created at: {output_file}")


def cleanup_loose_files(target_dir, keep_files):
    for f in os.listdir(target_dir):
        if f.endswith(".xlsx") and f not in keep_files:
            p = os.path.join(target_dir, f)
            try:
                os.remove(p)
                print(f"Removed redundant file: {f}")
            except Exception as e:
                print(f"Error removing {f}: {e}")


def main():
    target_dir = r"C:\Users\KOMPUTERKU2\Documents\English SMA\3-Program Tahunan dan Program Semester"
    os.makedirs(target_dir, exist_ok=True)

    single_file_name = "Program_Tahunan_dan_Program_Semester_Bahasa_Inggris_SMA.xlsx"
    single_file_path = os.path.join(target_dir, single_file_name)

    print("Building single consolidated PROTA and PROSEM workbook...")
    create_single_unified_prota_prosem_workbook(single_file_path)

    print("\nCleaning up loose separate Excel files...")
    cleanup_loose_files(target_dir, keep_files=[single_file_name])

    print("\nTask completed successfully!")

if __name__ == "__main__":
    main()
