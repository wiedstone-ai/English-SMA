# -*- coding: utf-8 -*-
"""
Script to generate official SILABUS PEMBELAJARAN BAHASA INGGRIS SMA (WAJIB & TINGKAT LANJUT)
in a single, beautifully structured Excel workbook.
Emphasizes Constructive Alignment:
- Indikator Ketercapaian Tujuan Pembelajaran (IKTP / KKTP) aligned with specific assessment techniques & instruments.
- Covers all modules across Semester 1 to Semester 6 for both Wajib and Tingkat Lanjut.
Conforms to Permendikdasmen No. 13 Tahun 2025, CP 2025/2026, and Panduan Mapel Bahasa Inggris Puskurjar 2025.
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
# DATA DEFINITIONS FOR SILABUS (POKOK-POKOK MODUL AJAR & KESELARASAN ASESMEN)
# ==============================================================================

# KELAS X WAJIB (FASE E - SEMESTER 1 & 2)
DATA_SILABUS_X_WAJIB = [
    # (Modul_No, Sem, Kode_TP, Elemen, TP, Materi_Pokok_Genre, IKTP_Indikator, Asesmen_Teknik_Instrumen, Bukti_Produk, Langkah_GBA, JP_Waktu, Media_Sumber)
    ("Modul 1", "Semester 1", "TP 10.1", "Menyimak-Berbicara & Membaca-Memirsa",
     "Menganalisis fungsi sosial, struktur teks, dan unsur kebahasaan teks deskriptif lisan dan tulis tentang atlet berprestasi dan destinasi wisata dunia secara kritis dan mandiri.",
     "Descriptive Text:\n• Identification & Description\n• Adjectives of physical/personality traits\n• Simple Present Tense & sensory details",
     "IKTP 1.1: Mengidentifikasi 5 ide pokok dan 8 informasi rinci tersurat/tersirat dari monolog deskripsi atlet berprestasi.\nIKTP 1.2: Menganalisis fungsi sosial dan kosakata sifat pada teks deskripsi tempat wisata dunia.\nIKTP 1.3: Membedakan penggunaan to be dan verb-1 dalam mendeskripsikan fakta aktual.",
     "• Asesmen Diagnostik: Kuis kosakata visual atlet/tempat wisata.\n• Asesmen Formatif: Lembar Kerja Menyimak (Audio Listening Comprehension) & Graphic Organizer struktur identifikasi-deskripsi.\n• Asesmen Sumatif: Tes Tulis Analisis Teks Deskriptif (Pilihan ganda HOTS & uraian singkat).",
     "Lembar kerja analisis teks & lembar jawaban tes menyimak audio terstruktur.",
     "1. BKoF: Menonton cuplikan video biografi atlet olimpiade Indonesia.\n2. MoT: Membaca teks model deskriptif & membedah unsur kebahasaan.\n3. JCoT: Berpasangan melengkapi profil atlet dalam bagan klasifikasi.\n4. ICoT: Menjawab pertanyaan inferensial mandiri & refleksi belajar.",
     "10 JP (5 Pertemuan x 2 JP)",
     "Video YouTube Atlet Hebat, Modul Puskurjar 2025, Foto Candi Borobudur/Raja Ampat, Lembar Kerja Murid"),

    ("Modul 2", "Semester 1", "TP 10.2", "Membaca-Memirsa & Menulis-Mempresentasikan",
     "Memproduksi teks deskriptif multimodal berbentuk brosur digital atau poster promosi cagar budaya dan destinasi wisata nusantara dengan struktur teks yang kohesif.",
     "Descriptive Text (Multimodal Production):\n• Layouting brosur promosi wisata\n• Sensory adjectives & persuasive appeal\n• Spatial ordering & descriptive noun phrases",
     "IKTP 2.1: Merancang kerangka deskripsi objek wisata nusantara dengan struktur identifikasi dan deskripsi terperinci.\nIKTP 2.2: Memadukan gambar otentik dengan takarir teks deskriptif yang koheren.\nIKTP 2.3: Mempresentasikan brosur wisata secara lisan dengan artikulasi dan kontak mata yang percaya diri.",
     "• Asesmen Diagnostik: Pemetaan destinasi wisata lokal yang pernah dikunjungi murid.\n• Asesmen Formatif: Peer-review draf brosur menggunakan daftar ceklis bahasa deskripsi.\n• Asesmen Sumatif: Penilaian Kinerja Produk (Brosur Digital Wisata) & Unjuk Kerja Presentasi Lisan menggunakan Rubrik Analitik Menulis dan Berbicara Puskurjar (Skala 1-4).",
     "Brosur promosi wisata digital (Canva/Poster) & rekaman video presentasi lisan singkat.",
     "1. BKoF: Eksplorasi 5 destinasi super-prioritas Indonesia.\n2. MoT: Dekonstruksi brosur wisata profesional (headline, body text, visual).\n3. JCoT: Kolaborasi kelompok merancang draf brosur tempat wisata lokal.\n4. ICoT: Menyelesaikan brosur mandiri dan mempresentasikannya di depan kelas.",
     "10 JP (5 Pertemuan x 2 JP)",
     "Platform Canva/Google Docs, Brosur Kemenparekraf Wonderful Indonesia, Rubrik Penilaian Multimodal"),

    ("Modul 3", "Semester 1", "TP 10.3", "Menyimak-Berbicara & Menulis-Mempresentasikan",
     "Mengidentifikasi alur kronologis peristiwa sejarah dan pengalaman pribadi masa lalu (Recount Text), serta menyusun teks recount personal dan faktual dengan runtut dan akurat.",
     "Recount Text (Historical & Personal):\n• Orientation, Series of Events, Re-orientation\n• Simple Past Tense & Past Continuous\n• Temporal linkers (first, then, after that, finally)",
     "IKTP 3.1: Menyusun kembali urutan acak peristiwa sejarah nasional ke dalam linimasa kronologis yang tepat.\nIKTP 3.2: Mengidentifikasi kata kerja aksi bentuk lampau (regular/irregular verbs) dalam teks recount.\nIKTP 3.3: Menulis teks recount pribadi/sejarah sepanjang 150-200 kata dengan tata bahasa lampau yang akurat.",
     "• Asesmen Diagnostik: Kuis pemilahan kata kerja Present vs Past (V1 vs V2).\n• Asesmen Formatif: LKPD penyusunan Timeline Bagan Sejarah & latihan grammar in context.\n• Asesmen Sumatif: Portofolio Tulisan Recount Text (Personal Journal/Historical Event) dinilai dengan Rubrik Analitik Menulis (Organisasi, Tata Bahasa, Diksi, Ejaan).",
     "Kompilasi lembar linimasa sejarah & draf esai teks recount pribadi (Learning Journal).",
     "1. BKoF: Diskusi foto peristiwa bersejarah Hari Pahlawan / Kemerdekaan.\n2. MoT: Menganalisis teks recount 'The Battle of Surabaya' (orientasi, rentetan peristiwa, reorientasi).\n3. JCoT: Menuliskan cerita pengalaman kunjungan edukasi bersama kelompok.\n4. ICoT: Menulis jurnal memoar pribadi secara mandiri & penilaian teman sejawat.",
     "10 JP (5 Pertemuan x 2 JP)",
     "Teks Sejarah Perjuangan Nasional, Flashcard Irregular Verbs, Audio Podcast Sejarah"),

    ("Modul 4", "Semester 2", "TP 10.4", "Menyimak-Berbicara & Membaca-Memirsa",
     "Menganalisis nilai-nilai moral, alur naratif, dan konflik dalam cerita rakyat Indonesia dan legenda dunia (Narrative Text), serta mendemonstrasikan pembacaan cerita naratif (storytelling).",
     "Narrative Text (Folklore & Legends):\n• Orientation, Complication, Resolution, Coda\n• Direct & Indirect Speech, Action verbs\n• Expressive storytelling intonation & moral values",
     "IKTP 4.1: Menemukan klimaks komplikasi dan solusi konflik dalam teks legenda nusantara/dunia.\nIKTP 4.2: Mengubah kalimat ujaran langsung (direct speech) menjadi kalimat tidak langsung (indirect speech).\nIKTP 4.3: Menampilkan pembacaan cerita rakyat secara ekspresif dengan intonasi, jeda, dan artikulasi vokal yang tepat.",
     "• Asesmen Diagnostik: Tanya jawab cerita rakyat favorit masa kecil.\n• Asesmen Formatif: Lembar telaah unsur intrinsik (tokoh, latar, plot, amanat) & latihan direct/indirect speech.\n• Asesmen Sumatif: Unjuk Kerja Praktik Storytelling lisan (live performance / video rekaman) dinilai dengan Rubrik Storytelling Baku Puskurjar (Fluency, Pronunciation, Expression).",
     "Rekaman video unjuk kerja storytelling & lembar kerja bedah karakter tokoh cerita.",
     "1. BKoF: Menyimak penceritaan dongeng Malin Kundang / Roro Jonggrang.\n2. MoT: Dekonstruksi struktur komplikasi dan resolusi serta pola dialog langsung.\n3. JCoT: Latihan membaca berpasangan dengan saling memberikan koreksi intonasi.\n4. ICoT: Penampilan storytelling mandiri & refleksi nilai kebajikan moral.",
     "10 JP (5 Pertemuan x 2 JP)",
     "Buku Kumpulan Cerita Rakyat Nusantara, Rekaman Audio Storytelling Profesional, Kostum/Properti Sederhana"),

    ("Modul 5", "Semester 2", "TP 10.5", "Membaca-Memirsa & Menulis-Mempresentasikan",
     "Memahami dan menyusun teks prosedur lisan dan tulis terkait pola hidup sehat (healthy living), resep kuliner nusantara, serta panduan keselamatan berinternet (cyber safety) secara sistematis.",
     "Procedure Text:\n• Goal/Aim, Materials/Ingredients, Steps/Methods\n• Imperative sentences, action verbs, sequencers\n• Adverbials of time, manner, and sequence",
     "IKTP 5.1: Mengidentifikasi tujuan, peralatan/bahan yang diperlukan, dan urutan instruksi dari teks prosedur hidup sehat.\nIKTP 5.2: Menggunakan kalimat imperatif dan konjungsi sekuensial (first, second, next, finally) secara tepat.\nIKTP 5.3: Merancang dan memproduksi teks prosedur infografis atau video demo tutorial panduan keselamatan digital.",
     "• Asesmen Diagnostik: Brainstorming tips menjaga kesehatan fisik & tips aman bermedia sosial.\n• Asesmen Formatif: Lembar kerja pengurutan langkah acak (jumbled steps) & latihan kalimat perintah imperatif.\n• Asesmen Sumatif: Penilaian Produk Tutorial (Poster Infografis Prosedur / Video Demo 2 Menit) dinilai dengan Rubrik Produk Teks Prosedur.",
     "Infografis panduan digital safety / video panduan resep sehat berdurasi 2 menit.",
     "1. BKoF: Diskusi bahaya hoaks di medsos & pentingnya gaya hidup sehat remaja.\n2. MoT: Bedah struktur teks manual 'How to Create a Strong Password' & resep jus sehat.\n3. JCoT: Menyusun draft prosedur tips kesehatan kelompok.\n4. ICoT: Membuat infografis/video panduan tutorial mandiri & asesmen antar-teman.",
     "10 JP (5 Pertemuan x 2 JP)",
     "Panduan Literasi Digital Kominfo, Resep Kuliner Nusantara Sehat, Templat Desain Canva"),

    ("Modul 6", "Semester 2", "TP 10.6", "Menyimak-Berbicara, Membaca-Memirsa & Menulis",
     "Menganalisis laporan observasi ilmiah (Report Text) mengenai fauna endemik dan inovasi sains, serta mengevaluasi argumen awal teks eksposisi analitis bertema kelestarian lingkungan hidup.",
     "Report & Intro to Analytical Exposition:\n• General Classification & Description of parts/habits\n• Timeless present tense, general nouns, passive voice\n• Thesis statement & causal connectors on eco-issues",
     "IKTP 6.1: Mengklasifikasikan deskripsi bagian, habitat, dan perilaku satwa terancam punah Indonesia dari teks report.\nIKTP 6.2: Mengidentifikasi kalimat pasif (Passive Voice) dan kosakata teknis sains dalam teks faktual.\nIKTP 6.3: Menyusun laporan observasi ilmiah 3 paragraf dan merumuskan tesis pendukung pelestarian satwa/lingkungan.",
     "• Asesmen Diagnostik: Kuis tebak fauna endemik Indonesia (Komodo, Orangutan, Cendrawasih).\n• Asesmen Formatif: Lembar telaah Passive Voice dalam konteks laporan sains & Graphic Organizer klasifikasi.\n• Asesmen Sumatif: Ujian Akhir Semester / Portofolio Laporan Ilmiah Mini dinilai dengan Rubrik Laporan Sains Faktual.",
     "Makalah laporan observasi fauna ilmiah mini & poster kampanye perlindungan satwa.",
     "1. BKoF: Menonton dokumenter konservasi satwa langka Indonesia.\n2. MoT: Dekonstruksi teks report ilmiah 'The Sumatran Tiger' & pola Passive Voice.\n3. JCoT: Menyusun draf klasifikasi satwa endemik bersama kelompok.\n4. ICoT: Menulis laporan ilmiah mandiri dan memamerkannya dalam pameran sains kelas.",
     "10 JP (5 Pertemuan x 2 JP)",
     "Dokumenter Fauna Indonesia, Ensiklopedia Sains Remaja, Modul Puskurjar Bab C")
]

# KELAS XI WAJIB (FASE F UMUM - SEMESTER 3 & 4)
DATA_SILABUS_XI_WAJIB = [
    ("Modul 1", "Semester 3", "TP 11.W.1", "Membaca-Memirsa & Menulis-Mempresentasikan",
     "Menganalisis fungsi sosial, struktur teks, dan unsur kebahasaan teks eksposisi analitis tentang isu sosial dan literasi digital di era kecerdasan buatan, serta menyusun esai argumentatif logis.",
     "Analytical Exposition Text:\n• Thesis statement, Series of Arguments, Reiteration\n• Mental verbs (feel, believe, realize), Causal connectives\n• Distinguishing facts from subjective opinions",
     "IKTP 1.1: Menemukan kalimat tesis (thesis statement) dan 3 argumen pendukung teks eksposisi isu digital.\nIKTP 1.2: Membedakan data faktual berbasis bukti dari opini subjektif tanpa dasar.\nIKTP 1.3: Menyusun esai eksposisi analitis tertulis 4-5 paragraf yang kohesif menggunakan konjungsi kausal (because, consequently, therefore).",
     "• Asesmen Diagnostik: Survei pandangan siswa mengenai dampak medsos pada fokus belajar.\n• Asesmen Formatif: Pemetaan Argumen (Argument Mapping Worksheet) & latihan mental verbs.\n• Asesmen Sumatif: Penilaian Esai Argumentatif Analitis dinilai dengan Rubrik Analitik Menulis Esai (Thesis Clarity, Argument Validity, Cohesion, Grammar).",
     "Naskah esai eksposisi analitis 4 paragraf bertema etika digital / literasi media.",
     "1. BKoF: Diskusi fenomena kecanduan gawai dan kebocoran data pribadi.\n2. MoT: Dekonstruksi teks 'Why Digital Literacy Matters' (tesis, argumen 1-3, reiterasi).\n3. JCoT: Menyusun draf kerangka esai berpasangan dengan saling memverifikasi fakta.\n4. ICoT: Menulis esai utuh secara independen & evaluasi diri (Self-assessment check).",
     "10 JP (5 Pertemuan x 2 JP)",
     "Artikel Opini The Jakarta Post, Modul Ajar SMA Puskurjar, Slide Presentasi Teks Eksposisi"),

    ("Modul 2", "Semester 3", "TP 11.W.2", "Menyimak-Berbicara & Menulis-Mempresentasikan",
     "Mengidentifikasi argumen dan rekomendasi persuasif dalam teks eksposisi hortatori terkait kesehatan mental remaja dan keselamatan publik, serta mempresentasikan pidato ajakan bertindak.",
     "Hortatory Exposition Text:\n• Thesis, Arguments, Urgent Recommendation\n• Modals of advice & obligation (should, ought to, must)\n• Emotive language, rhetorical questions, call to action",
     "IKTP 2.1: Mengidentifikasi bagian rekomendasi (recommendation) dan membedakannya dari reiterasi eksposisi analitis.\nIKTP 2.2: Menggunakan kata modalitas (should, ought to) dan bahasa emotif ajakan persuasif.\nIKTP 2.3: Membawakan pidato persuasif kampanye kesehatan mental berdurasi 2-3 menit secara meyakinkan.",
     "• Asesmen Diagnostik: Kuis interaktif membedakan teks analitis vs hortatori.\n• Asesmen Formatif: Lembar analisis rekomendasi kebijakan & latihan menulis kalimat saran kuat.\n• Asesmen Sumatif: Unjuk Kerja Pidato Persuasif (Speech Presentation) dinilai dengan Rubrik Berbicara Pidato Persuasif (Konten Tesis/Saran, Pelafalan, Kontak Mata, Daya Bujuk).",
     "Teks pidato persuasif kampanye remaja & rekaman audio visual penampilan pidato.",
     "1. BKoF: Curah pendapat isu stres akademik dan pentingnya kesehatan mental siswa.\n2. MoT: Membedah teks 'Schools Should Provide Mental Health Days' & pola rekomendasi.\n3. JCoT: Menulis draf teks kampanye kelompok dengan kalimat ajakan yang kuat.\n4. ICoT: Menampilkan pidato persuasif di depan kelas & penilaian kawan sebaya (Peer review).",
     "10 JP (5 Pertemuan x 2 JP)",
     "Teks Pidato Kampanye UNESCO, Panduan Kesehatan Jiwa Remaja Kemenkes, Rubrik Pidato"),

    ("Modul 3", "Semester 3", "TP 11.W.3", "Menyimak-Berbicara & Membaca-Memirsa",
     "Menganalisis struktur narasi, dinamika konflik karakter, dan pesan moral dalam teks naratif kontemporer dan fabel modern (Fractured Fairy Tales), serta menyusun adaptasi narasi kreatif.",
     "Contemporary Narrative & Fractured Tales:\n• Orientation, Complication, Evaluation, Resolution\n• Dialogue formatting, descriptive adjectives, shift of perspective\n• Moral nuance & alternative storytelling viewpoints",
     "IKTP 3.1: Mengidentifikasi pergeseran sudut pandang (point of view) dan ironi dalam dongeng modern yang dimodifikasi.\nIKTP 3.2: Menganalisis motivasi psikologis karakter antagonis yang memiliki sisi humanis.\nIKTP 3.3: Menulis ulang penggalan cerita rakyat dengan sudut pandang tokoh yang berbeda (misal: sudut pandang serigala pada Si Tudung Merah).",
     "• Asesmen Diagnostik: Tanya jawab tentang dongeng klasik yang alur akhirnya ingin diubah.\n• Asesmen Formatif: Lembar analisis karakterisasi (Character trait chart) & bedah dialog langsung.\n• Asesmen Sumatif: Portofolio Cerita Narasi Adaptasi Kreatif dinilai dengan Rubrik Menulis Naratif (Kreativitas Alur, Kompleksitas Karakter, Struktur Bahasa, Ejaan).",
     "Kumpulan naskah cerpen adaptasi dongeng modern siswa (Antologi Kelas).",
     "1. BKoF: Menonton cuplikan animasi fractured fairy tale (misal: kisah Shrek/Tangled).\n2. MoT: Membaca teks naratif 'The True Story of the Three Little Pigs' & analisis perspektif.\n3. JCoT: Kelompok merancang ulang komplikasi dan resolusi alternatif cerita Timun Mas.\n4. ICoT: Menulis cerpen naratif adaptasi mandiri & publikasi karya di mading sekolah.",
     "10 JP (5 Pertemuan x 2 JP)",
     "Buku Kumpulan Cerpen Sastra Remaja, Buku Fractured Fairy Tales, Lembar Kerja Murid"),

    ("Modul 4", "Semester 4", "TP 11.W.4", "Membaca-Memirsa & Menulis-Mempresentasikan",
     "Menganalisis proses kausal dan kronologis terjadinya fenomena alam dan sosial budaya (Explanation Text), serta mempresentasikan diagram alir penjelasan ilmiah populer.",
     "Explanation Text (Sequential & Causal):\n• General Statement, Sequenced Explanation, Concluding Remark\n• Cause-and-effect linkers (due to, as a consequence, leads to)\n• Passive voice, technical nouns, chronological flowcharts",
     "IKTP 4.1: Mengidentifikasi tahapan siklus proses sebab-akibat (how & why) terjadinya fenomena alam/sosial.\nIKTP 4.2: Menggunakan kata penghubung kausalitas dan kalimat pasif formal secara akurat.\nIKTP 4.3: Menyusun teks eksplanasi ilmiah populer disertai diagram alir proses (flowchart) informatif.",
     "• Asesmen Diagnostik: Kuis tebak proses ilmiah (terjadinya pelangi, gempa bumi, inflasi).\n• Asesmen Formatif: Lembar Kerja Diagram Alir Proses Ilmiah & latihan transformasi kalimat aktif ke pasif.\n• Asesmen Sumatif: Penilaian Proyek Artikel Eksplanasi Multimodal (Teks + Diagram Skematis) dinilai dengan Rubrik Teks Eksplanasi Baku Puskurjar.",
     "Artikel eksplanasi ilmiah berdiagram alur proses (Infografis Proses Sains).",
     "1. BKoF: Menonton video simulasi proses terjadinya efek rumah kaca & pemanasan global.\n2. MoT: Membedah teks 'How Acid Rain is Formed' (pernyataan umum, deret penjelas kausal, penutup).\n3. JCoT: Berpasangan melengkapi diagram alur terjadinya tsunami atau urbanisasi.\n4. ICoT: Menulis teks eksplanasi mandiri bertopik fenomena pilihan & penilaian portofolio.",
     "10 JP (5 Pertemuan x 2 JP)",
     "Video Edukasi Sains National Geographic, Artikel Eksplanasi Ilmiah Populer, Papan Diagram Alur"),

    ("Modul 5", "Semester 4", "TP 11.W.5", "Menyimak-Berbicara & Membaca-Memirsa",
     "Menimbang dua sudut pandang berimbang pro dan kontra terkait isu keseharian remaja dan lingkungan sekolah (Discussion Text), serta berpartisipasi aktif dalam diskusi kelompok terarah.",
     "Discussion Text (Everyday Teen Issues):\n• Issue, Arguments For, Arguments Against, Recommendation\n• Contrastive conjunctions (on the other hand, however, conversely)\n• Hedging expressions (it may be argued, some people suggest)",
     "IKTP 5.1: Mengidentifikasi argumen pro dan kontra yang berbobot mengenai penggunaan gawai di sekolah.\nIKTP 5.2: Menggunakan konjungsi pertentangan dan bahasa kesantunan dalam mengemukakan ketidaksepakatan.\nIKTP 5.3: Berpartisipasi aktif dalam diskusi meja bundar (round-table discussion) dengan menyampaikan opini berimbang.",
     "• Asesmen Diagnostik: Polling kelas pro vs kontra gawai di lingkungan belajar.\n• Asesmen Formatif: Bagan T-Chart Pro & Con Arguments & lembar observasi partisipasi diskusi.\n• Asesmen Sumatif: Unjuk Kerja Diskusi Meja Bundar & Penulisan Teks Diskusi Berimbang dinilai dengan Rubrik Diskusi & Argumentasi Objektif.",
     "Bagan T-Chart argumen pro-kontra & esai teks diskusi berimbang siswa sepanjang 200 kata.",
     "1. BKoF: Diskusi pro-kontra pemberlakuan jam malam dan penggunaan smartphone di sekolah.\n2. MoT: Bedah teks diskusi 'Pros and Cons of Homework for High Schoolers' & bahasa berimbang.\n3. JCoT: Simulasi diskusi kelompok kecil membagi peran tim pro dan tim kontra.\n4. ICoT: Menuliskan sintesis rekomendasi mandiri & evaluasi kesantunan berbicara.",
     "10 JP (5 Pertemuan x 2 JP)",
     "Artikel Diskusi Masalah Remaja, Modul Puskurjar Bab C, Lembar Observasi Diskusi Terarah"),

    ("Modul 6", "Semester 4", "TP 11.W.6", "Membaca-Memirsa & Menulis-Mempresentasikan",
     "Mengevaluasi dan menyusun teks resensi kritis (Review Text) terhadap film edukatif, novel remaja, atau karya seni budaya lokal nusantara dengan bahasa pertimbangan mutu yang objektif.",
     "Review Text (Arts & Literature):\n• Orientation, Interpretative Recount, Evaluation, Summary\n• Evaluative adjectives (compelling, thought-provoking, subpar)\n• Critical aesthetic judgment & constructive recommendation",
     "IKTP 6.1: Menemukan interpretative recount dan kriteria evaluasi karya pada resensi film/buku.\nIKTP 6.2: Menggunakan kosakata evaluasi mutu artistik secara tepat dan santun.\nIKTP 6.3: Menulis resensi kritis film/novel inspiratif 250 kata dan mempublikasikannya di blog/majalah sekolah.",
     "• Asesmen Diagnostik: Tanya jawab film atau buku terbaik yang pernah ditonton/dibaca murid.\n• Asesmen Formatif: Lembar Telaah Resensi Karya (Plot, Karakter, Sinematografi/Bahasa, Nilai).\n• Asesmen Sumatif: Penilaian Produk Resensi Kritis (Published Review Article) dinilai dengan Rubrik Menulis Resensi Seni Baku Puskurjar.",
     "Artikel resensi film/novel yang siap terbit di mading atau blog edukatif sekolah.",
     "1. BKoF: Menonton cuplikan film inspiratif Indonesia (Laskar Pelangi / Sokola Rimba).\n2. MoT: Membedah struktur teks resensi film internasional (orientasi, tafsiran, evaluasi, rangkuman).\n3. JCoT: Berpasangan menuliskan draf evaluasi kekuatan dan kelemahan film.\n4. ICoT: Menulis resensi lengkap secara mandiri dan menggelar pameran ulasan karya (Review Fair).",
     "10 JP (5 Pertemuan x 2 JP)",
     "Koleksi Resensi Film Rotten Tomatoes/Kompas, Cuplikan Film Nasional, Rubrik Penilaian Resensi")
]

# KELAS XI TINGKAT LANJUT (FASE F TL - SEMESTER 3 & 4 - 5 JP/MINGGU)
DATA_SILABUS_XI_TL = [
    ("Modul 1", "Semester 3", "TP 11.TL.1", "Membaca-Memirsa & Menulis-Mempresentasikan",
     "Mengevaluasi wacana argumentatif kritis tingkat lanjut dalam teks eksposisi analitis bertema disrupsi kecerdasan buatan (AI) dan transformasi masyarakat digital, serta menyusun makalah posisi akademik.",
     "Advanced Analytical Exposition:\n• Complex thesis formulation & multi-tiered arguments\n• Abstract nominalization (digitization, ethical implication)\n• Modal adverbs of certainty, conditional & concessive linkers",
     "IKTP 1.1: Menganalisis counter-arguments dan fallacy dalam artikel ilmiah populer tentang etika kecerdasan buatan.\nIKTP 1.2: Menggunakan nominalisasi abstrak tingkat lanjut untuk meningkatkan kepadatan leksikal akademik.\nIKTP 1.3: Menyusun academic position paper 500-600 kata dengan argumen terstruktur dan sitasi bukti empiris.",
     "• Asesmen Diagnostik: Kuis wawasan AI di era Generative Artificial Intelligence.\n• Asesmen Formatif: Matriks Dekonstruksi Argumen Kompleks & latihan pembentukan kata nominalisasi.\n• Asesmen Sumatif: Penilaian Makalah Posisi Akademik (Position Paper) & Podcast Analitis dinilai dengan Rubrik Analitik Menulis Tingkat Lanjut Puskurjar.",
     "Makalah posisi akademik 500 kata & rekaman audio podcast diskusi analitis 5 menit.",
     "1. BKoF: Menelaah infografis perkembangan AI generatif dan tantangan etika plagiarisme.\n2. MoT: Analisis teks akademik 'The Perils and Promises of Artificial Intelligence' (nominalisasi, modalitas).\n3. JCoT: Menyusun peta argumen dan counter-argument bersama kelompok riset mini.\n4. ICoT: Penulisan makalah posisi mandiri dan rekaman podcast analitis.",
     "20 JP (4 Minggu x 5 JP)",
     "Jurnal Ilmiah Populer MIT Tech Review, Modul Puskurjar 2025, Software Perekam Audio Podcast"),

    ("Modul 2", "Semester 3", "TP 11.TL.2", "Menyimak-Berbicara & Menulis-Mempresentasikan",
     "Menganalisis dan menyusun teks eksposisi hortatori tingkat lanjut yang mendesak tindakan solutif terkait transisi energi bersih dan keadilan iklim global, serta mempresentasikan policy pitch.",
     "Advanced Hortatory Exposition:\n• Multi-stakeholder thesis, evaluative evidence, urgent policy recommendations\n• Advanced rhetoric (Ethos, Pathos, Logos), strong modal auxiliaries\n• Policy pitch & persuasive diplomatic advocacy",
     "IKTP 2.1: Menganalisis strategi retorika persuasif (Ethos, Pathos, Logos) pada pidato pembuat kebijakan iklim dunia.\nIKTP 2.2: Mengintegrasikan data statistik dan rekomendasi regulasi konkret dalam teks eksposisi hortatori.\nIKTP 2.3: Menyampaikan presentasi policy pitch 3 menit di hadapan panel penilai simulasi KTT Iklim.",
     "• Asesmen Diagnostik: Tanya jawab tentang krisis iklim dan energi baru terbarukan di Indonesia.\n• Asesmen Formatif: Lembar Telaah Retorika Pidato Pemimpin Dunia & draf lembar advokasi kebijakan.\n• Asesmen Sumatif: Unjuk Kerja Presentasi Policy Pitch & Dokumen Policy Brief Siswa dinilai dengan Rubrik Retorika & Advokasi Kebijakan.",
     "Dokumen ringkasan kebijakan (Policy Brief) 2 halaman & rekaman presentasi policy pitch.",
     "1. BKoF: Mengkaji data emisi karbon global dan target Net Zero Emission 2060 Indonesia.\n2. MoT: Membedah pidato aktivis lingkungan di forum PBB (strategi Ethos, Pathos, Logos, rekomendasi).\n3. JCoT: Merumuskan draf naskah advokasi energi surya kelompok dengan rekomendasi terukur.\n4. ICoT: Membawakan simulasi policy pitch mandiri di depan panel penguji.",
     "25 JP (5 Minggu x 5 JP)",
     "Laporan IPCC PBB, Data Kementerian ESDM, Video Pidato Greta Thunberg / Sekjen PBB"),

    ("Modul 3", "Semester 3", "TP 11.TL.3", "Menyimak-Berbicara & Membaca-Memirsa",
     "Mengapresiasi, menganalisis kedalaman psikologis karakter, dan merekonstruksi narasi sastra fiksi/cerpen klasik dan kontemporer dengan alur bertingkat dan dilema moral.",
     "Advanced Literary Narrative & Short Stories:\n• Narrative arc, foreshadowing, symbolism, stream of consciousness\n• Literary devices: metaphor, irony, character introspection\n• Comparative character analysis & reader's theatre",
     "IKTP 3.1: Mengidentifikasi teknik narasi sastra canggih (foreshadowing, symbolism, internal monologue) dalam cerpen pilihan.\nIKTP 3.2: Menganalisis konflik psikologis dan pertimbangan etika tokoh utama dalam menghadapi dilema moral.\nIKTP 3.3: Menyusun esai apresiasi sastra (literary analysis essay) atau mementaskan pembacaan dramatik naskah cerpen.",
     "• Asesmen Diagnostik: Kuis pemahaman elemen intrinsik sastra tingkat lanjut.\n• Asesmen Formatif: Peta Karakter Psikologis (Character Psychology Chart) & jurnal apresiasi cerita.\n• Asesmen Sumatif: Portofolio Esai Kritik Sastra / Unjuk Kerja Reader's Theatre dinilai dengan Rubrik Analisis Sastra Baku Puskurjar.",
     "Esai analisis karakter cerpen 400 kata & penampilan pembacaan dramatik sastra (Reader's Theatre).",
     "1. BKoF: Diskusi tema dilema moral dalam karya sastra dunia (misal karya Edgar Allan Poe / O. Henry).\n2. MoT: Menganalisis cerpen 'The Gift of the Magi' / 'The Tell-Tale Heart' (simbolisme, ironi situasi).\n3. JCoT: Berpasangan membedah kutipan dialog bermakna subteks psikologis.\n4. ICoT: Menulis esai telaah sastra mandiri dan pementasan pembacaan naskah dramatik.",
     "25 JP (5 Minggu x 5 JP)",
     "Buku Antologi Cerpen Sastra Bahasa Inggris Klasik & Modern, Rubrik Reader's Theatre"),

    ("Modul 4", "Semester 3", "TP 11.TL.4", "Menyimak-Berbicara & Membaca-Memirsa",
     "Menganalisis gaya bahasa stilistika, makna simbolik, dan kritik sosial dalam puisi bahasa Inggris dan lirik lagu dunia berpengaruh, serta mendeklamasikannya.",
     "Stylistic & Societal Reading of Poetry & Protest Songs:\n• Meter, stanza structure, rhyme schemes, imagery, tone, subtext\n• Figuring social inequality, peace, and human dignity in verses\n• Expressive recitation & comparative poetic commentary",
     "IKTP 4.1: Menemukan majas metafora, personifikasi, dan citraan sensorik pada puisi/lirik lagu kritik sosial.\nIKTP 4.2: Menafsirkan tone pengarang dan pesan tersembunyi mengenai perdamaian dan martabat manusia.\nIKTP 4.3: Mendeklamasikan puisi bahasa Inggris pilihan dengan intonasi ekspresif dan penghayatan makna yang mendalam.",
     "• Asesmen Diagnostik: Mendengarkan lagu bertema kemanusiaan dan mendiskusikan kesan emosional.\n• Asesmen Formatif: Lembar analisis bait puisi (Stanza annotation worksheet) & bedah majas.\n• Asesmen Sumatif: Unjuk Kerja Deklamasi Puisi (Poetry Recitation) & Ulasan Stilistika Tertulis dinilai dengan Rubrik Apresiasi Puisi.",
     "Rekaman deklamasi puisi ekspresif & lembar anotasi kritik sosial lirik lagu.",
     "1. BKoF: Mendengarkan lagu legendaris 'Imagine' / 'Blowin' in the Wind' & puisi Robert Frost.\n2. MoT: Dekonstruksi struktur bait, rima, dan citraan puitis puisi 'Invictus' / 'Still I Rise'.\n3. JCoT: Menelaah makna baris demi baris bersama kelompok dan memberi catatan kritis.\n4. ICoT: Deklamasi puisi mandiri di panggung kelas & penyusunan esai interpretasi puitik.",
     "10 JP (2 Minggu x 5 JP)",
     "Antologi Puisi Bahasa Inggris Pilihan, Rekaman Audio Pembacaan Puisi Profesional, Rubrik Resitasi"),

    ("Modul 5", "Semester 4", "TP 11.TL.5", "Menyimak-Berbicara & Menulis-Mempresentasikan",
     "Mengonstruksi dan menimbang dua sudut pandang berimbang dalam teks diskusi mengenai kontroversi e-sports/gaming kompetitif dan AI (Model Resmi Bab D Panduan Puskurjar 2025), serta berdebat gaya Oxford.",
     "Discussion Text (Official Puskurjar Bab D Model):\n• Issue, Arguments For, Arguments Against, Balanced Recommendation\n• Contrasting connectives (on the contrary, whereas, conversely)\n• Hedging/cautious academic language & live parliamentary debate",
     "IKTP 5.1: Mengidentifikasi argumen pro dan kontra serta data pendukung dari teks model resmi Puskurjar 'Online Games'.\nIKTP 5.2: Menggunakan ragam bahasa kehati-hatian akademik (hedging: evidence suggests, it could be argued).\nIKTP 5.3: Melaksanakan simulasi debat formal gaya parlemen/Oxford dan menyusun esai sintesis berimbang 500 kata.",
     "• Asesmen Diagnostik: Pemetaan posisi awal siswa terkait esports dan kecanduan bermain gim.\n• Asesmen Formatif: Lembar Kerja LKPD Resmi Bab D Panduan Puskurjar 2025 & draf argumen debat.\n• Asesmen Sumatif: Unjuk Kerja Debat Formal Gaya Oxford & Esai Diskusi Sintesis dinilai dengan Rubrik Analitik Diskusi Bab D Puskurjar (Skala 1-4).",
     "Transkrip argumen debat kelompok & esai teks diskusi berimbang 500 kata (Model Bab D).",
     "1. BKoF: Curah pendapat fenomena industri game online bernilai miliaran dolar vs isu kesehatan mental.\n2. MoT: Bedah teks resmi Bab D Panduan 2025 'Pros and Cons of Playing Online Games' & fitur bahasa hedging.\n3. JCoT: Simulasi debat formal 3 tim afirmasi vs 3 tim oposisi dengan mosi terstandar.\n4. ICoT: Menulis esai sintesis rekomendasi mandiri & refleksi metakognitif (Learning log).",
     "25 JP (5 Minggu x 5 JP)",
     "Modul Ajar Bab D Panduan Mapel Bahasa Inggris 2025 (hlm 158-166), Panduan Debat Oxford, LKPD Resmi"),

    ("Modul 6", "Semester 4", "TP 11.TL.6", "Membaca-Memirsa & Menulis-Mempresentasikan",
     "Menganalisis mekanisme kausalitas ilmiah kompleks dan urutan proses terjadinya fenomena teknologi mutakhir dan astrofisika dalam teks eksplanasi ilmiah tingkat tinggi.",
     "Complex Causal Explanation Text:\n• Phenomena statement, multi-layered causal explanation, scientific summary\n• Multi-branching cause-effect pathways, technical passive voice, nominalization\n• Technical schematics & scientific magazine layout",
     "IKTP 6.1: Memetakan rantai kausalitas bercabang banyak (multi-branching causality) fenomena kuantum/astrofisika.\nIKTP 6.2: Menganalisis kalimat pasif dengan frase preposisi kompleks dan terminologi sains mutakhir.\nIKTP 6.3: Menyusun artikel eksplanasi ilmiah populer 500 kata dilengkapi bagan skematis teknis untuk majalah sains.",
     "• Asesmen Diagnostik: Kuis konsep fenomena sains frontier (black hole, computing quantum, blockchain).\n• Asesmen Formatif: Lembar pemetaan jaring-jaring sebab-akibat ilmiah (Causal Network Mapping).\n• Asesmen Sumatif: Penilaian Artikel Majalah Sains Populer dinilai dengan Rubrik Menulis Eksplanasi Saintifik.",
     "Artikel sains populer berlayout majalah (Scientific Magazine Article) karya mandiri siswa.",
     "1. BKoF: Menonton animasi mekanisme kerja komputer kuantum dan penjelajahan lubang hitam.\n2. MoT: Dekonstruksi teks 'How Quantum Computers Break Classical Limits' & analisis kepadatan leksikal.\n3. JCoT: Berpasangan menyusun diagram alir sebab-akibat fenomena fusi nuklir atau AI superkomputer.\n4. ICoT: Menulis artikel ilmiah mandiri dan mendesain tata letak visual infografis teknis.",
     "25 JP (5 Minggu x 5 JP)",
     "Jurnal Ilmiah Populer Scientific American / Nature, Software Desain Grafis, Rubrik Artikel Sains"),

    ("Modul 7", "Semester 4", "TP 11.TL.7", "Membaca-Memirsa & Menulis-Mempresentasikan",
     "Mengevaluasi rekam jejak perjuangan, kepemimpinan transformasional, dan kontribusi historis tokoh-tokoh visioner dunia dalam teks biografi mendalam dan memoar sejarah.",
     "Comprehensive Biography & Historical Memoirs:\n• Historical milestones, obstacles overcome, transformational leadership impact\n• Past Perfect & Past Perfect Continuous, complex temporal clauses\n• Character introspection monograph & biographical podcast",
     "IKTP 7.1: Mengidentifikasi titik balik kehidupan (turning points) dan dampak kepemimpinan tokoh dunia dari teks biografi.\nIKTP 7.2: Menggunakan tenses lampau kompleks (Past Perfect Tense) untuk menyatakan urutan dua peristiwa lampau.\nIKTP 7.3: Menulis mini-monograf biografi 500 kata atau memproduksi rekaman podcast biografi tokoh inspiratif.",
     "• Asesmen Diagnostik: Identifikasi tokoh-tokoh pembawa perubahan dunia yang dikagumi siswa.\n• Asesmen Formatif: Bagan Linimasa Titik Balik Tokoh (Milestone Turning Points) & latihan Past Perfect.\n• Asesmen Sumatif: Penilaian Produk Mini-Monograf Biografi / Podcast Biografi dinilai dengan Rubrik Biografi Analitis.",
     "Naskah monograf biografi tokoh visioner & file rekaman podcast berdurasi 7 menit.",
     "1. BKoF: Diskusi kepemimpinan tokoh visioner (Nelson Mandela, Marie Curie, B.J. Habibie, Steve Jobs).\n2. MoT: Bedah buku memoar sejarah (orientasi, krisis kehidupan, pencapaian transformatif, evaluasi sejarah).\n3. JCoT: Kelompok meriset data arsip sejarah tokoh dan memverifikasi keabsahan sumber.\n4. ICoT: Menulis naskah monograf biografi mandiri dan merekam podcast naratif informatif.",
     "20 JP (4 Minggu x 5 JP)",
     "Buku Biografi Tokoh Dunia, Arsip Pidato Sejarah, Templat Naskah Podcast"),

    ("Modul 8", "Semester 4", "TP 11.TL.8", "Membaca-Memirsa & Menulis-Mempresentasikan",
     "Menyusun resensi dan kritik estetika mendalam (Review Text) terhadap novel sastra dunia atau karya film berkualitas dengan rubrik evaluasi artistik objektif.",
     "Critical Review of International Literature & Cinema:\n• Orientation, interpretative recount, multi-criteria aesthetic evaluation, summative punch\n• Aesthetic vocabulary, cinematic terminology, cultural commentary\n• Critical review portfolio publication",
     "IKTP 8.1: Menilai aspek sinematografi, akting, atau gaya penulisan novel secara mendalam dan berimbang.\nIKTP 8.2: Menggunakan kosakata evaluasi estetika tingkat tinggi (masterpiece, nuanced, compelling, flawed).\nIKTP 8.3: Menulis artikel kritik seni/sastra 400 kata dan mempresentasikannya dalam gelar wicara budaya kelas.",
     "• Asesmen Diagnostik: Kuis membedakan sinopsis biasa vs resensi kritis analitis.\n• Asesmen Formatif: Matriks Evaluasi Karya Multi-Kriteria (Plot, Karakter, Tema, Estetika Sinema).\n• Asesmen Sumatif: Portofolio Artikel Resensi Kritis yang siap terbit dinilai dengan Rubrik Kritik Seni.",
     "Artikel resensi kritis berstandar jurnalisme budaya & gelar wicara mini presentasi resensi.",
     "1. BKoF: Menonton cuplikan film pemenang penghargaan internasional (Oscar / Cannes Film Festival).\n2. MoT: Menganalisis artikel kritik film The New York Times (istilah sinematik, argumen apresiasi).\n3. JCoT: Berpasangan mengevaluasi kelebihan dan kekurangan sebuah karya sastra klasik.\n4. ICoT: Menulis naskah resensi kritis secara mandiri & pameran portofolio resensi di akhir fase.",
     "10 JP (2 Minggu x 5 JP)",
     "Jurnal Kritik Film & Sastra Internasional, Cuplikan Film Pemenang Festival, Rubrik Kritik Seni")
]

# KELAS XII WAJIB (FASE F UMUM - SEMESTER 5 & 6)
DATA_SILABUS_XII_WAJIB = [
    ("Modul 1", "Semester 5", "TP 12.W.1", "Menyimak-Berbicara & Membaca-Memirsa",
     "Menganalisis teks diskusi mengenai disrupsi dunia kerja masa depan, etika kecerdasan buatan, dan keberlanjutan lingkungan, serta menyusun teks diskusi berimbang.",
     "Discussion Text (Future Work & Environment):\n• Issue, Arguments Pro & Con, Balanced Recommendation\n• Discourse markers of concession (although, even though, despite)\n• Hedging and objective stance in evaluating workplace automation",
     "IKTP 1.1: Menganalisis dampak otomasi AI terhadap lapangan kerja dari berbagai sudut pandang pemangku kepentingan.\nIKTP 1.2: Menggunakan konjungsi konsesif (despite, in spite of) dan ragam bahasa netral objektif.\nIKTP 1.3: Menyusun teks diskusi tertulis 300 kata yang menimbang sisi positif dan negatif otomasi kerja.",
     "• Asesmen Diagnostik: Polling karir masa depan impian siswa dan kekhawatiran digantikan AI.\n• Asesmen Formatif: Bagan Analisis Pro-Kontra multi-sektor & latihan konjungsi konsesif.\n• Asesmen Sumatif: Esai Teks Diskusi Berimbang dinilai dengan Rubrik Penilaian Teks Diskusi.",
     "Esai teks diskusi berimbang bertema masa depan dunia kerja dan kecerdasan buatan.",
     "1. BKoF: Menonton video laporan World Economic Forum tentang 'Future of Jobs Report'.\n2. MoT: Menganalisis teks diskusi 'Will AI Replace Human Workers?' & membedah struktur argumen.\n3. JCoT: Berkelompok menyusun matriks solusi transisi tenaga kerja di era digital.\n4. ICoT: Menulis esai diskusi mandiri dan melakukan peer-review antarteman.",
     "10 JP (5 Pertemuan x 2 JP)",
     "Laporan World Economic Forum, Artikel Kompas/BBC, Lembar Kerja Murid"),

    ("Modul 2", "Semester 5", "TP 12.W.2", "Membaca-Memirsa & Menulis-Mempresentasikan",
     "Mengidentifikasi fakta utama, latar belakang peristiwa, dan sumber berita dalam teks berita aktual (News Item Text), serta memproduksi siaran berita (news broadcast) informatif.",
     "News Item Text & Media Literacy:\n• Newsworthy event, background events, sources\n• Headlines, reporting verbs (stated, announced, claimed), direct quotes\n• News anchor presentation & video broadcasting simulation",
     "IKTP 2.1: Menemukan 5W+1H pada teks berita aktual dari kantor berita resmi terpercaya.\nIKTP 2.2: Menggunakan kata kerja pelaporan (reporting verbs) dan penulisan kutipan langsung narasumber yang benar.\nIKTP 2.3: Mempraktikkan siaran berita lisan (news anchor) secara berpasangan dalam rekaman video 2 menit.",
     "• Asesmen Diagnostik: Kuis membedakan berita fakta vs berita hoaks/umpan klik (clickbait).\n• Asesmen Formatif: Lembar Kerja Analisis Struktur Berita 5W+1H & latihan kalimat kutipan.\n• Asesmen Sumatif: Penilaian Praktik Siaran Berita (News Broadcast Video) dinilai dengan Rubrik News Casting (Kelancaran, Diksi, Intonasi Suara, Pembawaan Formal).",
     "Video rekaman simulasi presenter berita TV (News Casting) berdurasi 2 menit & naskah berita.",
     "1. BKoF: Menonton siaran berita bahasa Inggris (CNA / BBC World News / Antara News).\n2. MoT: Dekonstruksi struktur berita cetak & siaran televisi (lead, detail, kutipan narasumber).\n3. JCoT: Menulis naskah berita liputan sekolah/daerah secara berpasangan.\n4. ICoT: Merekam video penampilan pembaca berita mandiri dan penilaian unjuk kerja.",
     "10 JP (5 Pertemuan x 2 JP)",
     "Rekaman Berita Televisi Aktual, Naskah Berita Kantor Berita Antara, Mikrofon/Kamera Ponsel"),

    ("Modul 3", "Semester 5", "TP 12.W.3", "Membaca-Memirsa & Menulis-Mempresentasikan",
     "Mengevaluasi dan menyusun ulasan kritis (Review Text) terhadap produk kreatif nusantara, buku, atau film lokal untuk promosi kekayaan budaya Indonesia secara multimodal.",
     "Review Text (Indonesian Creative Products & Heritage):\n• Orientation, evaluation, interpretative recount, evaluative summary\n• Evaluative adjectives, cultural appreciation terms, commercial appeal\n• Multimodal product review blog post & photo layout",
     "IKTP 3.1: Menganalisis nilai keunikan dan kualitas estetika produk kreatif nusantara (batik, tenun, kriya, kuliner).\nIKTP 3.2: Menggunakan kosakata evaluasi bernada positif dan konstruktif untuk mempromosikan produk lokal.\nIKTP 3.3: Menulis artikel ulasan produk kreatif nusantara disertai foto pendukung dan mempublikasikannya di blog/media sosial.",
     "• Asesmen Diagnostik: Brainstorming produk kreatif unggulan daerah masing-masing siswa.\n• Asesmen Formatif: Lembar telaah kriteria penilaian produk kreatif (keaslian, estetika, fungsionalitas).\n• Asesmen Sumatif: Penilaian Artikel Ulasan Produk Budaya (Blog Review Post) dinilai dengan Rubrik Menulis Resensi Produk.",
     "Artikel blog ulasan produk kreatif lokal berfoto estetik untuk promosi UMKM/budaya nusantara.",
     "1. BKoF: Mengamati pameran produk ekonomi kreatif dan kriya tradisional nusantara.\n2. MoT: Membedah artikel ulasan produk kriya lokal berstandar ekspor pada majalah gaya hidup.\n3. JCoT: Berpasangan menyusun draf ulasan produk kreatif lokal pilihan kelompok.\n4. ICoT: Menulis artikel ulasan mandiri dan mempublikasikannya secara digital.",
     "10 JP (5 Pertemuan x 2 JP)",
     "Katalog Produk Kreatif Kemenparekraf, Foto Produk Kriya Nusantara, Platform Blog/Media Sosial"),

    ("Modul 4", "Semester 6", "TP 12.W.4", "Membaca-Memirsa & Menulis-Mempresentasikan",
     "Menyusun surat lamaran pekerjaan formal (Job Application Letter) dan daftar riwayat hidup modern (Curriculum Vitae / Resume) sesuai standar profesional dunia kerja modern.",
     "Job Application Letter & Modern Curriculum Vitae:\n• Formal business letter format, salutation, body paragraphs, polite closing\n• ATS-friendly CV design, summary statement, highlighting hard/soft skills\n• Professional action verbs, error-free formatting & executive tone",
     "IKTP 4.1: Menganalisis iklan lowongan kerja (job vacancy) dan menyesuaikan kualifikasi diri yang relevan.\nIKTP 4.2: Menyusun surat lamaran pekerjaan formal bahasa Inggris dengan tata bahasa dan ejaan baku.\nIKTP 4.3: Merancang Curriculum Vitae ringkas, terstruktur, dan ramah sistem pemindaian pelamar kerja (ATS-friendly).",
     "• Asesmen Diagnostik: Kuis membedakan format surat informal vs formal business letter.\n• Asesmen Formatif: Peer-editing draf surat lamaran kerja menggunakan panduan ceklis profesionalitas.\n• Asesmen Sumatif: Penilaian Berkas Dossier Lamaran Lengkap (Cover Letter + CV) dinilai dengan Rubrik Dokumen Karir Profesional Puskurjar.",
     "Berkas surat lamaran kerja resmi & Curriculum Vitae (CV) ATS-compliant siap pakai.",
     "1. BKoF: Mengkaji tren lowongan pekerjaan masa kini di portal karir internasional/nasional.\n2. MoT: Membedah contoh surat lamaran formal profesional dan format CV modern 1 halaman.\n3. JCoT: Berpasangan saling mereview kesesuaian kualifikasi dan keterampilan dalam draf CV.\n4. ICoT: Menyusun berkas lamaran kerja mandiri lengkap dengan portofolio pencapaian diri.",
     "10 JP (5 Pertemuan x 2 JP)",
     "Iklan Lowongan Kerja Asli, Templat CV ATS-Friendly, Panduan Karir Profesional"),

    ("Modul 5", "Semester 6", "TP 12.W.5", "Menyimak-Berbicara",
     "Mempraktikkan percakapan wawancara kerja formal (Job Interview Simulation) dan keterampilan presentasi diri secara percaya diri, santun, dan runtut.",
     "Job Interview Simulation & Oral Presentation:\n• Responding to common interview queries (strengths, weakness, future goals)\n• Asking clarifying questions, active listening, professional executive presence\n• Voice modulation, polite formal register, positive body language",
     "IKTP 5.1: Mengidentifikasi pertanyaan umum wawancara kerja dan strategi menjawab dengan percaya diri.\nIKTP 5.2: Menggunakan bahasa tubuh profesional (eye contact, posture) dan intonasi suara yang mantap.\nIKTP 5.3: Mendemonstrasikan simulasi wawancara kerja di hadapan panel penilai dengan respon spontan dan santun.",
     "• Asesmen Diagnostik: Identifikasi kekuatan diri (strengths) dan area pengembangan diri (weaknesses).\n• Asesmen Formatif: Latihan berpasangan tanya-jawab wawancara kerja (Mock Interview Drill).\n• Asesmen Sumatif: Ujian Praktik Sekolah / Simulasi Wawancara Kerja Lisan dinilai dengan Rubrik Wawancara Kerja Formal Puskurjar.",
     "Lembar catatan wawancara panelis & rekaman simulasi wawancara kerja formal siswa.",
     "1. BKoF: Menonton video simulasi wawancara kerja profesional berstandar internasional.\n2. MoT: Bedah pola pertanyaan wawancara perilaku (behavioral interview questions) & strategi menjawab.\n3. JCoT: Simulasi wawancara bergantian dalam kelompok kecil 3 orang (pewawancara, pelamar, observer).\n4. ICoT: Penampilan wawancara kerja mandiri di hadapan guru penguji & umpan balik individual.",
     "8 JP (4 Pertemuan x 2 JP)",
     "Daftar Pertanyaan Wawancara Kerja Standar, Lembar Penilaian Observasi Wawancara, Ruang Simulasi"),

    ("Modul 6", "Semester 6", "TP 12.W.6", "Membaca-Memirsa & Menulis-Mempresentasikan",
     "Menafsirkan makna dan pesan tersirat dalam takarir (Caption), grafik, dan infografis multimodal, serta menyajikan ringkasan visual ke dalam bahasa Inggris yang komunikatif.",
     "Captions, Data Graphs, & Multimodal Infographics:\n• Caption structure (lead sentence, explanatory detail, photo credit)\n• Interpreting statistical charts (bar, line, pie charts) & trends\n• Writing concise, punchy, and accurate informative photo captions",
     "IKTP 6.1: Menemukan makna pesan tersirat dan konteks sosial foto jurnalistik dari teks takarir (caption).\nIKTP 6.2: Membaca tren kenaikan/penurunan data pada grafik batang, garis, dan diagram lingkaran.\nIKTP 6.3: Menulis takarir informatif 2-3 kalimat untuk foto peristiwa atau infografis data ilmiah.",
     "• Asesmen Diagnostik: Kuis mencocokkan foto jurnalistik dengan takarir yang sesuai.\n• Asesmen Formatif: Lembar analisis data grafik statistik (Describing trends worksheet).\n• Asesmen Sumatif: Asesmen Akhir Jenjang / Penilaian Portofolio Infografis Bertakarir dinilai dengan Rubrik Literasi Multimodal.",
     "Kompilasi poster infografis bertakarir komunikatif & lembar jawaban analisis grafik.",
     "1. BKoF: Mengamati foto-foto pemenang World Press Photo dan diagram tren data sosial.\n2. MoT: Bedah kaidah penulisan caption foto jurnalistik (present tense untuk aksi, konteks, sumber foto).\n3. JCoT: Berpasangan menuliskan takarir untuk 3 gambar fenomena sosial yang berbeda.\n4. ICoT: Menulis takarir mandiri untuk infografis karya siswa & verifikasi kelulusan akhir jenjang.",
     "6 JP (3 Pertemuan x 2 JP)",
     "Koleksi Foto Jurnalistik Berita, Grafik Data Statistik BPS/WHO, Rubrik Penilaian Caption")
]

# KELAS XII TINGKAT LANJUT (FASE F TL - SEMESTER 5 & 6 - 5 JP/MINGGU)
DATA_SILABUS_XII_TL = [
    ("Modul 1", "Semester 5", "TP 12.TL.1", "Menyimak-Berbicara & Membaca-Memirsa",
     "Menganalisis wacana debat akademik tingkat tinggi (Advanced Discussion) mengenai bioetika medis, rekayasa genetika, eksplorasi antariksa, dan merumuskan sikap etis ilmiah secara komprehensif.",
     "Global Academic Discussion on Bioethics & Frontier Technologies:\n• Multi-layered argumentation, philosophical nuances, ethical stances\n• Advanced stance markers, concessive clauses, academic collocations\n• Scholarly dialogue / Policy deliberation roundtable",
     "IKTP 1.1: Mengidentifikasi argumen filosofis dan implikasi moral dari rekayasa genetika CRISPR dan bioetika medis.\nIKTP 1.2: Menggunakan penanda sikap akademik (stance markers: arguably, it is fundamentally imperative) dan klausa konsesif.\nIKTP 1.3: Memimpin dan berpartisipasi dalam dialog meja bundar akademik serta menulis makalah posisi bioetika 500 kata.",
     "• Asesmen Diagnostik: Pemetaan wawasan siswa tentang rekayasa genetika dan batas etika sains.\n• Asesmen Formatif: Lembar Analisis Argumen Filosofis & latihan klausa konsesif akademik.\n• Asesmen Sumatif: Unjuk Kerja Diskusi Meja Bundar Akademik & Makalah Posisi Bioetika dinilai dengan Rubrik Diskusi Akademik Tingkat Tinggi.",
     "Transkrip dialog meja bundar ilmiah & makalah posisi bioetika akademik 500 kata.",
     "1. BKoF: Menonton dokumenter sains pengeditan gen manusia dan penjelajahan koloni Mars.\n2. MoT: Membedah jurnal bioetika 'The Ethics of Human Genome Editing' (analisis stance markers, hedging).\n3. JCoT: Simulasi sidang komite etika medis universitas memecahkan studi kasus kontroversial.\n4. ICoT: Menulis esai posisi bioetika mandiri dan refleksi integritas moral ilmuwan.",
     "25 JP (5 Minggu x 5 JP)",
     "Jurnal Bioetika Internasional, Video Dokumenter CRISPR/Bioethics, Rubrik Diskusi Meja Bundar"),

    ("Modul 2", "Semester 5", "TP 12.TL.2", "Membaca-Memirsa & Menulis-Mempresentasikan",
     "Menyusun esai penelitian akademik (Academic Synthesis Essay) 1.000 kata dengan memadukan multi-referensi, kutipan langsung/tidak langsung, dan daftar pustaka standar internasional (APA/MLA).",
     "Academic Research Essay Writing & Source Synthesis:\n• Formulating working thesis, literature synthesis, topic sentences\n• In-text citation protocols (APA/MLA style), paraphrasing, avoiding plagiarism\n• Cohesive devices, academic hedging, formal objective register",
     "IKTP 2.1: Merumuskan working thesis statement dan menyintesis minimal 3 sumber jurnal ilmiah terakreditasi.\nIKTP 2.2: Menerapkan teknik sitasi dalam teks (in-text citation) gaya APA/MLA dan teknik parafrasa bebas plagiarisme.\nIKTP 2.3: Menyusun esai sintesis akademik utuh 1.000 kata berstruktur 5 paragraf dengan daftar pustaka yang akurat.",
     "• Asesmen Diagnostik: Uji pemahaman perbedaan kutipan langsung vs parafrasa & bahaya plagiarisme.\n• Asesmen Formatif: Matriks Sintesis Literatur (Literature Synthesis Matrix) & draf sitasi daftar pustaka.\n• Asesmen Sumatif: Penilaian Makalah Penelitian Akademik 1.000 Kata dinilai dengan Rubrik Esai Akademik Internasional (Thesis, Synthesis, In-text Citations, Coherence, Grammar).",
     "Makalah sintesis penelitian akademik 1.000 kata lengkap dengan daftar pustaka gaya APA/MLA.",
     "1. BKoF: Pelatihan pencarian referensi jurnal ilmiah di Google Scholar / ScienceDirect.\n2. MoT: Bedah esai akademik mahasiswa universitas (sintesis 3 sumber, topic sentence, sitasi APA).\n3. JCoT: Latihan berpasangan memparafrasa penggalan teks jurnal tanpa mengubah makna asli.\n4. ICoT: Penulisan draf mandiri, peer-review ketat, dan finalisasi makalah akademik 1.000 kata.",
     "25 JP (5 Minggu x 5 JP)",
     "Database Jurnal Google Scholar, Panduan Sitasi APA Style 7th Edition, Perangkat Lunak Plagiarisme"),

    ("Modul 3", "Semester 5", "TP 12.TL.3", "Menyimak-Berbicara & Membaca-Memirsa",
     "Menganalisis teks naskah drama klasik/modern, menafsirkan subteks psikologis, serta memperagakan adegan monolog/dialog dramatik dengan intonasi dan pelafalan standar profesional.",
     "Dramatic Literature & Theatrical Script Performance:\n• Dramatic structure, conflict escalation, stage directions, subtext\n• Voice modulation, stress, intonation, emotional projection\n• Script adaptation & theatrical live performance / Dramatic monologue recording",
     "IKTP 3.1: Menganalisis petunjuk panggung (stage directions), ketegangan dramatik, dan subteks emosional tokoh naskah sandiwara.\nIKTP 3.2: Menginterpretasikan konflik batin karakter sastra ke dalam variasi modulasi suara dan tekanan kata.\nIKTP 3.3: Mementaskan pertunjukan monolog dramatik atau adegan sandiwara singkat berdurasi 3-5 menit secara profesional.",
     "• Asesmen Diagnostik: Membaca teks dialog pendek dengan berbagai variasi emosi (marah, sedih, bangga).\n• Asesmen Formatif: Lembar Bedah Naskah Teater (Subtext & Stage Directions) & latihan vokal berpasangan.\n• Asesmen Sumatif: Unjuk Kerja Pementasan Drama / Rekaman Monolog Teatrikal dinilai dengan Rubrik Penilaian Seni Peran Teater (Vocal Delivery, Emotional Projection, Blocking/Movement).",
     "Rekaman video penampilan monolog dramatik siswa & naskah sandiwara yang dianotasi.",
     "1. BKoF: Menonton pertunjukan teater Broadway / pementasan sandiwara Shakespeare modern.\n2. MoT: Menganalisis naskah sandiwara klasik (Hamlet / Death of a Salesman) dan subteks tersembunyi.\n3. JCoT: Latihan blocking panggung dan pembacaan naskah dialog intensif bersama kelompok.\n4. ICoT: Pementasan monolog mandiri di panggung mini kelas dan evaluasi pertunjukan.",
     "20 JP (4 Minggu x 5 JP)",
     "Naskah Drama Sastra Dunia, Rekaman Teater Profesional, Kostum & Tata Panggung Sederhana"),

    ("Modul 4", "Semester 5", "TP 12.TL.4", "Menyimak-Berbicara & Menulis-Mempresentasikan",
     "Merancang dan membawakan pidato persuasif tingkat tinggi (Persuasive Oratory) dan mempraktikkan debat formal gaya parlemen dengan retorika persuasif yang meyakinkan.",
     "Persuasive Public Speaking & Parliamentary Debating:\n• Classical rhetoric: Ethos (credibility), Pathos (emotion), Logos (logic)\n• Parallelism, antithesis, persuasive prosody, rebuttal strategies\n• Formal speech delivery in international forum simulation",
     "IKTP 4.1: Mengaplikasikan perangkat retorika klasik (Ethos, Pathos, Logos) dan gaya bahasa pengulangan (parallelism, antithesis).\nIKTP 4.2: Menyusun naskah pidato formal 400 kata yang memiliki tesis menggugah, argumen kuat, dan seruan bertindak.\nIKTP 4.3: Menyampaikan pidato oratori di podium forum internasional simulasi dengan retorika vokal memukau.",
     "• Asesmen Diagnostik: Analisis kutipan pidato terkenal dunia ('I Have a Dream' - Martin Luther King Jr.).\n• Asesmen Formatif: Lembar Telaah Perangkat Retorika & latihan penyusunan kalimat antithesis.\n• Asesmen Sumatif: Unjuk Kerja Pidato Oratori Formal (Persuasive Oratory Contest) dinilai dengan Rubrik Oratori Publik (Kekuatan Argumen, Bahasa Retoris, Daya Tarik Emosi, Karisma Panggung).",
     "Naskah pidato oratori persuasif resmi & video rekaman penampilan pidato di podium.",
     "1. BKoF: Mengkaji pidato-pidato bersejarah pemimpin dunia yang mengubah peradaban manusia.\n2. MoT: Dekonstruksi retorika pidato (teknik pembukaan yang memukau, pembuktian tiga pilar, klimaks seruan).\n3. JCoT: Berpasangan saling melatih intonasi retoris dan jeda teatrikal pidato.\n4. ICoT: Penampilan oratori pidato formal mandiri di hadapan seluruh siswa kelas.",
     "10 JP (2 Minggu x 5 JP)",
     "Rekaman Pidato Bersejarah Dunia, Naskah Pidato Oratori, Podium Kelas & Mikrofon"),

    ("Modul 5", "Semester 6", "TP 12.TL.5", "Membaca-Memirsa & Menulis-Mempresentasikan",
     "Menyusun korespondensi kerja profesional dan akademik internasional: Surat Lamaran Kerja, Curriculum Vitae (CV) ATS-compliant, dan Pernyataan Motivasi Beasiswa/Kuliah (Personal Statement).",
     "Professional Career & Academic University Admissions Portfolio:\n• Formal business letter format, salutation, professional tone\n• Modern ATS-friendly Curriculum Vitae / Resume design\n• Compelling personal statement for scholarships / college admissions",
     "IKTP 5.1: Menganalisis persyaratan beasiswa internasional dan kriteria seleksi perguruan tinggi top dunia.\nIKTP 5.2: Menulis Personal Statement motivasi diri yang otentik, memikat, dan menonjolkan keunikan pencapaian pribadi.\nIKTP 5.3: Menyusun dossier pendaftaran kuliah/beasiswa lengkap (Cover Letter, ATS CV, Personal Statement) bebas kesalahan format.",
     "• Asesmen Diagnostik: Refleksi visi masa depan, pilihan jurusan perguruan tinggi, dan beasiswa impian.\n• Asesmen Formatif: Peer-review Personal Statement menggunakan rubrik esai pendaftaran universitas global.\n• Asesmen Sumatif: Penilaian Berkas Dossier Akademik Lengkap dinilai dengan Rubrik Admissions Portfolio Internasional.",
     "Paket berkas pendaftaran universitas/beasiswa (Cover Letter, CV ATS, Personal Statement 600 kata).",
     "1. BKoF: Mengkaji panduan esai pendaftaran beasiswa (LPDP, Chevening, Fulbright, MEXT).\n2. MoT: Membedah contoh Personal Statement sukses mahasiswa Indonesia di universitas dunia (Ivy League/Oxbridge).\n3. JCoT: Berpasangan menyusun kerangka cerita narasi perjalanan hidup yang menginspirasi.\n4. ICoT: Menulis naskah Personal Statement mandiri, revisi draf berulang, dan finalisasi berkas.",
     "20 JP (4 Minggu x 5 JP)",
     "Panduan Aplikasi Universitas Internasional, Contoh Esai Personal Statement Sukses, Templat CV ATS"),

    ("Modul 6", "Semester 6", "TP 12.TL.6", "Menyimak-Berbicara",
     "Mempraktikkan wawancara kerja profesional (Job Interview Simulation) dan pertahanan karya ilmiah lisan (Oral Defense) dengan respons spontan, taktis, dan register bahasa formal.",
     "Job Interview & Academic Oral Defense Simulation:\n• STAR Method (Situation, Task, Action, Result) in behavioral questions\n• Defending research methodology and conclusions under questioning\n• Professional executive presence, composure, and tone",
     "IKTP 6.1: Menerapkan formula STAR (Situation, Task, Action, Result) dalam menjawab studi kasus wawancara seleksi.\nIKTP 6.2: Menjawab dan mempertahankan metodologi serta temuan karya ilmiah dari sanggahan kritis penguji.\nIKTP 6.3: Menampilkan ketenangan emosional, artikulasi bahasa Inggris presisi, dan kehadiran eksekutif selama sesi wawancara.",
     "• Asesmen Diagnostik: Latihan menjawab pertanyaan tak terduga dalam waktu 30 detik.\n• Asesmen Formatif: Lembar Kerja Formulasi Jawaban STAR Method & simulasi tanya-jawab antar-kawan.\n• Asesmen Sumatif: Ujian Praktik Sekolah / Sidang Pertahanan Karya Ilmiah Lisan (Oral Defense) dinilai dengan Rubrik Sidang Wawancara Eksekutif.",
     "Lembar berita acara sidang oral defense & rekaman video wawancara eksekutif siswa.",
     "1. BKoF: Menelaah video sidang skripsi/tesis universitas dan wawancara kerja perusahaan multinasional.\n2. MoT: Analisis strategi STAR Method dan kesantunan menanggapi pertanyaan kritik dewan penguji.\n3. JCoT: Simulasi sidang mini dengan saling bergantian menjadi dewan penguji dan kandidat.\n4. ICoT: Ujian sidang oral defense mandiri di hadapan guru penguji resmi.",
     "15 JP (3 Minggu x 5 JP)",
     "Daftar Pertanyaan Sidang Karya Ilmiah, Panduan STAR Method, Ruang Sidang Simulasi"),

    ("Modul 7", "Semester 6", "TP 12.TL.7", "Membaca-Memirsa & Menulis-Mempresentasikan",
     "Menyusun abstrak laporan penelitian ilmiah (Research Abstract) 250 kata dan mendesain presentasi infografis saintifik untuk diseminasi publik dalam simposium ilmiah siswa.",
     "Mini Academic Paper Abstract & Scientific Infographic Poster:\n• 250-word structured abstract (Background, Method, Results, Conclusion)\n• Scientific reporting verbs, nominalization, data visualization synthesis\n• Academic poster symposium session",
     "IKTP 7.1: Merangkum intisari penelitian ilmiah (latar belakang, metode, hasil, simpulan) dalam 250 kata berstruktur baku.\nIKTP 7.2: Menggunakan terminologi ilmiah presisi dan kata kerja laporan ilmiah (revealed, demonstrated, concluded).\nIKTP 7.3: Mendesain poster infografis saintifik dan mempresentasikannya dalam ajang simposium ilmiah siswa.",
     "• Asesmen Diagnostik: Kuis membedakan abstrak penelitian yang baik vs abstrak yang tidak lengkap.\n• Asesmen Formatif: Lembar Panduan Penulisan Abstrak 250 Kata (Abstract Template Drafting).\n• Asesmen Sumatif: Penilaian Produk Poster Simposium Ilmiah & Abstrak Cetak dinilai dengan Rubrik Simposium Ilmiah Internasional.",
     "Naskah abstrak ilmiah 250 kata siap publikasi & poster infografis saintifik ukuran A1/digital.",
     "1. BKoF: Mengamati pameran poster ilmiah simposium penelitian internasional.\n2. MoT: Bedah struktur abstrak jurnal ilmiah internasional terindeks Scopus (IMRaD structure).\n3. JCoT: Berpasangan saling mengoreksi keringkasan kata dan ketepatan diksi dalam abstrak.\n4. ICoT: Finalisasi naskah abstrak mandiri dan presentasi poster dalam simposium ilmiah kelas.",
     "15 JP (3 Minggu x 5 JP)",
     "Koleksi Abstrak Jurnal Ilmiah Terindeks, Templat Desain Poster Simposium, Papan Pameran Poster"),

    ("Modul 8", "Semester 6", "TP 12.TL.8", "Menyimak, Membaca, Menulis & Berbicara",
     "Mengembangkan kesiapan literasi bahasa Inggris standar internasional (International Academic Literacy: IELTS/TOEFL/CEFR B2+ reading & critical analytical writing).",
     "Standardized English Proficiency Readiness for Higher Education:\n• Advanced skimming, scanning, synthesizing inferential questions\n• Cohesion, coherence, and lexical precision under timed test conditions\n• Benchmark diagnostic proficiency assessment",
     "IKTP 8.1: Menerapkan teknik membaca cepat (skimming/scanning) dan penarikan kesimpulan inferensial pada teks bacaan kompleks.\nIKTP 8.2: Menganalisis data grafik/tabel dan menyusun ringkasan laporan analitis tertulis 150 kata dalam waktu 20 menit.\nIKTP 8.3: Mengerjakan tes simulasi kemahiran bahasa Inggris standar internasional dan mencapai target benchmark level B2.",
     "• Asesmen Diagnostik: Pre-test diagnostik membaca dan tata bahasa standar IELTS/TOEFL.\n• Asesmen Formatif: Latihan soal pemahaman inferensial & penulisan esai deskripsi grafik di bawah batasan waktu.\n• Asesmen Sumatif: Asesmen Sumatif Akhir Jenjang / Benchmark Standardized English Proficiency Test dinilai dengan standar skor resmi CEFR B2.",
     "Lembar hasil skor simulasi kemahiran bahasa Inggris standar & berkas esai analisis grafik di bawah waktu ujian.",
     "1. BKoF: Penjelasan format tes kemahiran bahasa Inggris internasional (IELTS Academic / TOEFL iBT / Duolingo English Test).\n2. MoT: Membedah strategi menjawab soal bacaan inferensial panjang dan strategi penulisan esai waktu terbatas.\n3. JCoT: Latihan berkelompok menganalisis grafik kompleks dan merumuskan kalimat pengantar laporan.\n4. ICoT: Pelaksanaan tes simulasi mandiri terstandar sebagai evaluasi akhir kelulusan tingkat lanjut.",
     "10 JP (2 Minggu x 5 JP)",
     "Bank Soal Simulasi Standar Internasional (IELTS/TOEFL style), Perangkat Audio Listening, Lembar Jawaban Komputer")
]


# ==============================================================================
# FUNCTION TO BUILD SILABUS WORKSHEET TABLE
# ==============================================================================

def add_silabus_table_sheet(wb, sheet_name, page_title, class_name, fase_name, subject_name, alokasi_jp, silabus_data):
    ws = wb.create_sheet(title=sheet_name)
    apply_title(ws, f"SILABUS PEMBELAJARAN (BLUEPRINT MODUL AJAR) - {class_name.upper()} ({fase_name.upper()})",
                f"Mata Pelajaran: {subject_name} | Alokasi: {alokasi_jp} JP / Tahun | Prinsip: Constructive Alignment (IKTP <-> Asesmen)", 12)

    school_data = [
        ("Satuan Pendidikan", "Sekolah Menengah Atas (SMA)"),
        ("Mata Pelajaran", subject_name),
        ("Fase / Kelas", f"{fase_name} / {class_name}"),
        ("Tahun Pelajaran", "2025/2026"),
        ("Beban Belajar", f"{alokasi_jp} Jam Pelajaran (JP) / Tahun"),
        ("Prinsip Utama", "Keselarasan Konstruktif (Indikator Ketercapaian diukur secara presisi oleh Asesmen Formatif & Sumatif)")
    ]
    r_start = add_school_identity(ws, school_data, start_row=4)

    headers = [
        "No Modul", "Semester", "Kode TP", "Elemen CP",
        "Tujuan Pembelajaran (TP)",
        "Materi Pokok & Ruang Lingkup Genre",
        "Indikator Ketercapaian (IKTP / KKTP)",
        "Rencana Asesmen Autentik (Diagnostik, Formatif, Sumatif)",
        "Bukti Belajar / Produk Autentik (Evidence)",
        "Pokok Langkah Pembelajaran (Siklus GBA & Deep Learning)",
        "Alokasi Waktu", "Media & Sumber Belajar"
    ]

    ws.row_dimensions[r_start].height = 28
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=r_start, column=col_idx, value=h)
        cell.font = FONT_TBL_HEADER
        cell.fill = FILL_NAVY
        cell.alignment = ALIGN_CENTER
        cell.border = CELL_BORDER

    curr_row = r_start + 1
    for row_val in silabus_data:
        modul_no, sem, kode, elemen, tp, materi, iktp, asesmen, bukti, gba_steps, jp, media = row_val
        ws.row_dimensions[curr_row].height = 80

        ws.cell(row=curr_row, column=1, value=modul_no).font = FONT_DATA_BOLD
        ws.cell(row=curr_row, column=1).alignment = ALIGN_CENTER
        ws.cell(row=curr_row, column=1).fill = FILL_LIGHT_BLUE
        ws.cell(row=curr_row, column=1).border = CELL_BORDER

        ws.cell(row=curr_row, column=2, value=sem).font = FONT_DATA
        ws.cell(row=curr_row, column=2).alignment = ALIGN_CENTER
        ws.cell(row=curr_row, column=2).border = CELL_BORDER

        ws.cell(row=curr_row, column=3, value=kode).font = FONT_DATA_BOLD
        ws.cell(row=curr_row, column=3).alignment = ALIGN_CENTER
        ws.cell(row=curr_row, column=3).border = CELL_BORDER

        ws.cell(row=curr_row, column=4, value=elemen).font = FONT_DATA_BOLD
        ws.cell(row=curr_row, column=4).alignment = ALIGN_CENTER
        ws.cell(row=curr_row, column=4).fill = FILL_VERY_LIGHT
        ws.cell(row=curr_row, column=4).border = CELL_BORDER

        ws.cell(row=curr_row, column=5, value=tp).font = FONT_DATA
        ws.cell(row=curr_row, column=5).alignment = ALIGN_TOP_LEFT
        ws.cell(row=curr_row, column=5).border = CELL_BORDER

        ws.cell(row=curr_row, column=6, value=materi).font = FONT_DATA
        ws.cell(row=curr_row, column=6).alignment = ALIGN_TOP_LEFT
        ws.cell(row=curr_row, column=6).border = CELL_BORDER

        # Indikator Ketercapaian (IKTP)
        c_iktp = ws.cell(row=curr_row, column=7, value=iktp)
        c_iktp.font = FONT_DATA_BOLD
        c_iktp.alignment = ALIGN_TOP_LEFT
        c_iktp.fill = FILL_ACCENT_GREEN
        c_iktp.border = CELL_BORDER

        # Asesmen (Keselarasan Langsung dengan IKTP)
        c_asesmen = ws.cell(row=curr_row, column=8, value=asesmen)
        c_asesmen.font = FONT_DATA
        c_asesmen.alignment = ALIGN_TOP_LEFT
        c_asesmen.fill = FILL_ACCENT_AMBER
        c_asesmen.border = CELL_BORDER

        ws.cell(row=curr_row, column=9, value=bukti).font = FONT_DATA_BOLD
        ws.cell(row=curr_row, column=9).alignment = ALIGN_TOP_LEFT
        ws.cell(row=curr_row, column=9).border = CELL_BORDER

        ws.cell(row=curr_row, column=10, value=gba_steps).font = FONT_DATA
        ws.cell(row=curr_row, column=10).alignment = ALIGN_TOP_LEFT
        ws.cell(row=curr_row, column=10).border = CELL_BORDER

        ws.cell(row=curr_row, column=11, value=jp).font = FONT_DATA_NUMBER
        ws.cell(row=curr_row, column=11).alignment = ALIGN_CENTER
        ws.cell(row=curr_row, column=11).border = CELL_BORDER

        ws.cell(row=curr_row, column=12, value=media).font = FONT_ITALIC_DESC
        ws.cell(row=curr_row, column=12).alignment = ALIGN_TOP_LEFT
        ws.cell(row=curr_row, column=12).border = CELL_BORDER

        curr_row += 1

    add_signatures(ws, start_row=curr_row + 2, max_col=12)

    ws.column_dimensions['A'].width = 10
    ws.column_dimensions['B'].width = 13
    ws.column_dimensions['C'].width = 11
    ws.column_dimensions['D'].width = 22
    ws.column_dimensions['E'].width = 34
    ws.column_dimensions['F'].width = 32
    ws.column_dimensions['G'].width = 38
    ws.column_dimensions['H'].width = 38
    ws.column_dimensions['I'].width = 28
    ws.column_dimensions['J'].width = 38
    ws.column_dimensions['K'].width = 16
    ws.column_dimensions['L'].width = 28
    ws.views.sheetView[0].showGridLines = True


# ==============================================================================
# SHEET: PANDUAN & MATRIKS KESELARASAN IKTP DENGAN ASESMEN (CONSTRUCTIVE ALIGNMENT)
# ==============================================================================

def add_guide_and_alignment_sheet(wb):
    ws = wb.create_sheet(title="Panduan_&_Alignment_Asesmen")
    apply_title(ws, "PANDUAN PENYUSUNAN SILABUS & MATRIKS KESELARASAN KONSTRUKTIF (CONSTRUCTIVE ALIGNMENT)",
                "Integrasi Pokok Modul Ajar: Keselarasan Presisi antara Indikator Ketercapaian (IKTP) dan Asesmen Pengukur", 7)

    ws.cell(row=4, column=1, value="I. HAKIKAT & POKOK-POKOK MODUL AJAR DALAM SILABUS KURIKULUM MERDEKA").font = FONT_SEC_HEADER
    ws.merge_cells("A4:G4")
    for c in range(1, 8):
        ws.cell(row=4, column=c).fill = FILL_NAVY
    ws.row_dimensions[4].height = 24

    points = [
        ("Kedudukan Silabus", "Silabus berfungsi sebagai cetak biru (blueprint) kurikuler yang merangkum pokok-pokok seluruh Modul Ajar selama 1 semester hingga 1 tahun. Silabus memastikan bahwa guru memiliki peta navigasi yang jelas sebelum menyusun rincian LKPD dan skenario pertemuan harian."),
        ("Prinsip Keselarasan Konstruktif (Constructive Alignment)", "Indikator Ketercapaian Tujuan Pembelajaran (IKTP / KKTP) TIDAK BOLEH berdiri sendiri tanpa alat ukur yang jelas. Setiap indikator yang dirumuskan WAJIB berpasangan secara presisi dengan bentuk dan teknik asesmen yang mengukurnya (misal: indikator reseptif diukur dengan lembar kerja pemahaman, indikator lisan diukur dengan unjuk kerja berbicara, indikator tulis diukur dengan rubrik penulisan teks)."),
        ("Integrasi 3 Jenis Asesmen", "Silabus memuat asesmen komprehensif:\n1. Asesmen Diagnostik (awal pembelajaran): mendeteksi kesiapan dan pengetahuan awal murid.\n2. Asesmen Formatif (selama proses belajar): memberikan umpan balik berkelanjutan (feedback) tanpa nilai rapor.\n3. Asesmen Sumatif (akhir modul/fase): mengukur ketuntasan kompetensi murid dengan rubrik analitik 4 skala baku Puskurjar."),
        ("Pendekatan Berbasis Teks & Pembelajaran Mendalam (Deep Learning)", "Aktivitas belajar dalam silabus dirancang menurut Teaching Learning Cycle (TLC) Genre-based Approach: Building Knowledge of the Field (BKoF - Meaningful), Modelling of the Text (MoT - Mindful), Joint Construction of the Text (JCoT - Joyful/Collaborative), dan Independent Construction of the Text (ICoT - Independent)."),
        ("Bukti Autentik (Evidence of Learning)", "Penilaian tidak semata-mata mengandalkan tes kertas-pensil (paper-and-pencil test), melainkan menuntut bukti nyata performa berbahasa murid: brosur promosi, esai argumentatif, rekaman podcast, simulasi debat, naskah drama, abstrak ilmiah, hingga dossier lamaran kerja profesional.")
    ]

    r = 5
    for lbl, txt in points:
        ws.row_dimensions[r].height = 38
        c1 = ws.cell(row=r, column=1, value=lbl)
        c1.font = FONT_DATA_BOLD
        c1.fill = FILL_LIGHT_BLUE
        c1.alignment = ALIGN_LEFT
        c1.border = CELL_BORDER

        c2 = ws.cell(row=r, column=2, value=txt)
        c2.font = FONT_DATA
        c2.alignment = ALIGN_TOP_LEFT
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        for c in range(2, 8):
            ws.cell(row=r, column=c).border = CELL_BORDER
            if r % 2 == 1:
                ws.cell(row=r, column=c).fill = FILL_ZEBRA
        r += 1

    r += 1
    # Matriks Keselarasan Jenis Kompetensi vs Teknik Asesmen
    ws.cell(row=r, column=1, value="II. MATRIKS STANDAR KESELARASAN IKTP DENGAN TEKNIK & INSTRUMEN ASESMEN").font = FONT_SEC_HEADER
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    for c in range(1, 8):
        ws.cell(row=r, column=c).fill = FILL_ROYAL
    ws.row_dimensions[r].height = 24

    r += 1
    m_headers = ["Ranah Kemampuan", "Fokus Elemen CP", "Karakteristik Rumusan IKTP", "Teknik Asesmen yang Tepat", "Bentuk Instrumen Asesmen", "Wujud Bukti Autentik (Evidence)", "Kriteria Ketercapaian"]
    for i, h in enumerate(m_headers, 1):
        cell = ws.cell(row=r, column=i, value=h)
        cell.font = FONT_TBL_HEADER
        cell.fill = FILL_NAVY
        cell.alignment = ALIGN_CENTER
        cell.border = CELL_BORDER
    ws.row_dimensions[r].height = 22

    matrix_data = [
        ("Reseptif Lisan", "Menyimak (Listening)", "Mampu mengidentifikasi ide pokok, informasi rinci, dan intonasi pembicara dari audio monolog/dialog.", "Tes Lisan / Tes Menyimak Terstruktur", "Lembar Kerja Menyimak (Audio Listening Sheet) & Kuis Isian Singkat", "Lembar jawaban menyimak & peta konsep informasi audio", "Menjawab benar minimal 75% butir informasi rinci"),
        ("Reseptif Tulis & Visual", "Membaca - Memirsa (Reading - Viewing)", "Mampu menganalisis struktur teks, gagasan tersirat, fungsi sosial, dan tata bahasa teks model.", "Tes Tulis Pemahaman Membaca & Telaah Teks", "Graphic Organizer Struktur Teks, Soal HOTS Pilihan Ganda & Analisis Teks", "Bagan analisis struktur teks & lembar jawaban HOTS", "Mampu mengurai elemen struktur teks secara lengkap dan logis"),
        ("Produktif Lisan", "Berbicara (Speaking)", "Mampu menyampaikan gagasan, berdebat, berpidato, atau melakukan wawancara kerja secara lisan.", "Asesmen Kinerja / Unjuk Kerja Performa Lisan", "Rubrik Analitik Berbicara 4 Skala Puskurjar (Fluency, Pronunciation, Grammar, Task Achievement)", "Rekaman video presentasi, simulasi wawancara, atau live performance", "Mencapai skor minimal kategori 'Meets Standards' (Skala 3 dari 4)"),
        ("Produktif Tulis", "Menulis (Writing)", "Mampu menyusun draf teks, mengembangkan argumen, dan merevisi tulisan dengan tata bahasa akurat.", "Penilaian Portofolio & Penugasan Esai Mandiri", "Rubrik Analitik Menulis 4 Skala Puskurjar (Content, Organization, Vocabulary, Language Use, Mechanics)", "Esai tertulis, brosur wisata, naskah drama, abstrak ilmiah, atau CV", "Memenuhi standar struktur genre dan kohesi paragraf (Skala 3 dari 4)"),
        ("Multimodal & Kolaboratif", "Mempresentasikan (Presenting)", "Mampu memadukan teks verbal, gambar/grafik visual, dan gestur tubuh dalam gelar wicara/poster.", "Penilaian Produk Multimodal & Observasi Partisipasi", "Rubrik Produk Poster/Video & Lembar Observasi Kolaborasi Kelompok", "Poster infografis digital, video tutorial, podcast, atau stan simposium", "Kesesuaian pesan visual dengan teks narasi pendukung")
    ]

    r += 1
    for r_val in matrix_data:
        ws.row_dimensions[r].height = 36
        for c_idx, val in enumerate(r_val, 1):
            cell = ws.cell(row=r, column=c_idx, value=val)
            cell.font = FONT_DATA_BOLD if c_idx in [1, 2] else FONT_DATA
            cell.alignment = ALIGN_CENTER if c_idx in [1, 2] else ALIGN_TOP_LEFT
            cell.border = CELL_BORDER
            if r % 2 == 1:
                cell.fill = FILL_ZEBRA
        r += 1

    auto_fit_cols(ws, min_w=12, max_w=46)


# ==============================================================================
# MAIN FUNCTION TO GENERATE COMPLETE SILABUS EXCEL
# ==============================================================================

def generate_silabus_workbook(output_file):
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove default sheet

    # 1. Sheet 1: Panduan & Matriks Keselarasan Asesmen
    add_guide_and_alignment_sheet(wb)

    # 2. Sheet 2: Silabus Kelas X Wajib (Semester 1 & 2 - 72 JP)
    add_silabus_table_sheet(
        wb, "Silabus_Kelas_X_Wajib", "Silabus Kelas X Wajib",
        "Kelas X", "Fase E", "Bahasa Inggris Umum (Wajib)", 72, DATA_SILABUS_X_WAJIB
    )

    # 3. Sheet 3: Silabus Kelas XI Wajib (Semester 3 & 4 - 72 JP)
    add_silabus_table_sheet(
        wb, "Silabus_Kelas_XI_Wajib", "Silabus Kelas XI Wajib",
        "Kelas XI", "Fase F Umum", "Bahasa Inggris Umum (Wajib)", 72, DATA_SILABUS_XI_WAJIB
    )

    # 4. Sheet 4: Silabus Kelas XI Tingkat Lanjut (Semester 3 & 4 - 180 JP)
    add_silabus_table_sheet(
        wb, "Silabus_Kelas_XI_TL", "Silabus Kelas XI Tingkat Lanjut",
        "Kelas XI", "Fase F Tingkat Lanjut", "Bahasa Inggris Tingkat Lanjut (Pilihan)", 180, DATA_SILABUS_XI_TL
    )

    # 5. Sheet 5: Silabus Kelas XII Wajib (Semester 5 & 6 - 64 JP)
    add_silabus_table_sheet(
        wb, "Silabus_Kelas_XII_Wajib", "Silabus Kelas XII Wajib",
        "Kelas XII", "Fase F Umum", "Bahasa Inggris Umum (Wajib)", 64, DATA_SILABUS_XII_WAJIB
    )

    # 6. Sheet 6: Silabus Kelas XII Tingkat Lanjut (Semester 5 & 6 - 160 JP)
    add_silabus_table_sheet(
        wb, "Silabus_Kelas_XII_TL", "Silabus Kelas XII Tingkat Lanjut",
        "Kelas XII", "Fase F Tingkat Lanjut", "Bahasa Inggris Tingkat Lanjut (Pilihan)", 160, DATA_SILABUS_XII_TL
    )

    wb.save(output_file)
    print(f"SILABUS Unified Workbook successfully created at: {output_file}")


def main():
    target_dir = r"C:\Users\KOMPUTERKU2\Documents\English SMA\4-Silabus"
    os.makedirs(target_dir, exist_ok=True)
    out_file = os.path.join(target_dir, "Silabus_Bahasa_Inggris_SMA_Wajib_dan_Tingkat_Lanjut.xlsx")
    
    print("Generating comprehensive Silabus workbook...")
    generate_silabus_workbook(out_file)

if __name__ == "__main__":
    main()
