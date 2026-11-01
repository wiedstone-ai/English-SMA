# -*- coding: utf-8 -*-
"""
Script to generate a comprehensive, highly styled Excel workbook summarizing
PANDUAN MATA PELAJARAN BAHASA INGGRIS REVISI 2025 (Puskurjar Kemendikdasmen)
Fokus: Kurikulum, Capaian Pembelajaran, Pedagogi Deep Learning & GBA, Modul Ajar SMA
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_panduan_bing_excel(output_path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove default sheet

    font_family = "Segoe UI"
    
    # Fonts
    font_main_title = Font(name=font_family, size=15, bold=True, color="FFFFFF")
    font_subtitle = Font(name=font_family, size=10, italic=True, color="E2E8F0")
    font_section_header = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    font_tbl_header = Font(name=font_family, size=10, bold=True, color="FFFFFF")
    font_group_header = Font(name=font_family, size=10, bold=True, color="0F2942")
    font_subgroup = Font(name=font_family, size=10, bold=True, color="1E3A8A")
    font_total = Font(name=font_family, size=10, bold=True, color="0F2942")
    font_data = Font(name=font_family, size=10, bold=False, color="1E293B")
    font_data_bold = Font(name=font_family, size=10, bold=True, color="1E293B")
    font_english = Font(name=font_family, size=9, italic=True, color="334155")
    font_callout_title = Font(name=font_family, size=10, bold=True, color="92400E")
    font_callout_body = Font(name=font_family, size=9, bold=False, color="78350F")

    # Fills
    fill_main_title = PatternFill(start_color="0F2942", end_color="0F2942", fill_type="solid")     # Deep Navy
    fill_subtitle = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")       # Dark Blue
    fill_tbl_header = PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid")     # Royal Blue
    fill_card_header = PatternFill(start_color="2563EB", end_color="2563EB", fill_type="solid")    # Accent Blue
    fill_group = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")          # Soft Light Blue
    fill_subgroup = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")       # Very Light Blue
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")          # Slate White
    fill_total = PatternFill(start_color="BFDBFE", end_color="BFDBFE", fill_type="solid")          # Highlight Total Blue
    fill_callout = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")        # Warm Amber Note
    fill_callout_green = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")  # Soft Emerald Green

    # Borders
    border_thin_gray = Side(style='thin', color="CBD5E1")
    border_double_bottom = Side(style='double', color="0F2942")
    border_thick_top = Side(style='medium', color="1E40AF")
    
    cell_border = Border(left=border_thin_gray, right=border_thin_gray, top=border_thin_gray, bottom=border_thin_gray)
    total_border = Border(left=border_thin_gray, right=border_thin_gray, top=border_thin_gray, bottom=border_double_bottom)

    # Alignments
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    align_top_left = Alignment(horizontal='left', vertical='top', wrap_text=True)
    align_top_center = Alignment(horizontal='center', vertical='top', wrap_text=True)

    def apply_title_banner(ws, title, subtitle, max_col=6):
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
        
        c1 = ws.cell(row=1, column=1, value=title)
        c1.font = font_main_title
        c1.fill = fill_main_title
        c1.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 32

        c2 = ws.cell(row=2, column=1, value=subtitle)
        c2.font = font_subtitle
        c2.fill = fill_subtitle
        c2.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[2].height = 20

        for col in range(1, max_col + 1):
            ws.cell(row=1, column=col).fill = fill_main_title
            ws.cell(row=2, column=col).fill = fill_subtitle

    def auto_fit_columns(ws, min_width=10, max_width=55):
        ws.views.sheetView[0].showGridLines = True
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                if cell.row in [1, 2, 3]:
                    continue
                val = str(cell.value or '')
                if '\n' in val:
                    lines = val.split('\n')
                    val_len = max(len(l) for l in lines)
                else:
                    val_len = len(val)
                if val_len > max_len:
                    max_len = val_len
            ws.column_dimensions[col_letter].width = max(min(max_len + 4, max_width), min_width)

    # ==========================================
    # SHEET 1: RINGKASAN PANDUAN MAPEL 2025
    # ==========================================
    ws1 = wb.create_sheet(title="Ringkasan Panduan Mapel 2025")
    apply_title_banner(
        ws1,
        "PANDUAN MATA PELAJARAN BAHASA INGGRIS REVISI 2025 (FASE B - F & F TINGKAT LANJUT)",
        "Dokumen Resmi Pusat Kurikulum dan Pembelajaran (Puskurjar) BSKAP Kemendikdasmen Edisi Revisi 3 (September 2025)",
        max_col=6
    )

    ws1.cell(row=4, column=1, value="1. INFORMASI RESMI DOKUMEN PANDUAN").font = font_section_header
    ws1.merge_cells("A4:F4")
    for c in range(1, 7):
        ws1.cell(row=4, column=c).fill = fill_tbl_header
    ws1.row_dimensions[4].height = 24

    doc_info = [
        ("Judul Panduan", "Panduan Mata Pelajaran Bahasa Inggris Fase B – Fase F dan Fase F Tingkat Lanjut (Edisi Revisi 3, September 2025)"),
        ("Instansi Penerbit", "Pusat Kurikulum dan Pembelajaran, Badan Standar, Kurikulum, dan Asesmen Pendidikan (BSKAP) Kemendikdasmen RI"),
        ("Tujuan Penerbitan", "Memberikan panduan operasional implementatif bagi guru Bahasa Inggris dalam merancang pembelajaran berbasis teks (Genre-based Approach) yang selaras dengan pendekatan Pembelajaran Mendalam (Deep Learning)."),
        ("Fokus Khusus Jenjang SMA", "Menjangkau Fase E (Kelas X SMA), Fase F Umum/Wajib (Kelas XI-XII SMA), dan Fase F Tingkat Lanjut (Mata Pelajaran Pilihan Kelas XI-XII SMA)."),
        ("Isu Kritis yang Diselesaikan", "Mengatasi miskonsepsi pengajaran teks 'sekali jadi'. Pembelajaran teks menuntut siklus terencana (Teaching Learning Cycle: BKoF, MoT, JCoT, ICoT) selama beberapa pertemuan hingga murid benar-benar menguasai kompetensi.")
    ]

    r = 5
    for item, val in doc_info:
        ws1.cell(row=r, column=1, value=item).font = font_data_bold
        ws1.cell(row=r, column=1).fill = fill_group
        ws1.cell(row=r, column=1).alignment = align_left
        ws1.cell(row=r, column=1).border = cell_border
        
        ws1.cell(row=r, column=2, value=val).font = font_data
        ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws1.cell(row=r, column=2).alignment = align_left
        for col in range(2, 7):
            ws1.cell(row=r, column=col).border = cell_border
            if r % 2 == 1:
                ws1.cell(row=r, column=col).fill = fill_zebra
        ws1.row_dimensions[r].height = 26
        r += 1

    r += 1
    # Matriks Keselarasan Deep Learning & GBA di Bab A
    ws1.cell(row=r, column=1, value="2. KESELARASAN PEMBELAJARAN MENDALAM DENGAN PENGAJARAN BERBASIS TEKS (GENRE)").font = font_section_header
    ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws1.cell(row=r, column=c).fill = fill_card_header
    ws1.row_dimensions[r].height = 24
    r += 1

    headers_tbl1 = ["Prinsip Pembelajaran Mendalam", "Prinsip Pengajaran Berbasis Teks (Tabel 1 Panduan)", "Tahap Pengalaman Belajar", "Operasionalisasi di Kelas Bahasa Inggris SMA", "Fungsi Pedagogis bagi Guru", "Output Kinerja Murid"]
    for i, h in enumerate(headers_tbl1, 1):
        cell = ws1.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = cell_border
    ws1.row_dimensions[r].height = 26
    r += 1

    tbl1_data = [
        ("Bermakna (Meaningful)",
         "Pengajaran berfokus pada teks, yakni satu kesatuan bahasa yang bermakna dan mempunyai tujuan sosial nyata dalam interaksi manusia.",
         "MEMAHAMI (Understanding Context & Purpose)",
         "Setiap aktivitas belajar mengaitkan teks dengan isu riil di sekitar murid (lingkungan, sains, sosial, budaya); bukan sekadar menghafal rumus gramatika lepas konteks.",
         "Menyajikan stimulus video otentik, artikel isu terkini, dan studi kasus nyata.",
         "Murid memahami mengapa dan untuk situasi apa sebuah teks digunakan."),
        ("Berkesadaran (Mindful)",
         "Semua tindakan berbahasa dipikirkan dengan matang, dilakukan tahap demi tahap dengan penuh konsentrasi dan pemikiran terencana.",
         "MEMAHAMI & MENGAPLIKASI (Analyzing Architecture)",
         "Murid secara sadar membedah arsitektur struktur teks (generic structure), tenses, kosa kata kunci, serta menyadari proses berpikirnya saat menyusun kalimat.",
         "Memandu analisis model teks (MoT) dan membimbing refleksi metakognisi murid.",
         "Murid mengendalikan pilihan kosakata dan struktur kalimat secara akurat."),
        ("Menggembirakan (Joyful)",
         "Kemampuan menggunakan bahasa Inggris untuk mencapai tujuan tertentu dan mencipta karya teks lisan/tulis memunculkan rasa senang dan bangga.",
         "MENGAPLIKASI (Creating & Expressing Voice)",
         "Suasana kelas positif, suportif, saling memuliakan teman yang sedang berbicara; bebas celaan aksen; terhubung secara emosional melalui bermain peran, drama, atau presentasi.",
         "Membangun iklim kelas aman psikologis di mana kesalahan dipandang sebagai bagian wajar belajar.",
         "Murid antusias berekspresi, percaya diri berbicara, dan bangga atas karya teksnya."),
        ("Merefleksi Berkelanjutan",
         "Dilakukan di setiap tahapan pembelajaran, saat murid dan pendidik menilai apa yang dilakukan, keberhasilan, serta tantangan yang dihadapi.",
         "MEREFLEKSI (Self-Assessment & Metacognition)",
         "Murid mengisi lembar evaluasi diri (3 minutes writing), mencatat perbendaharaan kata baru dalam learning log, dan menyusun strategi perbaikan.",
         "Memberikan umpan balik deskriptif kualitatif yang membangun arah perbaikan.",
         "Murid berkembang menjadi pembelajar mandiri sepanjang hayat (autonomous learners).")
    ]

    for item in tbl1_data:
        prs, gba, exp, ops, fun, out = item
        ws1.cell(row=r, column=1, value=prs).alignment = align_left
        ws1.cell(row=r, column=2, value=gba).alignment = align_left
        ws1.cell(row=r, column=3, value=exp).alignment = align_center
        ws1.cell(row=r, column=4, value=ops).alignment = align_left
        ws1.cell(row=r, column=5, value=fun).alignment = align_left
        ws1.cell(row=r, column=6, value=out).alignment = align_left

        ws1.cell(row=r, column=1).font = font_data_bold
        ws1.cell(row=r, column=2).font = font_data
        ws1.cell(row=r, column=3).font = font_data_bold
        for c in range(4, 7):
            ws1.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws1.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws1.cell(row=r, column=c).fill = fill_zebra
        ws1.row_dimensions[r].height = 54
        r += 1

    auto_fit_columns(ws1, min_width=10, max_width=45)
    ws1.column_dimensions['A'].width = 20
    ws1.column_dimensions['B'].width = 34
    ws1.column_dimensions['C'].width = 24
    ws1.column_dimensions['D'].width = 34
    ws1.column_dimensions['E'].width = 30
    ws1.column_dimensions['F'].width = 30
    ws1.freeze_panes = 'A5'

    # ==========================================
    # SHEET 2: KOMPONEN & KARAKTERISTIK CP
    # ==========================================
    ws2 = wb.create_sheet(title="Komponen & Karakteristik CP")
    apply_title_banner(
        ws2,
        "KOMPONEN & KARAKTERISTIK CAPAIAN PEMBELAJARAN BAHASA INGGRIS",
        "Penjabaran Bab B Panduan Puskurjar 2025: Rasional, Karakteristik Teks, Elemen & Standar CEFR",
        max_col=6
    )

    ws2.cell(row=4, column=1, value="1. TIGA (3) PASANG ELEMEN KOMPETENSI BAHASA INGGRIS (BAB B.2.c)").font = font_section_header
    ws2.merge_cells("A4:F4")
    for c in range(1, 7):
        ws2.cell(row=4, column=c).fill = fill_tbl_header
    ws2.row_dimensions[4].height = 24

    headers_elemen = ["Pasangan Elemen", "Jenis Keterampilan Bahasa", "Karakteristik & Hakikat Reseptif / Produktif", "Cakupan Modalitas Teks", "Fokus Aktivitas di SMA", "Kaidah Penyusunan ATP"]
    r = 5
    for i, h in enumerate(headers_elemen, 1):
        cell = ws2.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws2.row_dimensions[5].height = 26
    r += 1

    elemen_details = [
        ("Menyimak - Berbicara\n(Listening - Speaking)",
         "Bahasa Lisan (Spoken Language) Reseptif & Produktif",
         "Menyimak adalah kemampuan memahami alur informasi lisan; Berbicara adalah kemampuan memproduksi tuturan lisan untuk berinteraksi, berdialog, dan berargumen.",
         "Teks lisan tatap muka, audio rekaman, podcast, video percakapan, instruksi lisan, debat formal.",
         "Melatih kemampuan merespons pertanyaan terbuka, mempertahankan giliran bicara (turn-taking), dan menyampaikan sanggahan argumen santun.",
         "Penomoran elemen (1, 2, 3) BUKAN urutan hierarkis kaku. Elemen dapat dipadukan dalam siklus pembelajaran teks."),
        ("Membaca - Memirsa\n(Reading - Viewing)",
         "Bahasa Tulis & Visual Reseptif",
         "Membaca adalah memahami informasi teks tertulis; Memirsa (Viewing) adalah kemampuan menginterpretasi dan mengevaluasi teks visual/audiovisual multimodal.",
         "Teks cetak, artikel digital, infografis, poster, diagram, video dokumenter, halaman web.",
         "Menganalisis ide utama, mengevaluasi pesan tersurat/tersirat, membedakan fakta vs opini, dan mengidentifikasi sudut pandang penulis esai.",
         "Membaca dan memirsa menjadi input pemodelan (Modelling) sebelum murid diminta menulis dan mempresentasikan."),
        ("Menulis - Mempresentasikan\n(Writing - Presenting)",
         "Bahasa Tulis & Visual Produktif",
         "Menulis adalah memproduksi teks tertulis terstruktur; Mempresentasikan adalah menyajikan ide melalui berbagai media cetak/digital secara efektif di hadapan audiens.",
         "Esai tulis, laporan hasil riset, naskah pidato, slide presentasi digital (Canva/PPT), video rekaman.",
         "Menyusun esai eksposisi analitis berbobot, menulis cerita pendek kreatif, dan menyajikan paparan presentasi yang memikat.",
         "Merupakan muara penguasaan bahasa di mana murid menunjukkan kemandirian dan orisinalitas gagasannya.")
    ]

    for item in elemen_details:
        el, jk, kh, md, fa, kd = item
        ws2.cell(row=r, column=1, value=el).alignment = align_center
        ws2.cell(row=r, column=2, value=jk).alignment = align_left
        ws2.cell(row=r, column=3, value=kh).alignment = align_left
        ws2.cell(row=r, column=4, value=md).alignment = align_left
        ws2.cell(row=r, column=5, value=fa).alignment = align_left
        ws2.cell(row=r, column=6, value=kd).alignment = align_left

        ws2.cell(row=r, column=1).font = font_data_bold
        ws2.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 7):
            ws2.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws2.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws2.cell(row=r, column=c).fill = fill_zebra
        ws2.row_dimensions[r].height = 54
        r += 1

    r += 1
    # Standar CEFR per Fase
    ws2.cell(row=r, column=1, value="2. TARGET TINGKAT KEMAHIRAN BAHASA INGGRIS SESUAI CEFR (BAB B.2.a)").font = font_section_header
    ws2.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws2.cell(row=r, column=c).fill = fill_tbl_header
    ws2.row_dimensions[r].height = 24
    r += 1

    headers_cefr = ["Fase Pembelajaran", "Jenjang Pendidikan", "Target Level CEFR", "Karakteristik Kemahiran Bahasa", "Cakupan Kompleksitas Teks", "Kemampuan Mandiri Murid"]
    for i, h in enumerate(headers_cefr, 1):
        cell = ws2.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws2.row_dimensions[r].height = 26
    r += 1

    cefr_data = [
        ("Fase B", "Kelas 3 - 4 SD/MI", "A1 (Breakthrough)", "Mengenali dan menggunakan ungkapan dasar sehari-hari yang sangat sederhana dengan bantuan visual.", "Kata tunggal, frasa pendek, instruksi kelas sederhana bergambar.", "Perlu bimbingan penuh dan stimulus visual konstan."),
        ("Fase C", "Kelas 5 - 6 SD/MI", "A1 - A2 (Waystage)", "Memahami kalimat dan ungkapan lazim seputar keluarga, lingkungan sekolah, dan rutinitas harian.", "Kalimat sederhana, deskripsi benda sekitar, cerita fabel sangat pendek.", "Dapat berinteraksi singkat pada kebutuhan rutin harian."),
        ("Fase D", "Kelas 7 - 9 SMP/MTs", "A2 (Waystage - Pre-Intermediate)", "Memahami poin utama pesan sederhana; mendeskripsikan pengalaman masa lampau dan rencana sederhana.", "Teks deskripsi, recount, prosedur, dan narasi sederhana.", "Mampu berkomunikasi dalam tugas-tugas rutin yang memerlukan pertukaran informasi langsung."),
        ("Fase E (SMA)", "Kelas X SMA/MA", "B1 (Threshold - Intermediate)", "Memahami ide pokok teks standar mengenai topik familiar dan isu terkini; mengungkapkan pendapat dan alasan.", "Teks faktual, deskripsi analitis, laporan pengalaman, pesan fungsional multimodal.", "Mampu memulai, mempertahankan, dan menutup percakapan; menulis paragraf terpadu."),
        ("Fase F Wajib (SMA)", "Kelas XI - XII SMA/MA", "B1 - B2 (Vantage)", "Memahami alur ide kompleks teks konkret & abstrak; berinteraksi dengan tingkat kelancaran wajar tanpa ketegangan.", "Teks eksposisi analitis, teks eksposisi hortatori, teks diskusi pro-kontra, narasi sastra.", "Mampu mempertahankan argumen dengan data, mengevaluasi fakta vs opini, dan self-editing."),
        ("Fase F Tingkat Lanjut", "Kelas XI - XII SMA Pilihan", "B2 (Vantage - Independent User)", "Memahami gagasan utama teks kompleks termasuk diskusi teknis akademis; mengekspresikan diri secara fasih dan spontan.", "Esai akademis berbasis riset empiris, cerpen fiksi orisinal, diskursus isu kontroversial global.", "Mampu menghasilkan teks yang jelas dan rinci tentang berbagai subjek; berdebat formal secara matang.")
    ]

    for item in cefr_data:
        fs, jnj, cfr, chr, ckp, mnd = item
        ws2.cell(row=r, column=1, value=fs).alignment = align_center
        ws2.cell(row=r, column=2, value=jnj).alignment = align_left
        ws2.cell(row=r, column=3, value=cfr).alignment = align_center
        ws2.cell(row=r, column=4, value=chr).alignment = align_left
        ws2.cell(row=r, column=5, value=ckp).alignment = align_left
        ws2.cell(row=r, column=6, value=mnd).alignment = align_left

        is_sma = "SMA" in jnj or "Fase E" in fs or "Fase F" in fs
        ws2.cell(row=r, column=1).font = font_data_bold
        ws2.cell(row=r, column=2).font = font_data_bold if is_sma else font_data
        ws2.cell(row=r, column=3).font = font_data_bold
        for c in range(4, 7):
            ws2.cell(row=r, column=c).font = font_data
            
        fill_curr = fill_callout_green if is_sma else (fill_zebra if r % 2 == 1 else PatternFill(fill_type=None))
        for c in range(1, 7):
            ws2.cell(row=r, column=c).border = cell_border
            if fill_curr.fill_type:
                ws2.cell(row=r, column=c).fill = fill_curr
        ws2.row_dimensions[r].height = 42
        r += 1

    auto_fit_columns(ws2, min_width=10, max_width=45)
    ws2.column_dimensions['A'].width = 18
    ws2.column_dimensions['B'].width = 24
    ws2.column_dimensions['C'].width = 20
    ws2.column_dimensions['D'].width = 32
    ws2.column_dimensions['E'].width = 32
    ws2.column_dimensions['F'].width = 34
    ws2.freeze_panes = 'A5'

    # ==========================================
    # SHEET 3: CP FASE E & F (WAJIB & TL)
    # ==========================================
    ws3 = wb.create_sheet(title="CP Fase E & F (Wajib & TL)")
    apply_title_banner(
        ws3,
        "MATRIKS RESMI CAPAIAN PEMBELAJARAN BAHASA INGGRIS SMA (BAB B.2.d PANDUAN)",
        "Rumusan Capaian Pembelajaran Fase E (Kelas X), Fase F Wajib (Kelas XI-XII), dan Fase F Tingkat Lanjut",
        max_col=6
    )

    ws3.cell(row=4, column=1, value="PERBANDINGAN ELEMEN CAPAIAN PEMBELAJARAN SMA (FASE E vs F WAJIB vs F TINGKAT LANJUT)").font = font_section_header
    ws3.merge_cells("A4:F4")
    for c in range(1, 7):
        ws3.cell(row=4, column=c).fill = fill_tbl_header
    ws3.row_dimensions[4].height = 24

    headers_cp_comp = ["Elemen Kompetensi", "Fase E (Kelas X SMA)", "Fase F Umum / Wajib (Kelas XI - XII SMA)", "Fase F Tingkat Lanjut (Pilihan SMA)", "Kunci Peningkatan Kognitif", "Implikasi Tugas Kelas"]
    r = 5
    for i, h in enumerate(headers_cp_comp, 1):
        cell = ws3.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws3.row_dimensions[5].height = 26
    r += 1

    cp_comp_data = [
        ("Menyimak - Berbicara\n(Listening - Speaking)",
         "Berkomunikasi dalam berbagai situasi dan tujuan; menggunakan & merespons pertanyaan terbuka; memulai, mempertahankan, menutup percakapan; memahami ide utama dan detail relevan teks lisan fiksi & nonfiksi.",
         "Memahami alur informasi secara keseluruhan dari teks lisan fiksi & nonfiksi tentang isu terkini atau topik terkait mata pelajaran lain; menggunakan bahasa Inggris untuk mengungkapkan pendapat dan mempertahankan argumen.",
         "Memahami alur cerita/informasi dari teks narasi, eksposisi, dan diskusi tentang isu terkini dan topik lintas mapel secara lisan/multimodal; menggunakan bahasa Inggris untuk mempertahankan argumen kompleks berbasis bukti empiris.",
         "Dari percakapan interaksi sehari-hari -> mempertahankan argumen lintas mapel -> berdebat ilmiah & menganalisis alur narasi kompleks.",
         "Fase E: Percakapan dialog lisan.\nFase F: Presentasi esai & debat argumen.\nFase F TL: Model United Nations (MUN) & diskusi isu kontroversial."),
        ("Membaca - Memirsa\n(Reading - Viewing)",
         "Membaca dan merespons berbagai teks fiksi & nonfiksi tertulis atau multimodal tentang isu terkini; mengidentifikasi tujuan, ide utama, dan informasi rinci; mengevaluasi detail spesifik & menarik kesimpulan.",
         "Mengevaluasi informasi tersurat dan tersirat dari berbagai jenis teks fiksi & nonfiksi tertulis atau teks multimodal tentang isu terkini atau topik terkait mata pelajaran lain.",
         "Mengevaluasi dan merefleksi informasi tersurat dan tersirat dari teks narasi, eksposisi, dan diskusi tentang isu terkini (nasional & global) atau topik terkait mata pelajaran lain secara tertulis/multimodal.",
         "Dari menarik kesimpulan (inferring) -> mengevaluasi makna tersirat lintas disiplin -> merefleksikan bias wacana kritis (Critical Discourse Analysis).",
         "Fase E: Membedah brosur & teks recount.\nFase F: Analisis fakta vs opini teks eksposisi.\nFase F TL: Analisis gaya bahasa cerpen & artikel jurnal ilmiah."),
        ("Menulis - Mempresentasikan\n(Writing - Presenting)",
         "Mengomunikasikan ide & pengalaman melalui paragraf sederhana & terstruktur dalam teks fiksi/nonfiksi multimodal; merencanakan, menulis, merevisi teks; mempertahankan argumen tentang topik harian atau isu terkini.",
         "Mengomunikasikan gagasan & pengalaman dalam teks fiksi/nonfiksi kompleks dengan struktur tepat menggunakan media presentasi cetak/digital; menunjukkan strategi koreksi mandiri; mempertahankan argumen isu terkini.",
         "Mengomunikasikan gagasan dalam teks narasi, eksposisi, diskusi secara detail menggunakan berbagai moda; menulis cerita pendek (creative writing); menulis esai argumen berbasis data dan fakta; membahas isu kontroversial pro-kontra.",
         "Dari paragraf terstruktur (150 kata) -> esai eksposisi kompleks (300-400 kata) -> cerpen sastra & esai akademis berbasis kutipan data riset (500-600 kata).",
         "Fase E: Poster digital & teks deskriptif.\nFase F: Esai analitis & presentasi infografis.\nFase F TL: Publikasi cerpen sastra & esai analisis isu kontroversial.")
    ]

    for item in cp_comp_data:
        el, fe, ff, ftl, kci, imp = item
        ws3.cell(row=r, column=1, value=el).alignment = align_center
        ws3.cell(row=r, column=2, value=fe).alignment = align_top_left
        ws3.cell(row=r, column=3, value=ff).alignment = align_top_left
        ws3.cell(row=r, column=4, value=ftl).alignment = align_top_left
        ws3.cell(row=r, column=5, value=kci).alignment = align_top_left
        ws3.cell(row=r, column=6, value=imp).alignment = align_top_left

        ws3.cell(row=r, column=1).font = font_data_bold
        ws3.cell(row=r, column=2).font = font_data
        ws3.cell(row=r, column=3).font = font_data
        ws3.cell(row=r, column=4).font = font_data
        ws3.cell(row=r, column=5).font = font_data_bold
        ws3.cell(row=r, column=6).font = font_data
        
        for c in range(1, 7):
            ws3.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws3.cell(row=r, column=c).fill = fill_zebra
        ws3.row_dimensions[r].height = 80
        r += 1

    auto_fit_columns(ws3, min_width=10, max_width=45)
    ws3.column_dimensions['A'].width = 20
    ws3.column_dimensions['B'].width = 34
    ws3.column_dimensions['C'].width = 34
    ws3.column_dimensions['D'].width = 36
    ws3.column_dimensions['E'].width = 28
    ws3.column_dimensions['F'].width = 32
    ws3.freeze_panes = 'A5'

    # ==========================================
    # SHEET 4: MATERI ESENSIAL & KONTEKSTUAL
    # ==========================================
    ws4 = wb.create_sheet(title="Materi Esensial & Kontekstual")
    apply_title_banner(
        ws4,
        "PEMETAAN MATERI ESENSIAL & KONTEKSTUALISASI PEMBELAJARAN MENDALAM",
        "Penjabaran Bab C Panduan Puskurjar 2025: Lingkup Materi Esensial, Struktur Teks, & Contoh Projek Otentik",
        max_col=6
    )

    ws4.cell(row=4, column=1, value="PEMETAAN MATERI ESENSIAL BAHASA INGGRIS SMA (BAB C PANDUAN PUSKURJAR 2025)").font = font_section_header
    ws4.merge_cells("A4:F4")
    for c in range(1, 7):
        ws4.cell(row=4, column=c).fill = fill_tbl_header
    ws4.row_dimensions[4].height = 24

    headers_mat = ["Fase Pembelajaran", "Lingkup Materi Esensial", "Fungsi Sosial & Kompetensi yang Dituju", "Fitur Bahasa Kunci (Language Features)", "Model Projek Pembelajaran Otentik", "Koneksi Konteks Nyata Siswa"]
    r = 5
    for i, h in enumerate(headers_mat, 1):
        cell = ws4.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws4.row_dimensions[5].height = 26
    r += 1

    mat_details = [
        ("Fase E (Kelas X)",
         "Opening & Closing Conversation, Congratulating & Complimenting",
         "Membangun interaksi sosial yang ramah, santun, dan saling menghargai baik formal maupun informal.",
         "Greeting formulas (formal/informal), complimenting expressions ('What a great performance!'), gratitude phrases.",
         "Roleplay percakapan lisan di koridor sekolah, video greeting singkat untuk siswa baru.",
         "Membiasakan budaya saling memuliakan dan menyapa ramah di lingkungan sekolah."),
        ("Fase E (Kelas X)",
         "Asking & Giving Opinion, Agreement & Disagreement",
         "Mengungkapkan pandangan pribadi, meminta opini orang lain, serta menyatakan persetujuan atau ketidaksetujuan secara logis.",
         "Opinion verbs (I think, I believe, in my view), Agreement phrases (I agree, exactly), Disagreement phrases (I'm afraid I disagree, I see your point but...).",
         "Projek 'Opinion Wall' (Murid memilih isu sekolah, menulis kartu opini, menempel di papan opini kelas, dan mendiskusikannya).",
         "Musyawarah kelas, pemilihan topik kegiatan pentas seni, pembahasan isu kebersihan sekolah."),
        ("Fase E (Kelas X)",
         "Procedure Text (Food, Manual, Digital Safety)",
         "Memberikan instruksi tahapan membuat atau mengoperasikan sesuatu secara runtut dan efektif.",
         "Imperative verbs (cut, press, avoid, click), Temporal connectives (first, next, then, finally), Adverbials of manner.",
         "Video 'How-To Guide' berdurasi 2 menit, infografis langkah menjaga keamanan sandi di media sosial.",
         "Edukasi literasi digital remaja, resep minuman sehat tradisional lokal."),
        ("Fase F Wajib (XI-XII)",
         "Analytical Exposition Text",
         "Meyakinkan audiens bahwa sebuah isu/fenomena penting dan memerlukan perhatian serius melalui argumen logis.",
         "Thesis statement, Point & Elaboration, Causal conjunctions (consequently, because, as a result), Evaluative language.",
         "Projek 'School Environmental Campaign': Menulis esai analitis tentang bahaya sampah plastik dan mempublikasikannya di blog/majalah sekolah.",
         "Isu pemanasan global, urgensi gizi seimbang bagi konsentrasi belajar murid SMA."),
        ("Fase F Wajib (XI-XII)",
         "Hortatory Exposition Text",
         "Mempengaruhi dan membujuk audiens agar melakukan atau tidak melakukan suatu tindakan spesifik.",
         "Modal verbs of recommendation (should, ought to, must), emotive adjectives, action verbs in recommendation.",
         "Naskah Pidato Persuasif (Speech) yang direkam dalam format audio/video tentang pentingnya tidur cukup dan manajemen stres.",
         "Kampanye anti-bullying di sekolah, gerakan 7 kebiasaan anak Indonesia hebat."),
        ("Fase F Wajib (XI-XII)",
         "Discussion Text (Balancing Perspectives)",
         "Menyajikan dua sudut pandang berbeda (pro dan kontra) secara berimbang sebelum menarik kesimpulan objektif.",
         "Contrasting connectives (however, on the other hand, in contrast, despite), Passive voice, Modality, Impartial tone.",
         "Simulasi 'Oxford Debate' di kelas tentang pro-kontra penggunaan AI dalam tugas sekolah, dilanjutkan menyusun teks diskusi berimbang.",
         "Kecerdasan Buatan (AI) dalam pendidikan, sistem pembelajaran hybrid."),
        ("Fase F Tingkat Lanjut",
         "Literary Narrative & Creative Writing (Cerpen)",
         "Mengeksplorasi nilai estetika bahasa Inggris, dilema moral, dan kedalaman karakter melalui penulisan fiksi kreatif.",
         "Past tenses narrative, figurative language (metaphor, simile, personification), dialogue tags, theme development.",
         "Penerbitan 'Class Anthology of Short Stories' (Buku antologi cerpen bahasa Inggris karya siswa kelas XI/XII peminatan bahasa).",
         "Eksplorasi emosi remaja, pencarian identitas diri, apresiasi karya sastra internasional."),
        ("Fase F Tingkat Lanjut",
         "Academic Discussion & Controversial Issues",
         "Menganalisis diskursus isu kontroversial secara ilmiah dengan argumen yang didukung data dan fakta empiris terpercaya.",
         "Academic Word List (AWL), hedging language (it suggests, data indicates), complex nominal syntax, in-text referencing.",
         "Mini Research Paper: Menyusun laporan analisis pro-kontra isu global (misal: transisi energi hijau atau kecerdasan buatan) lengkap dengan data grafik.",
         "Penyiapan studi lanjut ke perguruan tinggi bereputasi dan kompetensi riset internasional.")
    ]

    for item in mat_details:
        fs, lk, fsoc, fit, prj, ctx = item
        ws4.cell(row=r, column=1, value=fs).alignment = align_center
        ws4.cell(row=r, column=2, value=lk).alignment = align_left
        ws4.cell(row=r, column=3, value=fsoc).alignment = align_left
        ws4.cell(row=r, column=4, value=fit).alignment = align_left
        ws4.cell(row=r, column=5, value=prj).alignment = align_left
        ws4.cell(row=r, column=6, value=ctx).alignment = align_left

        ws4.cell(row=r, column=1).font = font_data_bold
        ws4.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 7):
            ws4.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws4.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws4.cell(row=r, column=c).fill = fill_zebra
        ws4.row_dimensions[r].height = 54
        r += 1

    auto_fit_columns(ws4, min_width=10, max_width=45)
    ws4.column_dimensions['A'].width = 18
    ws4.column_dimensions['B'].width = 28
    ws4.column_dimensions['C'].width = 30
    ws4.column_dimensions['D'].width = 32
    ws4.column_dimensions['E'].width = 34
    ws4.column_dimensions['F'].width = 30
    ws4.freeze_panes = 'A5'

    # ==========================================
    # SHEET 5: SIKLUS GBA & DEEP LEARNING
    # ==========================================
    ws5 = wb.create_sheet(title="Siklus GBA & Deep Learning")
    apply_title_banner(
        ws5,
        "EMPAT TAHAP TEACHING LEARNING CYCLE (TLC) BERBASIS PEMBELAJARAN MENDALAM",
        "Panduan Metodologi Pengajaran Berbasis Teks Bab D Puskurjar 2025: BKoF -> MoT -> JCoT -> ICoT",
        max_col=6
    )

    ws5.cell(row=4, column=1, value="LANGKAH OPERASIONAL TEACHING LEARNING CYCLE (TLC) DI KELAS SMA").font = font_section_header
    ws5.merge_cells("A4:F4")
    for c in range(1, 7):
        ws5.cell(row=4, column=c).fill = fill_tbl_header
    ws5.row_dimensions[4].height = 24

    headers_tlc = ["Tahapan TLC", "Prinsip Pembelajaran Mendalam", "Fokus Aktivitas Guru (Pendidik)", "Aktivitas Otentik Murid SMA", "Teknik Asesmen Formatif yang Cocok", "Durasi & Alokasi Waktu Khas"]
    r = 5
    for i, h in enumerate(headers_tlc, 1):
        cell = ws5.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws5.row_dimensions[5].height = 26
    r += 1

    tlc_data = [
        ("Tahap 1: Building Knowledge of the Field (BKoF)",
         "Bermakna (Meaningful Learning)\n& Pengalaman Memahami",
         "Membangun skema awal murid, mengenalkan konteks sosial-budaya di mana teks digunakan, menayangkan video pemantik, mengaitkan topik dengan pengalaman nyata murid.",
         "Menyimak video/podcast, menjawab pertanyaan pemantik terbuka, berdiskusi tentang relevansi topik dalam kehidupan sehari-hari, eksplorasi kosa kata tematik awal.",
         "Asesmen diagnostik non-tes, tanya jawab interaktif, mind-mapping kosa kata awal.",
         "Pertemuan 1 (Sekitar 2-3 JP tatap muka)."),
        ("Tahap 2: Modelling of the Text (MoT)",
         "Berkesadaran (Mindful Learning)\n& Pengalaman Memahami",
         "Menyajikan teks model utuh; membedah fungsi sosial, generic structure, dan language features secara metakognitif; mendemonstrasikan cara berpikir penulis/pembicara.",
         "Menganalisis teks model bersama guru, memberi label bagian struktur teks (misal: Thesis, Arguments, Reiteration), mengidentifikasi pola kalimat tenses dan kata hubung.",
         "Lembar kerja dekonstruksi teks (text deconstruction worksheet), kuis pemahaman konsep.",
         "Pertemuan 2 (Sekitar 2-3 JP tatap muka)."),
        ("Tahap 3: Joint Construction of the Text (JCoT)",
         "Menggembirakan (Joyful) & Kolaboratif\n& Pengalaman Mengaplikasi",
         "Memberikan perancah (scaffolding); mendampingi penyusunan teks bersama; memfasilitasi kerja kelompok kolaboratif; memberikan umpan balik langsung.",
         "Bekerja dalam kelompok kecil (3-4 orang) menyusun teks serupa dengan topik baru; saling bertukar ide; peer-editing konstruktif antarteman kelompok.",
         "Observasi rubrik kerja sama tim, lembar penyuntingan sejawat (peer-review checklist).",
         "Pertemuan 3 (Sekitar 2-3 JP tatap muka)."),
        ("Tahap 4: Independent Construction of the Text (ICoT)",
         "Kemandirian & Aktualisasi Diri\n& Pengalaman Mengaplikasi",
         "Berperan sebagai fasilitator dan evaluator; memberikan kebebasan murid mengekspresikan gagasan otentiknya; memantau proses penulisan mandiri.",
         "Menulis atau menyajikan teks secara mandiri tanpa bantuan kelompok; melakukan self-editing draf; mempublikasikan karya dalam bentuk tulisan, rekaman, atau presentasi.",
         "Rubrik asesmen sumatif tugas (Writing/Speaking Rubric), lembar evaluasi diri.",
         "Pertemuan 4 (Sekitar 2-3 JP tatap muka)."),
        ("Tahap Terintegrasi: Refleksi Metakognisi",
         "Reflektif (Lifelong Learning)\n& Pengalaman Merefleksi",
         "Memandu murid mengevaluasi proses belajar, mendeteksi tantangan yang dihadapi, dan merumuskan rencana peningkatan di masa depan.",
         "Menjawab pertanyaan refleksi diri (3 minutes reflection), menulis language learning log, mendeskripsikan strategi yang berhasil membantu pemahamannya.",
         "Exit ticket reflektif, jurnal belajar mingguan murid.",
         "10-15 menit di setiap akhir pertemuan KBM.")
    ]

    for item in tlc_data:
        th, prs, act_g, act_m, asm, dur = item
        ws5.cell(row=r, column=1, value=th).alignment = align_left
        ws5.cell(row=r, column=2, value=prs).alignment = align_center
        ws5.cell(row=r, column=3, value=act_g).alignment = align_left
        ws5.cell(row=r, column=4, value=act_m).alignment = align_left
        ws5.cell(row=r, column=5, value=asm).alignment = align_left
        ws5.cell(row=r, column=6, value=dur).alignment = align_center

        ws5.cell(row=r, column=1).font = font_data_bold
        ws5.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 7):
            ws5.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws5.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws5.cell(row=r, column=c).fill = fill_zebra
        ws5.row_dimensions[r].height = 56
        r += 1

    auto_fit_columns(ws5, min_width=10, max_width=45)
    ws5.column_dimensions['A'].width = 24
    ws5.column_dimensions['B'].width = 24
    ws5.column_dimensions['C'].width = 34
    ws5.column_dimensions['D'].width = 34
    ws5.column_dimensions['E'].width = 32
    ws5.column_dimensions['F'].width = 20
    ws5.freeze_panes = 'A5'

    # ==========================================
    # SHEET 6: DESAIN MODUL AJAR SMA RESMI
    # ==========================================
    ws6 = wb.create_sheet(title="Desain Modul Ajar SMA Resmi")
    apply_title_banner(
        ws6,
        "BEDAH CONTOH PERENCANAAN PEMBELAJARAN (MODUL AJAR) RESMI PUSKURJAR 2025",
        "Rangkuman Contoh Modul Ajar Fase F & Fase F Tingkat Lanjut (Halaman 134-166 Panduan Resmi)",
        max_col=6
    )

    ws6.cell(row=4, column=1, value="CONTOH MODUL AJAR RESMI: FASE F TINGKAT LANJUT (DISCUSSION TEXT - 20 JP)").font = font_section_header
    ws6.merge_cells("A4:F4")
    for c in range(1, 7):
        ws6.cell(row=4, column=c).fill = fill_tbl_header
    ws6.row_dimensions[4].height = 24

    modul_info = [
        ("Judul Pembelajaran", "Pros and Cons of Online Games: Exploring Balanced Arguments (Discussion Text)"),
        ("Fase & Alokasi Waktu", "Fase F Tingkat Lanjut (Kelas XI/XII SMA) | Alokasi: 20 JP (4 Pertemuan Tatap Muka x 5 JP)"),
        ("Tujuan Pembelajaran (TP)", "Menulis pendapat atau pandangan tentang suatu topik atau isu dengan mengungkapkan argumen yang didukung data dan fakta (Discussion Text) secara tertulis atau multimodal."),
        ("Dimensi Profil Lulusan", "Penalaran Kritis (Critical Thinking), Kolaborasi (Collaboration), dan Komunikasi (Communication)."),
        ("Identifikasi Kesiapan Murid", "Asesmen diagnostik awal: Mengecek pengetahuan awal murid tentang ungkapan agreement & disagreement serta pemahaman isu game online.")
    ]

    r = 5
    for item, val in modul_info:
        ws6.cell(row=r, column=1, value=item).font = font_data_bold
        ws6.cell(row=r, column=1).fill = fill_group
        ws6.cell(row=r, column=1).alignment = align_left
        ws6.cell(row=r, column=1).border = cell_border
        
        ws6.cell(row=r, column=2, value=val).font = font_data
        ws6.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws6.cell(row=r, column=2).alignment = align_left
        for col in range(2, 7):
            ws6.cell(row=r, column=col).border = cell_border
        ws6.row_dimensions[r].height = 24
        r += 1

    r += 1
    # Rincian Langkah Pertemuan Modul Ajar Resmi
    ws6.cell(row=r, column=1, value="RINCIAN AKTIVITAS PERTEMUAN 1 S.D. 4 (MODEL GBA & DEEP LEARNING PUSKURJAR)").font = font_section_header
    ws6.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws6.cell(row=r, column=c).fill = fill_card_header
    ws6.row_dimensions[r].height = 24
    r += 1

    headers_modul_steps = ["Pertemuan", "Fase Siklus TLC", "Fokus Kegiatan Pembelajaran", "Bahan & Media Digital yang Digunakan", "Langkah Asesmen Formatif", "Tindak Lanjut Diferensiasi"]
    for i, h in enumerate(headers_modul_steps, 1):
        cell = ws6.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = cell_border
    ws6.row_dimensions[r].height = 26
    r += 1

    modul_steps_data = [
        ("Pertemuan 1 (5 JP)",
         "Building Knowledge of the Field (BKoF)",
         "1. Mengamati gambar infografis game online dan mendiskusikan pertanyaan pemantik.\n2. Melakukan asesmen diagnostik awal kesiapan kosa kata dan opini.\n3. Menonton video dokumenter dampak video game pada remaja.\n4. Memetakan pro dan kontra dalam grafik T-Chart.",
         "Infografis Game Quitters, video YouTube edukatif, lembar T-Chart (Pros vs Cons).",
         "Tanya jawab lisan spontan; penilaian diagnostik pemahaman awal murid.",
         "Murid mahir langsung mengelaborasi argumen tingkat tinggi; murid menengah dipandu frasa pembuka opini."),
        ("Pertemuan 2 (5 JP)",
         "Modelling of the Text (MoT)",
         "1. Membaca teks model 'The Pros and Cons of Playing Online Games'.\n2. Menganalisis Generic Structure: Issue, Arguments For, Arguments Against, Conclusion.\n3. Membedah Language Features: Contrastive Conjunctions (However, Although, In contrast) dan Adjectives.",
         "Lembar kerja dekonstruksi teks model, teks artikel rujukan, presentasi interaktif.",
         "Kuis identifikasi konjungsi pertentangan dan fungsi sosial teks diskusi.",
         "Pendidik memberikan bimbingan khusus pada murid yang masih kesulitan membedakan argumen pro vs kontra."),
        ("Pertemuan 3 (5 JP)",
         "Joint Construction of the Text (JCoT)",
         "1. Memilih topik kontroversial baru (misal: AI in Schools, Fast Food Ban, Electric Vehicles).\n2. Kerja kelompok menyusun draf teks diskusi bersama.\n3. Melakukan peer-review antarkelompok menggunakan rubrik panduan.",
         "Papan tulis kolaboratif/Padlet, lembar rubrik penyuntingan sejawat, referensi data.",
         "Penilaian draf tulisan kelompok dan keterlibatan aktif kolaborasi murid.",
         "Guru melakukan intervensi bimbingan intensif pada kelompok yang kesulitan menyusun transisi kohesif."),
        ("Pertemuan 4 (5 JP)",
         "Independent Construction (ICoT) & Reflection",
         "1. Murid menulis draf teks diskusi secara mandiri berdasarkan pilihan isu personal.\n2. Mengedit draf mandiri (self-editing).\n3. Mempublikasikan esai final (Canva poster/blog).\n4. Melakukan refleksi 3-minutes writing tentang proses berpikir.",
         "Platform publikasi digital sekolah, lembar rubrik evaluasi mandiri (self-rubric).",
         "Asesmen Sumatif tugas penulisan teks diskusi menggunakan rubrik standar Puskurjar.",
         "Karya terbaik dipamerkan di papan majalah dinding digital sekolah; murid belum tuntas mendapat bimbingan remidi.")
    ]

    for item in modul_steps_data:
        pt, fs, fks, bhn, asm, tnd = item
        ws6.cell(row=r, column=1, value=pt).alignment = align_center
        ws6.cell(row=r, column=2, value=fs).alignment = align_left
        ws6.cell(row=r, column=3, value=fks).alignment = align_left
        ws6.cell(row=r, column=4, value=bhn).alignment = align_left
        ws6.cell(row=r, column=5, value=asm).alignment = align_left
        ws6.cell(row=r, column=6, value=tnd).alignment = align_left

        ws6.cell(row=r, column=1).font = font_data_bold
        ws6.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 7):
            ws6.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws6.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws6.cell(row=r, column=c).fill = fill_zebra
        ws6.row_dimensions[r].height = 58
        r += 1

    auto_fit_columns(ws6, min_width=10, max_width=45)
    ws6.column_dimensions['A'].width = 18
    ws6.column_dimensions['B'].width = 24
    ws6.column_dimensions['C'].width = 38
    ws6.column_dimensions['D'].width = 30
    ws6.column_dimensions['E'].width = 30
    ws6.column_dimensions['F'].width = 32
    ws6.freeze_panes = 'A5'

    # ==========================================
    # SHEET 7: INSTRUMEN & RUBRIK ASESMEN
    # ==========================================
    ws7 = wb.create_sheet(title="Instrumen & Rubrik Asesmen")
    apply_title_banner(
        ws7,
        "RUBRIK ASESMEN RESMI PUSKURJAR 2025 UNTUK KETERAMPILAN BAHASA INGGRIS",
        "Rujukan Rubrik Speaking & Writing Baku Edisi Revisi 2025 (Halaman 165-166 Buku Panduan)",
        max_col=6
    )

    ws7.cell(row=4, column=1, value="1. RUBRIK PENILAIAN MENULIS RESMI (WRITING RUBRIC - 4 TINGKAT PUSKURJAR 2025)").font = font_section_header
    ws7.merge_cells("A4:F4")
    for c in range(1, 7):
        ws7.cell(row=4, column=c).fill = fill_tbl_header
    ws7.row_dimensions[4].height = 24

    headers_wrt_std = ["Kriteria Penilaian", "Category 4: Above Standards (Sangat Mahir)", "Category 3: Meets Standards (Mahir)", "Category 2: Approaching Standards (Cukup)", "Category 1: Below Standards (Perlu Bimbingan)", "Bobot Penilaian"]
    r = 5
    for i, h in enumerate(headers_wrt_std, 1):
        cell = ws7.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws7.row_dimensions[5].height = 24
    r += 1

    wrt_std_data = [
        ("Vocabulary",
         "Excellent vocabulary choice. Errors in vocabulary choice are absent or extremely rare; rich academic and contextual terms used.",
         "Good in vocabulary choice. Errors in vocabulary choice are few and do not interfere with understanding.",
         "Errors in vocabulary choice are frequent and sometimes they interfere with clear understanding.",
         "Many serious errors in vocabulary choice that severely interfere with understanding and communication.",
         "25%"),
        ("Grammar & Syntax",
         "Exceptional grammar control. Complex sentences used accurately with correct clauses and appropriate hedging language.",
         "Good in grammar. Errors in grammar choice are few and do not interfere with understanding.",
         "Errors in grammar choice are noticeable and sometimes they interfere with sentence coherence.",
         "Pervasive errors in grammar choice that severely disrupt understanding.",
         "25%"),
        ("Structure (Generic Structure)",
         "Structure of the text is completely well-developed and highly logical; all generic elements fully present with smooth transitions.",
         "Structure of the text is clear and meets standard genre requirements; main elements well-organized.",
         "Structure of the text is incomplete or partially disordered; arguments not well-connected.",
         "Structure of the text is disorganized and lacks essential generic components.",
         "30%"),
        ("Mechanics (Spelling & Punctuation)",
         "Flawless spelling and punctuation; formatting adheres to professional publication standards.",
         "Minor errors in spelling and punctuation that do not impede reader comprehension.",
         "Frequent errors in spelling and punctuation that cause reader distraction.",
         "Severe errors in spelling and punctuation that make the text difficult to decipher.",
         "20%")
    ]

    for item in wrt_std_data:
        kr, cat4, cat3, cat2, cat1, bob = item
        ws7.cell(row=r, column=1, value=kr).alignment = align_left
        ws7.cell(row=r, column=2, value=cat4).alignment = align_left
        ws7.cell(row=r, column=3, value=cat3).alignment = align_left
        ws7.cell(row=r, column=4, value=cat2).alignment = align_left
        ws7.cell(row=r, column=5, value=cat1).alignment = align_left
        ws7.cell(row=r, column=6, value=bob).alignment = align_center

        ws7.cell(row=r, column=1).font = font_data_bold
        for c in range(2, 6):
            ws7.cell(row=r, column=c).font = font_data
        ws7.cell(row=r, column=6).font = font_data_bold
        
        for c in range(1, 7):
            ws7.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws7.cell(row=r, column=c).fill = fill_zebra
        ws7.row_dimensions[r].height = 45
        r += 1

    r += 1
    # Rubrik Speaking Resmi
    ws7.cell(row=r, column=1, value="2. RUBRIK PENILAIAN FORMATIF BERBICARA (SPEAKING RUBRIC PUSKURJAR 2025)").font = font_section_header
    ws7.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws7.cell(row=r, column=c).fill = fill_tbl_header
    ws7.row_dimensions[r].height = 24
    r += 1

    headers_spk_std = ["Kriteria Penilaian", "Excellent (Sangat Baik)", "Good (Baik)", "Needs Improvements (Perlu Bimbingan)", "Strategi Intervensi Guru", "Panduan Asesmen Lisan"]
    for i, h in enumerate(headers_spk_std, 1):
        cell = ws7.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws7.row_dimensions[r].height = 24
    r += 1

    spk_std_data = [
        ("Pronunciation",
         "Pronunciation is accurate, with correct inflections, number of syllables, and natural nuance of spoken English.",
         "Pronunciation is satisfactory; however words sometimes have incorrect inflections or are sometimes hard to understand.",
         "Pronunciation is frequently inaccurate; severe mispronunciation that impedes basic communication.",
         "Latihan pengucapan vokal/konsonan kunci lewat phonetic drills dan shadowing audio penutur asli.",
         "Dinilai saat murid berdialog spontan dan presentasi lisan."),
        ("Fluency",
         "Speaks fluently with spontaneous flow; minimal unnatural hesitations; thoughts expressed with natural speech rhythm.",
         "Speaks relatively smoothly with occasional pauses to search for words; message remains understandable.",
         "Speaks with heavy hesitation, stammering, and long awkward silences; communication breaks down.",
         "Membiasakan interaksi transaksional harian di kelas dalam suasana aman tanpa celaan aksen.",
         "Fokus pada kelancaran alur pikiran, bukan menghukum kesalahan kecil."),
        ("Interactive Communication",
         "Actively initiates conversation, responds appropriately to open-ended questions, and sustains debate turn-taking.",
         "Responds adequately to questions; maintains conversation with moderate encouragement.",
         "Passive participant; gives only one-word answers; unable to sustain interaction without direct prompts.",
         "Menggunakan kartu panduan frasa pembuka percakapan (conversation prompt cards).",
         "Menilai kemampuan murid berkolaborasi dan mendengarkan aktif lawan bicara.")
    ]

    for item in spk_std_data:
        kr, exc, gd, nd, strg, pnd = item
        ws7.cell(row=r, column=1, value=kr).alignment = align_left
        ws7.cell(row=r, column=2, value=exc).alignment = align_left
        ws7.cell(row=r, column=3, value=gd).alignment = align_left
        ws7.cell(row=r, column=4, value=nd).alignment = align_left
        ws7.cell(row=r, column=5, value=strg).alignment = align_left
        ws7.cell(row=r, column=6, value=pnd).alignment = align_left

        ws7.cell(row=r, column=1).font = font_data_bold
        for c in range(2, 7):
            ws7.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws7.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws7.cell(row=r, column=c).fill = fill_zebra
        ws7.row_dimensions[r].height = 45
        r += 1

    auto_fit_columns(ws7, min_width=10, max_width=45)
    ws7.column_dimensions['A'].width = 22
    ws7.column_dimensions['B'].width = 30
    ws7.column_dimensions['C'].width = 30
    ws7.column_dimensions['D'].width = 30
    ws7.column_dimensions['E'].width = 32
    ws7.column_dimensions['F'].width = 26
    ws7.freeze_panes = 'A5'

    # ==========================================
    # SHEET 8: GLOSARIUM ISTILAH PENGAJARAN
    # ==========================================
    ws8 = wb.create_sheet(title="Glosarium Istilah Pengajaran")
    apply_title_banner(
        ws8,
        "GLOSARIUM ISTILAH KUNCI PENGAJARAN BAHASA INGGRIS KURIKULUM NASIONAL",
        "Daftar Definisi Istilah Pedagogis dan Linguistik Buku Panduan Puskurjar Edisi Revisi 2025 (Halaman 167)",
        max_col=5
    )

    ws8.cell(row=4, column=1, value="DAFTAR ISTILAH RESMI & DEFINISI OPERASIONAL PANDUAN PUSKURJAR").font = font_section_header
    ws8.merge_cells("A4:E4")
    for c in range(1, 6):
        ws8.cell(row=4, column=c).fill = fill_tbl_header
    ws8.row_dimensions[4].height = 24

    headers_glo = ["No", "Istilah Kunci", "Kepanjangan / Padanan", "Definisi Resmi Buku Panduan (Halaman 167)", "Relevansi Praktis bagi Guru SMA"]
    r = 5
    for i, h in enumerate(headers_glo, 1):
        cell = ws8.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws8.row_dimensions[5].height = 26
    r += 1

    glossary_data = [
        (1, "CEFR", "Common European Framework of Reference for Languages", "Suatu penentuan berbagai tingkat kecakapan bahasa yang disusun oleh Dewan Eropa (A1, A2, B1, B2, C1, C2).", "Standar internasional pemetaan target kemahiran bahasa Inggris lulusan SMA (Target Fase E = B1, Fase F = B1-B2)."),
        (2, "Genre-Based Approach (GBA)", "Pendekatan Berbasis Teks (Genre)", "Pendekatan pembelajaran bahasa yang menekankan pada penguasaan teks utuh bermakna sesuai fungsi sosialnya melalui tahapan Teaching Learning Cycle.", "Metodologi wajib dalam kurikulum Bahasa Inggris Indonesia sejak 2004 hingga kurikulum saat ini."),
        (3, "BKoF", "Building Knowledge of the Field", "Tahap pertama TLC untuk membangun pengetahuan latar belakang murid mengenai topik dan konteks sosial budaya teks.", "Membantu murid memahami konteks nyata teks sebelum belajar menyusun kalimat."),
        (4, "MoT", "Modelling of the Text", "Tahap kedua TLC di mana pendidik dan murid mendekonstruksi struktur generik dan unsur kebahasaan dari contoh teks model.", "Murid secara sadar (mindful) membedah pola bahasa dan gramatika teks."),
        (5, "JCoT", "Joint Construction of the Text", "Tahap ketiga TLC di mana pendidik dan murid secara kolaboratif menyusun teks baru dengan topik serupa.", "Memberikan perancah (scaffolding) yang aman sehingga murid tidak langsung dilepas sendirian."),
        (6, "ICoT", "Independent Construction of the Text", "Tahap akhir TLC di mana murid secara mandiri menyusun teks lisan atau tulis untuk tujuan komunikatif otentik.", "Muara kemandirian dan kreativitas murid dalam mengekspresikan suaranya."),
        (7, "Multimodal Text", "Teks Multimoda", "Teks yang menggabungkan dua atau lebih sistem semiotik (linguistik/teks, visual/gambar, audio/suara, gestur, dan spasial).", "Siswa SMA tidak hanya membaca teks cetak, tetapi juga menganalisis infografis, video, podcast, dan halaman web."),
        (8, "Familiar Texts", "Teks Akrab / Teks Dikenal", "Teks yang memuat kosakata, struktur, dan topik yang sudah dikenal baik oleh murid dalam kehidupan sehari-hari.", "Pintu masuk pembelajaran pada awal fase agar murid merasa nyaman dan percaya diri."),
        (9, "Interactive Text", "Teks Interaktif", "Teks yang melibatkan partisipasi aktif dan respons langsung dari audiens atau pembaca.", "Dialog lisan, forum diskusi daring, media sosial, dan presentasi interaktif."),
        (10, "Invented Spelling", "Ejaan Rekaan", "Usaha murid dalam menuliskan ejaan dari kata yang didengar atau dilihat berdasarkan pemahamannya sebelum menguasai ejaan baku.", "Tahapan wajar perkembangan fonologis murid dalam belajar menulis bahasa Inggris."),
        (11, "Miskonsepsi Pengajaran Teks", "Misconception of Text-Based Teaching", "Anggapan keliru bahwa pengajaran teks dapat diselesaikan dalam satu pertemuan tatap muka tanpa melalui tahapan siklus belajar.", "Pendidik harus merancang pembelajaran teks dalam beberapa pertemuan terencana (4 tahap TLC)."),
        (12, "Hedging Language", "Bahasa Pembatas / Kehati-hatian", "Penggunaan kata-kata untuk menyatakan kehati-hatian atau probabilitas ilmiah (misal: suggests, indicates, tends to, appears to).", "Fitur bahasa kunci dalam esai akademis dan diskusi kontroversial Fase F Tingkat Lanjut.")
    ]

    for item in glossary_data:
        no, term, acr, defn, rel = item
        ws8.cell(row=r, column=1, value=no).alignment = align_center
        ws8.cell(row=r, column=2, value=term).alignment = align_left
        ws8.cell(row=r, column=3, value=acr).alignment = align_left
        ws8.cell(row=r, column=4, value=defn).alignment = align_left
        ws8.cell(row=r, column=5, value=rel).alignment = align_left

        ws8.cell(row=r, column=1).font = font_data_bold
        ws8.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 6):
            ws8.cell(row=r, column=c).font = font_data
            
        for c in range(1, 6):
            ws8.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws8.cell(row=r, column=c).fill = fill_zebra
        ws8.row_dimensions[r].height = 42
        r += 1

    auto_fit_columns(ws8, min_width=10, max_width=45)
    ws8.column_dimensions['A'].width = 5
    ws8.column_dimensions['B'].width = 24
    ws8.column_dimensions['C'].width = 30
    ws8.column_dimensions['D'].width = 38
    ws8.column_dimensions['E'].width = 34
    ws8.freeze_panes = 'A5'

    # Save workbook
    wb.save(output_path)
    print(f"Panduan Mapel Bahasa Inggris Workbook successfully created and saved to: {output_path}")

if __name__ == "__main__":
    out_file = r"C:\Users\KOMPUTERKU2\Documents\English SMA\Peraturan\Panduan_Mata_Pelajaran_Bahasa_Inggris_Revisi_2025.xlsx"
    create_panduan_bing_excel(out_file)
