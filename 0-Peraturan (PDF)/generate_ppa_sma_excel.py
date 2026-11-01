# -*- coding: utf-8 -*-
"""
Script to generate a comprehensive, highly styled Excel workbook summarizing
PANDUAN PEMBELAJARAN DAN ASESMEN (PPA) EDISI REVISI TAHUN 2025
specifically customized for JENJANG SMA (Sekolah Menengah Atas).
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_ppa_sma_excel(output_path):
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
    font_note = Font(name=font_family, size=9, italic=True, color="475569")
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
    fill_success = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")        # Soft Emerald Green

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
    # SHEET 1: RINGKASAN PPA SMA 2025
    # ==========================================
    ws1 = wb.create_sheet(title="Ringkasan PPA SMA 2025")
    apply_title_banner(
        ws1,
        "PANDUAN PEMBELAJARAN DAN ASESMEN (PPA) JENJANG SMA - EDISI REVISI 2025",
        "Rujukan Resmi Badan Standar, Kurikulum, dan Asesmen Pendidikan (BSKAP) Kemendikdasmen RI",
        max_col=6
    )

    ws1.cell(row=4, column=1, value="1. IDENTITAS DOKUMEN & LANDASAN KEBIJAKAN RESMI").font = font_section_header
    ws1.merge_cells("A4:F4")
    for c in range(1, 7):
        ws1.cell(row=4, column=c).fill = fill_tbl_header
    ws1.row_dimensions[4].height = 24

    doc_info = [
        ("Nama Dokumen", "Panduan Pembelajaran dan Asesmen Pendidikan Anak Usia Dini, Pendidikan Dasar, dan Pendidikan Menengah (Edisi Revisi Ke-3, 2025)"),
        ("Instansi Penerbit", "Badan Standar, Kurikulum, dan Asesmen Pendidikan (BSKAP) Kementerian Pendidikan Dasar dan Menengah RI"),
        ("Fokus Dokumen", "Pedoman komprehensif bagi pendidik, kepala sekolah, dan pengawas dalam merencanakan, melaksanakan, mengolah, dan melaporkan pembelajaran serta asesmen."),
        ("Fokus Jenjang SMA", "Implementasi pada Fase E (Kelas X SMA) dan Fase F (Kelas XI dan XII SMA) baik mata pelajaran wajib maupun mata pelajaran pilihan."),
        ("Regulasi Payung Terkait", "Permendikdasmen No. 10/2025 (SKL), Permendikdasmen No. 12/2025 (Standar Isi), Permendikbudristek No. 16/2022 (Standar Proses), Permendikbudristek No. 21/2022 (Standar Penilaian), dan Permendikdasmen No. 13/2025 (Struktur Kurikulum).")
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
        ws1.row_dimensions[r].height = 24
        r += 1

    r += 1
    # 5 Prinsip Asesmen Edisi Revisi 2025
    ws1.cell(row=r, column=1, value="2. LIMA (5) PRINSIP ASESMEN PEMBELAJARAN EDISI REVISI 2025").font = font_section_header
    ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws1.cell(row=r, column=c).fill = fill_card_header
    ws1.row_dimensions[r].height = 24
    r += 1

    headers_prinsip = ["No", "Prinsip Asesmen (PPA 2025)", "Makna & Konsep Esensial", "Penerapan Praktis di SMA", "Hal yang Harus Dihindari Guru", "Dampak pada Murid SMA"]
    for i, h in enumerate(headers_prinsip, 1):
        cell = ws1.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = cell_border
    ws1.row_dimensions[r].height = 26
    r += 1

    prinsip_data = [
        (1, "Bagian Terpadu dari Pembelajaran", "Asesmen memfasilitasi pembelajaran dan menyediakan informasi holistik sebagai umpan balik pendidik, murid, dan orang tua.", "Mengintegrasikan asesmen diagnostik di awal materi, serta asesmen formatif sepanjang proses belajar untuk memandu scaffolding.", "Melakukan asesmen hanya di akhir bab/semester semata-mata untuk pengisian angka rapor.", "Murid memahami posisi belajarnya dan mendapatkan bantuan belajar yang tepat sasaran."),
        (2, "Dirancang Sesuai Fungsi Asesmen", "Asesmen formatif mendiagnosis dan memberi umpan balik proses; asesmen sumatif mengonfirmasi capaian hasil belajar.", "Formatif tidak digabung/dirata-rata ke dalam nilai rapor. Nilai akhir rapor murni diambil dari asesmen sumatif lingkup materi.", "Menghitung nilai tugas harian (formatif) ke dalam nilai rapor sehingga murid takut salah saat berlatih.", "Murid berani mencoba, bereksplorasi, dan belajar dari kesalahan tanpa takut nilai rapor jatuh."),
        (3, "Adil, Proporsional, Valid & Reliabel", "Menjelaskan kemajuan belajar secara objektif, menentukan keputusan tindak lanjut yang tepat dan terpercaya.", "Menyediakan instrumen asesmen yang beragam (kinerja, proyek, portofolio, tes tulis) sesuai karakteristik materi dan target CP.", "Menggunakan satu metode tes pilihan ganda untuk mengukur seluruh kompetensi penalaran kritis.", "Penilaian mencerminkan kompetensi riil murid secara komprehensif dan tidak diskriminatif."),
        (4, "Laporan Sederhana & Informatif", "Memberikan gambaran jelas tentang pencapaian karakter dan kompetensi murid serta strategi tindak lanjut yang konkret.", "Menuliskan deskripsi rapor yang memuat 1-2 capaian tertinggi yang dikuasai dan kompetensi yang perlu bimbingan lanjut.", "Menulis deskripsi rapor yang panjang berbelit-belit atau hanya berupa angka tanpa narasi bermakna.", "Orang tua dan murid memahami secara spesifik kekuatan akademis serta area yang harus ditingkatkan."),
        (5, "Bahan Refleksi Peningkatan Mutu", "Hasil asesmen digunakan murid, pendidik, dan pimpinan sekolah untuk merefleksikan dan memperbaiki kualitas KBM.", "Guru menganalisis butir soal atau tugas yang gagal dipahami mayoritas murid untuk mengevaluasi strategi pengajarannya.", "Menyalahkan murid saat hasil asesmen rendah tanpa mengoreksi metode pembelajaran guru.", "Terbangun kultur mutu reflektif di mana guru terus mengasah strategi pedagogisnya.")
    ]

    for item in prinsip_data:
        no, p_name, p_makna, p_app, p_avoid, p_impact = item
        ws1.cell(row=r, column=1, value=no).alignment = align_center
        ws1.cell(row=r, column=2, value=p_name).alignment = align_left
        ws1.cell(row=r, column=3, value=p_makna).alignment = align_left
        ws1.cell(row=r, column=4, value=p_app).alignment = align_left
        ws1.cell(row=r, column=5, value=p_avoid).alignment = align_left
        ws1.cell(row=r, column=6, value=p_impact).alignment = align_left

        ws1.cell(row=r, column=1).font = font_data_bold
        ws1.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 7):
            ws1.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws1.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws1.cell(row=r, column=c).fill = fill_zebra
        ws1.row_dimensions[r].height = 54
        r += 1

    auto_fit_columns(ws1, min_width=10, max_width=45)
    ws1.column_dimensions['A'].width = 5
    ws1.column_dimensions['B'].width = 24
    ws1.column_dimensions['C'].width = 34
    ws1.column_dimensions['D'].width = 34
    ws1.column_dimensions['E'].width = 32
    ws1.column_dimensions['F'].width = 32
    ws1.freeze_panes = 'A5'

    # ==========================================
    # SHEET 2: ANALISIS CP, TP & ATP SMA
    # ==========================================
    ws2 = wb.create_sheet(title="Analisis CP, TP & ATP SMA")
    apply_title_banner(
        ws2,
        "PANDUAN PERUMUSAN TUJUAN PEMBELAJARAN (TP) DAN ALURNYA (ATP) DI SMA",
        "Penjabaran Operasional Bab 3 PPA 2025: Analisis Capaian Pembelajaran Fase E & Fase F",
        max_col=6
    )

    ws2.cell(row=4, column=1, value="1. TIGA (3) ALTERNATIF MERUMUSKAN TUJUAN PEMBELAJARAN (TP) DI SMA").font = font_section_header
    ws2.merge_cells("A4:F4")
    for c in range(1, 7):
        ws2.cell(row=4, column=c).fill = fill_tbl_header
    ws2.row_dimensions[4].height = 24

    headers_tp_alt = ["Alternatif", "Nama Pendekatan Perumusan TP", "Langkah Kerja Guru SMA", "Karakteristik & Kelebihan", "Kapan Tepat Digunakan di SMA?", "Contoh Implementasi Mata Pelajaran"]
    r = 5
    for i, h in enumerate(headers_tp_alt, 1):
        cell = ws2.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws2.row_dimensions[5].height = 26
    r += 1

    tp_alternatives = [
        ("Alternatif 1", "Merumuskan Langsung dari Kalimat CP", "Pendidik langsung memecah kalimat Capaian Pembelajaran menjadi beberapa Tujuan Pembelajaran operasional tanpa analisis kata kerja terpisah.", "Praktis, cepat, dan menjaga keselarasan langsung dengan rumusan CP pemerintah.", "Sangat cocok untuk guru yang materi ajar dalam CP-nya sudah terdeskripsikan secara terinci dan eksplisit.", "Bahasa Indonesia atau Bahasa Inggris Fase E saat menyusun TP dari elemen membaca/menyimak."),
        ("Alternatif 2", "Menganalisis Kompetensi & Lingkup Materi", "Pendidik membedah kalimat CP menjadi dua komponen utama: (1) Kompetensi (kata kerja/kemampuan yang dituju) dan (2) Lingkup Materi (konten keilmuan).", "Sistematis, terukur, memudahkan penentuan level kognitif taksonomi Bloom, dan materi ajar tertata rapi.", "Rekomendasi utama untuk mapel sains (Fisika, Kimia, Biologi) dan Matematika yang memiliki struktur konsep bertingkat.", "Biologi Fase E: Kompetensi = 'menganalisis', Materi = 'perubahan lingkungan dan bioteknologi konvensional'."),
        ("Alternatif 3", "Merumuskan TP Lintas Elemen CP", "Pendidik menggabungkan beberapa kompetensi dari elemen CP yang berbeda ke dalam satu rangkaian Tujuan Pembelajaran terpadu.", "Holistik, menghindari fragmentasi kompetensi, mendukung pembelajaran berbasis proyek/tema kolaboratif.", "Sangat ideal untuk mapel seni, bahasa (memadukan menyimak, berbicara, membaca, menulis), dan Informatika.", "Bahasa Inggris Fase F: memadukan elemen Menyimak (memahami podcast) dengan Berbicara (debat mempertahankan argumen).")
    ]

    for item in tp_alternatives:
        alt, name, step, char, when, ex = item
        ws2.cell(row=r, column=1, value=alt).alignment = align_center
        ws2.cell(row=r, column=2, value=name).alignment = align_left
        ws2.cell(row=r, column=3, value=step).alignment = align_left
        ws2.cell(row=r, column=4, value=char).alignment = align_left
        ws2.cell(row=r, column=5, value=when).alignment = align_left
        ws2.cell(row=r, column=6, value=ex).alignment = align_left

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
    # 6 Teknik Mengurutkan ATP
    ws2.cell(row=r, column=1, value="2. ENAM (6) CARA MENGURUTKAN ALUR TUJUAN PEMBELAJARAN (ATP) DI SMA").font = font_section_header
    ws2.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws2.cell(row=r, column=c).fill = fill_tbl_header
    ws2.row_dimensions[r].height = 24
    r += 1

    headers_atp_sort = ["Metode Urutan", "Prinsip Penyusunan Alur", "Mekanisme Progresi Belajar", "Kesesuaian Karakteristik Mapel SMA", "Contoh Progresi Materi di Kelas X / XI SMA", "Manfaat bagi Daya Serap Murid"]
    for i, h in enumerate(headers_atp_sort, 1):
        cell = ws2.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws2.row_dimensions[r].height = 26
    r += 1

    atp_sort_data = [
        ("Konkret ke Abstrak", "Memulai dari fenomena nyata yang dapat diamati panca indera menuju konsep teoritis yang abstrak.", "Fenomena visual/nyata -> Konseptualisasi -> Generalisasi teori formal.", "Fisika, Kimia, Geografi, Sosiologi.", "Fisika Kelas X: Mengamati gerak benda jatuh bebas di sekitar kelas baru menurunkan rumus matematis gerak lurus berubah beraturan (GLBB).", "Menghindari kejutan kognitif akibat rumus rumit tanpa pijakan empiris."),
        ("Deduktif (Umum ke Khusus)", "Menjelaskan konsep umum/definisi induk terlebih dahulu sebelum menguraikan spesifikasi bagian-bagian detailnya.", "Konsep makro -> Pembagian kategori -> Studi kasus mikro spesifik.", "Ekonomi, Biologi, Sejarah.", "Ekonomi Kelas XI: Menjelaskan konsep makro pertumbuhan ekonomi sebelum membedah komponen pendapatan per kapita dan pengangguran.", "Murid memiliki kerangka berpikir (framework) komprehensif sejak awal."),
        ("Mudah ke Lebih Sulit", "Mengurutkan materi dari kompetensi dasar berbeban kognitif ringan menuju kompetensi kompleks berbeban tinggi.", "C1-C2 (Mengingat/Memahami) -> C3-C4 (Menerapkan/Menganalisis) -> C5-C6 (Mengevaluasi/Mengkreasi).", "Matematika, Informatika, Bahasa Asing.", "Matematika Tingkat Lanjut: Menyelesaikan sistem persamaan linear dua variabel sebelum melangkah ke matriks dan determinan ordo tinggi.", "Membangun rasa percaya diri murid dan mencegah kecemasan belajar."),
        ("Hierarkis (Prasyarat)", "Mengurutkan kompetensi di mana sebuah tujuan pembelajaran baru bisa dicapai jika tujuan sebelumnya telah tuntas dikuasai.", "Kompetensi Fondasi (Prasyarat) -> Kompetensi Lanjutan -> Kompetensi Mahir.", "Matematika, Kimia Struktur Atom, Fisika Mekanika.", "Kimia Kelas X: Memahami konfigurasi elektron dan tabel periodik unsur sebelum mempelajari ikatan kimia kovalen/ionik.", "Mencegah terjadinya 'learning loss' dan gap pemahaman bertumpuk."),
        ("Prosedural", "Mengajarkan tahapan kerja standar yang harus ditempuh langkah demi langkah secara berurutan.", "Langkah 1 (Persiapan) -> Langkah 2 (Eksperimen) -> Langkah 3 (Analisis Data) -> Langkah 4 (Pelaporan).", "PJOK, Prakarya & Kewirausahaan (PKWU), Praktikum Sains.", "PKWU Kelas XI: Tahap riset pasar produk kerajinan -> Perancangan prototipe -> Proses produksi -> Pengemasan -> Pemasaran digital.", "Murid memahami sistematika kerja profesional dunia industri."),
        ("Scaffolding", "Pemberian bimbingan intensif oleh guru di awal yang secara bertahap dikurangi seiring meningkatnya kemandirian murid.", "I Do (Guru mendemonstrasikan) -> We Do (Guru & Murid bersama) -> You Do (Murid mandiri).", "Bahasa Inggris (Teaching Learning Cycle), Matematika, Seni Rupa.", "Bahasa Inggris Fase F: BKoF & MoT (bimbingan guru penuh) -> JCoT (kerja kelompok) -> ICoT (menulis esai mandiri tanpa bantuan).", "Menumbuhkan kemandirian belajar otentik tanpa rasa ditinggalkan.")
    ]

    for item in atp_sort_data:
        met, prs, mek, mapel, ex, bnf = item
        ws2.cell(row=r, column=1, value=met).alignment = align_center
        ws2.cell(row=r, column=2, value=prs).alignment = align_left
        ws2.cell(row=r, column=3, value=mek).alignment = align_left
        ws2.cell(row=r, column=4, value=mapel).alignment = align_left
        ws2.cell(row=r, column=5, value=ex).alignment = align_left
        ws2.cell(row=r, column=6, value=bnf).alignment = align_left

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

    auto_fit_columns(ws2, min_width=10, max_width=45)
    ws2.column_dimensions['A'].width = 18
    ws2.column_dimensions['B'].width = 24
    ws2.column_dimensions['C'].width = 30
    ws2.column_dimensions['D'].width = 24
    ws2.column_dimensions['E'].width = 36
    ws2.column_dimensions['F'].width = 32
    ws2.freeze_panes = 'A5'

    # ==========================================
    # SHEET 3: MODUL AJAR & DIFERENSIASI
    # ==========================================
    ws3 = wb.create_sheet(title="Modul Ajar & Diferensiasi")
    apply_title_banner(
        ws3,
        "PERENCANAAN PEMBELAJARAN SMA & DIFERENSIASI BERBASIS DEEP LEARNING",
        "Komponen RPP / Modul Ajar SMA dan Strategi Diferensiasi (Konten, Proses, Produk) PPA 2025",
        max_col=6
    )

    ws3.cell(row=4, column=1, value="1. KOMPONEN MINIMUM VS LENGKAP DOKUMEN PERENCANAAN PEMBELAJARAN").font = font_section_header
    ws3.merge_cells("A4:F4")
    for c in range(1, 7):
        ws3.cell(row=4, column=c).fill = fill_tbl_header
    ws3.row_dimensions[4].height = 24

    doc_plans = [
        ("Komponen Minimum (Cukup 1-2 Halaman)", "Sesuai Standar Proses Permendikbudristek No. 16/2022, rencana pembelajaran sekurang-kurangnya memuat 3 komponen inti:\n1. Tujuan Pembelajaran (TP) yang jelas dan terukur\n2. Langkah-langkah Kegiatan Pembelajaran (Pendahuluan, Inti, Penutup) yang bermakna dan berkesadaran\n3. Rencana Asesmen Pembelajaran (Asesmen awal, formatif proses, dan sumatif)"),
        ("Komponen Lengkap (Modul Ajar SMA)", "Guru dapat melengkapi dokumen perencanaan menjadi Modul Ajar utuh yang mencakup:\n• Identitas Sekolah, Mata Pelajaran, Fase/Kelas, Alokasi Waktu\n• Profil Lulusan yang dituju (8 Dimensi)\n• Sarana, Prasarana, dan Pemanfaatan Media Digital\n• Pemahaman Bermakna (Meaningful Concept) & Pertanyaan Pemantik (Driving Questions)\n• Rancangan Pembelajaran Berdiferensiasi (Konten, Proses, Produk)\n• Lampiran: Lembar Aktivitas Murid (LKPD), Bahan Bacaan Guru & Murid, Glosarium, Daftar Pustaka")
    ]

    r = 5
    for title, desc in doc_plans:
        ws3.cell(row=r, column=1, value=title).font = font_data_bold
        ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws3.cell(row=r, column=1).fill = fill_group
        ws3.cell(row=r, column=1).alignment = align_left
        
        ws3.cell(row=r, column=3, value=desc).font = font_data
        ws3.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
        ws3.cell(row=r, column=3).alignment = align_left

        for c in range(1, 7):
            ws3.cell(row=r, column=c).border = cell_border
            if c >= 3 and r % 2 == 1:
                ws3.cell(row=r, column=c).fill = fill_zebra
        ws3.row_dimensions[r].height = 54
        r += 1

    r += 1
    # Diferensiasi Pembelajaran di SMA
    ws3.cell(row=r, column=1, value="2. STRATEGI PEMBELAJARAN BERDIFERENSIASI DI SMA (PPA REVISI 2025)").font = font_section_header
    ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws3.cell(row=r, column=c).fill = fill_card_header
    ws3.row_dimensions[r].height = 24
    r += 1

    headers_dif = ["Aspek Diferensiasi", "Definisi Konseptual", "Dasar Pemetaan (Asesmen Diagnostik)", "Bentuk Aksi Guru di Kelas SMA", "Contoh Nyata di Mata Pelajaran SMA", "Prinsip Keadilan & Kesetaraan"]
    for i, h in enumerate(headers_dif, 1):
        cell = ws3.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = cell_border
    ws3.row_dimensions[r].height = 26
    r += 1

    dif_data = [
        ("Diferensiasi Konten", "Menyesuaikan materi, bahan ajar, dan tingkat kedalaman akses informasi dengan kesiapan belajar murid.", "Kesiapan belajar murid (Readiness Level) & Kompleksitas konsep prasyarat.", "Menyediakan bacaan berjenjang (teks sederhana, artikel populer, jurnal ilmiah) dan variasi format teks/video/infografis.", "Bahasa Inggris: Siswa yang belum lancar membaca diberikan teks deskripsi bergambar dengan glosarium; siswa mahir diberikan teks artikel berita asli tanpa bantuan.", "Tujuan pembelajaran tetap sama, tetapi akses pintu masuk materi disesuaikan."),
        ("Diferensiasi Proses", "Menyesuaikan cara murid memahami informasi, berlatih, dan memproses gagasan selama KBM.", "Gaya/profil belajar murid (Visual, Auditori, Kinestetik) dan kebutuhan pendampingan guru.", "Memberikan bimbingan berjenjang (scaffolding intensif pada kelompok yang kesulitan, eksplorasi mandiri pada kelompok mahir), variasi diskusi kelompok/individu.", "Matematika SMA: Kelompok murid yang belum paham diberikan manipulatif benda konkret/tutor sebaya; kelompok mahir langsung menyelesaikan tantangan soal HOTS kontekstual.", "Bukan membedakan kelas pintar vs lambat, melainkan memberi scaffolding adil."),
        ("Diferensiasi Produk", "Memberi keleluasaan bagi murid untuk mendemonstrasikan pemahaman dan hasil belajarnya dalam berbagai bentuk karya.", "Minat, bakat, kecerdasan majemuk, dan modalitas ekspresi murid.", "Menyediakan opsi tugas sumatif: boleh memilih menulis esai analitis, membuat video podcast, infografis digital Canva, atau presentasi lisan di depan kelas.", "Sejarah SMA: Untuk mendemonstrasikan analisis peristiwa Proklamasi, siswa boleh membuat video pendek dokumenter, komik sejarah digital, atau esai historiografi tertulis.", "Rubrik penilaian menilai substansi kompetensi yang sama meskipun format produk berbeda.")
    ]

    for item in dif_data:
        asp, defn, dsr, act, ex, prs = item
        ws3.cell(row=r, column=1, value=asp).alignment = align_center
        ws3.cell(row=r, column=2, value=defn).alignment = align_left
        ws3.cell(row=r, column=3, value=dsr).alignment = align_left
        ws3.cell(row=r, column=4, value=act).alignment = align_left
        ws3.cell(row=r, column=5, value=ex).alignment = align_left
        ws3.cell(row=r, column=6, value=prs).alignment = align_left

        ws3.cell(row=r, column=1).font = font_data_bold
        ws3.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 7):
            ws3.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws3.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws3.cell(row=r, column=c).fill = fill_zebra
        ws3.row_dimensions[r].height = 58
        r += 1

    auto_fit_columns(ws3, min_width=10, max_width=45)
    ws3.column_dimensions['A'].width = 20
    ws3.column_dimensions['B'].width = 26
    ws3.column_dimensions['C'].width = 26
    ws3.column_dimensions['D'].width = 34
    ws3.column_dimensions['E'].width = 36
    ws3.column_dimensions['F'].width = 30
    ws3.freeze_panes = 'A5'

    # ==========================================
    # SHEET 4: KRITERIA KETUNTASAN (KKTP)
    # ==========================================
    ws4 = wb.create_sheet(title="Kriteria Ketuntasan (KKTP)")
    apply_title_banner(
        ws4,
        "PANDUAN PENETAPAN KRITERIA KETERCAPAIAN TUJUAN PEMBELAJARAN (KKTP) SMA",
        "Tiga (3) Pendekatan Penentuan KKTP Berdasarkan Bab 3 & Bab 5 PPA Edisi Revisi 2025",
        max_col=6
    )

    ws4.cell(row=4, column=1, value="1. TIGA (3) PENDEKATAN MERUMUSKAN KKTP DI JENJANG SMA").font = font_section_header
    ws4.merge_cells("A4:F4")
    for c in range(1, 7):
        ws4.cell(row=4, column=c).fill = fill_tbl_header
    ws4.row_dimensions[4].height = 24

    headers_kktp = ["Pendekatan KKTP", "Format & Karakteristik Penilaian", "Struktur Penilaian", "Kriteria Ketuntasan Murid", "Tindak Lanjut Pembelajaran", "Kelebihan Pendekatan di SMA"]
    r = 5
    for i, h in enumerate(headers_kktp, 1):
        cell = ws4.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws4.row_dimensions[5].height = 26
    r += 1

    kktp_approaches = [
        ("Pendekatan 1: Deskripsi Kriteria", "Kualitatif deskriptif sederhana dengan daftar kriteria ketuntasan ya/tidak (sudah muncul vs belum muncul).", "Menyusun 3-5 indikator ketercapaian; guru mencentang status 'Memadai' atau 'Belum Memadai'.", "Murid dianggap tuntas mencapai TP jika minimal memenuhi batas jumlah kriteria memadai yang ditetapkan guru (misal 3 dari 4 kriteria memadai).", "Kriteria yang belum memadai menjadi fokus remedial spesifik sebelum melangkah ke tujuan berikutnya.", "Sangat praktis, mudah digunakan untuk tugas harian, dan fokus pada perilaku teramati."),
        ("Pendekatan 2: Rubrik Berskala", "Kualitatif bergradasi tingkat pencapaian (misal 4 jenjang: Baru Berkembang, Layak, Cakap, Mahir).", "Setiap kriteria dijelaskan indikator perilakunya secara rinci di masing-masing tingkat kemahiran.", "Murid tuntas jika berada minimal pada kategori 'Layak' atau 'Cakap' pada seluruh kriteria utama.", "Murid pada level Baru Berkembang mendapat bimbingan intensif; murid pada level Mahir mendapat tantangan pengayaan (*enrichment*).", "Sangat informatif, memandu murid melakukan evaluasi diri (*self-assessment*), dan memotivasi peningkatan mutu karya."),
        ("Pendekatan 3: Interval Nilai", "Kuantitatif persentase/skor yang dibagi ke dalam beberapa rentang interval ketercapaian.", "Menghitung persentase kriteria yang dipenuhi atau skor tes kuantitatif (rentang skala 0 - 100).", "Sekolah/guru menetapkan interval batas tuntas, misalnya:\n• 0 - 40%: Belum tuntas, remedial menyeluruh\n• 41 - 65%: Belum tuntas, remedial pada bagian tertentu\n• 66 - 85%: Sudah tuntas mencapai TP\n• 86 - 100%: Sudah tuntas, perlu pengayaan", "Remedial dilakukan spesifik pada indikator yang belum tuntas; nilai dapat dikonversi ke angka rapor.", "Mudah diintegrasikan dengan sistem penilaian rapor angka dan mudah dipahami oleh orang tua murid SMA.")
    ]

    for item in kktp_approaches:
        pdk, fmt, strc, tnt, tnd, adv = item
        ws4.cell(row=r, column=1, value=pdk).alignment = align_center
        ws4.cell(row=r, column=2, value=fmt).alignment = align_left
        ws4.cell(row=r, column=3, value=strc).alignment = align_left
        ws4.cell(row=r, column=4, value=tnt).alignment = align_left
        ws4.cell(row=r, column=5, value=tnd).alignment = align_left
        ws4.cell(row=r, column=6, value=adv).alignment = align_left

        ws4.cell(row=r, column=1).font = font_data_bold
        ws4.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 7):
            ws4.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws4.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws4.cell(row=r, column=c).fill = fill_zebra
        ws4.row_dimensions[r].height = 58
        r += 1

    r += 1
    # Contoh Simulasi Nyata KKTP Interval Nilai di SMA
    ws4.cell(row=r, column=1, value="2. CONTOH SIMULASI PENERAPAN KKTP INTERVAL NILAI DI SMA").font = font_section_header
    ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws4.cell(row=r, column=c).fill = fill_tbl_header
    ws4.row_dimensions[r].height = 24
    r += 1

    headers_sim = ["Rentang Interval Nilai", "Status Ketercapaian TP", "Kategori Kinerja Murid", "Tindakan Intervensi Pedagogis Guru", "Perlakuan Remedial di SMA", "Keterangan Angka Rapor"]
    for i, h in enumerate(headers_sim, 1):
        cell = ws4.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws4.row_dimensions[r].height = 24
    r += 1

    sim_data = [
        ("0 - 40%", "Belum Mencapai Ketuntasan", "Perlu Bimbingan Khusus (Sangat Rendah)", "Guru mengulang penjelasan konsep dasar secara individual atau kelompok kecil; pendampingan intensif.", "Remedial menyeluruh di seluruh indikator tujuan pembelajaran.", "Nilai diinput ulang setelah murid menyelesaikan program remedial."),
        ("41 - 65%", "Belum Mencapai Ketuntasan", "Hampir Mencapai (Perlu Penguatan)", "Guru memberikan bimbingan atau tugas terarah hanya pada materi/indikator yang belum dipahami murid.", "Remedial parsial pada indikator yang keliru saja (tidak perlu mengulang seluruh tes/tugas).", "Nilai disesuaikan berdasarkan pencapaian pasca-remedial indikator terkait."),
        ("66 - 85%", "Sudah Mencapai Ketuntasan", "Cakap / Memadai (Tuntas)", "Murid telah menguasai kompetensi yang diharapkan dan dapat melanjutkan ke materi berikutnya.", "Tidak memerlukan remedial; murid melanjutkan pembelajaran ke TP selanjutnya.", "Nilai diakui sebagai skor capaian kompetensi murni tanpa remedial."),
        ("86 - 100%", "Sudah Mencapai Ketuntasan", "Mahir / Sangat Baik (Tuntas Unggul)", "Guru memberikan tantangan lebih tinggi berupa pengayaan, studi kasus kontekstual, atau tugas kreatif berbasis proyek.", "Diberikan materi pengayaan (*enrichment*) untuk mengoptimalkan potensi akademisnya.", "Nilai menjadi dasar deskripsi keunggulan/kekuatan tertinggi di rapor.")
    ]

    for item in sim_data:
        rt, st, kt, intv, rem, rpr = item
        ws4.cell(row=r, column=1, value=rt).alignment = align_center
        ws4.cell(row=r, column=2, value=st).alignment = align_left
        ws4.cell(row=r, column=3, value=kt).alignment = align_left
        ws4.cell(row=r, column=4, value=intv).alignment = align_left
        ws4.cell(row=r, column=5, value=rem).alignment = align_left
        ws4.cell(row=r, column=6, value=rpr).alignment = align_left

        ws4.cell(row=r, column=1).font = font_data_bold
        ws4.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 7):
            ws4.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws4.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws4.cell(row=r, column=c).fill = fill_zebra
        ws4.row_dimensions[r].height = 42
        r += 1

    auto_fit_columns(ws4, min_width=10, max_width=45)
    ws4.column_dimensions['A'].width = 18
    ws4.column_dimensions['B'].width = 24
    ws4.column_dimensions['C'].width = 28
    ws4.column_dimensions['D'].width = 34
    ws4.column_dimensions['E'].width = 34
    ws4.column_dimensions['F'].width = 32
    ws4.freeze_panes = 'A5'

    # ==========================================
    # SHEET 5: PENGOLAHAN NILAI & RAPOR
    # ==========================================
    ws5 = wb.create_sheet(title="Pengolahan Nilai & Rapor")
    apply_title_banner(
        ws5,
        "TATA CARA PENGOLAHAN NILAI AKHIR DAN PENYUSUNAN DESKRIPSI RAPOR SMA",
        "Penetapan Nilai Rapor Berdasarkan Bab 5 PPA 2025: Pemisahan Formatif & Sumatif Serta Narasi Deskriptif",
        max_col=6
    )

    ws5.cell(row=4, column=1, value="1. PRINSIP PEMISAHAN ASESMEN FORMATIF DAN SUMATIF DALAM NILAI RAPOR").font = font_section_header
    ws5.merge_cells("A4:F4")
    for c in range(1, 7):
        ws5.cell(row=4, column=c).fill = fill_tbl_header
    ws5.row_dimensions[4].height = 24

    ws5.cell(row=5, column=1, value="ATURAN UTAMA PPA REVISI 2025 TERKAIT PENGOLAHAN NILAI:").font = font_callout_title
    ws5.cell(row=6, column=1, value=(
        "1. Pendidik TIDAK MENGGABUNGKAN hasil asesmen formatif dengan asesmen sumatif ke dalam nilai akhir rapor murid.\n"
        "2. Asesmen Formatif murni berfungsi sebagai 'Assessment for Learning' & 'Assessment as Learning' untuk mendiagnosis kendala dan memberikan umpan balik perbaikan.\n"
        "3. Nilai Akhir Rapor Murni Diperoleh dari ASESMEN SUMATIF LINGKUP MATERI (Sumatif 1, 2, 3...) yang dapat dipadukan dengan Sumatif Akhir Semester (SAS).\n"
        "4. Satuan pendidikan memiliki keleluasaan menentukan bobot penilaian antara Sumatif Lingkup Materi dan Sumatif Akhir Semester (misal: 60% Sumatif Materi + 40% SAS, atau 100% Rerata Sumatif Materi tanpa ujian akhir semester tertulis)."
    )).font = font_callout_body
    ws5.merge_cells("A5:F5")
    ws5.merge_cells("A6:F6")
    ws5.cell(row=5, column=1).fill = fill_callout
    ws5.cell(row=6, column=1).fill = fill_callout
    ws5.cell(row=6, column=1).alignment = align_left
    for r_i in [5, 6]:
        for c in range(1, 7):
            ws5.cell(row=r_i, column=c).border = cell_border
            ws5.cell(row=r_i, column=c).fill = fill_callout
    ws5.row_dimensions[5].height = 22
    ws5.row_dimensions[6].height = 68

    r = 8
    ws5.cell(row=r, column=1, value="2. CONTOH SIMULASI PENGHITUNGAN NILAI AKHIR RAPOR SMA").font = font_section_header
    ws5.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws5.cell(row=r, column=c).fill = fill_card_header
    ws5.row_dimensions[r].height = 24
    r += 1

    headers_sim_nilai = ["Nama Murid SMA", "Sumatif Lingkup Materi 1 (TP 1-2)", "Sumatif Lingkup Materi 2 (TP 3-4)", "Sumatif Lingkup Materi 3 (TP 5-6)", "Sumatif Akhir Semester (SAS)", "Nilai Akhir Rapor (Rerata / Bobot)"]
    for i, h in enumerate(headers_sim_nilai, 1):
        cell = ws5.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = cell_border
    ws5.row_dimensions[r].height = 26
    r += 1

    sample_scores = [
        ("Aditya Pratama (Kelas X SMA)", 85, 78, 92, 88, "86 (Skor Akhir Rapor)"),
        ("Bunga Citra Lestari (Kelas X SMA)", 92, 90, 95, 94, "93 (Skor Akhir Rapor)"),
        ("Dimas Arya Putra (Kelas XI SMA)", 72, 68, 75, 74, "72 (Skor Akhir Rapor)"),
        ("Fathir Muhammad (Kelas XI SMA)", 65, 80, 85, 78, "77 (Skor Akhir Rapor)")
    ]

    for item in sample_scores:
        nm, s1, s2, s3, sas, tot = item
        ws5.cell(row=r, column=1, value=nm).alignment = align_left
        ws5.cell(row=r, column=2, value=s1).alignment = align_center
        ws5.cell(row=r, column=3, value=s2).alignment = align_center
        ws5.cell(row=r, column=4, value=s3).alignment = align_center
        ws5.cell(row=r, column=5, value=sas).alignment = align_center
        ws5.cell(row=r, column=6, value=tot).alignment = align_center

        ws5.cell(row=r, column=1).font = font_data_bold
        for c in range(2, 6):
            ws5.cell(row=r, column=c).font = font_data
        ws5.cell(row=r, column=6).font = font_data_bold
        
        for c in range(1, 7):
            ws5.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws5.cell(row=r, column=c).fill = fill_zebra
        ws5.row_dimensions[r].height = 24
        r += 1

    r += 1
    # Tata Cara Penulisan Deskripsi Rapor
    ws5.cell(row=r, column=1, value="3. TATA CARA PENYUSUNAN DESKRIPSI CAPAIAN KOMPETENSI DI RAPOR SMA").font = font_section_header
    ws5.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws5.cell(row=r, column=c).fill = fill_tbl_header
    ws5.row_dimensions[r].height = 24
    r += 1

    headers_desk = ["Komponen Deskripsi", "Aturan Baku PPA 2025", "Cara Menentukan", "Contoh Rumusan Kalimat di Rapor SMA", "Kata Kunci yang Dianjurkan", "Hal yang Dilarang"]
    for i, h in enumerate(headers_desk, 1):
        cell = ws5.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws5.row_dimensions[r].height = 24
    r += 1

    desk_rules = [
        ("Kompetensi Tertinggi (Kekuatan Murid)", "Menjelaskan capaian kompetensi yang telah dikuasai murid dengan sangat baik.", "Dilihat dari skor sumatif tertinggi atau indikator KKTP yang dicapai dengan predikat Mahir.", "Mencapai kompetensi dengan sangat baik dalam menganalisis struktur dan kebahasaan teks eksposisi analitis serta menyajikan argumen berbasis bukti empiris.", "'Menunjukkan penguasaan sangat baik dalam...', 'Sangat terampil dalam...', 'Mencapai kompetensi prima dalam...'", "Menggunakan kata-kata yang terlalu umum seperti 'Pintar' atau 'Bagus' tanpa menyebut lingkup materi."),
        ("Kompetensi yang Perlu Peningkatan", "Menjelaskan capaian kompetensi yang masih perlu bimbingan atau pendampingan lebih lanjut.", "Dilihat dari skor sumatif terendah atau indikator KKTP yang belum memadai/remedial.", "Perlu peningkatan dan pendampingan dalam mengevaluasi informasi tersirat serta membedakan fakta dan opini pada teks diskusi kontroversial.", "'Perlu pendampingan dalam...', 'Perlu peningkatan pada hal...', 'Masih perlu berlatih untuk menguasai...'", "Menggunakan kata-kata yang merendahkan atau melabeli murid seperti 'Malas', 'Bodoh', atau 'Gagal'.")
    ]

    for item in desk_rules:
        komp, rule, how, ex, kw, ban = item
        ws5.cell(row=r, column=1, value=komp).alignment = align_left
        ws5.cell(row=r, column=2, value=rule).alignment = align_left
        ws5.cell(row=r, column=3, value=how).alignment = align_left
        ws5.cell(row=r, column=4, value=ex).alignment = align_left
        ws5.cell(row=r, column=5, value=kw).alignment = align_left
        ws5.cell(row=r, column=6, value=ban).alignment = align_left

        ws5.cell(row=r, column=1).font = font_data_bold
        for c in range(2, 7):
            ws5.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws5.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws5.cell(row=r, column=c).fill = fill_zebra
        ws5.row_dimensions[r].height = 54
        r += 1

    auto_fit_columns(ws5, min_width=10, max_width=45)
    ws5.column_dimensions['A'].width = 24
    ws5.column_dimensions['B'].width = 28
    ws5.column_dimensions['C'].width = 26
    ws5.column_dimensions['D'].width = 36
    ws5.column_dimensions['E'].width = 28
    ws5.column_dimensions['F'].width = 30
    ws5.freeze_panes = 'A5'

    # ==========================================
    # SHEET 6: FORMAT RAPOR SMA (X, XI, XII)
    # ==========================================
    ws6 = wb.create_sheet(title="Format Rapor SMA (X, XI, XII)")
    apply_title_banner(
        ws6,
        "FORMAT RESMI LAPORAN HASIL BELAJAR (RAPOR) JENJANG SMA",
        "Lampiran PPA Edisi Revisi 2025 (Halaman 148-151) | Format Fase E (Kelas X) & Fase F (Kelas XI - XII)",
        max_col=5
    )

    ws6.cell(row=4, column=1, value="SIMULASI FORMAT RAPOR SMA KELAS XI (FASE F) - EDISI REVISI 2025").font = font_section_header
    ws6.merge_cells("A4:E4")
    for c in range(1, 6):
        ws6.cell(row=4, column=c).fill = fill_tbl_header
    ws6.row_dimensions[4].height = 24

    bio_rapor = [
        ("Nama Murid", "Ahmad Fauzan Syahputra", "Kelas", "XI-A (Fase F)"),
        ("NISN", "0078912345", "Semester", "1 (Ganjil)"),
        ("Nama Sekolah", "SMA Negeri Unggulan", "Tahun Ajaran", "2025/2026")
    ]

    r = 5
    for l_label, l_val, r_label, r_val in bio_rapor:
        ws6.cell(row=r, column=1, value=l_label).font = font_data_bold
        ws6.cell(row=r, column=1).fill = fill_group
        ws6.cell(row=r, column=2, value=l_val).font = font_data
        ws6.cell(row=r, column=3, value=r_label).font = font_data_bold
        ws6.cell(row=r, column=3).fill = fill_group
        ws6.cell(row=r, column=4, value=r_val).font = font_data
        ws6.merge_cells(start_row=r, start_column=4, end_row=r, end_column=5)

        for col in range(1, 6):
            ws6.cell(row=r, column=col).border = cell_border
        ws6.row_dimensions[r].height = 22
        r += 1

    r += 1
    headers_tbl_rapor = ["No", "Kelompok & Mata Pelajaran", "Nilai Akhir", "Capaian Kompetensi (Deskripsi Keberhasilan & Peningkatan)", "Status KKTP"]
    for i, h in enumerate(headers_tbl_rapor, 1):
        cell = ws6.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws6.row_dimensions[r].height = 26
    r += 1

    rapor_data = [
        ("SECTION", "A. KELOMPOK MATA PELAJARAN WAJIB", "", "", ""),
        (1, "Pendidikan Agama Islam dan Budi Pekerti", 88, "Mencapai kompetensi sangat baik dalam menganalisis hukum bacaan Al-Qur'an dan mengamalkan nilai akhlak mulia dalam pergaulan remaja.\nPerlu peningkatan dalam menjelaskan sejarah perkembangan peradaban Islam di era modern.", "Tuntas"),
        (2, "Pendidikan Pancasila", 85, "Mencapai kompetensi sangat baik dalam menganalisis peluang dan tantangan penerapan nilai Pancasila di era digital.\nPerlu peningkatan dalam mengkaji dinamika tata kelola sistem pemerintahan Indonesia.", "Tuntas"),
        (3, "Bahasa Indonesia", 90, "Mencapai kompetensi sangat baik dalam mengapresiasi teks sastra puisi dan menyusun teks artikel opini secara kreatif dan logis.\nPerlu pendampingan dalam menyunting ejaan baku dan tanda baca pada karya esai panjang.", "Tuntas"),
        (4, "Matematika", 82, "Mencapai kompetensi sangat baik dalam memodelkan masalah sehari-hari menggunakan konsep barisan dan deret geometri.\nPerlu bimbingan lebih lanjut dalam menyelesaikan pembuktian rumus trigonometri analitik.", "Tuntas"),
        (5, "Bahasa Inggris", 92, "Mencapai kompetensi sangat baik dalam mengomunikasikan gagasan secara fasih pada presentasi teks eksposisi dan mempertahankan argumen berbasis bukti.\nPerlu peningkatan dalam mengevaluasi informasi tersirat teks diskusi kontroversial.", "Tuntas"),
        (6, "Pendidikan Jasmani Olahraga dan Kesehatan", 86, "Mencapai kompetensi sangat baik dalam mempraktikkan keterampilan gerak spesifik permainan bola basket dan pola hidup sehat.\nPerlu peningkatan dalam menyusun program latihan kebugaran jasmani pribadi teratur.", "Tuntas"),
        (7, "Sejarah", 88, "Mencapai kompetensi sangat baik dalam merekonstruksi peristiwa pergerakan nasional dan berpikir kritis terhadap sumber historiografi.\nPerlu peningkatan dalam mengidentifikasi dampak ekonomi kebijakan masa kolonial.", "Tuntas"),
        (8, "Seni Budaya (Seni Rupa)", 89, "Mencapai kompetensi sangat baik dalam merancang karya seni rupa dua dimensi bertema kearifan lokal dengan teknik media campuran.\nPerlu peningkatan dalam menyusun esai kritik seni rupa komparatif.", "Tuntas"),
        ("SECTION", "B. KELOMPOK MATA PELAJARAN PILIHAN (Pilihan Minat Siswa Fase F)", "", "", ""),
        (9, "Bahasa Inggris Tingkat Lanjut", 94, "Mencapai kompetensi sangat baik dalam menulis cerita pendek (creative writing) dengan plot kompleks dan memoderasi debat bahasa Inggris formal.\nPerlu peningkatan dalam analisis wacana kritis teks pidato internasional.", "Tuntas Unggul"),
        (10, "Matematika Tingkat Lanjut", 84, "Mencapai kompetensi sangat baik dalam mengoperasikan matriks dan menyelesaikan sistem persamaan linear multi-variabel.\nPerlu peningkatan dalam pemahaman konsep limit fungsi aljabar tak hingga.", "Tuntas"),
        (11, "Fisika", 85, "Mencapai kompetensi sangat baik dalam menganalisis hukum fluida dinamis dan melaksanakan eksperimen gerak harmonik sederhana.\nPerlu peningkatan dalam menerapkan hukum termodinamika pada siklus mesin pendingin.", "Tuntas"),
        (12, "Informatika", 91, "Mencapai kompetensi sangat baik dalam merancang algoritma pemrograman Python dan menganalisis keamanan jaringan komputer.\nPerlu bimbingan dalam pengolahan basis data SQL relasional kompleks.", "Tuntas Unggul"),
        (13, "Koding dan Kecerdasan Artifisial (AI)", 93, "Mencapai kompetensi sangat baik dalam melatih model machine learning sederhana dan mengidentifikasi etika penggunaan AI di sekolah.\nPerlu peningkatan dalam optimalisasi hyperparameter algoritma neural network.", "Tuntas Unggul")
    ]

    for item in rapor_data:
        code = item[0]
        if code == "SECTION":
            ws6.cell(row=r, column=1, value="")
            ws6.cell(row=r, column=2, value=item[1]).font = font_group_header
            ws6.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
            for col in range(1, 6):
                ws6.cell(row=r, column=col).fill = fill_group
                ws6.cell(row=r, column=col).border = cell_border
            ws6.row_dimensions[r].height = 24
        else:
            no, mapel, nilai, desk, stat = item
            ws6.cell(row=r, column=1, value=no).alignment = align_center
            ws6.cell(row=r, column=2, value=mapel).alignment = align_left
            ws6.cell(row=r, column=3, value=nilai).alignment = align_center
            ws6.cell(row=r, column=4, value=desk).alignment = align_left
            ws6.cell(row=r, column=5, value=stat).alignment = align_center

            ws6.cell(row=r, column=1).font = font_data_bold
            ws6.cell(row=r, column=2).font = font_data_bold
            ws6.cell(row=r, column=3).font = font_data_bold
            ws6.cell(row=r, column=4).font = font_data
            ws6.cell(row=r, column=5).font = font_data_bold
            
            for col in range(1, 6):
                ws6.cell(row=r, column=col).border = cell_border
                if r % 2 == 1:
                    ws6.cell(row=r, column=col).fill = fill_zebra
            ws6.row_dimensions[r].height = 42
        r += 1

    r += 1
    # Bagian Kokurikuler & Ekstrakurikuler Rapor
    ws6.cell(row=r, column=1, value="C. KOKURIKULER & EKSTRAKURIKULER").font = font_section_header
    ws6.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    for c in range(1, 6):
        ws6.cell(row=r, column=c).fill = fill_tbl_header
    ws6.row_dimensions[r].height = 24
    r += 1

    koku_rapor = [
        ("Kokurikuler (Pembelajaran Kolaboratif Lintas Disiplin / Gerakan 7 Kebiasaan)",
         "Ananda menunjukkan perkembangan sangat baik pada dimensi Kolaborasi dan Penalaran Kritis dalam projek 'Inovasi Pengelolaan Sampah Sekolah'. Mampu mengidentifikasi alternatif solusi daur ulang yang efektif. Masih perlu berlatih dalam dimensi Komunikasi untuk mempresentasikan laporan di depan publik."),
        ("Ekstrakurikuler: Pramuka (Wajib Disediakan)", "Terampil, berdisiplin tinggi, dan aktif dalam latihan kepramukaan gugus depan serta giat bakti sosial."),
        ("Ekstrakurikuler: English Club / Debate", "Sangat aktif sebagai debater utama; berhasil memenangkan kompetisi debat bahasa Inggris antar-SMA.")
    ]

    for title, desc in koku_rapor:
        ws6.cell(row=r, column=1, value=title).font = font_data_bold
        ws6.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws6.cell(row=r, column=1).fill = fill_group
        ws6.cell(row=r, column=1).alignment = align_left

        ws6.cell(row=r, column=3, value=desc).font = font_data
        ws6.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        ws6.cell(row=r, column=3).alignment = align_left

        for c in range(1, 6):
            ws6.cell(row=r, column=c).border = cell_border
            if c >= 3 and r % 2 == 1:
                ws6.cell(row=r, column=c).fill = fill_zebra
        ws6.row_dimensions[r].height = 36
        r += 1

    auto_fit_columns(ws6, min_width=10, max_width=50)
    ws6.column_dimensions['A'].width = 5
    ws6.column_dimensions['B'].width = 32
    ws6.column_dimensions['C'].width = 14
    ws6.column_dimensions['D'].width = 45
    ws6.column_dimensions['E'].width = 16
    ws6.freeze_panes = 'A5'

    # ==========================================
    # SHEET 7: KENAIKAN KELAS & KELULUSAN
    # ==========================================
    ws7 = wb.create_sheet(title="Kenaikan Kelas & Kelulusan")
    apply_title_banner(
        ws7,
        "MEKANISME KENAIKAN KELAS DAN KELULUSAN SMA (PPA 2025)",
        "Ketentuan Bab 5 PPA Edisi Revisi 2025: Prinsip Non-Tinggal Kelas Parsial, Kriteria Kenaikan & Syarat Kelulusan",
        max_col=6
    )

    ws7.cell(row=4, column=1, value="1. PRINSIP FILOSOFIS KENAIKAN KELAS DI SMA (PPA REVISI 2025)").font = font_section_header
    ws7.merge_cells("A4:F4")
    for c in range(1, 7):
        ws7.cell(row=4, column=c).fill = fill_tbl_header
    ws7.row_dimensions[4].height = 24

    ws7.cell(row=5, column=1, value="PENEGASAN REGULASI TERKAIT KENAIKAN KELAS MURID SMA:").font = font_callout_title
    ws7.cell(row=6, column=1, value=(
        "1. Penentuan kenaikan kelas TIDAK DITENTUKAN oleh kegagalan parsial pada 1-2 tujuan pembelajaran atau mata pelajaran tertentu.\n"
        "2. Pendidik wajib melakukan monitoring dan intervensi sepanjang tahun ajaran (bukan baru memvonis murid di akhir tahun saat pembagian rapor).\n"
        "3. Murid yang belum tuntas diberikan program bantuan belajar khusus (remedial, tutor sebaya, penyesuaian strategi) secara berkelanjutan.\n"
        "4. Kenaikan kelas didasarkan pada pertimbangan menyeluruh: (a) Laporan kemajuan belajar semua mapel, kokurikuler, dan ekstrakurikuler, (b) Rekam kehadiran murid, dan (c) Catatan karakter dan perkembangan sikap."
    )).font = font_callout_body
    ws7.merge_cells("A5:F5")
    ws7.merge_cells("A6:F6")
    ws7.cell(row=5, column=1).fill = fill_callout
    ws7.cell(row=6, column=1).fill = fill_callout
    ws7.cell(row=6, column=1).alignment = align_left
    for r_i in [5, 6]:
        for c in range(1, 7):
            ws7.cell(row=r_i, column=c).border = cell_border
            ws7.cell(row=r_i, column=c).fill = fill_callout
    ws7.row_dimensions[5].height = 22
    ws7.row_dimensions[6].height = 68

    r = 8
    ws7.cell(row=r, column=1, value="2. ASPEK PERTIMBANGAN KENAIKAN KELAS & KELULUSAN SMA").font = font_section_header
    ws7.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws7.cell(row=r, column=c).fill = fill_card_header
    ws7.row_dimensions[r].height = 24
    r += 1

    headers_kenaikan = ["Kategori Mekanisme", "Aspek Penilaian Utama", "Kriteria Regulasi Resmi PPA 2025", "Bentuk Bukti & Dokumen", "Peran Rapat Dewan Guru SMA", "Catatan Khusus Satuan Pendidikan"]
    for i, h in enumerate(headers_kenaikan, 1):
        cell = ws7.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = cell_border
    ws7.row_dimensions[r].height = 26
    r += 1

    kenaikan_data = [
        ("Kenaikan Kelas X ke XI (Fase E ke F)", "Pencapaian Kompetensi Fase E", "Murid telah mengikuti seluruh program pembelajaran kelas X dan memperoleh laporan hasil belajar lengkap.", "Buku Laporan Hasil Belajar (Rapor Semester 1 dan 2).", "Mempertimbangkan kesiapan murid memilih mata pelajaran pilihan di Fase F.", "Jika ada mapel belum tuntas, murid tetap naik kelas dengan catatan program intervensi di kelas XI."),
        ("Kenaikan Kelas XI ke XII (Fase F)", "Pencapaian Kompetensi Fase F Awal", "Murid menyelesaikan seluruh pembelajaran kelas XI pada kelompok wajib dan kelompok pilihan.", "Rapor semester 3 dan 4, portofolio tugas, catatan kehadiran.", "Mengevaluasi apakah ada perubahan minat mapel pilihan (paling lambat semester 2 kelas XI).", "Sekolah memfasilitasi pendampingan konseling karir dan studi lanjut."),
        ("Kelulusan SMA (Kelas XII)", "1. Menyelesaikan Seluruh Program Pembelajaran", "Murid telah menuntaskan seluruh proses pembelajaran dari semester 1 sampai semester 6 yang dibuktikan dengan rapor lengkap.", "Laporan lengkap penilaian hasil belajar semester 1-6.", "Menetapkan kelulusan murid secara resmi melalui Keputusan Kepala Sekolah.", "Murid yang lulus diberikan Ijazah resmi dari satuan pendidikan terakreditasi."),
        ("Kelulusan SMA (Kelas XII)", "2. Mengikuti Ujian Satuan Pendidikan", "Murid wajib mengikuti ujian akhir yang diselenggarakan oleh sekolah (bukan ujian nasional pemerintah).", "Portofolio, Penugasan, Tes Tertulis, dan/atau bentuk asesmen lain yang ditetapkan sekolah.", "Sekolah disarankan memadukan beragam bentuk ujian secara utuh.", "Pengolahan nilai ijazah diserahkan kepada kewenangan satuan pendidikan."),
        ("Kelulusan SMA (Kelas XII)", "3. Kehadiran & Kepribadian Baik", "Tingkat kehadiran konsisten mencerminkan kedisiplinan dan tanggung jawab; catatan karakter/akhlak baik.", "Rekapitulasi absensi tahunan, jurnal BK, dan catatan pembiasaan karakter.", "Rapat pleno dewan pendidik mempertimbangkan rekam jejak murid secara komprehensif.", "Batas minimal kehadiran dan kriteria sikap ditetapkan dalam KSP sekolah.")
    ]

    for item in kenaikan_data:
        kat, asp, reg, dok, rpt, cat = item
        ws7.cell(row=r, column=1, value=kat).alignment = align_center
        ws7.cell(row=r, column=2, value=asp).alignment = align_left
        ws7.cell(row=r, column=3, value=reg).alignment = align_left
        ws7.cell(row=r, column=4, value=dok).alignment = align_left
        ws7.cell(row=r, column=5, value=rpt).alignment = align_left
        ws7.cell(row=r, column=6, value=cat).alignment = align_left

        ws7.cell(row=r, column=1).font = font_data_bold
        ws7.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 7):
            ws7.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws7.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws7.cell(row=r, column=c).fill = fill_zebra
        ws7.row_dimensions[r].height = 48
        r += 1

    auto_fit_columns(ws7, min_width=10, max_width=45)
    ws7.column_dimensions['A'].width = 22
    ws7.column_dimensions['B'].width = 26
    ws7.column_dimensions['C'].width = 34
    ws7.column_dimensions['D'].width = 30
    ws7.column_dimensions['E'].width = 32
    ws7.column_dimensions['F'].width = 32
    ws7.freeze_panes = 'A5'

    # ==========================================
    # SHEET 8: REFLEKSI & BUDAYA MUTU SMA
    # ==========================================
    ws8 = wb.create_sheet(title="Refleksi & Budaya Mutu SMA")
    apply_title_banner(
        ws8,
        "KERANGKA REFLEKSI DAN EVALUASI BERKELANJUTAN PEMBELAJARAN SMA",
        "Penjabaran Bab 6 PPA 2025: Budaya Belajar Reflektif Guru, Teman Sejawat, Kepala Sekolah, dan Murid",
        max_col=6
    )

    ws8.cell(row=4, column=1, value="EMPAT (4) PILAR REFLEKSI PEMBELAJARAN DAN ASESMEN DI SMA").font = font_section_header
    ws8.merge_cells("A4:F4")
    for c in range(1, 7):
        ws8.cell(row=4, column=c).fill = fill_tbl_header
    ws8.row_dimensions[4].height = 24

    headers_ref = ["Pilar Refleksi", "Pelaksana Refleksi", "Frekuensi Pelaksanaan", "Pertanyaan Pemantik / Fokus Refleksi", "Bentuk Instrumen & Tindak Lanjut", "Dampak bagi Mutu Sekolah"]
    r = 5
    for i, h in enumerate(headers_ref, 1):
        cell = ws8.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws8.row_dimensions[5].height = 26
    r += 1

    ref_data = [
        ("Refleksi Diri Guru", "Guru Mata Pelajaran Secara Mandiri", "Minimal 1 kali dalam 1 semester",
         "1. Apa yang berhasil dan disukai dari KBM semester ini?\n2. Hal apa yang paling perlu ditingkatkan?\n3. Apa tantangan terbesar dan strategi solusinya?\n4. Apa yang akan saya ubah untuk semester depan?",
         "Jurnal refleksi mandiri, catatan anekdot kelas, checklist perbaikan modul ajar.",
         "Mencegah kejenuhan mengajar, mendorong inovasi metodologi pengajaran guru."),
        ("Refleksi Sesama Pendidik (Peer Reflection)", "Rekan Sejawat Guru / Komunitas Belajar (MGMP)", "Minimal 1 kali dalam 1 semester",
         "1. Bagaimana efektivitas manajemen kelas kolega?\n2. Apakah langkah pembelajaran sudah berpusat pada murid?\n3. Apakah asesmen yang dirancang adil dan proporsional?",
         "Observasi pembelajaran terbuka (open class), diskusi refleksi dialogis nonjudgmental pasca-observasi.",
         "Membangun kultur kolaborasi suportif antarguru; saling berbagi inspirasi praktik baik."),
        ("Refleksi oleh Kepala Satuan Pendidikan", "Kepala Sekolah / Wakil Kepala Bidang Kurikulum", "Secara berkala & terencana",
         "1. Apakah iklim belajar sekolah aman, nyaman, dan menggembirakan?\n2. Fasilitas dan dukungan apa yang dibutuhkan guru?\n3. Bagaimana progres capaian literasi dan numerasi sekolah?",
         "Supervisi klinis suportif, observasi acak suasana kelas, forum evaluasi pembelajaran bulanan.",
         "Kepala sekolah berfungsi sebagai pemimpin pembelajaran (instructional leader) bukan sekadar administrator."),
        ("Refleksi oleh Murid SMA", "Peserta Didik (Siswa Kelas X, XI, XII)", "Minimal 1 kali dalam 1 semester (akhir materi/semester)",
         "1. Bagian materi mana yang paling mudah dan paling sulit dipahami?\n2. Apakah cara mengajar guru menyenangkan dan membantu?\n3. Suasana belajar seperti apa yang diharapkan?",
         "Kuesioner umpan balik murid (Google Form/kertas survei), kartu refleksi 'Exit Ticket' harian.",
         "Menumbuhkan kemandirian, nalar kritis murid, dan memberi masukan nyata bagi guru.")
    ]

    for item in ref_data:
        pil, pel, frq, qst, ins, imp = item
        ws8.cell(row=r, column=1, value=pil).alignment = align_center
        ws8.cell(row=r, column=2, value=pel).alignment = align_left
        ws8.cell(row=r, column=3, value=frq).alignment = align_center
        ws8.cell(row=r, column=4, value=qst).alignment = align_left
        ws8.cell(row=r, column=5, value=ins).alignment = align_left
        ws8.cell(row=r, column=6, value=imp).alignment = align_left

        ws8.cell(row=r, column=1).font = font_data_bold
        ws8.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 7):
            ws8.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws8.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws8.cell(row=r, column=c).fill = fill_zebra
        ws8.row_dimensions[r].height = 56
        r += 1

    auto_fit_columns(ws8, min_width=10, max_width=45)
    ws8.column_dimensions['A'].width = 20
    ws8.column_dimensions['B'].width = 24
    ws8.column_dimensions['C'].width = 18
    ws8.column_dimensions['D'].width = 38
    ws8.column_dimensions['E'].width = 34
    ws8.column_dimensions['F'].width = 32
    ws8.freeze_panes = 'A5'

    # Save workbook
    wb.save(output_path)
    print(f"PPA SMA Workbook successfully created and saved to: {output_path}")

if __name__ == "__main__":
    out_file = r"C:\Users\KOMPUTERKU2\Documents\English SMA\Peraturan\Panduan_Pembelajaran_dan_Asesmen_SMA_2025.xlsx"
    create_ppa_sma_excel(out_file)
