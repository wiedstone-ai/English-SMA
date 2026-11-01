# -*- coding: utf-8 -*-
"""
Script to generate a comprehensive, highly styled Excel workbook summarizing
PERMENDIKDASMEN No. 13 Tahun 2025 regarding SMA Curriculum.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_kurikulum_excel(output_path):
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Styles definition
    font_family = "Segoe UI"
    
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
    font_badge = Font(name=font_family, size=9, bold=True, color="047857")

    fill_main_title = PatternFill(start_color="0F2942", end_color="0F2942", fill_type="solid")     # Deep Navy
    fill_subtitle = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")       # Dark Blue
    fill_tbl_header = PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid")     # Royal Blue
    fill_group = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")          # Soft Light Blue
    fill_subgroup = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")       # Very Light Blue
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")          # Slate White
    fill_total = PatternFill(start_color="BFDBFE", end_color="BFDBFE", fill_type="solid")          # Highlight Total Blue
    fill_callout = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")        # Warm Amber Note
    fill_card_header = PatternFill(start_color="3B82F6", end_color="3B82F6", fill_type="solid")    # Card Accent Blue

    border_thin_gray = Side(style='thin', color="CBD5E1")
    border_double_bottom = Side(style='double', color="0F2942")
    border_thick_top = Side(style='medium', color="1E40AF")
    
    cell_border = Border(left=border_thin_gray, right=border_thin_gray, top=border_thin_gray, bottom=border_thin_gray)
    total_border = Border(left=border_thin_gray, right=border_thin_gray, top=border_thin_gray, bottom=border_double_bottom)
    header_border = Border(left=border_thin_gray, right=border_thin_gray, top=border_thick_top, bottom=border_thick_top)

    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    align_right = Alignment(horizontal='right', vertical='center')

    def apply_title_banner(ws, title, subtitle, max_col=7):
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

        # Fill merged cells background
        for col in range(1, max_col + 1):
            ws.cell(row=1, column=col).fill = fill_main_title
            ws.cell(row=2, column=col).fill = fill_subtitle

    def auto_fit_columns(ws, min_width=12, max_width=60):
        ws.views.sheetView[0].showGridLines = True
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                # ignore merged cells in row 1-3 for width
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
    # SHEET 1: RINGKASAN EKSEKUTIF
    # ==========================================
    ws1 = wb.create_sheet(title="Ringkasan Eksekutif")
    apply_title_banner(
        ws1,
        "RINGKASAN EKSEKUTIF: KURIKULUM SMA",
        "Berdasarkan PERMENDIKDASMEN RI Nomor 13 Tahun 2025 (Perubahan atas Permendikbudristek No. 12 Tahun 2024)",
        max_col=6
    )

    ws1.cell(row=4, column=1, value="INFORMASI REGULASI RESMI").font = font_section_header
    ws1.merge_cells("A4:F4")
    for c in range(1, 7):
        ws1.cell(row=4, column=c).fill = fill_tbl_header
        ws1.cell(row=4, column=c).alignment = align_left
    ws1.row_dimensions[4].height = 24

    reg_info = [
        ("Nama Peraturan", "Peraturan Menteri Pendidikan Dasar dan Menengah Republik Indonesia Nomor 13 Tahun 2025"),
        ("Tentang", "Perubahan atas Peraturan Menteri Pendidikan, Kebudayaan, Riset, dan Teknologi Nomor 12 Tahun 2024 tentang Kurikulum pada Pendidikan Anak Usia Dini, Jenjang Pendidikan Dasar, dan Jenjang Pendidikan Menengah"),
        ("Pejabat Penetap", "Menteri Pendidikan Dasar dan Menengah Republik Indonesia (Abdul Mu'ti)"),
        ("Tanggal Ditetapkan", "11 Juli 2025 (Diundangkan dan mulai berlaku sejak tanggal pengundangan)"),
        ("Tujuan Pokok", "Penyempurnaan kerangka dasar, struktur kurikulum, kokurikuler, dan ekstrakurikuler melalui pendekatan Pembelajaran Mendalam (Deep Learning)"),
        ("Fokus Jenjang", "SMA / MA (Sekolah Menengah Atas / Madrasah Aliyah) Fase E (Kelas X) & Fase F (Kelas XI - XII)")
    ]

    r = 5
    for item, desc in reg_info:
        ws1.cell(row=r, column=1, value=item).font = font_data_bold
        ws1.cell(row=r, column=1).fill = fill_group
        ws1.cell(row=r, column=1).border = cell_border
        ws1.cell(row=r, column=1).alignment = align_left
        
        ws1.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        c_desc = ws1.cell(row=r, column=2, value=desc)
        c_desc.font = font_data
        c_desc.alignment = align_left
        for col in range(2, 7):
            ws1.cell(row=r, column=col).border = cell_border
            if r % 2 == 1:
                ws1.cell(row=r, column=col).fill = fill_zebra
        ws1.row_dimensions[r].height = 24
        r += 1

    r += 1
    ws1.cell(row=r, column=1, value="6 PERUBAHAN UTAMA DALAM PERMENDIKDASMEN 13 TAHUN 2025 UNTUK SMA").font = font_section_header
    ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws1.cell(row=r, column=c).fill = fill_tbl_header
        ws1.cell(row=r, column=c).alignment = align_left
    ws1.row_dimensions[r].height = 24
    r += 1

    headers_perubahan = ["No", "Aspek Regulasi", "Ketentuan Baru (Permendikdasmen 13/2025)", "Dasar Hukum", "Dampak Praktis pada SMA", "Catatan Implementasi"]
    for i, h in enumerate(headers_perubahan, 1):
        cell = ws1.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws1.row_dimensions[r].height = 26
    r += 1

    perubahan_data = [
        (1, "Pendekatan Pembelajaran", "Mengadopsi resmi 'Pembelajaran Mendalam (Deep Learning)' dengan 3 prinsip: Berkesadaran (Mindful), Bermakna (Meaningful), Menggembirakan (Joyful) dan siklus Memahami-Mengaplikasi-Merefleksi.", "Pasal 3 ayat (2) & Lampiran I", "Guru mengarahkan siswa aktif meregulasi diri, mengaitkan konten dengan konteks nyata, dan menghilangkan suasana belajar tertekan.", "Didukung semangat saling memuliakan (K.H. Hasyim Asy'ari) & Sistem Among (Ki Hajar Dewantara)."),
        (2, "Profil Lulusan (8 Dimensi)", "Penguatan profil lulusan berfokus pada 8 dimensi kompetensi utuh: Keimanan & Ketakwaan, Kewargaan, Penalaran Kritis, Kreativitas, Kolaborasi, Kemandirian, Kesehatan, dan Komunikasi.", "Pasal 17 ayat (1) & Lampiran I", "Menggantikan formulasi 6 dimensi P5 sebelumnya; kini menambahkan dimensi 'Kesehatan' dan 'Komunikasi' serta 'Kewargaan'.", "Alur perkembangan kompetensi ditetapkan oleh pimpinan tinggi madya bidang kurikulum."),
        (3, "Format Baru Kokurikuler", "Kokurikuler dilaksanakan dalam bentuk: (1) Pembelajaran kolaboratif lintas disiplin ilmu, (2) Gerakan 7 Kebiasaan Anak Indonesia Hebat, dan/atau (3) Cara lainnya. Alokasi waktu dirumuskan 1 tahun.", "Pasal 16, 18, 19", "Projek kokurikuler tidak lagi kaku, sekolah memiliki fleksibilitas tinggi memadukan proyek lintas mapel atau pembiasaan karakter 7 kebiasaan.", "Tema dikembangkan satuan pendidikan sesuai konteks sosial budaya dan karakteristik siswa."),
        (4, "Mapel Baru: Koding & AI", "Mata pelajaran pilihan 'Koding dan Kecerdasan Artifisial' resmi masuk struktur kurikulum SMA mulai Tahun Ajaran 2025-2026 secara bertahap.", "Pasal 32A & Lampiran II (Bagian D)", "Kelas X: pilihan 2 JP/minggu (72 JP/tahun). Kelas XI & XII: pilihan 5 JP/minggu (180 JP kelas XI / 160 JP kelas XII).", "Disediakan oleh Satuan Pendidikan sesuai sumber daya yang dimiliki dan dipilih sesuai minat murid."),
        (5, "Layanan Ekstrakurikuler", "Satuan pendidikan jalur formal jenjang dikdasmen sekurang-kurangnya menyediakan Ekstrakurikuler Kepramukaan atau Kepanduan lainnya.", "Pasal 22 & Lampiran III", "Pramuka/Kepanduan wajib disiapkan oleh sekolah, namun keikutsertaan siswa bersifat sukarela sesuai minat dan potensi individual.", "Penilaian ekstrakurikuler bersifat kualitatif dan dideskripsikan di rapor."),
        (6, "Implementasi Bertahap", "Penerapan kurikulum baru pada jenjang SMA/MA dapat dilakukan secara bertahap mulai dari Kelas X pada tahun ajaran baru.", "Pasal 32 huruf c", "Memberikan ruang adaptasi terstruktur bagi guru dan pihak sekolah tanpa gejolak administratif.", "Tahun Ajaran 2025/2026 menjadi titik awal penerapan kurikulum penyempurnaan ini.")
    ]

    for row_item in perubahan_data:
        no, aspek, ketentuan, hukum, dampak, catatan = row_item
        ws1.cell(row=r, column=1, value=no).alignment = align_center
        ws1.cell(row=r, column=2, value=aspek).alignment = align_left
        ws1.cell(row=r, column=3, value=ketentuan).alignment = align_left
        ws1.cell(row=r, column=4, value=hukum).alignment = align_center
        ws1.cell(row=r, column=5, value=dampak).alignment = align_left
        ws1.cell(row=r, column=6, value=catatan).alignment = align_left
        
        ws1.cell(row=r, column=1).font = font_data_bold
        ws1.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 7):
            ws1.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws1.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws1.cell(row=r, column=c).fill = fill_zebra
        ws1.row_dimensions[r].height = 48
        r += 1

    auto_fit_columns(ws1, min_width=10, max_width=45)
    ws1.column_dimensions['A'].width = 6
    ws1.column_dimensions['B'].width = 22
    ws1.column_dimensions['C'].width = 38
    ws1.column_dimensions['D'].width = 20
    ws1.column_dimensions['E'].width = 34
    ws1.column_dimensions['F'].width = 32

    # ==========================================
    # SHEET 2: MATRIKS STRUKTUR SMA (MASTER)
    # ==========================================
    ws2 = wb.create_sheet(title="Matriks Struktur SMA")
    apply_title_banner(
        ws2,
        "MATRIKS STRUKTUR KURIKULUM SMA / MA",
        "Komparasi Alokasi Waktu Beban Belajar Kelas X (Fase E), Kelas XI (Fase F), dan Kelas XII (Fase F)",
        max_col=13
    )

    headers_m1 = ["No", "Kelompok & Mata Pelajaran", "Kategori", "KELAS X (FASE E - 36 Mgg)", "", "", "KELAS XI (FASE F - 36 Mgg)", "", "", "KELAS XII (FASE F - 32 Mgg)", "", "", "Keterangan & Ketentuan Teknis"]
    headers_m2 = ["", "", "", "Intra/Th", "Koku/Th", "Total JP (JP/Mgg)", "Intra/Th", "Koku/Th", "Total JP (JP/Mgg)", "Intra/Th", "Koku/Th", "Total JP (JP/Mgg)", ""]

    r = 4
    for i, h in enumerate(headers_m1, 1):
        cell = ws2.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = cell_border

    r = 5
    for i, h in enumerate(headers_m2, 1):
        cell = ws2.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = cell_border

    # Merges for 2-row headers
    ws2.merge_cells("A4:A5")
    ws2.merge_cells("B4:B5")
    ws2.merge_cells("C4:C5")
    ws2.merge_cells("D4:F4")
    ws2.merge_cells("G4:I4")
    ws2.merge_cells("J4:L4")
    ws2.merge_cells("M4:M5")
    ws2.row_dimensions[4].height = 22
    ws2.row_dimensions[5].height = 22
    r = 6

    # Master Table Data: (No, Subject, Category, X_intra, X_koku, X_tot, XI_intra, XI_koku, XI_tot, XII_intra, XII_koku, XII_tot, Notes)
    master_data = [
        # --- KELOMPOK WAJIB ---
        ("SECTION", "I. KELOMPOK MATA PELAJARAN WAJIB (Diikuti oleh seluruh peserta didik)", "", "", "", "", "", "", "", "", "", "", ""),
        (1, "Pendidikan Agama dan Budi Pekerti", "Wajib", 72, 36, "108 (3)", 72, 36, "108 (3)", 64, 32, "96 (3)", "Diikuti sesuai agama/kepercayaan masing-masing."),
        (2, "Pendidikan Pancasila", "Wajib", 72, 0, "72 (2)", 72, 0, "72 (2)", 64, 0, "64 (2)", "Fokus pada penguatan kewargaan dan ideologi Pancasila."),
        (3, "Bahasa Indonesia", "Wajib", 108, 36, "144 (4)", 108, 36, "144 (4)", 96, 32, "128 (4)", "Mencakup kompetensi literasi, komunikasi, dan apresiasi karya."),
        (4, "Matematika", "Wajib", 108, 36, "144 (4)", 108, 36, "144 (4)", 96, 32, "128 (4)", "Matematika umum untuk fondasi penalaran kritis dan numerasi."),
        (5, "Bahasa Inggris", "Wajib", 108, 0, "108 (3)", 108, 0, "108 (3)", 96, 0, "96 (3)", "Mata pelajaran wajib intrakurikuler (3 JP/mgg di semua jenjang)."),
        (6, "Pendidikan Jasmani Olahraga dan Kesehatan (PJOK)", "Wajib", 72, 36, "108 (3)", 72, 36, "108 (3)", 64, 32, "96 (3)", "Menekankan kebugaran jasmani dan pembiasaan gaya hidup sehat."),
        (7, "Sejarah", "Wajib", "-", "-", "Termasuk IPS", 72, 0, "72 (2)", 64, 0, "64 (2)", "Di kelas X terintegrasi dalam rumpun IPS; di kelas XI-XII berdiri sendiri."),
        (8, "Seni dan Budaya / Prakarya", "Wajib", 72, 0, "72 (2)", 72, 0, "72 (2)", 64, 0, "64 (2)", "Di kelas X: Seni (Musik/Rupa/Teater/Tari) atau Prakarya. Di kelas XI-XII: Seni Budaya."),
        (9, "Ilmu Pengetahuan Alam (IPA: Fisika, Kimia, Biologi)", "Wajib", 216, 108, "324 (9)", "-", "-", "-", "-", "-", "-", "Hanya di Kelas X. Model terpadu, blok, atau paralel inkuiri."),
        (10, "Ilmu Pengetahuan Sosial (IPS: Sosiologi, Ekonomi, Sejarah, Geografi)", "Wajib", 288, 144, "432 (12)", "-", "-", "-", "-", "-", "-", "Hanya di Kelas X. Model terpadu, blok, atau paralel inkuiri."),
        (11, "Informatika", "Wajib", 72, 0, "72 (2)", "-", "-", "-", "-", "-", "-", "Mata pelajaran wajib di Kelas X (Fase E). Menjadi mapel pilihan di Fase F."),
        ("SUBTOTAL_WAJIB", "SUBTOTAL MATA PELAJARAN WAJIB", "Subtotal", 1188, 396, "1.584 (44)", 684, 144, "828 (23)", 608, 128, "736 (23)", "Beban wajib: Kelas X = 44 JP/mgg, Kelas XI = 23 JP/mgg, Kelas XII = 23 JP/mgg"),
        
        # --- KELOMPOK PILIHAN ---
        ("SECTION", "II. KELOMPOK MATA PELAJARAN PILIHAN (Satuan pendidikan wajib membuka min. 7 mapel pilihan)", "", "", "", "", "", "", "", "", "", "", ""),
        (12, "Koding dan Kecerdasan Artifisial (AI)", "Pilihan", 72, 0, "72 (2)", 180, 0, "180 (5)", 160, 0, "160 (5)", "MAPEL BARU Permendikdasmen 13/2025. Mulai TA 2025/2026 secara bertahap."),
        (13, "Bahasa Inggris Tingkat Lanjut", "Pilihan", "-", "-", "-", 180, 0, "180 (5)", 160, 0, "160 (5)", "Mata pelajaran pilihan peminatan bahasa di Fase F (5 JP/mgg)."),
        (14, "Matematika Tingkat Lanjut", "Pilihan", "-", "-", "-", 180, 0, "180 (5)", 160, 0, "160 (5)", "Mata pelajaran pilihan peminatan MIPA di Fase F (5 JP/mgg)."),
        (15, "Fisika", "Pilihan", "-", "-", "-", 180, 0, "180 (5)", 160, 0, "160 (5)", "Mata pelajaran pilihan peminatan MIPA di Fase F (5 JP/mgg)."),
        (16, "Kimia", "Pilihan", "-", "-", "-", 180, 0, "180 (5)", 160, 0, "160 (5)", "Mata pelajaran pilihan peminatan MIPA di Fase F (5 JP/mgg)."),
        (17, "Biologi", "Pilihan", "-", "-", "-", 180, 0, "180 (5)", 160, 0, "160 (5)", "Mata pelajaran pilihan peminatan MIPA di Fase F (5 JP/mgg)."),
        (18, "Informatika (Tingkat Lanjut/Pilihan)", "Pilihan", "-", "-", "-", 180, 0, "180 (5)", 160, 0, "160 (5)", "Mata pelajaran pilihan di Fase F (5 JP/mgg)."),
        (19, "Geografi", "Pilihan", "-", "-", "-", 180, 0, "180 (5)", 160, 0, "160 (5)", "Mata pelajaran pilihan peminatan IPS di Fase F (5 JP/mgg)."),
        (20, "Sosiologi", "Pilihan", "-", "-", "-", 180, 0, "180 (5)", 160, 0, "160 (5)", "Mata pelajaran pilihan peminatan IPS di Fase F (5 JP/mgg)."),
        (21, "Ekonomi", "Pilihan", "-", "-", "-", 180, 0, "180 (5)", 160, 0, "160 (5)", "Mata pelajaran pilihan peminatan IPS di Fase F (5 JP/mgg)."),
        (22, "Sejarah Tingkat Lanjut", "Pilihan", "-", "-", "-", 180, 0, "180 (5)", 160, 0, "160 (5)", "Mata pelajaran pilihan peminatan IPS/Humaniora di Fase F (5 JP/mgg)."),
        (23, "Antropologi", "Pilihan", "-", "-", "-", 180, 0, "180 (5)", 160, 0, "160 (5)", "Mata pelajaran pilihan peminatan IPS/Humaniora di Fase F (5 JP/mgg)."),
        (24, "Bahasa Indonesia Tingkat Lanjut", "Pilihan", "-", "-", "-", 180, 0, "180 (5)", 160, 0, "160 (5)", "Mata pelajaran pilihan peminatan Bahasa di Fase F (5 JP/mgg)."),
        (25, "Bahasa Asing (Arab, Jepang, Jerman, Korea, Mandarin, Prancis)", "Pilihan", "-", "-", "-", 180, 0, "180 (5)", 160, 0, "160 (5)", "Tersedia 6 pilihan bahasa asing resmi (masing-masing 5 JP/mgg)."),
        (26, "Prakarya dan Kewirausahaan (PKWU)", "Pilihan", "-", "-", "-", 72, 0, "72 (2)", 64, 0, "64 (2)", "Pilihan budi daya, kerajinan, rekayasa, atau pengolahan (2 JP/mgg)."),
        (27, "Mata Pelajaran Lainnya (Pengembangan Sekolah)", "Pilihan", "-", "-", "-", "Sesuai", 0, "20-25 JP", "Sesuai", 0, "20-25 JP", "Dapat dibuka sesuai sumber daya & diversifikasi kurikulum sekolah."),
        ("SUBTOTAL_PILIHAN", "TOTAL MATA PELAJARAN PILIHAN YANG DIAMBIL SISWA", "Subtotal", 0, 0, "0-72 (0-2)", "720-900", 0, "720-900 (20-25)", "640-800", 0, "640-800 (20-25)", "Siswa Fase F memilih 4 s.d. 5 mapel pilihan (20-25 JP/minggu)."),

        # --- MUATAN LOKAL & TOTAL GABUNGAN ---
        ("SECTION", "III. MUATAN LOKAL DAN TOTAL KESELURUHAN BEBAN BELAJAR", "", "", "", "", "", "", "", "", "", "", ""),
        (28, "Muatan Lokal (Maksimal)", "Mulok", 72, 0, "72 (2)", 72, 0, "72 (2)", 64, 0, "64 (2)", "Maksimal 2 JP/minggu (seni budaya, prakarya, PJOK, bahasa, teknologi)."),
        ("TOTAL_WAJIB_PILIHAN", "TOTAL WAJIB + PILIHAN", "Total", 1260, 396, "1.656 (46)", "1404-1584", 144, "1548-1728 (43-48)", "1248-1408", 128, "1376-1536 (43-48)", "Rentang beban reguler siswa tanpa muatan lokal mandiri."),
        ("TOTAL_AKHIR", "TOTAL KESELURUHAN (WAJIB + PILIHAN + MUATAN LOKAL)", "Grand Total", 1332, 396, "1.728 (48)", "1476-1656", 144, "1.620-1.800 (45-50)", "1312-1472", 128, "1.440-1.600 (45-50)", "Beban belajar maksimal per tahun ajaran.")
    ]

    for item in master_data:
        code = item[0]
        if code == "SECTION":
            ws2.cell(row=r, column=1, value="").fill = fill_group
            ws2.cell(row=r, column=2, value=item[1]).font = font_group_header
            ws2.merge_cells(start_row=r, start_column=2, end_row=r, end_column=13)
            for col in range(1, 14):
                ws2.cell(row=r, column=col).fill = fill_group
                ws2.cell(row=r, column=col).border = cell_border
            ws2.row_dimensions[r].height = 24
        elif code in ["SUBTOTAL_WAJIB", "SUBTOTAL_PILIHAN", "TOTAL_WAJIB_PILIHAN", "TOTAL_AKHIR"]:
            ws2.cell(row=r, column=1, value="")
            ws2.cell(row=r, column=2, value=item[1]).font = font_total
            ws2.cell(row=r, column=3, value=item[2]).font = font_total
            ws2.cell(row=r, column=4, value=item[3]).font = font_total
            ws2.cell(row=r, column=5, value=item[4]).font = font_total
            ws2.cell(row=r, column=6, value=item[5]).font = font_total
            ws2.cell(row=r, column=7, value=item[6]).font = font_total
            ws2.cell(row=r, column=8, value=item[7]).font = font_total
            ws2.cell(row=r, column=9, value=item[8]).font = font_total
            ws2.cell(row=r, column=10, value=item[9]).font = font_total
            ws2.cell(row=r, column=11, value=item[10]).font = font_total
            ws2.cell(row=r, column=12, value=item[11]).font = font_total
            ws2.cell(row=r, column=13, value=item[12]).font = font_data_bold
            
            fill_target = fill_total if "TOTAL" in code else fill_subgroup
            for col in range(1, 14):
                ws2.cell(row=r, column=col).fill = fill_target
                ws2.cell(row=r, column=col).border = total_border if "TOTAL" in code else cell_border
                if col in [1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]:
                    ws2.cell(row=r, column=col).alignment = align_center
                elif col == 2:
                    ws2.cell(row=r, column=col).alignment = align_left
                else:
                    ws2.cell(row=r, column=col).alignment = align_left
            ws2.row_dimensions[r].height = 24
        else:
            no, name, cat, x_in, x_ko, x_to, xi_in, xi_ko, xi_to, xii_in, xii_ko, xii_to, note = item
            ws2.cell(row=r, column=1, value=no).alignment = align_center
            ws2.cell(row=r, column=2, value=name).alignment = align_left
            ws2.cell(row=r, column=3, value=cat).alignment = align_center
            ws2.cell(row=r, column=4, value=x_in).alignment = align_center
            ws2.cell(row=r, column=5, value=x_ko).alignment = align_center
            ws2.cell(row=r, column=6, value=x_to).alignment = align_center
            ws2.cell(row=r, column=7, value=xi_in).alignment = align_center
            ws2.cell(row=r, column=8, value=xi_ko).alignment = align_center
            ws2.cell(row=r, column=9, value=xi_to).alignment = align_center
            ws2.cell(row=r, column=10, value=xii_in).alignment = align_center
            ws2.cell(row=r, column=11, value=xii_ko).alignment = align_center
            ws2.cell(row=r, column=12, value=xii_to).alignment = align_center
            ws2.cell(row=r, column=13, value=note).alignment = align_left

            ws2.cell(row=r, column=1).font = font_data_bold
            ws2.cell(row=r, column=2).font = font_data_bold if "Bahasa Inggris" in name or "Koding" in name else font_data
            for col in range(3, 14):
                ws2.cell(row=r, column=col).font = font_data
            
            for col in range(1, 14):
                ws2.cell(row=r, column=col).border = cell_border
                if r % 2 == 1:
                    ws2.cell(row=r, column=col).fill = fill_zebra
            ws2.row_dimensions[r].height = 22
        r += 1

    auto_fit_columns(ws2, min_width=8, max_width=50)
    ws2.column_dimensions['A'].width = 5
    ws2.column_dimensions['B'].width = 34
    ws2.column_dimensions['C'].width = 11
    ws2.column_dimensions['D'].width = 10
    ws2.column_dimensions['E'].width = 10
    ws2.column_dimensions['F'].width = 15
    ws2.column_dimensions['G'].width = 10
    ws2.column_dimensions['H'].width = 10
    ws2.column_dimensions['I'].width = 15
    ws2.column_dimensions['J'].width = 10
    ws2.column_dimensions['K'].width = 10
    ws2.column_dimensions['L'].width = 15
    ws2.column_dimensions['M'].width = 40
    ws2.freeze_panes = 'A6'

    # ==========================================
    # SHEET 3: KELAS X (FASE E)
    # ==========================================
    ws3 = wb.create_sheet(title="Kelas X (Fase E)")
    apply_title_banner(
        ws3,
        "STRUKTUR KURIKULUM KELAS X SMA / MA (FASE E)",
        "Tabel 8 Permendikdasmen No. 13 Tahun 2025 | Asumsi 1 Tahun = 36 Minggu, 1 Jam Pelajaran (JP) = 45 Menit",
        max_col=8
    )

    headers_x = ["No", "Mata Pelajaran", "Alokasi Intrakurikuler (JP/Tahun)", "Alokasi Kokurikuler (JP/Tahun)", "Total JP Per Tahun", "Alokasi JP Per Minggu", "Persentase Kokurikuler", "Keterangan & Opsi Implementasi"]
    r = 4
    for i, h in enumerate(headers_x, 1):
        cell = ws3.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = cell_border
    ws3.row_dimensions[4].height = 26
    r = 5

    data_x = [
        ("SECTION", "KELOMPOK MATA PELAJARAN WAJIB", "", "", "", "", "", ""),
        (1, "Pendidikan Agama dan Budi Pekerti (Islam/Kristen/Katolik/Buddha/Hindu/Khonghucu)", 72, 36, 108, "3 JP (2 Intra + 1 Koku)", "33.3%", "Diikuti sesuai agama/kepercayaan murid."),
        (2, "Pendidikan Pancasila", 72, 0, 72, "2 JP (Intra)", "0.0%", "Fokus penguatan kewargaan dan nilai-nilai konstitusi."),
        (3, "Bahasa Indonesia", 108, 36, 144, "4 JP (3 Intra + 1 Koku)", "25.0%", "Kemampuan literasi teks dan komunikasi lisan/tertulis."),
        (4, "Matematika", 108, 36, 144, "4 JP (3 Intra + 1 Koku)", "25.0%", "Fondasi numerasi, pemodelan matematis, logika."),
        (5, "Ilmu Pengetahuan Alam (IPA: Fisika, Kimia, Biologi)", 216, 108, 324, "9 JP (6 Intra + 3 Koku)", "33.3%", "Terdiri atas Fisika, Kimia, Biologi. Diajarkan terintegrasi/blok/paralel."),
        (6, "Ilmu Pengetahuan Sosial (IPS: Sosiologi, Ekonomi, Sejarah, Geografi)", 288, 144, 432, "12 JP (8 Intra + 4 Koku)", "33.3%", "Terdiri atas Sosiologi, Ekonomi, Sejarah, Geografi. Terintegrasi/blok/paralel."),
        (7, "Bahasa Inggris", 108, 0, 108, "3 JP (Intra)", "0.0%", "Kemampuan komunikasi global, menyimak, membaca, berbicara, menulis."),
        (8, "Pendidikan Jasmani, Olahraga, dan Kesehatan (PJOK)", 72, 36, 108, "3 JP (2 Intra + 1 Koku)", "33.3%", "Aktivitas jasmani, kesehatan fisik & mental, sportivitas."),
        (9, "Informatika", 72, 0, 72, "2 JP (Intra)", "0.0%", "Berpikir komputasional, literasi digital, sistem komputasi."),
        (10, "Seni, Budaya, dan Prakarya (Pilih salah 1):\n  • Seni Musik  • Seni Rupa  • Seni Teater  • Seni Tari\n  • Prakarya Budi Daya  • Kerajinan  • Rekayasa  • Pengolahan", 72, 0, 72, "2 JP (Intra)", "0.0%", "Sekolah wajib menyediakan min. 1 jenis seni/prakarya; murid memilih 1 jenis."),
        ("SUBTOTAL", "Total JP Mata Pelajaran Wajib", 1188, 396, 1584, "44 JP (33 Intra + 11 Koku)", "25.0%", "Beban wajib mingguan siswa kelas X."),
        ("SECTION", "MATA PELAJARAN PILIHAN & MUATAN LOKAL", "", "", "", "", "", ""),
        (11, "Koding dan Kecerdasan Artifisial (Mata Pelajaran Pilihan Baru)", 72, 0, 72, "2 JP (Intra)", "0.0%", "Disediakan bertahap mulai TA 2025/2026 sesuai kesiapan sekolah."),
        (12, "Muatan Lokal (Maksimal)", 72, 0, 72, "2 JP (Intra)", "0.0%", "Maksimal 2 JP/minggu. Potensi lokal bahasa daerah, seni, dll."),
        ("TOTAL1", "Total JP Mapel Wajib + Mapel Pilihan (Koding & AI) ATAU Muatan Lokal", 1260, 396, 1656, "46 JP", "23.9%", "Jika memilih Koding & AI ATAU Muatan Lokal"),
        ("TOTAL2", "Total JP Mapel Wajib + Mapel Pilihan (Koding & AI) + Muatan Lokal", 1332, 396, 1728, "48 JP", "22.9%", "Beban belajar maksimal kelas X jika mengambil keduanya")
    ]

    for item in data_x:
        code = item[0]
        if code == "SECTION":
            ws3.cell(row=r, column=1, value="")
            ws3.cell(row=r, column=2, value=item[1]).font = font_group_header
            ws3.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
            for col in range(1, 9):
                ws3.cell(row=r, column=col).fill = fill_group
                ws3.cell(row=r, column=col).border = cell_border
            ws3.row_dimensions[r].height = 24
        elif code in ["SUBTOTAL", "TOTAL1", "TOTAL2"]:
            ws3.cell(row=r, column=1, value="")
            ws3.cell(row=r, column=2, value=item[1]).font = font_total
            ws3.cell(row=r, column=3, value=item[2]).font = font_total
            ws3.cell(row=r, column=4, value=item[3]).font = font_total
            ws3.cell(row=r, column=5, value=item[4]).font = font_total
            ws3.cell(row=r, column=6, value=item[5]).font = font_total
            ws3.cell(row=r, column=7, value=item[6]).font = font_total
            ws3.cell(row=r, column=8, value=item[7]).font = font_data_bold
            
            fill_target = fill_total if "TOTAL" in code else fill_subgroup
            for col in range(1, 9):
                ws3.cell(row=r, column=col).fill = fill_target
                ws3.cell(row=r, column=col).border = total_border if "TOTAL" in code else cell_border
                if col in [1, 3, 4, 5, 6, 7]:
                    ws3.cell(row=r, column=col).alignment = align_center
                else:
                    ws3.cell(row=r, column=col).alignment = align_left
            ws3.row_dimensions[r].height = 24
        else:
            no, name, intra, koku, tot, mgg, pct, note = item
            ws3.cell(row=r, column=1, value=no).alignment = align_center
            ws3.cell(row=r, column=2, value=name).alignment = align_left
            ws3.cell(row=r, column=3, value=intra).alignment = align_center
            ws3.cell(row=r, column=4, value=koku).alignment = align_center
            ws3.cell(row=r, column=5, value=tot).alignment = align_center
            ws3.cell(row=r, column=6, value=mgg).alignment = align_center
            ws3.cell(row=r, column=7, value=pct).alignment = align_center
            ws3.cell(row=r, column=8, value=note).alignment = align_left

            ws3.cell(row=r, column=1).font = font_data_bold
            ws3.cell(row=r, column=2).font = font_data_bold if "Bahasa Inggris" in name or "Koding" in name else font_data
            for col in range(3, 9):
                ws3.cell(row=r, column=col).font = font_data
                
            for col in range(1, 9):
                ws3.cell(row=r, column=col).border = cell_border
                if r % 2 == 1:
                    ws3.cell(row=r, column=col).fill = fill_zebra
            ws3.row_dimensions[r].height = 36 if "\n" in name else 22
        r += 1

    r += 1
    # Penjelasan Model Pengorganisasian IPA & IPS Kelas X
    ws3.cell(row=r, column=1, value="PENGORGANISASIAN IPA & IPS DI KELAS X (FASE E)").font = font_section_header
    ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    for c in range(1, 9):
        ws3.cell(row=r, column=c).fill = fill_tbl_header
    ws3.row_dimensions[r].height = 24
    r += 1

    ipa_ips_notes = [
        ("Pendekatan 1: Terintegrasi (Integrated)", "Muatan Fisika, Kimia, dan Biologi (IPA) atau Sosiologi, Ekonomi, Sejarah, Geografi (IPS) diajarkan secara terpadu dalam tema-tema lintas disiplin oleh tim guru."),
        ("Pendekatan 2: Sistem Blok Bergantian (Block System)", "Muatan IPA atau IPS diajarkan secara bergantian dalam blok waktu yang terpisah (misal: blok semester 1 Fisika & Biologi, semester 2 Kimia; atau rotasi per 3 bulan)."),
        ("Pendekatan 3: Paralel dengan Unit Inkuiri (Parallel with Inquiry Unit)", "Muatan IPA atau IPS diajarkan secara paralel dengan JP terpisah seperti mapel mandiri, diikuti dengan unit pembelajaran inkuiri kolaboratif yang mengintegrasikan seluruh muatan tersebut.")
    ]
    for p_title, p_desc in ipa_ips_notes:
        ws3.cell(row=r, column=1, value=p_title).font = font_data_bold
        ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        ws3.cell(row=r, column=1).alignment = align_left
        ws3.cell(row=r, column=1).fill = fill_group

        ws3.cell(row=r, column=4, value=p_desc).font = font_data
        ws3.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)
        ws3.cell(row=r, column=4).alignment = align_left

        for c in range(1, 9):
            ws3.cell(row=r, column=c).border = cell_border
            if c >= 4 and r % 2 == 1:
                ws3.cell(row=r, column=c).fill = fill_zebra
        ws3.row_dimensions[r].height = 24
        r += 1

    auto_fit_columns(ws3, min_width=10, max_width=45)
    ws3.column_dimensions['A'].width = 5
    ws3.column_dimensions['B'].width = 36
    ws3.column_dimensions['C'].width = 16
    ws3.column_dimensions['D'].width = 16
    ws3.column_dimensions['E'].width = 14
    ws3.column_dimensions['F'].width = 18
    ws3.column_dimensions['G'].width = 14
    ws3.column_dimensions['H'].width = 38
    ws3.freeze_panes = 'A5'

    # ==========================================
    # SHEET 4: KELAS XI (FASE F)
    # ==========================================
    ws4 = wb.create_sheet(title="Kelas XI (Fase F)")
    apply_title_banner(
        ws4,
        "STRUKTUR KURIKULUM KELAS XI SMA / MA (FASE F)",
        "Tabel 9 Permendikdasmen No. 13 Tahun 2025 | Asumsi 1 Tahun = 36 Minggu, 1 Jam Pelajaran (JP) = 45 Menit",
        max_col=8
    )

    headers_xi = ["No", "Mata Pelajaran", "Alokasi Intrakurikuler (JP/Tahun)", "Alokasi Kokurikuler (JP/Tahun)", "Total JP Per Tahun", "Alokasi JP Per Minggu", "Kategori Mapel", "Aturan & Ketentuan Pemilihan"]
    r = 4
    for i, h in enumerate(headers_xi, 1):
        cell = ws4.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = cell_border
    ws4.row_dimensions[4].height = 26
    r = 5

    data_xi_wajib = [
        (1, "Pendidikan Agama dan Budi Pekerti", 72, 36, 108, "3 JP (2 Intra + 1 Koku)", "Wajib", "Diikuti sesuai agama/kepercayaan masing-masing murid."),
        (2, "Pendidikan Pancasila", 72, 0, 72, "2 JP (Intra)", "Wajib", "Penguatan kewargaan, konstitusi, norma berbangsa."),
        (3, "Bahasa Indonesia", 108, 36, 144, "4 JP (3 Intra + 1 Koku)", "Wajib", "Literasi akademis, teks argumentatif, riset ilmiah."),
        (4, "Matematika", 108, 36, 144, "4 JP (3 Intra + 1 Koku)", "Wajib", "Matematika wajib untuk semua peminatan siswa."),
        (5, "Bahasa Inggris", 108, 0, 108, "3 JP (Intra)", "Wajib", "Bahasa Inggris wajib (seluruh siswa wajib mengikuti)."),
        (6, "Pendidikan Jasmani Olahraga dan Kesehatan", 72, 36, 108, "3 JP (2 Intra + 1 Koku)", "Wajib", "Kebugaran fisik, gaya hidup sehat, pembiasaan gerak aktif."),
        (7, "Sejarah", 72, 0, 72, "2 JP (Intra)", "Wajib", "Mata pelajaran wajib sejarah nasional dan kontekstual."),
        (8, "Seni dan Budaya (Pilih salah 1: Musik, Rupa, Teater, Tari)", 72, 0, 72, "2 JP (Intra)", "Wajib", "Sekolah menyediakan min. 1 jenis seni; murid memilih 1."),
        ("SUBTOTAL", "Total JP Mata Pelajaran Wajib Kelas XI", 684, 144, 828, "23 JP (19 Intra + 4 Koku)", "Subtotal Wajib", "Wajib diikuti oleh seluruh peserta didik Kelas XI")
    ]

    ws4.cell(row=r, column=1, value="")
    ws4.cell(row=r, column=2, value="A. KELOMPOK MATA PELAJARAN WAJIB").font = font_group_header
    ws4.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    for col in range(1, 9):
        ws4.cell(row=r, column=col).fill = fill_group
        ws4.cell(row=r, column=col).border = cell_border
    ws4.row_dimensions[r].height = 24
    r += 1

    for item in data_xi_wajib:
        code = item[0]
        if code == "SUBTOTAL":
            ws4.cell(row=r, column=1, value="")
            ws4.cell(row=r, column=2, value=item[1]).font = font_total
            ws4.cell(row=r, column=3, value=item[2]).font = font_total
            ws4.cell(row=r, column=4, value=item[3]).font = font_total
            ws4.cell(row=r, column=5, value=item[4]).font = font_total
            ws4.cell(row=r, column=6, value=item[5]).font = font_total
            ws4.cell(row=r, column=7, value=item[6]).font = font_total
            ws4.cell(row=r, column=8, value=item[7]).font = font_data_bold
            for col in range(1, 9):
                ws4.cell(row=r, column=col).fill = fill_subgroup
                ws4.cell(row=r, column=col).border = cell_border
                if col in [1, 3, 4, 5, 6, 7]:
                    ws4.cell(row=r, column=col).alignment = align_center
                else:
                    ws4.cell(row=r, column=col).alignment = align_left
            ws4.row_dimensions[r].height = 24
        else:
            no, name, intra, koku, tot, mgg, cat, note = item
            ws4.cell(row=r, column=1, value=no).alignment = align_center
            ws4.cell(row=r, column=2, value=name).alignment = align_left
            ws4.cell(row=r, column=3, value=intra).alignment = align_center
            ws4.cell(row=r, column=4, value=koku).alignment = align_center
            ws4.cell(row=r, column=5, value=tot).alignment = align_center
            ws4.cell(row=r, column=6, value=mgg).alignment = align_center
            ws4.cell(row=r, column=7, value=cat).alignment = align_center
            ws4.cell(row=r, column=8, value=note).alignment = align_left
            ws4.cell(row=r, column=1).font = font_data_bold
            ws4.cell(row=r, column=2).font = font_data_bold if "Bahasa Inggris" in name else font_data
            for col in range(3, 9):
                ws4.cell(row=r, column=col).font = font_data
            for col in range(1, 9):
                ws4.cell(row=r, column=col).border = cell_border
                if r % 2 == 1:
                    ws4.cell(row=r, column=col).fill = fill_zebra
            ws4.row_dimensions[r].height = 22
        r += 1

    # Mapel Pilihan Kelas XI
    ws4.cell(row=r, column=1, value="")
    ws4.cell(row=r, column=2, value="B. KELOMPOK MATA PELAJARAN PILIHAN (Sekolah WAJIB buka min. 7 mapel; Siswa memilih 4 s.d. 5 mapel)").font = font_group_header
    ws4.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    for col in range(1, 9):
        ws4.cell(row=r, column=col).fill = fill_group
        ws4.cell(row=r, column=col).border = cell_border
    ws4.row_dimensions[r].height = 24
    r += 1

    data_xi_pilihan = [
        (9, "Matematika Tingkat Lanjut", 180, 0, 180, "5 JP", "Pilihan MIPA", "Alokasi 5 JP/minggu. Pendalaman kalkulus, aljabar linear, geometri."),
        (10, "Fisika", 180, 0, 180, "5 JP", "Pilihan MIPA", "Alokasi 5 JP/minggu. Mekanika, termodinamika, gelombang, listrik magnet."),
        (11, "Kimia", 180, 0, 180, "5 JP", "Pilihan MIPA", "Alokasi 5 JP/minggu. Struktur atom, ikatan kimia, stoikiometri, larutan."),
        (12, "Biologi", 180, 0, 180, "5 JP", "Pilihan MIPA", "Alokasi 5 JP/minggu. Biologi sel, genetika, metabolisme, ekosistem."),
        (13, "Informatika", 180, 0, 180, "5 JP", "Pilihan MIPA/Teknologi", "Alokasi 5 JP/minggu. Algoritma, analisis data, jaringan, rekayasa software."),
        (14, "Koding dan Kecerdasan Artifisial (AI)", 180, 0, 180, "5 JP", "Pilihan Teknologi", "MAPEL BARU Permendikdasmen 13/2025 (5 JP/mgg). Pembelajaran koding & AI praktis."),
        (15, "Geografi", 180, 0, 180, "5 JP", "Pilihan IPS", "Alokasi 5 JP/minggu. Geografi fisik, sosial, SIG, dinamika geosfer."),
        (16, "Sosiologi", 180, 0, 180, "5 JP", "Pilihan IPS", "Alokasi 5 JP/minggu. Struktur sosial, perubahan sosial, konflik, integrasi."),
        (17, "Ekonomi", 180, 0, 180, "5 JP", "Pilihan IPS", "Alokasi 5 JP/minggu. Makro-mikro ekonomi, moneter, fiskal, akuntansi."),
        (18, "Sejarah Tingkat Lanjut", 180, 0, 180, "5 JP", "Pilihan IPS/Humaniora", "Alokasi 5 JP/minggu. Sejarah dunia, historiografi, riset sejarah kritis."),
        (19, "Antropologi", 180, 0, 180, "5 JP", "Pilihan IPS/Humaniora", "Alokasi 5 JP/minggu. Kebudayaan, keragaman etnografi, dinamika tradisi."),
        (20, "Bahasa Indonesia Tingkat Lanjut", 180, 0, 180, "5 JP", "Pilihan Bahasa", "Alokasi 5 JP/minggu. Kajian sastra mendalam, linguistik, penulisan kreatif."),
        (21, "Bahasa Inggris Tingkat Lanjut", 180, 0, 180, "5 JP", "Pilihan Bahasa", "Alokasi 5 JP/minggu. Academic English, critical discourse, literary texts."),
        (22, "Bahasa Arab", 180, 0, 180, "5 JP", "Pilihan Bahasa Asing", "Alokasi 5 JP/minggu. Kemampuan komunikatif dan gramatika bahasa Arab."),
        (23, "Bahasa Jepang", 180, 0, 180, "5 JP", "Pilihan Bahasa Asing", "Alokasi 5 JP/minggu. Huruf Kana/Kanji, percakapan, kebudayaan Jepang."),
        (24, "Bahasa Jerman", 180, 0, 180, "5 JP", "Pilihan Bahasa Asing", "Alokasi 5 JP/minggu. Struktur bahasa dan komunikasi standar bahasa Jerman."),
        (25, "Bahasa Korea", 180, 0, 180, "5 JP", "Pilihan Bahasa Asing", "Alokasi 5 JP/minggu. Hangul, ungkapan komunikatif, pemahaman kultural."),
        (26, "Bahasa Mandarin", 180, 0, 180, "5 JP", "Pilihan Bahasa Asing", "Alokasi 5 JP/minggu. Pinyin, Hanzi, nada dan percakapan bahasa Mandarin."),
        (27, "Bahasa Prancis", 180, 0, 180, "5 JP", "Pilihan Bahasa Asing", "Alokasi 5 JP/minggu. Kosa kata, gramatika, dan komunikasi bahasa Prancis."),
        (28, "Prakarya dan Kewirausahaan (PKWU)", 72, 0, 72, "2 JP", "Pilihan Vokasi", "Alokasi 2 JP/minggu (72 JP/tahun). Pilihan: budi daya/kerajinan/rekayasa/pengolahan."),
        (29, "Mata Pelajaran Lainnya (Pengembangan Satuan Pendidikan)", 180, 0, 180, "5 JP", "Pilihan Lokal/Khusus", "Dikembangkan mandiri oleh sekolah sesuai sumber daya yang tersedia.")
    ]

    for item in data_xi_pilihan:
        no, name, intra, koku, tot, mgg, cat, note = item
        ws4.cell(row=r, column=1, value=no).alignment = align_center
        ws4.cell(row=r, column=2, value=name).alignment = align_left
        ws4.cell(row=r, column=3, value=intra).alignment = align_center
        ws4.cell(row=r, column=4, value=koku).alignment = align_center
        ws4.cell(row=r, column=5, value=tot).alignment = align_center
        ws4.cell(row=r, column=6, value=mgg).alignment = align_center
        ws4.cell(row=r, column=7, value=cat).alignment = align_center
        ws4.cell(row=r, column=8, value=note).alignment = align_left
        ws4.cell(row=r, column=1).font = font_data_bold
        ws4.cell(row=r, column=2).font = font_data_bold if "Bahasa Inggris" in name or "Koding" in name else font_data
        for col in range(3, 9):
            ws4.cell(row=r, column=col).font = font_data
        for col in range(1, 9):
            ws4.cell(row=r, column=col).border = cell_border
            if r % 2 == 1:
                ws4.cell(row=r, column=col).fill = fill_zebra
        ws4.row_dimensions[r].height = 22
        r += 1

    # Totals Kelas XI
    totals_xi = [
        ("TOTAL_PIL", "Total Beban Mapel Pilihan yang Dipilih Siswa (4 s.d. 5 Mapel)", "720 - 900", 0, "720 - 900", "20 - 25 JP", "Subtotal Pilihan", "Siswa memilih 4-5 mapel pilihan setara 20-25 JP/minggu."),
        ("TOTAL_W_P", "Total JP Mata Pelajaran Wajib + Mata Pelajaran Pilihan", "1.404 - 1.584", 144, "1.548 - 1.728", "43 - 48 JP", "Total Inti", "Rentang alokasi reguler kelas XI."),
        ("MULOK", "Muatan Lokal (Maksimal)", 72, 0, 72, "2 JP", "Muatan Lokal", "Paling banyak 2 JP per minggu (72 JP/tahun)."),
        ("GRAND_TOT", "TOTAL KESELURUHAN (WAJIB + PILIHAN + MUATAN LOKAL)", "1.476 - 1.656", 144, "1.620 - 1.800", "45 - 50 JP", "Grand Total", "Total beban belajar tahunan kelas XI.")
    ]

    for item in totals_xi:
        code, name, intra, koku, tot, mgg, cat, note = item
        ws4.cell(row=r, column=1, value="")
        ws4.cell(row=r, column=2, value=name).font = font_total
        ws4.cell(row=r, column=3, value=intra).font = font_total
        ws4.cell(row=r, column=4, value=koku).font = font_total
        ws4.cell(row=r, column=5, value=tot).font = font_total
        ws4.cell(row=r, column=6, value=mgg).font = font_total
        ws4.cell(row=r, column=7, value=cat).font = font_total
        ws4.cell(row=r, column=8, value=note).font = font_data_bold

        fill_target = fill_total if code in ["TOTAL_W_P", "GRAND_TOT"] else fill_subgroup
        for col in range(1, 9):
            ws4.cell(row=r, column=col).fill = fill_target
            ws4.cell(row=r, column=col).border = total_border if code in ["TOTAL_W_P", "GRAND_TOT"] else cell_border
            if col in [1, 3, 4, 5, 6, 7]:
                ws4.cell(row=r, column=col).alignment = align_center
            else:
                ws4.cell(row=r, column=col).alignment = align_left
        ws4.row_dimensions[r].height = 24
        r += 1

    auto_fit_columns(ws4, min_width=10, max_width=45)
    ws4.column_dimensions['A'].width = 5
    ws4.column_dimensions['B'].width = 36
    ws4.column_dimensions['C'].width = 16
    ws4.column_dimensions['D'].width = 16
    ws4.column_dimensions['E'].width = 15
    ws4.column_dimensions['F'].width = 16
    ws4.column_dimensions['G'].width = 16
    ws4.column_dimensions['H'].width = 38
    ws4.freeze_panes = 'A5'

    # ==========================================
    # SHEET 5: KELAS XII (FASE F)
    # ==========================================
    ws5 = wb.create_sheet(title="Kelas XII (Fase F)")
    apply_title_banner(
        ws5,
        "STRUKTUR KURIKULUM KELAS XII SMA / MA (FASE F)",
        "Tabel 10 Permendikdasmen No. 13 Tahun 2025 | Asumsi 1 Tahun = 32 Minggu (Waktu Ujian/Kelulusan), 1 JP = 45 Menit",
        max_col=8
    )

    headers_xii = ["No", "Mata Pelajaran", "Alokasi Intrakurikuler (JP/Tahun)", "Alokasi Kokurikuler (JP/Tahun)", "Total JP Per Tahun", "Alokasi JP Per Minggu", "Kategori Mapel", "Aturan & Ketentuan Pemilihan"]
    r = 4
    for i, h in enumerate(headers_xii, 1):
        cell = ws5.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = cell_border
    ws5.row_dimensions[4].height = 26
    r = 5

    data_xii_wajib = [
        (1, "Pendidikan Agama dan Budi Pekerti", 64, 32, 96, "3 JP (2 Intra + 1 Koku)", "Wajib", "Diikuti sesuai agama/kepercayaan masing-masing murid."),
        (2, "Pendidikan Pancasila", 64, 0, 64, "2 JP (Intra)", "Wajib", "Penguatan kewargaan, etika publik, bela negara."),
        (3, "Bahasa Indonesia", 96, 32, 128, "4 JP (3 Intra + 1 Koku)", "Wajib", "Literasi tingkat lanjut, esai ilmiah, karya tulis."),
        (4, "Matematika", 96, 32, 128, "4 JP (3 Intra + 1 Koku)", "Wajib", "Matematika wajib umum untuk persiapan jenjang berikutnya."),
        (5, "Bahasa Inggris", 96, 0, 96, "3 JP (Intra)", "Wajib", "Bahasa Inggris wajib (kemampuan komunikasi internasional)."),
        (6, "Seni dan Budaya (Pilih salah 1: Musik, Rupa, Teater, Tari)", 64, 0, 64, "2 JP (Intra)", "Wajib", "Sekolah menyediakan min. 1 jenis seni; murid memilih 1."),
        (7, "Pendidikan Jasmani Olahraga dan Kesehatan", 64, 32, 96, "3 JP (2 Intra + 1 Koku)", "Wajib", "Kebugaran jasmani, kesehatan reproduksi, gaya hidup aktif."),
        (8, "Sejarah", 64, 0, 64, "2 JP (Intra)", "Wajib", "Sejarah Indonesia kontemporer dan dinamika global."),
        ("SUBTOTAL", "Total JP Mata Pelajaran Wajib Kelas XII", 608, 128, 736, "23 JP (19 Intra + 4 Koku)", "Subtotal Wajib", "Wajib diikuti oleh seluruh peserta didik Kelas XII")
    ]

    ws5.cell(row=r, column=1, value="")
    ws5.cell(row=r, column=2, value="A. KELOMPOK MATA PELAJARAN WAJIB").font = font_group_header
    ws5.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    for col in range(1, 9):
        ws5.cell(row=r, column=col).fill = fill_group
        ws5.cell(row=r, column=col).border = cell_border
    ws5.row_dimensions[r].height = 24
    r += 1

    for item in data_xii_wajib:
        code = item[0]
        if code == "SUBTOTAL":
            ws5.cell(row=r, column=1, value="")
            ws5.cell(row=r, column=2, value=item[1]).font = font_total
            ws5.cell(row=r, column=3, value=item[2]).font = font_total
            ws5.cell(row=r, column=4, value=item[3]).font = font_total
            ws5.cell(row=r, column=5, value=item[4]).font = font_total
            ws5.cell(row=r, column=6, value=item[5]).font = font_total
            ws5.cell(row=r, column=7, value=item[6]).font = font_total
            ws5.cell(row=r, column=8, value=item[7]).font = font_data_bold
            for col in range(1, 9):
                ws5.cell(row=r, column=col).fill = fill_subgroup
                ws5.cell(row=r, column=col).border = cell_border
                if col in [1, 3, 4, 5, 6, 7]:
                    ws5.cell(row=r, column=col).alignment = align_center
                else:
                    ws5.cell(row=r, column=col).alignment = align_left
            ws5.row_dimensions[r].height = 24
        else:
            no, name, intra, koku, tot, mgg, cat, note = item
            ws5.cell(row=r, column=1, value=no).alignment = align_center
            ws5.cell(row=r, column=2, value=name).alignment = align_left
            ws5.cell(row=r, column=3, value=intra).alignment = align_center
            ws5.cell(row=r, column=4, value=koku).alignment = align_center
            ws5.cell(row=r, column=5, value=tot).alignment = align_center
            ws5.cell(row=r, column=6, value=mgg).alignment = align_center
            ws5.cell(row=r, column=7, value=cat).alignment = align_center
            ws5.cell(row=r, column=8, value=note).alignment = align_left
            ws5.cell(row=r, column=1).font = font_data_bold
            ws5.cell(row=r, column=2).font = font_data_bold if "Bahasa Inggris" in name else font_data
            for col in range(3, 9):
                ws5.cell(row=r, column=col).font = font_data
            for col in range(1, 9):
                ws5.cell(row=r, column=col).border = cell_border
                if r % 2 == 1:
                    ws5.cell(row=r, column=col).fill = fill_zebra
            ws5.row_dimensions[r].height = 22
        r += 1

    # Mapel Pilihan Kelas XII
    ws5.cell(row=r, column=1, value="")
    ws5.cell(row=r, column=2, value="B. KELOMPOK MATA PELAJARAN PILIHAN (Alokasi 5 JP/minggu = 160 JP/tahun; PKWU = 2 JP/minggu = 64 JP/tahun)").font = font_group_header
    ws5.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    for col in range(1, 9):
        ws5.cell(row=r, column=col).fill = fill_group
        ws5.cell(row=r, column=col).border = cell_border
    ws5.row_dimensions[r].height = 24
    r += 1

    data_xii_pilihan = [
        (9, "Matematika Tingkat Lanjut", 160, 0, 160, "5 JP", "Pilihan MIPA", "Alokasi 5 JP/minggu (160 JP/tahun di kelas XII)."),
        (10, "Fisika", 160, 0, 160, "5 JP", "Pilihan MIPA", "Alokasi 5 JP/minggu (160 JP/tahun di kelas XII)."),
        (11, "Kimia", 160, 0, 160, "5 JP", "Pilihan MIPA", "Alokasi 5 JP/minggu (160 JP/tahun di kelas XII)."),
        (12, "Biologi", 160, 0, 160, "5 JP", "Pilihan MIPA", "Alokasi 5 JP/minggu (160 JP/tahun di kelas XII)."),
        (13, "Informatika", 160, 0, 160, "5 JP", "Pilihan MIPA/Teknologi", "Alokasi 5 JP/minggu (160 JP/tahun di kelas XII)."),
        (14, "Koding dan Kecerdasan Artifisial (AI)", 160, 0, 160, "5 JP", "Pilihan Teknologi", "MAPEL BARU Permendikdasmen 13/2025. Alokasi 5 JP/mgg (160 JP/thn)."),
        (15, "Geografi", 160, 0, 160, "5 JP", "Pilihan IPS", "Alokasi 5 JP/minggu (160 JP/tahun di kelas XII)."),
        (16, "Sosiologi", 160, 0, 160, "5 JP", "Pilihan IPS", "Alokasi 5 JP/minggu (160 JP/tahun di kelas XII)."),
        (17, "Ekonomi", 160, 0, 160, "5 JP", "Pilihan IPS", "Alokasi 5 JP/minggu (160 JP/tahun di kelas XII)."),
        (18, "Sejarah Tingkat Lanjut", 160, 0, 160, "5 JP", "Pilihan IPS/Humaniora", "Alokasi 5 JP/minggu (160 JP/tahun di kelas XII)."),
        (19, "Antropologi", 160, 0, 160, "5 JP", "Pilihan IPS/Humaniora", "Alokasi 5 JP/minggu (160 JP/tahun di kelas XII)."),
        (20, "Bahasa Indonesia Tingkat Lanjut", 160, 0, 160, "5 JP", "Pilihan Bahasa", "Alokasi 5 JP/minggu (160 JP/tahun di kelas XII)."),
        (21, "Bahasa Inggris Tingkat Lanjut", 160, 0, 160, "5 JP", "Pilihan Bahasa", "Alokasi 5 JP/minggu (160 JP/tahun di kelas XII). Peminatan bahasa."),
        (22, "Bahasa Arab", 160, 0, 160, "5 JP", "Pilihan Bahasa Asing", "Alokasi 5 JP/minggu (160 JP/tahun di kelas XII)."),
        (23, "Bahasa Jepang", 160, 0, 160, "5 JP", "Pilihan Bahasa Asing", "Alokasi 5 JP/minggu (160 JP/tahun di kelas XII)."),
        (24, "Bahasa Jerman", 160, 0, 160, "5 JP", "Pilihan Bahasa Asing", "Alokasi 5 JP/minggu (160 JP/tahun di kelas XII)."),
        (25, "Bahasa Korea", 160, 0, 160, "5 JP", "Pilihan Bahasa Asing", "Alokasi 5 JP/minggu (160 JP/tahun di kelas XII)."),
        (26, "Bahasa Mandarin", 160, 0, 160, "5 JP", "Pilihan Bahasa Asing", "Alokasi 5 JP/minggu (160 JP/tahun di kelas XII)."),
        (27, "Bahasa Prancis", 160, 0, 160, "5 JP", "Pilihan Bahasa Asing", "Alokasi 5 JP/minggu (160 JP/tahun di kelas XII)."),
        (28, "Prakarya dan Kewirausahaan (PKWU)", 64, 0, 64, "2 JP", "Pilihan Vokasi", "Alokasi 2 JP/minggu (64 JP/tahun di kelas XII)."),
        (29, "Mata Pelajaran Lainnya (Pengembangan Satuan Pendidikan)", 160, 0, 160, "5 JP", "Pilihan Lokal/Khusus", "Dikembangkan mandiri sesuai sumber daya satuan pendidikan.")
    ]

    for item in data_xii_pilihan:
        no, name, intra, koku, tot, mgg, cat, note = item
        ws5.cell(row=r, column=1, value=no).alignment = align_center
        ws5.cell(row=r, column=2, value=name).alignment = align_left
        ws5.cell(row=r, column=3, value=intra).alignment = align_center
        ws5.cell(row=r, column=4, value=koku).alignment = align_center
        ws5.cell(row=r, column=5, value=tot).alignment = align_center
        ws5.cell(row=r, column=6, value=mgg).alignment = align_center
        ws5.cell(row=r, column=7, value=cat).alignment = align_center
        ws5.cell(row=r, column=8, value=note).alignment = align_left
        ws5.cell(row=r, column=1).font = font_data_bold
        ws5.cell(row=r, column=2).font = font_data_bold if "Bahasa Inggris" in name or "Koding" in name else font_data
        for col in range(3, 9):
            ws5.cell(row=r, column=col).font = font_data
        for col in range(1, 9):
            ws5.cell(row=r, column=col).border = cell_border
            if r % 2 == 1:
                ws5.cell(row=r, column=col).fill = fill_zebra
        ws5.row_dimensions[r].height = 22
        r += 1

    # Totals Kelas XII
    totals_xii = [
        ("TOTAL_PIL", "Total Beban Mapel Pilihan yang Dipilih Siswa (4 s.d. 5 Mapel)", "640 - 800", 0, "640 - 800", "20 - 25 JP", "Subtotal Pilihan", "Siswa memilih 4-5 mapel pilihan setara 20-25 JP/minggu."),
        ("TOTAL_W_P", "Total JP Mata Pelajaran Wajib + Mata Pelajaran Pilihan", "1.248 - 1.408", 128, "1.376 - 1.536", "43 - 48 JP", "Total Inti", "Rentang alokasi reguler kelas XII."),
        ("MULOK", "Muatan Lokal (Maksimal)", 64, 0, 64, "2 JP", "Muatan Lokal", "Paling banyak 2 JP per minggu (64 JP/tahun pada asumsi 32 minggu)."),
        ("GRAND_TOT", "TOTAL KESELURUHAN (WAJIB + PILIHAN + MUATAN LOKAL)", "1.312 - 1.472", 128, "1.440 - 1.600", "45 - 50 JP", "Grand Total", "Total beban belajar tahunan kelas XII.")
    ]

    for item in totals_xii:
        code, name, intra, koku, tot, mgg, cat, note = item
        ws5.cell(row=r, column=1, value="")
        ws5.cell(row=r, column=2, value=name).font = font_total
        ws5.cell(row=r, column=3, value=intra).font = font_total
        ws5.cell(row=r, column=4, value=koku).font = font_total
        ws5.cell(row=r, column=5, value=tot).font = font_total
        ws5.cell(row=r, column=6, value=mgg).font = font_total
        ws5.cell(row=r, column=7, value=cat).font = font_total
        ws5.cell(row=r, column=8, value=note).font = font_data_bold

        fill_target = fill_total if code in ["TOTAL_W_P", "GRAND_TOT"] else fill_subgroup
        for col in range(1, 9):
            ws5.cell(row=r, column=col).fill = fill_target
            ws5.cell(row=r, column=col).border = total_border if code in ["TOTAL_W_P", "GRAND_TOT"] else cell_border
            if col in [1, 3, 4, 5, 6, 7]:
                ws5.cell(row=r, column=col).alignment = align_center
            else:
                ws5.cell(row=r, column=col).alignment = align_left
        ws5.row_dimensions[r].height = 24
        r += 1

    auto_fit_columns(ws5, min_width=10, max_width=45)
    ws5.column_dimensions['A'].width = 5
    ws5.column_dimensions['B'].width = 36
    ws5.column_dimensions['C'].width = 16
    ws5.column_dimensions['D'].width = 16
    ws5.column_dimensions['E'].width = 15
    ws5.column_dimensions['F'].width = 16
    ws5.column_dimensions['G'].width = 16
    ws5.column_dimensions['H'].width = 38
    ws5.freeze_panes = 'A5'

    # ==========================================
    # SHEET 6: DEEP LEARNING & PROFIL LULUSAN
    # ==========================================
    ws6 = wb.create_sheet(title="Deep Learning & Profil")
    apply_title_banner(
        ws6,
        "KERANGKA DASAR KURIKULUM: PEMBELAJARAN MENDALAM & PROFIL LULUSAN",
        "Lampiran I Permendikdasmen No. 13 Tahun 2025 | Pendekatan Baru Pembelajaran & 8 Dimensi Profil Lulusan",
        max_col=5
    )

    ws6.cell(row=4, column=1, value="1. TIGA PRINSIP UTAMA PEMBELAJARAN MENDALAM (DEEP LEARNING)").font = font_section_header
    ws6.merge_cells("A4:E4")
    for c in range(1, 6):
        ws6.cell(row=4, column=c).fill = fill_tbl_header
    ws6.row_dimensions[4].height = 24

    headers_dl = ["No", "Prinsip Pembelajaran Mendalam", "Makna & Konseptualisasi", "Peran Pendidik (Guru)", "Pengalaman Belajar Siswa"]
    r = 5
    for i, h in enumerate(headers_dl, 1):
        cell = ws6.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws6.row_dimensions[5].height = 24
    r = 6

    dl_principles = [
        (1, "Berkesadaran (Mindful Learning)", "Peserta didik menjadi pemelajar yang aktif dan mampu meregulasi diri secara mandiri. Memahami tujuan belajar secara sadar, memiliki motivasi intrinsik kuat, dan menyusun strategi belajar efektif sebagai pembelajar sepanjang hayat.", "Memfasilitasi refleksi diri, memandu penetapan target belajar pribadi, menumbuhkan regulasi metakognisi, membangun kesadaran 'mengapa' materi dipelajari.", "Siswa tidak sekadar menerima instruksi pasif, melainkan menyadari tahapan berpikirnya dan mengontrol ritme pemahamannya."),
        (2, "Bermakna (Meaningful Learning)", "Pembelajaran terhubung langsung dengan isu nyata, konteks sosial, budaya, dan lingkungan peserta didik. Menolak hafalan konten semata; berorientasi pada retensi jangka panjang melalui penerapan pengetahuan di dunia nyata.", "Menghubungkan materi ajar dengan fenomena lokal/global, mendesain studi kasus otentik, melibatkan masyarakat/orang tua sebagai sumber belajar praktis.", "Siswa memahami relevansi ilmu bagi kehidupannya, mampu menerapkan konsep ke situasi baru, dan terdorong berkontribusi bagi lingkungan sekitar."),
        (3, "Menggembirakan (Joyful Learning)", "Menciptakan iklim belajar yang positif, menantang, menggembirakan, dan aman secara emosional. Menghilangkan suasana intimidatif dan rasa takut salah, sehingga rasa ingin tahu alami dan kreativitas tumbuh optimal.", "Menciptakan ruang kelas inklusif, memberi apresiasi tulus, merespons kebutuhan fisiologis & psikologis anak, mengintegrasikan metode interaktif inovatif.", "Siswa menikmati proses belajar, merasa dihargai dan aman berekspresi, berani mencoba hal baru, serta terhubung secara emosional dengan ilmu.")
    ]

    for item in dl_principles:
        no, name, makna, guru, siswa = item
        ws6.cell(row=r, column=1, value=no).alignment = align_center
        ws6.cell(row=r, column=2, value=name).alignment = align_left
        ws6.cell(row=r, column=3, value=makna).alignment = align_left
        ws6.cell(row=r, column=4, value=guru).alignment = align_left
        ws6.cell(row=r, column=5, value=siswa).alignment = align_left

        ws6.cell(row=r, column=1).font = font_data_bold
        ws6.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 6):
            ws6.cell(row=r, column=c).font = font_data
            
        for c in range(1, 6):
            ws6.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws6.cell(row=r, column=c).fill = fill_zebra
        ws6.row_dimensions[r].height = 54
        r += 1

    r += 1
    # 8 Dimensi Profil Lulusan
    ws6.cell(row=r, column=1, value="2. DELAPAN (8) DIMENSI PROFIL LULUSAN").font = font_section_header
    ws6.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    for c in range(1, 6):
        ws6.cell(row=r, column=c).fill = fill_tbl_header
    ws6.row_dimensions[r].height = 24
    r += 1

    headers_profil = ["No", "Dimensi Profil Lulusan", "Fokus Capaian Karakter & Kompetensi", "Keterkaitan dengan Mata Pelajaran", "Implementasi Nyata di SMA"]
    for i, h in enumerate(headers_profil, 1):
        cell = ws6.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws6.row_dimensions[r].height = 24
    r += 1

    profil_dimensions = [
        (1, "Keimanan dan Ketakwaan kepada Tuhan YME", "Membentuk akhlak mulia, integritas moral spiritual, penghayatan nilai agama, dan toleransi antarumat.", "Pendidikan Agama & Budi Pekerti, seluruh pembiasaan kultur sekolah.", "Ibadah harian tertib, adab saling memuliakan, integritas kejujuran akademik."),
        (2, "Kewargaan", "Kesadaran berbangsa, cinta tanah air, pemahaman hak dan kewajiban konstitusional, serta keterlibatan aktif dalam masyarakat.", "Pendidikan Pancasila, Sejarah, Sosiologi, Geografi.", "Partisipasi pemilihan OSIS, advokasi isu komunitas, kepatuhan norma hukum."),
        (3, "Penalaran Kritis", "Kemampuan memproses informasi secara objektif, menganalisis data, mengevaluasi bukti, dan memecahkan masalah kompleks.", "Matematika, IPA (Fisika, Kimia, Biologi), Informatika, Koding & AI.", "Riset berbasis inkuiri, pemecahan masalah ilmiah, analisis data objektif."),
        (4, "Kreativitas", "Kemampuan menghasilkan gagasan orisinal, solusi inovatif, daya cipta seni, serta kemampuan melihat perspektif baru.", "Seni & Budaya, Prakarya & Kewirausahaan (PKWU), Proyek Kokurikuler.", "Pembuatan prototipe karya, pameran seni, penulisan kreatif sastra."),
        (5, "Kolaborasi", "Kemampuan bekerja sama dalam tim lintas latar belakang, komunikasi terbuka, empati, dan gotong royong mencapai tujuan.", "Seluruh mata pelajaran, proyek kolaboratif lintas disiplin, olahraga.", "Kerja kelompok bergiliran peran, proyek tim lintas mapel, kegiatan sosial."),
        (6, "Kemandirian", "Tanggung jawab atas proses dan hasil belajar, regulasi emosi, kegigihan (resilience), dan inisiatif tanpa selalu diarahkan.", "Seluruh mata pelajaran, Bimbingan & Konseling, Ekstrakurikuler.", "Manajemen waktu mandiri, penyelesaian portofolio, penetapan target studi lanjut."),
        (7, "Kesehatan (Dimensi Baru)", "Kesadaran memelihara kesehatan raga, kebugaran fisik, pola makan sehat bergizi, dan kesejahteraan emosional (mental wellbeing).", "PJOK, IPA (Biologi), Gerakan 7 Kebiasaan Anak Indonesia Hebat.", "Pembiasaan senam/olahraga rutin, gizi seimbang, literasi kesehatan mental remaja."),
        (8, "Komunikasi (Dimensi Baru)", "Kecakapan menyampaikan gagasan secara jelas, santun, efektif baik lisan, tulisan, maupun multimodal digital dalam bahasa nasional dan internasional.", "Bahasa Indonesia, Bahasa Inggris, Bahasa Asing, Informatika.", "Presentasi publik, debat akademis, korespondensi formal, publikasi esai bahasa Inggris.")
    ]

    for item in profil_dimensions:
        no, name, fokus, mapel, imp = item
        ws6.cell(row=r, column=1, value=no).alignment = align_center
        ws6.cell(row=r, column=2, value=name).alignment = align_left
        ws6.cell(row=r, column=3, value=fokus).alignment = align_left
        ws6.cell(row=r, column=4, value=mapel).alignment = align_left
        ws6.cell(row=r, column=5, value=imp).alignment = align_left

        ws6.cell(row=r, column=1).font = font_data_bold
        ws6.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 6):
            ws6.cell(row=r, column=c).font = font_data
            
        for c in range(1, 6):
            ws6.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws6.cell(row=r, column=c).fill = fill_zebra
        ws6.row_dimensions[r].height = 36
        r += 1

    r += 1
    # Nilai Filosofis Pemuliaan
    ws6.cell(row=r, column=1, value="3. NILAI FILOSOFIS PENDUKUNG: SALING MEMULIAKAN & SISTEM AMONG").font = font_section_header
    ws6.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    for c in range(1, 6):
        ws6.cell(row=r, column=c).fill = fill_tbl_header
    ws6.row_dimensions[r].height = 24
    r += 1

    filosofi_data = [
        ("K.H. Hasyim Asy'ari (Saling Memuliakan)", "Menempatkan penghormatan sebagai inti proses pendidikan: Guru dihormati sebagai pembimbing penuh kasih sayang, teman sejawat dihargai dalam semangat kolaborasi, dan ilmu dirawat dengan sikap rendah hati (tawadhu'). Menghilangkan kekerasan verbal, fisik, dan relasi kuasa yang menekan."),
        ("Ki Hajar Dewantara (Sistem Among)", "Mencakup prinsip Asah (saling mencerdaskan), Asih (saling mengasihi), dan Asuh (saling membimbing). Pendidikan menuntun kodrat alami anak untuk mencapai keselamatan dan kebahagiaan setinggi-tingginya sebagai manusia dan anggota masyarakat.")
    ]
    for f_name, f_desc in filosofi_data:
        ws6.cell(row=r, column=1, value=f_name).font = font_data_bold
        ws6.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws6.cell(row=r, column=1).alignment = align_left
        ws6.cell(row=r, column=1).fill = fill_group

        ws6.cell(row=r, column=3, value=f_desc).font = font_data
        ws6.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        ws6.cell(row=r, column=3).alignment = align_left

        for c in range(1, 6):
            ws6.cell(row=r, column=c).border = cell_border
            if c >= 3 and r % 2 == 1:
                ws6.cell(row=r, column=c).fill = fill_zebra
        ws6.row_dimensions[r].height = 34
        r += 1

    auto_fit_columns(ws6, min_width=10, max_width=50)
    ws6.column_dimensions['A'].width = 5
    ws6.column_dimensions['B'].width = 30
    ws6.column_dimensions['C'].width = 38
    ws6.column_dimensions['D'].width = 32
    ws6.column_dimensions['E'].width = 35
    ws6.freeze_panes = 'A5'

    # ==========================================
    # SHEET 7: KOKURIKULER & EKSTRAKURIKULER
    # ==========================================
    ws7 = wb.create_sheet(title="Koku & Ekstrakurikuler")
    apply_title_banner(
        ws7,
        "PENGEMBANGAN KOKURIKULER & EKSTRAKURIKULER SMA",
        "Pasal 16-19, Pasal 22, dan Lampiran III Permendikdasmen No. 13 Tahun 2025",
        max_col=6
    )

    ws7.cell(row=4, column=1, value="1. TATA KELOLA KOKURIKULER (PASAL 16-19)").font = font_section_header
    ws7.merge_cells("A4:F4")
    for c in range(1, 7):
        ws7.cell(row=4, column=c).fill = fill_tbl_header
    ws7.row_dimensions[4].height = 24

    koku_info = [
        ("Komponen Kokurikuler", "Memuat: (a) Kompetensi (8 Dimensi Profil Lulusan), (b) Muatan Pembelajaran (Tema), dan (c) Beban Belajar (Alokasi 1 tahun ajaran)."),
        ("Bentuk Pelaksanaan Baru", "Dilaksanakan dalam bentuk: (1) Pembelajaran kolaboratif lintas disiplin ilmu, (2) Gerakan 7 (tujuh) kebiasaan anak Indonesia hebat, dan/atau (3) Cara lainnya."),
        ("Pengembangan Tema", "Tema berupa payung konsep kontekstual yang digunakan untuk merumuskan topik relevan dengan konteks sosial budaya lokal dan karakteristik peserta didik."),
        ("Alokasi Waktu Jam Pelajaran", "Beban waktu kokurikuler dirumuskan dalam 1 tahun ajaran (Kelas X = 396 JP/thn, Kelas XI = 144 JP/thn, Kelas XII = 128 JP/thn). Bersifat sangat fleksibel."),
        ("Model Penjadwalan", "Dapat diatur: (a) Terjadwal reguler 1-2 hari per minggu, (b) Sistem blok berkala (misal 2 minggu berturut-turut fokus proyek), atau (c) Terintegrasi harian.")
    ]

    r = 5
    for title, desc in koku_info:
        ws7.cell(row=r, column=1, value=title).font = font_data_bold
        ws7.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws7.cell(row=r, column=1).fill = fill_group
        ws7.cell(row=r, column=1).alignment = align_left
        
        ws7.cell(row=r, column=3, value=desc).font = font_data
        ws7.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
        ws7.cell(row=r, column=3).alignment = align_left

        for c in range(1, 7):
            ws7.cell(row=r, column=c).border = cell_border
            if c >= 3 and r % 2 == 1:
                ws7.cell(row=r, column=c).fill = fill_zebra
        ws7.row_dimensions[r].height = 24
        r += 1

    r += 1
    # 7 Kebiasaan Anak Indonesia Hebat
    ws7.cell(row=r, column=1, value="GERAKAN 7 (TUJUH) KEBIASAAN ANAK INDONESIA HEBAT (KOKURIKULER)").font = font_section_header
    ws7.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws7.cell(row=r, column=c).fill = fill_card_header
    ws7.row_dimensions[r].height = 24
    r += 1

    headers_7k = ["No", "Kebiasaan Karakter Hebat", "Makna Pembiasaan Positif", "Relevansi dengan 8 Profil Lulusan", "Bentuk Aktivitas di SMA", "Indikator Keberhasilan"]
    for i, h in enumerate(headers_7k, 1):
        cell = ws7.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = cell_border
    ws7.row_dimensions[r].height = 24
    r += 1

    data_7k = [
        (1, "Bangun Pagi", "Membangun kedisiplinan diri, menghargai waktu, dan kesiapan mengawali hari dengan energi positif.", "Kemandirian, Kesehatan", "Hadir di sekolah tepat waktu, rutinitas pagi mandiri tanpa ketergantungan orang tua.", "Tingkat ketepatan waktu hadir di kelas tinggi."),
        (2, "Beribadah", "Menjalankan kewajiban spiritual, memperkuat hubungan dengan Tuhan, memupuk integritas batin.", "Keimanan & Ketakwaan", "Shalat dhuha/dzuhur berjamaah, doa bersama pagi hari, kebaktian/retret sekolah.", "Ketaatan beribadah dan sikap saling menghormati agama."),
        (3, "Berolahraga", "Menjaga kebugaran fisik, mengasah sportivitas, melepaskan ketegangan belajar lewat aktivitas jasmani.", "Kesehatan, Kolaborasi", "Senam pagi bersama, jalan sehat berkala, turnamen antarkelas (class meeting), jogging rutin.", "Fisik bugar, stamina belajar prima, rendah angka sakit."),
        (4, "Makan Sehat & Bergizi", "Membiasakan pola konsumsi bernutrisi seimbang, menjauhi makanan instan berlebihan dan junk food.", "Kesehatan", "Membawa bekal sehat dari rumah, kantin sekolah menyediakan pangan higienis bernutrisi.", "Pola makan terjaga, paham prinsip gizi seimbang."),
        (5, "Gemar Belajar", "Menumbuhkan rasa ingin tahu ilmiah, budaya membaca (literasi), dan eksplorasi ilmu secara berkesadaran.", "Penalaran Kritis, Kreativitas", "Literasi 15 menit sebelum KBM, diskusi bedah buku, riset mandiri proyek sains/sosial.", "Tingginya kunjungan perpustakaan dan kualitas karya tulis."),
        (6, "Bermasyarakat", "Mengembangkan kepedulian sosial, toleransi, gotong royong, dan kemampuan beradaptasi di lingkungan warga.", "Kewargaan, Kolaborasi, Komunikasi", "Bakti sosial, aksi bersih lingkungan sekitar, pengabdian masyarakat sederhana oleh OSIS.", "Kepekaan sosial tinggi terhadap permasalahan lingkungan."),
        (7, "Istirahat Cukup (Tidur Cepat)", "Menjaga ritme sirkadian tubuh, membatasi penggunaan gawai (gadget/screen time) larut malam.", "Kesehatan, Kemandirian", "Edukasi sleep hygiene oleh BK, pengaturan waktu belajar malam yang seimbang.", "Siswa fokus dan tidak mengantuk saat mengikuti pelajaran di kelas.")
    ]

    for item in data_7k:
        no, name, makna, rel, act, ind = item
        ws7.cell(row=r, column=1, value=no).alignment = align_center
        ws7.cell(row=r, column=2, value=name).alignment = align_left
        ws7.cell(row=r, column=3, value=makna).alignment = align_left
        ws7.cell(row=r, column=4, value=rel).alignment = align_center
        ws7.cell(row=r, column=5, value=act).alignment = align_left
        ws7.cell(row=r, column=6, value=ind).alignment = align_left

        ws7.cell(row=r, column=1).font = font_data_bold
        ws7.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 7):
            ws7.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws7.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws7.cell(row=r, column=c).fill = fill_zebra
        ws7.row_dimensions[r].height = 36
        r += 1

    r += 1
    # Kerangka Ekstrakurikuler
    ws7.cell(row=r, column=1, value="2. PENGEMBANGAN LAYANAN EKSTRAKURIKULER (PASAL 22 & LAMPIRAN III)").font = font_section_header
    ws7.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws7.cell(row=r, column=c).fill = fill_tbl_header
    ws7.row_dimensions[r].height = 24
    r += 1

    headers_ekskul = ["Kategori", "Ketentuan Regulasi Resmi", "Rincian & Contoh Kegiatan di SMA", "Format Pelaksanaan", "Prinsip Pelaksanaan", "Penilaian / Asesmen"]
    for i, h in enumerate(headers_ekskul, 1):
        cell = ws7.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws7.row_dimensions[r].height = 24
    r += 1

    ekskul_data = [
        ("Krida (Wajib Disediakan)", "Satuan pendidikan WAJIB menyediakan layanan Kepramukaan atau Kepanduan lainnya (Pasal 22 ayat 2). Siswa memilih secara sukarela.", "Kepramukaan, Latihan Kepemimpinan Siswa (LKS), Palang Merah Remaja (PMR), UKS, Paskibra.", "Klasikal, kelompok, atau gabungan gugus depan.", "Individual, sukarela, etos kerja, keterlibatan aktif.", "Kualitatif deskriptif pada buku laporan capaian hasil belajar."),
        ("Karya Ilmiah", "Mengembangkan nalar kritis, kemampuan akademis terapan, dan metodologi riset.", "Kelompok Ilmiah Remaja (KIR), riset sains, karya inovasi teknologi, robotika, karya tulis sosial.", "Kelompok dan individual terbimbing guru pembina.", "Menyenangkan, menumbuhkan nalar kritis dan kreativitas.", "Portofolio karya riset, keikutsertaan kompetisi ilmiah."),
        ("Latihan Olah-Bakat / Minat", "Wadah penyaluran bakat spesifik non-akademik siswa di bidang olahraga, seni, dan teknologi.", "Futsal, Basket, Voli, Musik/Band, Paduan Suara, Tari Tradisional/Modern, Fotografi, Jurnalistik.", "Individual, tim, atau kelompok gabungan.", "Pengembangan kapasitas diri dan penyiapan karir profesi.", "Sertifikat capaian, penilaian kualitatif keaktifan."),
        ("Keagamaan", "Memperdalam pemahaman nilai religi, praktik ibadah ritual, dan penguatan moralitas siswa.", "Rohani Islam (Rohis), Pendalaman Alkitab/Persekutuan Doa (Rohkris/Katolik), Pesantren Kilat, Retret.", "Kelompok keagamaan dan klasikal.", "Kemanfaatan sosial, membangun keimanan dan keikhlasan.", "Deskripsi karakter kepribadian dan akhlak siswa."),
        ("Bentuk Lainnya", "Kegiatan positif terorganisir sesuai kearifan lokal atau kebutuhan minat komunitas siswa.", "Pecinta alam, English Club, Debate Society, Komunitas Robotika & AI, Coding Club.", "Lapangan, gabungan antarkelas, proyek terbuka.", "Etos kerja kolaboratif dan kemanfaatan lingkungan.", "Asesmen partisipasi dan kontribusi komunitas.")
    ]

    for item in ekskul_data:
        kat, reg, rinci, fmt, prs, pen = item
        ws7.cell(row=r, column=1, value=kat).alignment = align_left
        ws7.cell(row=r, column=2, value=reg).alignment = align_left
        ws7.cell(row=r, column=3, value=rinci).alignment = align_left
        ws7.cell(row=r, column=4, value=fmt).alignment = align_center
        ws7.cell(row=r, column=5, value=prs).alignment = align_left
        ws7.cell(row=r, column=6, value=pen).alignment = align_left

        ws7.cell(row=r, column=1).font = font_data_bold
        for c in range(2, 7):
            ws7.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws7.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws7.cell(row=r, column=c).fill = fill_zebra
        ws7.row_dimensions[r].height = 42
        r += 1

    auto_fit_columns(ws7, min_width=10, max_width=45)
    ws7.column_dimensions['A'].width = 24
    ws7.column_dimensions['B'].width = 34
    ws7.column_dimensions['C'].width = 34
    ws7.column_dimensions['D'].width = 22
    ws7.column_dimensions['E'].width = 32
    ws7.column_dimensions['F'].width = 32
    ws7.freeze_panes = 'A5'

    # ==========================================
    # SHEET 8: FOKUS BAHASA INGGRIS & BAHASA
    # ==========================================
    ws8 = wb.create_sheet(title="Fokus Bahasa Inggris")
    apply_title_banner(
        ws8,
        "ANALISIS KHUSUS MATA PELAJARAN BAHASA INGGRIS DI SMA",
        "Kajian Struktur Beban Belajar Bahasa Inggris (Wajib & Tingkat Lanjut) Menurut Permendikdasmen No. 13 Tahun 2025",
        max_col=6
    )

    ws8.cell(row=4, column=1, value="1. STRUKTUR BEBAN BELAJAR BAHASA INGGRIS DI SMA").font = font_section_header
    ws8.merge_cells("A4:F4")
    for c in range(1, 7):
        ws8.cell(row=4, column=c).fill = fill_tbl_header
    ws8.row_dimensions[4].height = 24

    headers_bi = ["Tingkatan / Fase", "Jenis Mata Pelajaran", "Alokasi Intra (JP/Thn)", "Alokasi Koku (JP/Thn)", "Total JP/Thn (JP/Mgg)", "Peran & Karakteristik Kurikuler"]
    r = 5
    for i, h in enumerate(headers_bi, 1):
        cell = ws8.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws8.row_dimensions[5].height = 24
    r = 6

    bi_structure = [
        ("Kelas X (Fase E)", "Bahasa Inggris Wajib", 108, 0, "108 JP (3 JP/mgg)", "Mata pelajaran wajib bagi SEMUA peserta didik. Seluruh alokasi dialokasikan untuk intrakurikuler (tanpa jam terpisah untuk kokurikuler)."),
        ("Kelas XI (Fase F)", "Bahasa Inggris Wajib", 108, 0, "108 JP (3 JP/mgg)", "Mata pelajaran wajib intrakurikuler bagi SELURUH siswa kelas XI untuk menjaga kontinuitas kemahiran berbahasa Inggris."),
        ("Kelas XI (Fase F)", "Bahasa Inggris Tingkat Lanjut (Pilihan)", 180, 0, "180 JP (5 JP/mgg)", "Mata pelajaran pilihan bagi murid yang berminat mendalami Bahasa Inggris (Akademik, Sastra, Komunikasi Global). Total belajar siswa peminat: 8 JP/mgg!"),
        ("Kelas XII (Fase F)", "Bahasa Inggris Wajib", 96, 0, "96 JP (3 JP/mgg)", "Mata pelajaran wajib intrakurikuler (asumsi 32 minggu efektif karena masa persiapan kelulusan/ujian)."),
        ("Kelas XII (Fase F)", "Bahasa Inggris Tingkat Lanjut (Pilihan)", 160, 0, "160 JP (5 JP/mgg)", "Mata pelajaran pilihan peminatan bahasa di kelas XII (asumsi 32 minggu x 5 JP). Penyiapan studi lanjut universitas dan kompetensi global."),
        ("Kombinasi Fase F", "Maksimal Bahasa Inggris Siswa Peminat (XI)", 288, 0, "288 JP (8 JP/mgg)", "Siswa yang memilih Bahasa Inggris Tingkat Lanjut di Kelas XI mendapatkan total 8 JP Bahasa Inggris per minggu!"),
        ("Kombinasi Fase F", "Maksimal Bahasa Inggris Siswa Peminat (XII)", 256, 0, "256 JP (8 JP/mgg)", "Siswa yang memilih Bahasa Inggris Tingkat Lanjut di Kelas XII mendapatkan total 8 JP Bahasa Inggris per minggu!")
    ]

    for item in bi_structure:
        fase, jenis, intra, koku, tot, peran = item
        ws8.cell(row=r, column=1, value=fase).alignment = align_center
        ws8.cell(row=r, column=2, value=jenis).alignment = align_left
        ws8.cell(row=r, column=3, value=intra).alignment = align_center
        ws8.cell(row=r, column=4, value=koku).alignment = align_center
        ws8.cell(row=r, column=5, value=tot).alignment = align_center
        ws8.cell(row=r, column=6, value=peran).alignment = align_left

        ws8.cell(row=r, column=1).font = font_data_bold
        ws8.cell(row=r, column=2).font = font_data_bold if "Kombinasi" in fase or "Tingkat Lanjut" in jenis else font_data
        for c in range(3, 7):
            ws8.cell(row=r, column=c).font = font_data_bold if "Kombinasi" in fase else font_data
            
        fill_item = fill_subgroup if "Kombinasi" in fase else (fill_zebra if r % 2 == 1 else PatternFill(fill_type=None))
        for c in range(1, 7):
            ws8.cell(row=r, column=c).border = cell_border
            if fill_item.fill_type:
                ws8.cell(row=r, column=c).fill = fill_item
        ws8.row_dimensions[r].height = 34
        r += 1

    r += 1
    # Implementasi Deep Learning pada Pembelajaran Bahasa Inggris
    ws8.cell(row=r, column=1, value="2. STRATEGI PEMBELAJARAN MENDALAM (DEEP LEARNING) UNTUK BAHASA INGGRIS").font = font_section_header
    ws8.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws8.cell(row=r, column=c).fill = fill_tbl_header
    ws8.row_dimensions[r].height = 24
    r += 1

    headers_dl_bi = ["Prinsip Deep Learning", "Penerapan Spesifik dalam Bahasa Inggris", "Praktik Pedagogis Guru", "Aktivitas Otentik Siswa", "Koneksi Dimensi Profil", "Bentuk Asesmen Otentik"]
    for i, h in enumerate(headers_dl_bi, 1):
        cell = ws8.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws8.row_dimensions[r].height = 24
    r += 1

    dl_bi_data = [
        ("Mindful Learning (Berkesadaran)", "Siswa menyadari proses pemerolehan bahasanya (metalinguistic awareness), berani mencoba tanpa takut keliru gramatika, dan meregulasi kecepatan belajarnya.", "Memberikan umpan balik konstruktif yang suportif, mendorong 'growth mindset' berbahasa, mendampingi refleksi kesulitan membaca/menulis.", "Menulis jurnal belajar harian (language learning log), self-assessment kemahiran berbicara, tracking target vocabulary.", "Kemandirian, Komunikasi", "Rubrik evaluasi diri (self-rubric) dan refleksi metakognitif berkala."),
        ("Meaningful Learning (Bermakna)", "Bahasa Inggris tidak dipelajari sebagai rumus teoritis kaku, melainkan sebagai instrumen komunikasi nyata dalam mengeksplorasi isu lokal dan global.", "Menyajikan teks otentik (artikel berita, podcast, video dokumenter), merancang proyek berbasis masalah nyata (PBL), integrasi isu sosial.", "Melakukan wawancara turis/penutur asing, menyusun kampanye isu lingkungan sekolah dalam bahasa Inggris, menganalisis hoax berbahasa asing.", "Penalaran Kritis, Kewargaan, Komunikasi", "Presentasi infografis, video campaign berdurasi 3 menit, esai argumentatif terpublikasi."),
        ("Joyful Learning (Menggembirakan)", "Menciptakan suasana kelas yang interaktif, penuh tawa, saling memuliakan teman yang sedang berbicara, bebas intimidasi atau celaan logat (accent).", "Memanfaatkan gamifikasi edukatif, simulasi debat santun, role-play situasi nyata, drama monolog, storytelling interaktif.", "Berpartisipasi dalam english games (Kahoot, Wordwall), simulasi podcasting, drama pendek komedi bahasa Inggris.", "Kreativitas, Kolaborasi, Kesehatan Mental", "Observasi antusiasme partisipasi, peer-feedback positif, performa panggung karya.")
    ]

    for item in dl_bi_data:
        prs, pen, ped, act, prf, asm = item
        ws8.cell(row=r, column=1, value=prs).alignment = align_left
        ws8.cell(row=r, column=2, value=pen).alignment = align_left
        ws8.cell(row=r, column=3, value=ped).alignment = align_left
        ws8.cell(row=r, column=4, value=act).alignment = align_left
        ws8.cell(row=r, column=5, value=prf).alignment = align_center
        ws8.cell(row=r, column=6, value=asm).alignment = align_left

        ws8.cell(row=r, column=1).font = font_data_bold
        for c in range(2, 7):
            ws8.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws8.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws8.cell(row=r, column=c).fill = fill_zebra
        ws8.row_dimensions[r].height = 54
        r += 1

    r += 1
    # Bahasa Asing Lainnya di SMA
    ws8.cell(row=r, column=1, value="3. PILIHAN BAHASA ASING LAINNYA DI FASE F (KELAS XI & XII)").font = font_section_header
    ws8.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws8.cell(row=r, column=c).fill = fill_tbl_header
    ws8.row_dimensions[r].height = 24
    r += 1

    headers_lang = ["No", "Mata Pelajaran Bahasa Asing Pilihan", "Alokasi Kelas XI (JP/Thn)", "Alokasi Kelas XII (JP/Thn)", "Alokasi Mingguan", "Prospek & Relevansi Masa Depan"]
    for i, h in enumerate(headers_lang, 1):
        cell = ws8.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws8.row_dimensions[r].height = 24
    r += 1

    languages_data = [
        (1, "Bahasa Arab", 180, 160, "5 JP/minggu", "Studi keislaman, hubungan diplomatik Timur Tengah, peluang beasiswa Timur Tengah."),
        (2, "Bahasa Jepang", 180, 160, "5 JP/minggu", "Industri teknologi, beasiswa Monbukagakusho/MEXT, pertukaran budaya pop."),
        (3, "Bahasa Jerman", 180, 160, "5 JP/minggu", "Studi teknik, sains, beasiswa DAAD, studi tinggi gratis di Jerman (Studienkolleg)."),
        (4, "Bahasa Korea", 180, 160, "5 JP/minggu", "Industri kreatif, beasiswa GKS, investasi teknologi Korea di Indonesia."),
        (5, "Bahasa Mandarin", 180, 160, "5 JP/minggu", "Perdagangan global, ekonomi internasional, sains dan teknologi Tiongkok/Taiwan."),
        (6, "Bahasa Prancis", 180, 160, "5 JP/minggu", "Diplomasi internasional (PBB), seni budaya Eropa, studi hukum/arsitektur Prancis.")
    ]

    for item in languages_data:
        no, lang, xi, xii, mgg, prospek = item
        ws8.cell(row=r, column=1, value=no).alignment = align_center
        ws8.cell(row=r, column=2, value=lang).alignment = align_left
        ws8.cell(row=r, column=3, value=xi).alignment = align_center
        ws8.cell(row=r, column=4, value=xii).alignment = align_center
        ws8.cell(row=r, column=5, value=mgg).alignment = align_center
        ws8.cell(row=r, column=6, value=prospek).alignment = align_left

        ws8.cell(row=r, column=1).font = font_data_bold
        ws8.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 7):
            ws8.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws8.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws8.cell(row=r, column=c).fill = fill_zebra
        ws8.row_dimensions[r].height = 24
        r += 1

    auto_fit_columns(ws8, min_width=10, max_width=45)
    ws8.column_dimensions['A'].width = 18
    ws8.column_dimensions['B'].width = 34
    ws8.column_dimensions['C'].width = 24
    ws8.column_dimensions['D'].width = 30
    ws8.column_dimensions['E'].width = 24
    ws8.column_dimensions['F'].width = 36
    ws8.freeze_panes = 'A5'

    # ==========================================
    # SHEET 9: KOMPARASI PERMEN 12 VS PERMEN 13
    # ==========================================
    ws9 = wb.create_sheet(title="Komparasi Permen 12 vs 13")
    apply_title_banner(
        ws9,
        "MATRIKS KOMPARASI: PERMEN 12 TAHUN 2024 VS PERMEN 13 TAHUN 2025",
        "Peta Perubahan Kebijakan Regulasi Kurikulum Sekolah Menengah Atas (SMA)",
        max_col=5
    )

    headers_comp = ["No", "Aspek Komparasi", "Sebelumnya (Permendikbudristek No. 12 Tahun 2024)", "Penyempurnaan (Permendikdasmen No. 13 Tahun 2025)", "Implikasi & Penyesuaian Satuan Pendidikan"]
    r = 4
    for i, h in enumerate(headers_comp, 1):
        cell = ws9.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = cell_border
    ws9.row_dimensions[4].height = 26
    r = 5

    comparisons = [
        (1, "Nomenklatur & Pejabat Kementerian", "Kementerian Pendidikan, Kebudayaan, Riset, dan Teknologi (Mendikbudristek Nadiem Makarim).", "Kementerian Pendidikan Dasar dan Menengah (Mendikdasmen Abdul Mu'ti).", "Penyesuaian kop surat, dokumen kurikulum (KSP), dan regulasi acuan resmi sekolah."),
        (2, "Filosofi & Pendekatan Utama", "Mengedepankan pembelajaran berpusat pada murid dengan fokus esensial materi dan fleksibilitas kurikuler.", "Mengadopsi resmi 'Pembelajaran Mendalam (Deep Learning)' yang Berkesadaran, Bermakna, Menggembirakan, serta nilai saling memuliakan (KH Hasyim Asy'ari & Ki Hajar Dewantara).", "Guru mengubah pendekatan mengajar dari mengejar ketuntasan konten menjadi kedalaman pemahaman, retensi bermakna, dan iklim belajar nyaman."),
        (3, "Profil Pelajar / Dimensi Lulusan", "6 Dimensi Profil Pelajar Pancasila: (1) Beriman & bertakwa, (2) Berkebinekaan global, (3) Bergotong royong, (4) Mandiri, (5) Bernalar kritis, (6) Kreatif.", "8 Dimensi Profil Lulusan: (1) Keimanan & Ketakwaan, (2) Kewargaan, (3) Penalaran Kritis, (4) Kreativitas, (5) Kolaborasi, (6) Kemandirian, (7) Kesehatan, (8) Komunikasi.", "Pembaruan rubrik asesmen karakter di rapor; integrasi aspek kesehatan fisik/mental dan komunikasi efektif ke dalam penilaian."),
        (4, "Pelaksanaan Kokurikuler", "Diidentikkan secara khusus dengan Projek Penguatan Profil Pelajar Pancasila (P5) dengan tema-tema panduan Kemendikbudristek.", "Dilaksanakan fleksibel dalam bentuk: (1) Pembelajaran kolaboratif lintas disiplin ilmu, (2) Gerakan 7 Kebiasaan Anak Indonesia Hebat, dan/atau (3) Cara lainnya.", "Sekolah memiliki fleksibilitas penuh; tidak lagi terbebani prosedur administratif P5 yang kaku; proyek dapat dikaitkan langsung dengan pembiasaan karakter 7 kebiasaan."),
        (5, "Mata Pelajaran Koding & AI", "Belum dimasukkan sebagai mata pelajaran tersendiri dalam struktur kurikulum standar SMA (hanya bagian materi Informatika).", "Resmi masuk struktur kurikulum sebagai Mata Pelajaran Pilihan 'Koding dan Kecerdasan Artifisial' mulai TA 2025/2026 secara bertahap (Pasal 32A).", "Sekolah mempersiapkan guru TIK/Informatika, lab komputer, dan perangkat kurikulum koding dan AI bagi siswa yang berminat."),
        (6, "Beban Waktu Koding & AI", "Tidak ada alokasi JP terpisah untuk Koding & AI.", "Kelas X: 2 JP/minggu (72 JP/tahun). Kelas XI: 5 JP/minggu (180 JP/tahun). Kelas XII: 5 JP/minggu (160 JP/tahun).", "Dapat dipilih oleh siswa kelas X (opsional) dan siswa kelas XI-XII sebagai bagian dari 4-5 mapel pilihan."),
        (7, "Ekstrakurikuler Kepramukaan", "Ketentuan Kepramukaan sebagai ekstrakurikuler wajib bagi sekolah diatur dinamis, menimbulkan perdebatan publik mengenai status wajibnya.", "Ditegaskan tegas di Pasal 22 ayat (2): Satuan pendidikan SEKUKANG-KURANGNYA MENYEDIAKAN Ekstrakurikuler Kepramukaan atau Kepanduan lainnya.", "Sekolah wajib mengalokasikan anggaran, sarana, dan pembina pramuka/kepanduan; siswa mengikuti secara sukarela sesuai minat."),
        (8, "Batas Penggantian Mapel Pilihan", "Dibatasi pada ketentuan panduan umum pemilihan mata pelajaran.", "Ditegaskan tegas di Lampiran II Bagian D nomor 3: Murid diperbolehkan mengganti mata pelajaran pilihan paling lambat KELAS XI SEMESTER 2 berdasarkan evaluasi ulang sekolah.", "Sekolah menyelenggarakan asesmen diagnostik minat bakat ulang di akhir semester 1 atau pertengahan semester 2 kelas XI bagi siswa yang ingin berpindah."),
        (9, "Masa Transisi & Pemberlakuan", "Mulai diterapkan serentak atau bertahap sesuai Kurikulum Merdeka mandiri.", "Diberlakukan secara bertahap mulai dari Kelas X pada tahun ajaran baru (Pasal 32 huruf c).", "Sekolah dapat menerapkan kurikulum penyempurnaan ini pada angkatan kelas X baru sembari menuntaskan angkatan kelas XI & XII.")
    ]

    for item in comparisons:
        no, aspek, lama, baru, dampak = item
        ws9.cell(row=r, column=1, value=no).alignment = align_center
        ws9.cell(row=r, column=2, value=aspek).alignment = align_left
        ws9.cell(row=r, column=3, value=lama).alignment = align_left
        ws9.cell(row=r, column=4, value=baru).alignment = align_left
        ws9.cell(row=r, column=5, value=dampak).alignment = align_left

        ws9.cell(row=r, column=1).font = font_data_bold
        ws9.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 6):
            ws9.cell(row=r, column=c).font = font_data
            
        for c in range(1, 6):
            ws9.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws9.cell(row=r, column=c).fill = fill_zebra
        ws9.row_dimensions[r].height = 48
        r += 1

    auto_fit_columns(ws9, min_width=10, max_width=45)
    ws9.column_dimensions['A'].width = 5
    ws9.column_dimensions['B'].width = 24
    ws9.column_dimensions['C'].width = 38
    ws9.column_dimensions['D'].width = 40
    ws9.column_dimensions['E'].width = 38
    ws9.freeze_panes = 'A5'

    # Save workbook
    wb.save(output_path)
    print(f"Excel workbook successfully saved to: {output_path}")

if __name__ == "__main__":
    out_file = r"C:\Users\KOMPUTERKU2\Documents\English SMA\Peraturan\Ringkasan_Kurikulum_SMA_Permendikdasmen_13_2025.xlsx"
    create_kurikulum_excel(out_file)
