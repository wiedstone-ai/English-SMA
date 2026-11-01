# -*- coding: utf-8 -*-
"""
Script to generate a comprehensive, highly styled Excel workbook summarizing
Capaian Pembelajaran (CP) and Kurikulum for BAHASA INGGRIS SMA
based on:
1. Kepka BKPDM No. 020 Tahun 2026 (Perubahan Kepka BSKAP No. 046/H/KR/2025)
2. Keputusan Kepala BSKAP No. 046/H/KR/2025 (Regulasi Induk CP Seluruh Mapel Non-Agama)
3. Panduan Mata Pelajaran Bahasa Inggris Revisi 2025 (Puskurjar Kemendikdasmen)
4. PERMENDIKDASMEN No. 13 Tahun 2025 (Struktur Kurikulum & Beban Belajar)
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_english_sma_excel(output_path):
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
    fill_callout_blue = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")   # Light Sky Blue

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
    # SHEET 1: STATUS REGULASI CP 2026
    # ==========================================
    ws1 = wb.create_sheet(title="Status Regulasi CP 2026")
    apply_title_banner(
        ws1,
        "STATUS REGULASI RESMI CAPAIAN PEMBELAJARAN (CP) BAHASA INGGRIS SMA",
        "Analisis Kedudukan Hukum Kepka BKPDM No. 020 Tahun 2026, Kepka BSKAP No. 046/H/KR/2025, dan Panduan Mapel 2025",
        max_col=6
    )

    ws1.cell(row=4, column=1, value="PENEGASAN HUKUM: STATUS CAPAIAN PEMBELAJARAN BAHASA INGGRIS SMA").font = font_section_header
    ws1.merge_cells("A4:F4")
    for c in range(1, 7):
        ws1.cell(row=4, column=c).fill = fill_tbl_header
    ws1.row_dimensions[4].height = 24

    ws1.cell(row=5, column=1, value="CATATAN KRUSIAL BAGI PENDIDIK & KURIKULUM BAHASA INGGRIS:").font = font_callout_title
    ws1.cell(row=6, column=1, value=(
        "1. Keputusan Kepala Badan Kebijakan Pendidikan Dasar dan Menengah (BKPDM) Kemendikdasmen No. 020 Tahun 2026 "
        "merupakan regulasi perubahan terbatas yang HANYA merevisi Capaian Pembelajaran untuk mata pelajaran Pendidikan Agama dan Budi Pekerti.\n"
        "2. Capaian Pembelajaran (CP) untuk MATA PELAJARAN BAHASA INGGRIS SMA (baik Bahasa Inggris Umum/Wajib maupun Bahasa Inggris Tingkat Lanjut) "
        "TIDAK MENGALAMI PERUBAHAN dalam Kepka BKPDM No. 020 Tahun 2026.\n"
        "3. CP Bahasa Inggris SMA secara yuridis dan substansial TETAP RESMI MENGACU pada Keputusan Kepala BSKAP No. 046/H/KR/2025 serta dokumen teknis operasional "
        "Panduan Mata Pelajaran Bahasa Inggris Revisi 2025 (Pusat Kurikulum dan Pembelajaran Kemendikdasmen).\n"
        "4. Struktur alokasi beban belajar Bahasa Inggris diatur dalam PERMENDIKDASMEN No. 13 Tahun 2025 (3 JP/minggu untuk Mapel Wajib dan 5 JP/minggu untuk Mapel Pilihan Tingkat Lanjut)."
    )).font = font_callout_body
    ws1.merge_cells("A5:F5")
    ws1.merge_cells("A6:F6")
    ws1.cell(row=5, column=1).fill = fill_callout
    ws1.cell(row=6, column=1).fill = fill_callout
    ws1.cell(row=6, column=1).alignment = align_left
    for r_i in [5, 6]:
        for c in range(1, 7):
            ws1.cell(row=r_i, column=c).border = cell_border
            ws1.cell(row=r_i, column=c).fill = fill_callout
    ws1.row_dimensions[5].height = 22
    ws1.row_dimensions[6].height = 70

    r = 8
    ws1.cell(row=r, column=1, value="MATRIKS KETERKAITAN 4 REGULASI UTAMA KURIKULUM BAHASA INGGRIS SMA").font = font_section_header
    ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws1.cell(row=r, column=c).fill = fill_card_header
    ws1.row_dimensions[r].height = 24
    r += 1

    headers_reg = ["No", "Nama Regulasi / Dokumen Resmi", "Penerbit & Tahun", "Ruang Lingkup Regulasi", "Status bagi Bahasa Inggris SMA", "Fungsi Acuan di Sekolah"]
    for i, h in enumerate(headers_reg, 1):
        cell = ws1.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = cell_border
    ws1.row_dimensions[r].height = 26
    r += 1

    reg_matrix = [
        (1, "Kepka BKPDM Kemendikdasmen No. 020 Tahun 2026", "Kepala BKPDM (2026)", "Perubahan atas Kepka BSKAP No. 046/H/KR/2025 khusus mata pelajaran Pendidikan Agama dan Budi Pekerti.", "Tidak mengubah CP Bahasa Inggris. Mapel Bahasa Inggris tetap mengacu pada Kepka 046/2025.", "Menjadi landasan kepastian bahwa CP Bahasa Inggris tetap stabil dan berlaku penuh."),
        (2, "Keputusan Kepala BSKAP Kemendikdasmen No. 046/H/KR/2025", "Kepala BSKAP (2025)", "Standar Capaian Pembelajaran Nasional untuk PAUD, Jenjang Pendidikan Dasar, dan Jenjang Pendidikan Menengah.", "REGULASI INDUK RESMI CP Bahasa Inggris Fase E (Kelas X), Fase F Wajib (XI-XII), dan Fase F Tingkat Lanjut.", "Rujukan legal formal penulisan Capaian Pembelajaran pada KSP, Modul Ajar, dan Rapor."),
        (3, "PERMENDIKDASMEN RI No. 13 Tahun 2025", "Menteri Dikdasmen Abdul Mu'ti (11 Juli 2025)", "Perubahan Permendikbudristek No. 12/2024 tentang Kerangka Dasar, Struktur Kurikulum, Kokurikuler, Ekstrakurikuler.", "Menetapkan struktur alokasi jam: Bahasa Inggris Wajib = 3 JP/mgg (108 JP di X-XI, 96 JP di XII); B. Inggris Tingkat Lanjut = 5 JP/mgg (180 JP di XI, 160 JP di XII).", "Rujukan resmi penyusunan Struktur Kurikulum Sekolah, Pembagian Beban Mengajar Guru, dan Jadwal Pelajaran."),
        (4, "Panduan Mata Pelajaran Bahasa Inggris Revisi 2025 (Revisi 3)", "Pusat Kurikulum & Pembelajaran (Puskurjar) 2025", "Pedoman operasional implementasi kurikulum Bahasa Inggris Fase B hingga Fase F Tingkat Lanjut berbasis Deep Learning & GBA.", "PANDUAN OPERASIONAL UTAMA: memuat penjabaran CP, elemen, materi esensial, Teaching Learning Cycle (TLC), ATP, modul ajar, dan instrumen asesmen otentik.", "Rujukan operasional harian bagi guru Bahasa Inggris SMA dalam merancang RPP/Modul Ajar dan asesmen di kelas.")
    ]

    for item in reg_matrix:
        no, name, pub, scope, stat, func = item
        ws1.cell(row=r, column=1, value=no).alignment = align_center
        ws1.cell(row=r, column=2, value=name).alignment = align_left
        ws1.cell(row=r, column=3, value=pub).alignment = align_center
        ws1.cell(row=r, column=4, value=scope).alignment = align_left
        ws1.cell(row=r, column=5, value=stat).alignment = align_left
        ws1.cell(row=r, column=6, value=func).alignment = align_left

        ws1.cell(row=r, column=1).font = font_data_bold
        ws1.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 7):
            ws1.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws1.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws1.cell(row=r, column=c).fill = fill_zebra
        ws1.row_dimensions[r].height = 45
        r += 1

    auto_fit_columns(ws1, min_width=10, max_width=45)
    ws1.column_dimensions['A'].width = 5
    ws1.column_dimensions['B'].width = 30
    ws1.column_dimensions['C'].width = 20
    ws1.column_dimensions['D'].width = 32
    ws1.column_dimensions['E'].width = 34
    ws1.column_dimensions['F'].width = 34
    ws1.freeze_panes = 'A5'

    # ==========================================
    # SHEET 2: STRUKTUR JP BAHASA INGGRIS
    # ==========================================
    ws2 = wb.create_sheet(title="Struktur JP Bahasa Inggris")
    apply_title_banner(
        ws2,
        "STRUKTUR BEBAN BELAJAR BAHASA INGGRIS SMA",
        "Berdasarkan PERMENDIKDASMEN No. 13 Tahun 2025 (Lampiran II Bagian D) | 1 JP = 45 Menit",
        max_col=8
    )

    headers_jp = ["Tingkatan / Fase", "Jenis Mata Pelajaran", "Alokasi Intra (JP/Thn)", "Alokasi Koku (JP/Thn)", "Total JP Per Tahun", "Beban JP Per Minggu", "Status Keikutsertaan", "Ketentuan Administratif & Dapodik"]
    r = 4
    for i, h in enumerate(headers_jp, 1):
        cell = ws2.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = cell_border
    ws2.row_dimensions[4].height = 26
    r = 5

    data_jp = [
        ("Kelas X (Fase E)", "Bahasa Inggris Wajib", 108, 0, 108, "3 JP/minggu", "Wajib diikuti oleh SELURUH murid Kelas X", "Alokasi murni intrakurikuler 3 JP/mgg (asumsi 36 minggu). Terdaftar di kelompok Mapel Wajib."),
        ("Kelas XI (Fase F)", "Bahasa Inggris Wajib", 108, 0, 108, "3 JP/minggu", "Wajib diikuti oleh SELURUH murid Kelas XI", "Alokasi murni intrakurikuler 3 JP/mgg (asumsi 36 minggu). Terdaftar di kelompok Mapel Wajib."),
        ("Kelas XI (Fase F)", "Bahasa Inggris Tingkat Lanjut (Pilihan)", 180, 0, 180, "5 JP/minggu", "Pilihan bagi murid peminat bahasa/sastra", "Alokasi 5 JP/mgg (180 JP/thn). Sekolah wajib membuka min. 7 mapel pilihan, murid memilih 4-5 mapel."),
        ("Kelas XII (Fase F)", "Bahasa Inggris Wajib", 96, 0, 96, "3 JP/minggu", "Wajib diikuti oleh SELURUH murid Kelas XII", "Alokasi murni intrakurikuler 3 JP/mgg (asumsi 32 minggu efektif karena masa kelulusan)."),
        ("Kelas XII (Fase F)", "Bahasa Inggris Tingkat Lanjut (Pilihan)", 160, 0, 160, "5 JP/minggu", "Pilihan bagi murid peminat bahasa/sastra", "Alokasi 5 JP/mgg (160 JP/thn pada asumsi 32 minggu). Penyiapan kemampuan akademik global."),
        ("KOMBINASI XI", "Total Beban Siswa Peminat Bahasa Inggris (Kelas XI)", 288, 0, 288, "8 JP/minggu", "Siswa mengambil B. Inggris Wajib + Tingkat Lanjut", "Maksimal pemaparan bahasa Inggris di sekolah (8 JP/minggu atau 6 jam pelajaran penuh per minggu)."),
        ("KOMBINASI XII", "Total Beban Siswa Peminat Bahasa Inggris (Kelas XII)", 256, 0, 256, "8 JP/minggu", "Siswa mengambil B. Inggris Wajib + Tingkat Lanjut", "Maksimal pemaparan bahasa Inggris di sekolah (8 JP/minggu). Mendukung kesiapan tes bahasa & universitas.")
    ]

    for item in data_jp:
        fase, jenis, intra, koku, tot, mgg, stat, ket = item
        ws2.cell(row=r, column=1, value=fase).alignment = align_center
        ws2.cell(row=r, column=2, value=jenis).alignment = align_left
        ws2.cell(row=r, column=3, value=intra).alignment = align_center
        ws2.cell(row=r, column=4, value=koku).alignment = align_center
        ws2.cell(row=r, column=5, value=tot).alignment = align_center
        ws2.cell(row=r, column=6, value=mgg).alignment = align_center
        ws2.cell(row=r, column=7, value=stat).alignment = align_left
        ws2.cell(row=r, column=8, value=ket).alignment = align_left

        is_kombinasi = "KOMBINASI" in fase
        ws2.cell(row=r, column=1).font = font_data_bold
        ws2.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 9):
            ws2.cell(row=r, column=c).font = font_total if is_kombinasi else font_data
            
        fill_curr = fill_total if is_kombinasi else (fill_zebra if r % 2 == 1 else PatternFill(fill_type=None))
        for c in range(1, 9):
            ws2.cell(row=r, column=c).border = total_border if is_kombinasi else cell_border
            if fill_curr.fill_type:
                ws2.cell(row=r, column=c).fill = fill_curr
        ws2.row_dimensions[r].height = 28 if is_kombinasi else 24
        r += 1

    r += 1
    # Penjelasan Khusus Beban Kokurikuler pada Bahasa Inggris
    ws2.cell(row=r, column=1, value="CATATAN KHUSUS IMPLEMENTASI KOKURIKULER & BAHASA INGGRIS").font = font_section_header
    ws2.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    for c in range(1, 9):
        ws2.cell(row=r, column=c).fill = fill_card_header
    ws2.row_dimensions[r].height = 24
    r += 1

    notes_jp = [
        ("Murni Intrakurikuler", "Di dalam struktur tabel resmi Permendikdasmen 13/2025, alokasi jam Bahasa Inggris Wajib tercantum 108 JP (X), 108 JP (XI), dan 96 JP (XII) seluruhnya berada pada kolom Intrakurikuler (Kokurikuler = -). Artinya guru Bahasa Inggris memiliki 3 JP tatap muka reguler per minggu murni di kelas."),
        ("Peluang Integrasi Kokurikuler Lintas Disiplin", "Meskipun tidak memiliki jam kokurikuler terpisah, guru Bahasa Inggris dapat berkolaborasi dalam proyek kokurikuler lintas disiplin ilmu (Pasal 16 ayat 2), misalnya membimbing siswa menyusun poster berbahasa Inggris, presentasi video podcast, atau mini riset bertema lingkungan/kesehatan."),
        ("Mata Pelajaran Pilihan Tingkat Lanjut", "Bahasa Inggris Tingkat Lanjut di Kelas XI (180 JP) dan Kelas XII (160 JP) berbobot 5 JP per minggu. Waktu 5 JP/minggu sangat leluasa bagi guru untuk menerapkan siklus pengajaran berbasis teks (Teaching Learning Cycle) secara mendalam.")
    ]
    for n_title, n_desc in notes_jp:
        ws2.cell(row=r, column=1, value=n_title).font = font_data_bold
        ws2.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        ws2.cell(row=r, column=1).alignment = align_left
        ws2.cell(row=r, column=1).fill = fill_group

        ws2.cell(row=r, column=4, value=n_desc).font = font_data
        ws2.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)
        ws2.cell(row=r, column=4).alignment = align_left

        for c in range(1, 9):
            ws2.cell(row=r, column=c).border = cell_border
            if c >= 4 and r % 2 == 1:
                ws2.cell(row=r, column=c).fill = fill_zebra
        ws2.row_dimensions[r].height = 28
        r += 1

    auto_fit_columns(ws2, min_width=10, max_width=45)
    ws2.column_dimensions['A'].width = 18
    ws2.column_dimensions['B'].width = 30
    ws2.column_dimensions['C'].width = 16
    ws2.column_dimensions['D'].width = 16
    ws2.column_dimensions['E'].width = 16
    ws2.column_dimensions['F'].width = 18
    ws2.column_dimensions['G'].width = 32
    ws2.column_dimensions['H'].width = 40
    ws2.freeze_panes = 'A5'

    # ==========================================
    # SHEET 3: CP FASE E (KELAS X SMA)
    # ==========================================
    ws3 = wb.create_sheet(title="CP Fase E (Kelas X SMA)")
    apply_title_banner(
        ws3,
        "CAPAIAN PEMBELAJARAN BAHASA INGGRIS FASE E (KELAS X SMA)",
        "Keputusan Kepala BSKAP No. 046/H/KR/2025 & Panduan Mata Pelajaran Bahasa Inggris Revisi 2025",
        max_col=6
    )

    ws3.cell(row=4, column=1, value="1. CAPAIAN PEMBELAJARAN UMUM FASE E (KELAS X)").font = font_section_header
    ws3.merge_cells("A4:F4")
    for c in range(1, 7):
        ws3.cell(row=4, column=c).fill = fill_tbl_header
    ws3.row_dimensions[4].height = 24

    cp_umum_e_id = (
        "Pada akhir Fase E, murid menggunakan bahasa Inggris untuk berkomunikasi dengan guru, teman sebaya, dan orang lain dalam berbagai macam situasi dan tujuan. "
        "Murid menggunakan dan merespons pertanyaan terbuka serta menggunakan strategi untuk memulai, mempertahankan, dan menyimpulkan percakapan dan diskusi. "
        "Murid memahami dan mengidentifikasi ide utama dan detail relevan dari teks lisan dan tertulis (fiksi dan nonfiksi) terkait isu terkini. "
        "Murid mengomunikasikan ide dan pengalaman secara tertulis dan multimodal dengan struktur teks dan unsur kebahasaan yang tepat, serta mengungkapkan pendapat dan mempertahankan argumen tentang topik sehari-hari atau isu terkini."
    )
    cp_umum_e_en = (
        "By the end of Phase E, students use English to communicate with teachers, peers, and others in a variety of situations and purposes. "
        "They use and respond to open-ended questions and use strategies to initiate, sustain, and close conversations and discussions. "
        "They understand and identify the main ideas and relevant details of spoken and written texts (fiction and nonfiction) related to current issues. "
        "They communicate ideas and experiences in written and multimodal formats with appropriate text structures and language features, and express opinions and defend arguments about everyday topics or current issues."
    )

    ws3.cell(row=5, column=1, value="Capaian Umum (Bahasa Indonesia):").font = font_data_bold
    ws3.cell(row=6, column=1, value=cp_umum_e_id).font = font_data
    ws3.merge_cells("A5:F5")
    ws3.merge_cells("A6:F6")
    ws3.cell(row=6, column=1).alignment = align_left

    ws3.cell(row=7, column=1, value="General Outcome (English):").font = font_data_bold
    ws3.cell(row=8, column=1, value=cp_umum_e_en).font = font_english
    ws3.merge_cells("A7:F7")
    ws3.merge_cells("A8:F8")
    ws3.cell(row=8, column=1).alignment = align_left

    for r_idx in [5, 6, 7, 8]:
        for c in range(1, 7):
            ws3.cell(row=r_idx, column=c).border = cell_border
            ws3.cell(row=r_idx, column=c).fill = fill_subgroup if r_idx in [5, 7] else fill_zebra
    ws3.row_dimensions[5].height = 20
    ws3.row_dimensions[6].height = 50
    ws3.row_dimensions[7].height = 20
    ws3.row_dimensions[8].height = 50

    r = 10
    ws3.cell(row=r, column=1, value="2. CAPAIAN PEMBELAJARAN PER ELEMEN FASE E (KELAS X)").font = font_section_header
    ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws3.cell(row=r, column=c).fill = fill_card_header
    ws3.row_dimensions[r].height = 24
    r += 1

    headers_elemen = ["No", "Elemen Capaian Pembelajaran", "Rumusan Capaian Pembelajaran (Bahasa Indonesia)", "English Formulation", "Fokus Keterampilan Bahasa", "Target Kinerja Murid"]
    for i, h in enumerate(headers_elemen, 1):
        cell = ws3.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = cell_border
    ws3.row_dimensions[r].height = 26
    r += 1

    elemen_e_data = [
        (1, "Menyimak - Berbicara\n(Listening - Speaking)",
         "Murid menggunakan bahasa Inggris untuk berkomunikasi dengan guru, teman sebaya dan orang lain dalam berbagai macam situasi dan tujuan; menggunakan dan merespons pertanyaan terbuka; menggunakan strategi untuk memulai, mempertahankan dan menyimpulkan percakapan dan diskusi; memahami dan mengidentifikasi ide utama dan detail relevan dalam teks lisan fiksi dan nonfiksi tentang topik yang dibahas.",
         "Students use English to communicate with teachers, peers and others in various situations and purposes; use and respond to open-ended questions; use strategies to start, sustain and conclude conversations and discussions; understand and identify main ideas and relevant details in spoken fiction and nonfiction texts about the topics discussed.",
         "Interpersonal & Transaksional, Pertanyaan Terbuka, Strategi Percakapan, Menyimak Teks Fiksi & Nonfiksi.",
         "Mampu berdialog aktif, mempertahankan percakapan 3-5 giliran, menyimpulkan pesan teks lisan, merespons pertanyaan terbuka secara spontan."),
        (2, "Membaca - Memirsa\n(Reading - Viewing)",
         "Murid membaca dan merespons berbagai macam teks fiksi dan nonfiksi tertulis atau teks multimodal tentang isu terkini; mengidentifikasi tujuan, ide utama dan informasi rinci dari teks tertulis atau teks multimodal; mengevaluasi detail spesifik dan menarik kesimpulan.",
         "Students read and respond to a variety of written fiction and nonfiction texts or multimodal texts about current issues; identify the purpose, main ideas, and detailed information of written or multimodal texts; evaluate specific details and draw conclusions.",
         "Literasi Teks Tertulis & Multimodal (Infografis/Video), Menemukan Ide Pokok, Menilai Kredibilitas Detail, Menarik Kesimpulan (Inference).",
         "Mampu menganalisis teks fungsional/deskriptif/prosedur/recount, menemukan makna tersirat, mengevaluasi pesan infografis multimodal."),
        (3, "Menulis - Mempresentasikan\n(Writing - Presenting)",
         "Murid mengomunikasikan ide dan pengalaman melalui paragraf sederhana dan terstruktur dalam berbagai macam teks fiksi dan nonfiksi (tulisan atau teks multimodal); merencanakan, menulis, dan merevisi teks dengan struktur teks dan unsur kebahasaan yang tepat untuk mencapai tujuan tertentu; mengungkapkan pendapat dan mempertahankan argumen tentang topik sehari-hari atau isu terkini.",
         "Students communicate ideas and experiences through simple and structured paragraphs in various fiction and nonfiction texts (written or multimodal texts); plan, write, and revise texts with appropriate text structures and language features to achieve specific goals; express opinions and defend arguments about everyday topics or current issues.",
         "Proses Menulis (Drafting-Revising-Publishing), Struktur Paragraf Padu, Pembuatan Teks Multimodal, Mempertahankan Opini Tertulis.",
         "Mampu menulis esai mini deskriptif/recount/prosedur yang koheren, merancang poster/slide presentasi berbahasa Inggris, menyajikan argumen terstruktur.")
    ]

    for item in elemen_e_data:
        no, el, id_desc, en_desc, fok, tar = item
        ws3.cell(row=r, column=1, value=no).alignment = align_center
        ws3.cell(row=r, column=2, value=el).alignment = align_left
        ws3.cell(row=r, column=3, value=id_desc).alignment = align_top_left
        ws3.cell(row=r, column=4, value=en_desc).alignment = align_top_left
        ws3.cell(row=r, column=5, value=fok).alignment = align_top_left
        ws3.cell(row=r, column=6, value=tar).alignment = align_top_left

        ws3.cell(row=r, column=1).font = font_data_bold
        ws3.cell(row=r, column=2).font = font_data_bold
        ws3.cell(row=r, column=3).font = font_data
        ws3.cell(row=r, column=4).font = font_english
        ws3.cell(row=r, column=5).font = font_data
        ws3.cell(row=r, column=6).font = font_data
        
        for c in range(1, 7):
            ws3.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws3.cell(row=r, column=c).fill = fill_zebra
        ws3.row_dimensions[r].height = 68
        r += 1

    r += 1
    # Materi Esensial Fase E
    ws3.cell(row=r, column=1, value="3. PEMETAAN MATERI ESENSIAL & FUNGSI BAHASA FASE E (KELAS X)").font = font_section_header
    ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws3.cell(row=r, column=c).fill = fill_tbl_header
    ws3.row_dimensions[r].height = 24
    r += 1

    headers_mat_e = ["Kategori Materi", "Ungkapan / Materi Esensial", "Jenis Teks Utama (Text Types)", "Struktur Teks Khas", "Unsur Kebahasaan Utama", "Contoh Konteks Nyata Siswa SMA"]
    for i, h in enumerate(headers_mat_e, 1):
        cell = ws3.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws3.row_dimensions[r].height = 24
    r += 1

    materi_e_data = [
        ("Teks Interpersonal & Transaksional", "• Introducing oneself & others\n• Greeting, Opening & Closing Conversation\n• Thanking & Responding\n• Congratulating & Complimenting\n• Asking & Giving Opinion\n• Agreement & Disagreement\n• Comparing & Contrasting\n• Justifying personal preferences", "Percakapan Lisan, Dialog Singkat, Wawancara, Forum Diskusi Kelas", "Initiation - Response - Follow up (IRF)", "Simple Present, Past Tense, Modal Auxiliaries (can, could, would), Opinion verbs (think, believe), Conjunctions of contrast (however, although)", "Pertemuan siswa baru, apresiasi karya teman di kelas, diskusi pemilihan ketua kelas/OSIS."),
        ("Teks Fungsional Pendek (Multimodal)", "• Announcement (Pengumuman)\n• Personal Letter / Email\n• Advertisement (Iklan Layanan/Produk)\n• Invitation (Undangan resmi/akrab)\n• Poster & Infografis Edukatif", "Teks Cetak & Media Digital (Media Sosial, Situs Sekolah, Papan Buletin)", "Headings, Salutation, Body, Call to action, Signature/Contact", "Imperative sentences, Date/Time prepositions, Persuasive adjectives, Formal/Informal greetings", "Pengumuman kegiatan sekolah di mading, undangan pentas seni, iklan kampanye hemat energi."),
        ("Teks Faktual & Informasi (Descriptive)", "• Describing people, tourist destinations, and historical places\n• Sensory details & identification", "Descriptive Text (lisan & tertulis)", "Identification - Description", "Specific nouns, Adjective phrases, Simple Present Tense, Degree of comparison", "Mempromosikan potensi wisata daerah asal, mendeskripsikan tokoh inspiratif idola remaja."),
        ("Teks Prosedur (Procedure)", "• Step-by-step instructions\n• Food recipes, manual user, safety protocols\n• Digital safety guidelines", "Procedure Text", "Goal/Aim - Materials/Ingredients - Steps/Methods", "Action verbs (cut, plug, click), Temporal conjunctions (first, then, finally), Adverbials of manner", "Panduan keamanan siber media sosial, resep makanan tradisional lokal, petunjuk penggunaan lab."),
        ("Teks Narasi Lampau (Recount)", "• Personal recount (pengalaman berkesan)\n• Historical recount (peristiwa bersejarah)\n• Factual recount (laporan kunjungan studi)", "Recount Text", "Orientation - Events in chronological order - Re-orientation", "Simple Past Tense, Action verbs in past, Time connectives (yesterday, after that, when)", "Buku harian/jurnal belajar siswa, laporan karya wisata ilmiah, kilas balik perjuangan pahlawan lokal.")
    ]

    for item in materi_e_data:
        kat, ung, jt, st, uk, ctx = item
        ws3.cell(row=r, column=1, value=kat).alignment = align_left
        ws3.cell(row=r, column=2, value=ung).alignment = align_left
        ws3.cell(row=r, column=3, value=jt).alignment = align_left
        ws3.cell(row=r, column=4, value=st).alignment = align_left
        ws3.cell(row=r, column=5, value=uk).alignment = align_left
        ws3.cell(row=r, column=6, value=ctx).alignment = align_left

        ws3.cell(row=r, column=1).font = font_data_bold
        for c in range(2, 7):
            ws3.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws3.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws3.cell(row=r, column=c).fill = fill_zebra
        ws3.row_dimensions[r].height = 54
        r += 1

    auto_fit_columns(ws3, min_width=10, max_width=45)
    ws3.column_dimensions['A'].width = 24
    ws3.column_dimensions['B'].width = 30
    ws3.column_dimensions['C'].width = 26
    ws3.column_dimensions['D'].width = 24
    ws3.column_dimensions['E'].width = 32
    ws3.column_dimensions['F'].width = 36
    ws3.freeze_panes = 'A5'

    # ==========================================
    # SHEET 4: CP FASE F WAJIB (XI-XII)
    # ==========================================
    ws4 = wb.create_sheet(title="CP Fase F Wajib (XI-XII)")
    apply_title_banner(
        ws4,
        "CAPAIAN PEMBELAJARAN BAHASA INGGRIS FASE F - MAPEL WAJIB (KELAS XI & XII)",
        "Keputusan Kepala BSKAP No. 046/H/KR/2025 & Panduan Mata Pelajaran Bahasa Inggris Revisi 2025",
        max_col=6
    )

    ws4.cell(row=4, column=1, value="1. CAPAIAN PEMBELAJARAN UMUM FASE F WAJIB").font = font_section_header
    ws4.merge_cells("A4:F4")
    for c in range(1, 7):
        ws4.cell(row=4, column=c).fill = fill_tbl_header
    ws4.row_dimensions[4].height = 24

    cp_umum_f_id = (
        "Pada akhir Fase F, murid memahami alur informasi secara keseluruhan dari teks lisan fiksi dan nonfiksi tentang isu terkini atau topik terkait mata pelajaran lain, "
        "serta menggunakan bahasa Inggris untuk mengungkapkan pendapat dan mempertahankan argumen tentang topik yang dibahas. "
        "Murid mengevaluasi informasi tersurat dan tersirat dari berbagai jenis teks fiksi dan nonfiksi tertulis atau teks multimodal. "
        "Murid mengomunikasikan gagasan dan pengalaman secara tertulis atau multimodal dalam berbagai jenis teks fiksi dan nonfiksi dengan struktur teks dan unsur kebahasaan yang tepat dan kompleks "
        "menggunakan berbagai media presentasi (cetak dan digital), menunjukkan strategi koreksi mandiri, serta mempertahankan argumen tentang isu terkini atau topik lintas disiplin ilmu."
    )
    cp_umum_f_en = (
        "By the end of Phase F, students understand the overall flow of information in fiction and nonfiction spoken texts about current issues or topics related to other subjects, "
        "and use English to express opinions and defend arguments about the topics discussed. "
        "They evaluate explicit and implicit information from a variety of fiction and nonfiction written or multimodal texts. "
        "They communicate ideas and experiences in written or multimodal formats in fiction and nonfiction texts with appropriate and complex text structures and language features, "
        "using various presentation media (print and digital), demonstrate self-correction strategies, and defend arguments about current issues or interdisciplinary topics."
    )

    ws4.cell(row=5, column=1, value="Capaian Umum (Bahasa Indonesia):").font = font_data_bold
    ws4.cell(row=6, column=1, value=cp_umum_f_id).font = font_data
    ws4.merge_cells("A5:F5")
    ws4.merge_cells("A6:F6")
    ws4.cell(row=6, column=1).alignment = align_left

    ws4.cell(row=7, column=1, value="General Outcome (English):").font = font_data_bold
    ws4.cell(row=8, column=1, value=cp_umum_f_en).font = font_english
    ws4.merge_cells("A7:F7")
    ws4.merge_cells("A8:F8")
    ws4.cell(row=8, column=1).alignment = align_left

    for r_idx in [5, 6, 7, 8]:
        for c in range(1, 7):
            ws4.cell(row=r_idx, column=c).border = cell_border
            ws4.cell(row=r_idx, column=c).fill = fill_subgroup if r_idx in [5, 7] else fill_zebra
    ws4.row_dimensions[5].height = 20
    ws4.row_dimensions[6].height = 54
    ws4.row_dimensions[7].height = 20
    ws4.row_dimensions[8].height = 54

    r = 10
    ws4.cell(row=r, column=1, value="2. CAPAIAN PEMBELAJARAN PER ELEMEN FASE F WAJIB").font = font_section_header
    ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws4.cell(row=r, column=c).fill = fill_card_header
    ws4.row_dimensions[r].height = 24
    r += 1

    for i, h in enumerate(headers_elemen, 1):
        cell = ws4.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = cell_border
    ws4.row_dimensions[r].height = 26
    r += 1

    elemen_f_data = [
        (1, "Menyimak - Berbicara\n(Listening - Speaking)",
         "Memahami alur informasi secara keseluruhan dari teks lisan fiksi dan non fiksi tentang isu terkini atau topik terkait mata pelajaran lain; menggunakan bahasa Inggris untuk mengungkapkan pendapat dan mempertahankan argumen tentang topik yang dibahas.",
         "Understand the entire flow of information in fiction and nonfiction spoken texts about current issues or topics related to other subjects; use English to express their opinions and defend their arguments about the topics discussed.",
         "Memahami Alur Informasi Makro, Mempertahankan Argumen Terstruktur, Diskusi Lintas Disiplin Ilmu.",
         "Mampu berdebat formal, menyimak ceramah/podcast ilmiah, menyajikan presentasi argumentatif dengan bukti empiris."),
        (2, "Membaca - Memirsa\n(Reading - Viewing)",
         "Mengevaluasi informasi tersurat dan tersirat dari berbagai jenis teks fiksi dan non fiksi tertulis atau teks multimodal tentang isu terkini or topik yang terkait mata pelajaran lain.",
         "Evaluate explicit and implicit information in fiction and nonfiction from written and multimodal texts about current issues or topics related to other subjects.",
         "Evaluasi Kritis Informasi Tersirat, Analisis Bias & Sudut Pandang, Komprehensi Teks Akademis Lintas Disiplin.",
         "Mampu membedakan fakta vs opini penulis, mengevaluasi kekuatan argumen esai editorial, menyimpulkan pesan tersirat teks naratif kompleks."),
        (3, "Menulis - Mempresentasikan\n(Writing - Presenting)",
         "Mengomunikasikan gagasan dan pengalaman secara tertulis atau multimodal dalam berbagai jenis teks fiksi dan nonfiksi dengan struktur teks dan unsur kebahasaan yang tepat dan kompleks dengan menggunakan berbagai media presentasi (cetak dan digital) untuk mencapai tujuan yang berbeda-beda; menunjukkan strategi koreksi dalam kaidah menulis baik dengan bantuan pendidik maupun mandiri; mengungkapkan pendapat dan mempertahankan argumen tentang isu terkini atau topik yang terkait mata pelajaran lain.",
         "Communicate ideas and experiences in written or multimodal texts in fiction and nonfiction texts with appropriate and complex text structures and language features, using different media of presentation (print and digital) to achieve different purposes; demonstrate correction strategies in writing conventions with the help of a teacher or independently; express their opinions and defend their arguments on current issues or topics related to other subjects.",
         "Penyusunan Esai Kompleks, Strategi Koreksi Mandiri (Self-Editing), Presentasi Digital Multimodal, Pembuktian Argumen.",
         "Mampu menulis esai eksposisi analitis & hortatori berbobot 300-400 kata, menyusun teks diskusi pro-kontra, menghasilkan presentasi multimedia profesional.")
    ]

    for item in elemen_f_data:
        no, el, id_desc, en_desc, fok, tar = item
        ws4.cell(row=r, column=1, value=no).alignment = align_center
        ws4.cell(row=r, column=2, value=el).alignment = align_left
        ws4.cell(row=r, column=3, value=id_desc).alignment = align_top_left
        ws4.cell(row=r, column=4, value=en_desc).alignment = align_top_left
        ws4.cell(row=r, column=5, value=fok).alignment = align_top_left
        ws4.cell(row=r, column=6, value=tar).alignment = align_top_left

        ws4.cell(row=r, column=1).font = font_data_bold
        ws4.cell(row=r, column=2).font = font_data_bold
        ws4.cell(row=r, column=3).font = font_data
        ws4.cell(row=r, column=4).font = font_english
        ws4.cell(row=r, column=5).font = font_data
        ws4.cell(row=r, column=6).font = font_data
        
        for c in range(1, 7):
            ws4.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws4.cell(row=r, column=c).fill = fill_zebra
        ws4.row_dimensions[r].height = 70
        r += 1

    r += 1
    # Materi Esensial Fase F Wajib
    ws4.cell(row=r, column=1, value="3. PEMETAAN MATERI ESENSIAL & TEKS UTAMA FASE F WAJIB (KELAS XI & XII)").font = font_section_header
    ws4.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws4.cell(row=r, column=c).fill = fill_tbl_header
    ws4.row_dimensions[r].height = 24
    r += 1

    for i, h in enumerate(headers_mat_e, 1):
        cell = ws4.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws4.row_dimensions[r].height = 24
    r += 1

    materi_f_data = [
        ("Teks Eksposisi Analitis\n(Analytical Exposition)", "• Defending opinions on global issues\n• Thesis statement formulation\n• Evidence-based arguments\n• Reiteration / Conclusion", "Esai Argumentatif, Artikel Opini, Surat Pembaca", "Thesis - Arguments (Point & Elaboration) - Reiteration", "Passive Voice, Causal Conjunctions (because, consequently, therefore), Emotive language, Evaluative adjectives, Present Tense", "Isu lingkungan (perubahan iklim, sampah plastik), urgensi literasi finansial remaja."),
        ("Teks Eksposisi Hortatori\n(Hortatory Exposition)", "• Persuading the audience that something should or should not be done\n• Call to action and recommendations", "Pidato Persuasif, Editorial Kebijakan, Kampanye Sosial", "Thesis - Arguments - Recommendation", "Modal verbs for recommendation (should, ought to, must), Thinking verbs, Connectors of addition & result", "Pentingnya sarapan sehat sebelum belajar, pembatasan penggunaan gawai saat larut malam."),
        ("Teks Diskusi\n(Discussion Text)", "• Presenting at least two perspectives on controversial issues\n• Weighing pros and cons\n• Impartial conclusion / recommendation", "Naskah Debat, Esai Komparatif Pro-Kontra, Artikel Analisis Isu", "Issue - Arguments For & Elaboration - Arguments Against & Elaboration - Conclusion/Recommendation", "Contrasting connectors (on the one hand, on the other hand, in contrast), Complex sentences, Passive voice, Modal verbs", "Pro-kontra penggunaan Kecerdasan Buatan (AI) dalam tugas sekolah, sistem transportasi publik ramah lingkungan."),
        ("Teks Naratif Kompleks\n(Narrative Text)", "• Exploring moral dilemmas & character development\n• Legends, short stories, contemporary narratives\n• Figurative language & symbolism", "Cerita Pendek, Teks Sastra Populer, Fabel Modern", "Orientation - Complication - Evaluation - Resolution - Coda/Moral Value", "Past Continuous, Past Perfect, Direct and Indirect Speech, Action verbs, Metaphor, Simile", "Kisah inspiratif tokoh kemanusiaan, adaptasi legenda lokal ke bahasa Inggris dengan pesan moral universal.")
    ]

    for item in materi_f_data:
        kat, ung, jt, st, uk, ctx = item
        ws4.cell(row=r, column=1, value=kat).alignment = align_left
        ws4.cell(row=r, column=2, value=ung).alignment = align_left
        ws4.cell(row=r, column=3, value=jt).alignment = align_left
        ws4.cell(row=r, column=4, value=st).alignment = align_left
        ws4.cell(row=r, column=5, value=uk).alignment = align_left
        ws4.cell(row=r, column=6, value=ctx).alignment = align_left

        ws4.cell(row=r, column=1).font = font_data_bold
        for c in range(2, 7):
            ws4.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws4.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws4.cell(row=r, column=c).fill = fill_zebra
        ws4.row_dimensions[r].height = 54
        r += 1

    auto_fit_columns(ws4, min_width=10, max_width=45)
    ws4.column_dimensions['A'].width = 24
    ws4.column_dimensions['B'].width = 30
    ws4.column_dimensions['C'].width = 26
    ws4.column_dimensions['D'].width = 24
    ws4.column_dimensions['E'].width = 32
    ws4.column_dimensions['F'].width = 36
    ws4.freeze_panes = 'A5'

    # ==========================================
    # SHEET 5: CP FASE F TINGKAT LANJUT
    # ==========================================
    ws5 = wb.create_sheet(title="CP Fase F Tingkat Lanjut")
    apply_title_banner(
        ws5,
        "CAPAIAN PEMBELAJARAN BAHASA INGGRIS FASE F TINGKAT LANJUT (MAPEL PILIHAN SMA)",
        "Keputusan Kepala BSKAP No. 046/H/KR/2025 & Panduan Mata Pelajaran Bahasa Inggris Revisi 2025",
        max_col=6
    )

    ws5.cell(row=4, column=1, value="1. CAPAIAN PEMBELAJARAN UMUM BAHASA INGGRIS TINGKAT LANJUT").font = font_section_header
    ws5.merge_cells("A4:F4")
    for c in range(1, 7):
        ws5.cell(row=4, column=c).fill = fill_tbl_header
    ws5.row_dimensions[4].height = 24

    cp_umum_adv_id = (
        "Pada akhir Fase F Tingkat Lanjut, murid memahami alur cerita atau informasi dari teks narasi, eksposisi, dan diskusi tentang isu terkini dan topik terkait mata pelajaran lain "
        "secara lisan atau multimodal; menggunakan bahasa Inggris untuk mengungkapkan pendapat dan mempertahankan argumen tentang topik yang dibahas. "
        "Murid mengevaluasi dan merefleksi informasi tersurat dan tersirat dari teks narasi, eksposisi dan diskusi tentang isu terkini baik nasional maupun global, atau topik terkait mata pelajaran lain. "
        "Murid mengomunikasikan gagasan dan pengalaman secara tertulis atau multimodal dalam teks narasi, eksposisi, dan diskusi dengan struktur teks dan unsur kebahasaan yang tepat, jelas, dan detail; "
        "menulis cerita pendek, pendapat atau pandangan tentang suatu topik/isu dengan mengungkapkan argumen yang didukung data dan fakta, atau membahas suatu isu dengan menjelaskan manfaat dan kelemahan tentang isu kontroversial."
    )
    cp_umum_adv_en = (
        "By the end of Phase F Advanced, students understand the storyline or information from narrative, exposition, and discussion texts on current issues and topics related to other subjects in spoken or multimodal formats; "
        "use English to express opinions and defend arguments about the topics discussed. "
        "They evaluate and reflect on explicit and implicit information from narrative, exposition, and discussion texts on current issues (national and global) or topics related to other subjects. "
        "They communicate ideas and experiences in written or multimodal texts in narrative, exposition, and discussion texts with clear and detailed text structures and language features; "
        "write short stories, express opinions using evidence-based arguments with data and facts, or discuss controversial current issues with in-depth arguments for and against."
    )

    ws5.cell(row=5, column=1, value="Capaian Umum Tingkat Lanjut (Bahasa Indonesia):").font = font_data_bold
    ws5.cell(row=6, column=1, value=cp_umum_adv_id).font = font_data
    ws5.merge_cells("A5:F5")
    ws5.merge_cells("A6:F6")
    ws5.cell(row=6, column=1).alignment = align_left

    ws5.cell(row=7, column=1, value="General Outcome - Advanced (English):").font = font_data_bold
    ws5.cell(row=8, column=1, value=cp_umum_adv_en).font = font_english
    ws5.merge_cells("A7:F7")
    ws5.merge_cells("A8:F8")
    ws5.cell(row=8, column=1).alignment = align_left

    for r_idx in [5, 6, 7, 8]:
        for c in range(1, 7):
            ws5.cell(row=r_idx, column=c).border = cell_border
            ws5.cell(row=r_idx, column=c).fill = fill_subgroup if r_idx in [5, 7] else fill_zebra
    ws5.row_dimensions[5].height = 20
    ws5.row_dimensions[6].height = 58
    ws5.row_dimensions[7].height = 20
    ws5.row_dimensions[8].height = 58

    r = 10
    ws5.cell(row=r, column=1, value="2. CAPAIAN PEMBELAJARAN PER ELEMEN BAHASA INGGRIS TINGKAT LANJUT").font = font_section_header
    ws5.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws5.cell(row=r, column=c).fill = fill_card_header
    ws5.row_dimensions[r].height = 24
    r += 1

    for i, h in enumerate(headers_elemen, 1):
        cell = ws5.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = cell_border
    ws5.row_dimensions[r].height = 26
    r += 1

    elemen_adv_data = [
        (1, "Menyimak - Berbicara\n(Listening - Speaking)",
         "Memahami alur cerita atau informasi dari teks narasi, eksposisi, dan diskusi tentang isu terkini dan topik terkait mata pelajaran lain secara lisan atau multimodal; menggunakan bahasa Inggris untuk mengungkapkan pendapat dan mempertahankan argumen tentang topik yang dibahas.",
         "Understand the storyline or information from narrative, exposition, and discussion texts on current issues and topics related to other subjects in spoken or multimodal texts; use English to express their opinions and defend their arguments about the topics being discussed.",
         "Debat Akademis Tingkat Tinggi, Wawancara Mendalam, Presentasi Multimodal Berbasis Data Ilmiah.",
         "Mampu menjadi pembicara dalam model United Nations (MUN), menyampaikan sanggahan kritis (rebuttal), menyajikan riset lisan."),
        (2, "Membaca - Memirsa\n(Reading - Viewing)",
         "Mengevaluasi dan merefleksi informasi tersurat dan tersirat dari teks narasi, eksposisi dan diskusi tentang isu terkini baik nasional maupun global, atau topik yang terkait mata pelajaran lain secara tertulis atau multimodal.",
         "Evaluate and reflect on explicit and implicit information from narrative, exposition, and discussion on current issues, both national and global, or topics related to other subjects in written or multimodal texts.",
         "Critical Discourse Analysis (CDA), Evaluasi Retorika Penulis, Refleksi Kritis terhadap Isu Global Kontroversial.",
         "Mampu mengidentifikasi argumen tersembunyi (hidden agenda), menganalisis gaya sastra cerpen berbahasa Inggris, merefleksikan implikasi moral."),
        (3, "Menulis - Mempresentasikan\n(Writing - Presenting)",
         "Mengomunikasikan gagasan dan pengalaman secara tertulis atau multimodal dalam teks narasi, eksposisi, dan diskusi tentang berbagai topik dengan struktur teks dan unsur kebahasaan yang tepat, jelas, dan detail dengan menggunakan berbagai moda presentasi (cetak dan digital); menulis cerita pendek, pendapat atau pandangan tentang suatu topik atau isu dengan mengungkapkan argumen yang didukung dengan data dan fakta, atau membahas suatu isu dengan menjelaskan manfaat dan kelemahan atau argumen yang mendukung dan menentang tentang suatu isu terkini yang kontroversial.",
         "Communicate ideas and experiences in written or multimodal texts in narrative, exposition, and discussion texts on various topics with clear and detailed text structure and language features using different modes of presentation (print and digital); write short stories, their opinions or viewpoints about a topic or an issue by using arguments supported with data and facts, or discuss an issue by explaining the advantages and disadvantages or arguments for and against controversial current issues.",
         "Creative Writing (Menulis Cerita Pendek), Academic Essay Writing (Data & Fakta), Menulis Analisis Isu Kontroversial Berimbang.",
         "Mampu menghasilkan cerpen orisinal berbahasa Inggris, esai ilmiah mini 500-600 kata dengan kutipan referensi akademis terpercaya.")
    ]

    for item in elemen_adv_data:
        no, el, id_desc, en_desc, fok, tar = item
        ws5.cell(row=r, column=1, value=no).alignment = align_center
        ws5.cell(row=r, column=2, value=el).alignment = align_left
        ws5.cell(row=r, column=3, value=id_desc).alignment = align_top_left
        ws5.cell(row=r, column=4, value=en_desc).alignment = align_top_left
        ws5.cell(row=r, column=5, value=fok).alignment = align_top_left
        ws5.cell(row=r, column=6, value=tar).alignment = align_top_left

        ws5.cell(row=r, column=1).font = font_data_bold
        ws5.cell(row=r, column=2).font = font_data_bold
        ws5.cell(row=r, column=3).font = font_data
        ws5.cell(row=r, column=4).font = font_english
        ws5.cell(row=r, column=5).font = font_data
        ws5.cell(row=r, column=6).font = font_data
        
        for c in range(1, 7):
            ws5.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws5.cell(row=r, column=c).fill = fill_zebra
        ws5.row_dimensions[r].height = 76
        r += 1

    r += 1
    # Perbedaan Esensial Wajib vs Tingkat Lanjut
    ws5.cell(row=r, column=1, value="3. DISTINKSI MENDALAM: BAHASA INGGRIS WAJIB VS TINGKAT LANJUT").font = font_section_header
    ws5.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws5.cell(row=r, column=c).fill = fill_tbl_header
    ws5.row_dimensions[r].height = 24
    r += 1

    headers_dist = ["Dimensi Pembeda", "Bahasa Inggris Wajib (General English)", "Bahasa Inggris Tingkat Lanjut (Advanced English)", "Kompleksitas Linguistik", "Ekspektasi Output Siswa", "Relevansi Karir & Studi Lanjut"]
    for i, h in enumerate(headers_dist, 1):
        cell = ws5.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws5.row_dimensions[r].height = 24
    r += 1

    dist_data = [
        ("Tujuan Pembelajaran", "Kemahiran komunikasi fungsional sehari-hari dan literasi teks umum.", "Kemahiran akademik (Academic English), apresiasi sastra (Literature), dan diskursus kritis.", "General Vocabulary vs Academic Word List (AWL)", "Paragraf opini dan esai sederhana.", "Kecakapan kerja umum & komunikasi global standar."),
        ("Jenis Teks (Genre)", "Eksposisi sederhana, prosedur, recount, deskriptif, teks fungsional.", "Narasi sastra mendalam (cerpen), Eksposisi analitis berbasis riset data, Diskusi isu kontroversial multi-sudut pandang.", "Simple/Compound Sentences vs Complex Nominal Clauses & Passive Structures", "Naskah cerpen, esai komparatif berbobot, riset mini multimodal.", "Kesiapan kuliah luar negeri, jurusan Hubungan Internasional, Sastra/Linguistik, Hukum Internasional."),
        ("Bobot Waktu Belajar", "3 JP/minggu (108 JP di Kelas XI / 96 JP di Kelas XII).", "5 JP/minggu (180 JP di Kelas XI / 160 JP di Kelas XII).", "Fokus kelancaran komunikasi (Fluency).", "Kedalaman analisis, akurasi, dan keindahan retorika (Accuracy & Aesthetics).", "Kemampuan skor tes internasional (IELTS / TOEFL / Duolingo English Test)."),
        ("Tuntutan Kognitif", "Memahami informasi umum dan menyimpulkan ide pokok tersurat/tersirat.", "Mengevaluasi argumen tersembunyi, merefleksi nilai moral kultural, dan menimbang bias informasi.", "Level Kognitif C2 - C4 (Bloom).", "Level Kognitif C4 - C6 (Menganalisis, Mengevaluasi, Mengkreasi).", "Kemampuan berpikir tingkat tinggi (HOTS) dan kepemimpinan global.")
    ]

    for item in dist_data:
        dp, bg, ba, kl, out, rel = item
        ws5.cell(row=r, column=1, value=dp).alignment = align_left
        ws5.cell(row=r, column=2, value=bg).alignment = align_left
        ws5.cell(row=r, column=3, value=ba).alignment = align_left
        ws5.cell(row=r, column=4, value=kl).alignment = align_left
        ws5.cell(row=r, column=5, value=out).alignment = align_left
        ws5.cell(row=r, column=6, value=rel).alignment = align_left

        ws5.cell(row=r, column=1).font = font_data_bold
        for c in range(2, 7):
            ws5.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws5.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws5.cell(row=r, column=c).fill = fill_zebra
        ws5.row_dimensions[r].height = 42
        r += 1

    auto_fit_columns(ws5, min_width=10, max_width=45)
    ws5.column_dimensions['A'].width = 22
    ws5.column_dimensions['B'].width = 30
    ws5.column_dimensions['C'].width = 32
    ws5.column_dimensions['D'].width = 26
    ws5.column_dimensions['E'].width = 30
    ws5.column_dimensions['F'].width = 34
    ws5.freeze_panes = 'A5'

    # ==========================================
    # SHEET 6: PEDAGOGI DEEP LEARNING & GBA
    # ==========================================
    ws6 = wb.create_sheet(title="Pedagogi Deep Learning & GBA")
    apply_title_banner(
        ws6,
        "KESELARASAN PEMBELAJARAN MENDALAM DENGAN PENDEKATAN BERBASIS TEKS (GENRE)",
        "Implementasi Teaching Learning Cycle (TLC) dalam Kerangka Deep Learning Permendikdasmen No. 13 Tahun 2025",
        max_col=6
    )

    ws6.cell(row=4, column=1, value="1. MATRIKS KESELARASAN: PEMBELAJARAN MENDALAM DENGAN GENRE-BASED APPROACH (GBA)").font = font_section_header
    ws6.merge_cells("A4:F4")
    for c in range(1, 7):
        ws6.cell(row=4, column=c).fill = fill_tbl_header
    ws6.row_dimensions[4].height = 24

    headers_gba = ["Prinsip Pembelajaran Mendalam", "Prinsip Pengajaran Berbasis Teks (Genre)", "Tahap Teaching Learning Cycle (TLC)", "Pengalaman Belajar", "Aktivitas Otentik Guru & Murid di Kelas SMA", "Indikator Ketercapaian"]
    r = 5
    for i, h in enumerate(headers_gba, 1):
        cell = ws6.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws6.row_dimensions[5].height = 26
    r = 6

    gba_matrix = [
        ("Bermakna (Meaningful Learning)",
         "Pengajaran berfokus pada teks utuh yang bermakna dan mempunyai tujuan sosial nyata dalam kehidupan.",
         "Stage 1: Building Knowledge of the Field (BKoF)",
         "MEMAHAMI (Understanding Context & Field)",
         "Membahas topik melalui video/gambar otentik, tanya jawab pemantik, eksplorasi isu budaya lokal, mengenalkan konteks sosial di mana teks digunakan.",
         "Murid memahami mengapa teks tersebut penting dalam interaksi manusia nyata."),
        ("Berkesadaran (Mindful Learning)",
         "Tindakan berbahasa dipikirkan dengan matang, mengenali struktur teks dan fitur bahasa secara terencana dan metakognitif.",
         "Stage 2: Modelling of the Text (MoT)",
         "MEMAHAMI (Analyzing Text Architecture)",
         "Membedah teks model bersama guru, memetakan Generic Structure (misal Thesis, Arguments, Reiteration), menganalisis Language Features dan tenses.",
         "Murid menyadari kaidah linguistik dan strategi penyampaian pesan yang efektif."),
        ("Menggembirakan (Joyful) & Kolaboratif",
         "Kerja sama saling memuliakan ide teman, membangun rasa percaya diri dalam mencoba bahasa baru tanpa takut salah.",
         "Stage 3: Joint Construction of the Text (JCoT)",
         "MENGAPLIKASI (Collaborative Construction)",
         "Menyusun teks secara kolaboratif dalam kelompok kecil atau bersama guru di papan tulis/dokumen digital bersama, saling memberi umpan balik suportif.",
         "Murid merasakan dukungan teman sejawat dan mengalami proses scaffolding yang aman."),
        ("Kemandirian & Aktualisasi Diri",
         "Mencipta karya teks lisan/tulis orisinal, mengekspresikan suara unik (student's voice) dengan rasa bangga dan percaya diri.",
         "Stage 4: Independent Construction of the Text (ICoT)",
         "MENGAPLIKASI (Independent Creation)",
         "Menulis esai mandiri, membuat rekaman video/podcast, menyusun cerpen, atau menyajikan presentasi publik di hadapan audiens nyata.",
         "Murid menguasai kompetensi secara mandiri dan bangga atas karya bahasa Inggrisnya."),
        ("Refleksi Berkesadaran",
         "Menilai keberhasilan, tantangan yang dihadapi, serta strategi perbaikan metakognisi untuk pembelajaran selanjutnya.",
         "Di Semua Tahapan (Integrated Reflection)",
         "MEREFLEKSI (Reflecting & Metacognition)",
         "Mengisi lembar evaluasi diri (self-assessment), jurnal refleksi belajar (learning log), merumuskan rencana tindak lanjut perbaikan.",
         "Murid berkembang menjadi pembelajar mandiri sepanjang hayat (lifelong learners).")
    ]

    for item in gba_matrix:
        prs, gba, stg, exp, act, ind = item
        ws6.cell(row=r, column=1, value=prs).alignment = align_left
        ws6.cell(row=r, column=2, value=gba).alignment = align_left
        ws6.cell(row=r, column=3, value=stg).alignment = align_left
        ws6.cell(row=r, column=4, value=exp).alignment = align_center
        ws6.cell(row=r, column=5, value=act).alignment = align_left
        ws6.cell(row=r, column=6, value=ind).alignment = align_left

        ws6.cell(row=r, column=1).font = font_data_bold
        ws6.cell(row=r, column=2).font = font_data
        ws6.cell(row=r, column=3).font = font_data_bold
        ws6.cell(row=r, column=4).font = font_data_bold
        ws6.cell(row=r, column=5).font = font_data
        ws6.cell(row=r, column=6).font = font_data
        
        for c in range(1, 7):
            ws6.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws6.cell(row=r, column=c).fill = fill_zebra
        ws6.row_dimensions[r].height = 54
        r += 1

    r += 1
    # Integrasi 8 Profil Lulusan dalam Bahasa Inggris
    ws6.cell(row=r, column=1, value="2. INTEGRASI 8 DIMENSI PROFIL LULUSAN DALAM KELAS BAHASA INGGRIS SMA").font = font_section_header
    ws6.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws6.cell(row=r, column=c).fill = fill_tbl_header
    ws6.row_dimensions[r].height = 24
    r += 1

    headers_p8 = ["No", "Dimensi Profil Lulusan", "Relevansi Khusus Pembelajaran Bahasa Inggris", "Bentuk Tugas / Aktivitas Kelas", "Indikator Sikap & Karakter Murid", "Bukti Asesmen"]
    for i, h in enumerate(headers_p8, 1):
        cell = ws6.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws6.row_dimensions[r].height = 24
    r += 1

    p8_data = [
        (1, "Komunikasi (Communication)", "Fungsi inti bahasa Inggris sebagai sarana komunikasi lisan, tulisan, dan multimodal global.", "Presentasi publik, debat terstruktur, wawancara, korespondensi email formal.", "Menyampaikan pesan secara lugas, santun, dan responsif terhadap mitra bicara.", "Rubrik kelancaran dan ketepatan berbicara."),
        (2, "Penalaran Kritis (Critical Thinking)", "Menganalisis teks argumentasi, menyaring hoaks berbahasa asing, mengevaluasi bukti empiris.", "Membedah esai analitis, menganalisis bias berita internasional, verifikasi fakta teks diskusi.", "Kritis terhadap klaim tanpa data, mampu menyusun premis dan konklusi logis.", "Lembar analisis teks kritis (critical reading worksheet)."),
        (3, "Kewargaan (Citizenship)", "Mengenalkan budaya Indonesia ke kancah global dan memahami dinamika masyarakat internasional.", "Menulis deskripsi warisan budaya lokal dalam bahasa Inggris, kampanye isu global PBB.", "Bangga akan jati diri bangsa sembari menghargai keberagaman budaya dunia.", "Proyek digital storytelling kearifan lokal."),
        (4, "Kreativitas (Creativity)", "Mencipta karya bahasa orisinal yang memikat, memadukan teks dengan unsur visual/audio digital.", "Menulis cerita pendek (creative writing), memproduksi video vlog/podcast interaktif.", "Gagasan orisinal, penggunaan metafora imajinatif, kemasan presentasi menarik.", "Portofolio cerpen dan rekaman media podcast."),
        (5, "Kolaborasi (Collaboration)", "Bekerja sama dalam proyek teks bersama (JCoT), simulasi sidang diplomatik, peer review.", "Penyusunan naskah drama/debat tim, saling menyunting draf tulisan teman secara konstruktif.", "Mendengarkan aktif, menghargai masukan teman, berbagi peran secara adil.", "Lembar observasi peer-review kolaboratif."),
        (6, "Kemandirian (Self-Reliance)", "Meregulasi proses belajar bahasa asing, menambah perbendaharaan kosakata secara otonom.", "Menyusun jurnal belajar kosa kata mandiri (vocabulary journal), self-editing naskah esai.", "Inisiatif mencari sumber belajar otentik tanpa harus selalu disuruh pendidik.", "Jurnal refleksi diri dan checklist revisi tulisan."),
        (7, "Kesehatan (Health & Wellbeing)", "Membahas topik gaya hidup sehat, kebugaran remaja, dan kesehatan mental dalam teks eksposisi.", "Menulis teks prosedur menu diet sehat remaja, diskusi bahaya screen time berlebih.", "Memahami pentingnya keseimbangan emosional dan fisik melalui diskursus teks.", "Esai kampanye gaya hidup sehat remaja."),
        (8, "Keimanan & Ketakwaan", "Menginternalisasi adab saling memuliakan (K.H. Hasyim Asy'ari) saat berbeda pendapat dalam debat.", "Menjaga etika berkomunikasi santun (netiquette) saat berinteraksi lisan maupun digital.", "Santun dalam menyanggah argumen, toleran terhadap perbedaan pandangan.", "Catatan anekdotal etika interaksi kelas.")
    ]

    for item in p8_data:
        no, dim, rel, act, ind, ev = item
        ws6.cell(row=r, column=1, value=no).alignment = align_center
        ws6.cell(row=r, column=2, value=dim).alignment = align_left
        ws6.cell(row=r, column=3, value=rel).alignment = align_left
        ws6.cell(row=r, column=4, value=act).alignment = align_left
        ws6.cell(row=r, column=5, value=ind).alignment = align_left
        ws6.cell(row=r, column=6, value=ev).alignment = align_left

        ws6.cell(row=r, column=1).font = font_data_bold
        ws6.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 7):
            ws6.cell(row=r, column=c).font = font_data
            
        for c in range(1, 7):
            ws6.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws6.cell(row=r, column=c).fill = fill_zebra
        ws6.row_dimensions[r].height = 42
        r += 1

    auto_fit_columns(ws6, min_width=10, max_width=45)
    ws6.column_dimensions['A'].width = 5
    ws6.column_dimensions['B'].width = 24
    ws6.column_dimensions['C'].width = 34
    ws6.column_dimensions['D'].width = 34
    ws6.column_dimensions['E'].width = 32
    ws6.column_dimensions['F'].width = 32
    ws6.freeze_panes = 'A5'

    # ==========================================
    # SHEET 7: PEMETAAN GENRE & MATERI SMA
    # ==========================================
    ws7 = wb.create_sheet(title="Pemetaan Genre & Materi SMA")
    apply_title_banner(
        ws7,
        "MATRIKS KOMPREHENSIF GENRE & TEKS BAHASA INGGRIS SMA",
        "Rujukan Analisis Struktur Teks, Unsur Kebahasaan, dan Capaian Fase E & Fase F",
        max_col=7
    )

    headers_genre = ["Fase / Jenjang", "Jenis Teks (Genre)", "Fungsi Sosial (Social Function)", "Struktur Teks (Generic Structure)", "Unsur Kebahasaan Kunci (Language Features)", "Contoh Topik Kontekstual di SMA", "Tolak Ukur Kelulusan Fase"]
    r = 4
    for i, h in enumerate(headers_genre, 1):
        cell = ws7.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = cell_border
    ws7.row_dimensions[4].height = 26
    r = 5

    master_genres = [
        ("Fase E (Kelas X)", "Descriptive Text", "Mendeskripsikan benda, orang, tempat wisata, atau bangunan bersejarah secara rinci dan hidup.", "1. Identification (pengenalan objek)\n2. Description (ciri fisik, kualitas, sifat)", "Simple Present, Adjectives, Adverbial phrases, Sensory language", "Destinasi wisata unggulan daerah, seniman/atlet inspiratif nasional.", "Mampu menyusun teks deskripsi 150-200 kata dengan perincian visual hidup."),
        ("Fase E (Kelas X)", "Recount Text", "Menceritakan kembali peristiwa masa lampau secara kronologis untuk menghibur atau memberi informasi.", "1. Orientation (latar tokoh, tempat, waktu)\n2. Events (rangkaian peristiwa kronologis)\n3. Re-orientation (kesimpulan/kesan)", "Simple Past Tense, Action verbs, Chronological connectors (first, then, after that)", "Pengalaman karyawisata ke museum sains, sejarah berdirinya sekolah, peristiwa bersejarah Sumpah Pemuda.", "Mampu merekonstruksi peristiwa masa lalu dengan urutan waktu yang logis dan kohesif."),
        ("Fase E (Kelas X)", "Procedure Text", "Menjelaskan cara membuat atau mengoperasikan sesuatu melalui tahapan teratur.", "1. Goal/Aim (tujuan kegiatan)\n2. Materials/Equipment (alat & bahan)\n3. Steps/Methods (langkah berurutan)", "Imperatives, Temporal connectives, Action verbs, Adverbials of time/manner", "Panduan keamanan bermedia sosial, langkah pertolongan pertama kecelakaan ringan, resep kuliner nusantara.", "Mampu mendemonstrasikan prosedur lisan dan menyusun teks panduan terstruktur."),
        ("Fase E (Kelas X)", "Teks Fungsional Pendek (Multimodal)", "Menyampaikan informasi praktis singkat kepada khalayak spesifik.", "Announcement, Letter, Advertisement, Invitation, Poster", "Singkat, padat, persuasif, font variatif, tata letak grafis menarik", "Pengumuman lomba antar-SMA, poster kampanye hemat air, undangan pembukaan pameran seni.", "Mampu merancang produk multimodal digital yang komunikatif dan estetik."),
        ("Fase F Wajib (XI-XII)", "Analytical Exposition", "Meyakinkan pembaca bahwa sebuah topik atau fenomena penting dan memerlukan perhatian serius.", "1. Thesis (pengenalan topik & posisi penulis)\n2. Arguments (rangkaian alasan + bukti fakta)\n3. Reiteration (penegasan kembali tesis)", "Simple Present, Causal connectors (consequently, because), Emotive words, Evaluative vocabulary", "Urgensi memilah sampah di sekolah, bahaya gaya hidup sedentary pada remaja, pentingnya literasi digital.", "Mampu menulis esai analitis 300 kata yang dilengkapi data pendukung objektif."),
        ("Fase F Wajib (XI-XII)", "Hortatory Exposition", "Mempengaruhi pembaca agar melakukan atau tidak melakukan sesuatu tindakan spesifik.", "1. Thesis (isu permasalahan)\n2. Arguments (alasan pendukung sikap)\n3. Recommendation (tindakan yang disarankan)", "Modal verbs (should, must, ought to), Action verbs, Connectors of cause-effect", "Saran pembatasan jam main game online, rekomendasi penyediaan jalur sepeda di sekolah.", "Mampu menyampaikan pidato persuasif yang memuat rekomendasi solutif yang realistis."),
        ("Fase F Wajib (XI-XII)", "Discussion Text", "Menyajikan tinjauan seimbang dari dua sudut pandang berbeda (pro dan kontra) mengenai isu kontroversial.", "1. Issue (isu yang diperdebatkan)\n2. Arguments For (argumen mendukung)\n3. Arguments Against (argumen menentang)\n4. Conclusion (sintesis berimbang)", "Contrasting connectives (however, nevertheless, on the contrary), Passive voice, Modality", "Penggunaan gawai di dalam ruang kelas, sistem zonasi sekolah, kecerdasan buatan dalam penulisan ilmiah.", "Mampu memoderasi debat kelas dan menyusun teks diskusi berimbang tanpa bias personal kaku."),
        ("Fase F Wajib (XI-XII)", "Narrative Text", "Menghibur pembaca sekaligus menanamkan nilai-nilai moral kehidupan melalui jalinan cerita dramatik.", "1. Orientation\n2. Complication (konflik berkembang)\n3. Evaluation (puncak ketegangan)\n4. Resolution (penyelesaian konflik)\n5. Coda (pesan moral)", "Past continuous, Past perfect, Direct speech, Metaphors, Descriptive imagery", "Cerita rakyat nusantara yang dialihbahasakan ke bahasa Inggris, cerpen modern tentang keteguhan hati remaja.", "Mampu menarasikan cerita secara memikat dengan intonasi lisan dramatis dan struktur baku."),
        ("Fase F Tingkat Lanjut", "Advanced Academic Essay", "Menganalisis diskursus akademis ilmiah dengan metodologi penalaran kritis tingkat tinggi.", "1. Introduction (hook, background, thesis)\n2. Body Paragraphs (topic sentence, evidence, analysis)\n3. Counterargument & Rebuttal\n4. Conclusion", "Academic Word List (AWL), Complex nominal syntax, In-text citations, Hedging language (suggests, indicates)", "Dampak globalisasi terhadap pelestarian bahasa daerah, transisi energi hijau di negara berkembang.", "Mampu menyusun esai akademis 500-600 kata dengan kutipan ilmiah dan sanggahan argumen kuat."),
        ("Fase F Tingkat Lanjut", "Creative Short Story Writing", "Mencipta karya fiksi orisinal dengan eksplorasi tema psikologis, dilema moral, dan kedalaman karakter.", "Plot diagram lengkap, character arc, theme development, literary devices", "Show don't tell, Dialogue tags, Symbolism, Foreshadowing, Tone and Mood", "Cerita pendek tentang pencarian jati diri remaja, fiksi spekulatif masa depan Indonesia tahun 2045.", "Mampu mempublikasikan cerpen bahasa Inggris pada majalah dinding, blog sekolah, atau antologi buku.")
    ]

    for item in master_genres:
        fase, jt, fs, st, uk, ctx, tlk = item
        ws7.cell(row=r, column=1, value=fase).alignment = align_center
        ws7.cell(row=r, column=2, value=jt).alignment = align_left
        ws7.cell(row=r, column=3, value=fs).alignment = align_left
        ws7.cell(row=r, column=4, value=st).alignment = align_left
        ws7.cell(row=r, column=5, value=uk).alignment = align_left
        ws7.cell(row=r, column=6, value=ctx).alignment = align_left
        ws7.cell(row=r, column=7, value=tlk).alignment = align_left

        ws7.cell(row=r, column=1).font = font_data_bold
        ws7.cell(row=r, column=2).font = font_data_bold
        for c in range(3, 8):
            ws7.cell(row=r, column=c).font = font_data
            
        for c in range(1, 8):
            ws7.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws7.cell(row=r, column=c).fill = fill_zebra
        ws7.row_dimensions[r].height = 56
        r += 1

    auto_fit_columns(ws7, min_width=10, max_width=45)
    ws7.column_dimensions['A'].width = 18
    ws7.column_dimensions['B'].width = 24
    ws7.column_dimensions['C'].width = 30
    ws7.column_dimensions['D'].width = 28
    ws7.column_dimensions['E'].width = 28
    ws7.column_dimensions['F'].width = 32
    ws7.column_dimensions['G'].width = 32
    ws7.freeze_panes = 'A5'

    # ==========================================
    # SHEET 8: ALUR TUJUAN PEMBELAJARAN (ATP)
    # ==========================================
    ws8 = wb.create_sheet(title="Alur Tujuan Belajar (ATP)")
    apply_title_banner(
        ws8,
        "MODEL ALUR TUJUAN PEMBELAJARAN (ATP) BAHASA INGGRIS SMA",
        "Rancangan Progresi Tujuan Pembelajaran (TP) Fase E (Kelas X) dan Fase F (Kelas XI - XII)",
        max_col=7
    )

    headers_atp = ["Fase & Kelas", "Kode TP", "Tujuan Pembelajaran (TP)", "Alokasi Waktu (JP)", "Lingkup Materi / Teks", "Aktivitas Pembelajaran Mendalam", "Asesmen Ketercapaian"]
    r = 4
    for i, h in enumerate(headers_atp, 1):
        cell = ws8.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_tbl_header
        cell.alignment = align_center
        cell.border = cell_border
    ws8.row_dimensions[4].height = 26
    r = 5

    atp_data = [
        ("SECTION", "ALUR TUJUAN PEMBELAJARAN FASE E (KELAS X - TOTAL 108 JP)", "", "", "", "", ""),
        ("Fase E (Kelas X)", "TP-E.1", "Mengidentifikasi fungsi sosial, struktur teks, dan unsur kebahasaan dalam teks interaksi transaksional lisan dan tulis yang melibatkan tindakan menyapa, memperkenalkan diri dan orang lain secara santun dan berkesadaran.", "12 JP", "Greetings & Self Introduction", "BKoF video interaktif, role-play simulasi pertemuan murid baru internasional, self-reflection log.", "Asesmen performa percakapan lisan berpasangan."),
        ("Fase E (Kelas X)", "TP-E.2", "Menganalisis ide pokok, informasi rinci, dan makna tersirat dalam teks deskriptif lisan dan tulis tentang tempat wisata dan bangunan bersejarah dengan penuh rasa ingin tahu.", "18 JP", "Descriptive Text (Tourism & Culture)", "Membedah brosur wisata digital, virtual tour, menganalisis fitur bahasa sensory visual.", "Tes pemahaman membaca & lembar kerja analisis teks."),
        ("Fase E (Kelas X)", "TP-E.3", "Menyusun teks deskriptif multimodal tentang potensi budaya atau destinasi lokal secara kreatif dengan struktur yang padu dan kosakata yang tepat.", "18 JP", "Descriptive Writing (Local Heritage)", "JCoT menulis deskripsi bersama, ICoT mendesain brosur/infografis Canva, pameran karya mini.", "Rubrik penilaian karya multimodal dan presentasi produk."),
        ("Fase E (Kelas X)", "TP-E.4", "Mengevaluasi kronologi peristiwa dan pesan moral dalam teks recount tentang pengalaman pribadi atau peristiwa bersejarah perjuangan bangsa.", "18 JP", "Recount Text (Personal & Historical)", "Menyimak podcast sejarah, memetakan timeline peristiwa masa lalu, analisis regular/irregular past verbs.", "Kuis komprehensif & penulisan jurnal pengalaman lampau."),
        ("Fase E (Kelas X)", "TP-E.5", "Mengidentifikasi langkah-langkah terstruktur dan kalimat perintah (imperatives) dalam teks prosedur serta mendemonstrasikannya secara tepat.", "15 JP", "Procedure Text (Safety & Practical Steps)", "Simulasi cooking show/manual lab, menganalisis instruksi keamanan siber media sosial.", "Demonstrasi prosedur lisan (spoken procedure presentation)."),
        ("Fase E (Kelas X)", "TP-E.6", "Merancang pesan fungsional singkat (announcement, advertisement, poster) secara multimodal untuk tujuan komunikasi di lingkungan sekolah.", "15 JP", "Short Functional Text (Multimodal)", "Eksplorasi kebutuhan mading sekolah, kolaborasi merancang poster kampanye kebersihan/kesehatan.", "Portofolio karya desain pesan visual digital."),
        ("Fase E (Kelas X)", "TP-E.7", "Mengungkapkan pendapat, setuju, dan tidak setuju serta memberikan alasan logis dalam diskusi kelompok kecil mengenai isu lingkungan hidup.", "12 JP", "Giving Opinion & Justification", "Debat mini di kelas dengan saling menghargai pendapat, menyusun peta argumen sederhana.", "Observasi keaktifan diskusi dan ketepatan argumen lisan."),

        ("SECTION", "ALUR TUJUAN PEMBELAJARAN FASE F WAJIB (KELAS XI & XII - TOTAL 108 JP / 96 JP)", "", "", "", "", ""),
        ("Fase F (Kelas XI)", "TP-F.1", "Menganalisis fungsi sosial, struktur tesis-argumen-penegasan, dan bahasa evaluatif dalam teks eksposisi analitis tentang isu lingkungan dan sosial.", "24 JP", "Analytical Exposition Text", "BKoF bedah artikel berita lingkungan, dekonstruksi teks model editorial, pemetaan argumen pendukung.", "Tes formatif membaca kritis & analisis teks editorial."),
        ("Fase F (Kelas XI)", "TP-F.2", "Menyusun esai eksposisi analitis tulis orisinal yang mempertahankan argumen berbasis data mengenai gaya hidup sehat remaja.", "24 JP", "Analytical Essay Writing", "JCoT menulis draf esai kelompok, ICoT menulis mandiri 300 kata, peer-editing berbasis rubrik.", "Rubrik penilaian esai argumentatif tertulis."),
        ("Fase F (Kelas XI)", "TP-F.3", "Mengevaluasi teks eksposisi hortatori dan menyajikan pidato persuasif yang merekomendasikan solusi atas permasalahan kedisiplinan belajar.", "24 JP", "Hortatory Exposition & Speech", "Menonton rekaman speech inspiratif (TED Talks), menyusun naskah pidato berstruktur, latihan retorika.", "Unjuk kerja rekaman pidato persuasif 3 menit."),
        ("Fase F (Kelas XI)", "TP-F.4", "Membandingkan dua sudut pandang yang bertentangan (pro dan kontra) dalam teks diskusi mengenai pemanfaatan teknologi digital.", "24 JP", "Discussion Text (Tech Controversy)", "Simulasi panel ahli, menyusun tabel perbandingan pro-kontra, menulis teks diskusi seimbang.", "Naskah teks diskusi berimbang dan refleksi imparsialitas."),
        ("Fase F (Kelas XII)", "TP-F.5", "Menganalisis konflik, resolusi, dan pesan moral dalam teks naratif sastra serta merefleksikannya dalam kehidupan sehari-hari.", "24 JP", "Narrative Text (Folktales & Fiction)", "Storytelling interaktif, membedah character arc dan koda moral, membandingkan fabel global & lokal.", "Tes apresiasi karya sastra & performa storytelling lisan."),
        ("Fase F (Kelas XII)", "TP-F.6", "Menyajikan presentasi multimodal profesional yang menggabungkan berbagai teks argumentasi untuk persiapan karir dan studi lanjut.", "24 JP", "Professional Multimodal Presentation", "Penyusunan portofolio digital resume dan personal statement bahasa Inggris, simulasi wawancara.", "Penilaian portofolio komprehensif kelulusan Fase F.")
    ]

    for item in atp_data:
        code = item[0]
        if code == "SECTION":
            ws8.cell(row=r, column=1, value="")
            ws8.cell(row=r, column=2, value=item[1]).font = font_group_header
            ws8.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
            for col in range(1, 8):
                ws8.cell(row=r, column=col).fill = fill_group
                ws8.cell(row=r, column=col).border = cell_border
            ws8.row_dimensions[r].height = 24
        else:
            fase, kd, tp, jp, mat, act, asm = item
            ws8.cell(row=r, column=1, value=fase).alignment = align_center
            ws8.cell(row=r, column=2, value=kd).alignment = align_center
            ws8.cell(row=r, column=3, value=tp).alignment = align_left
            ws8.cell(row=r, column=4, value=jp).alignment = align_center
            ws8.cell(row=r, column=5, value=mat).alignment = align_left
            ws8.cell(row=r, column=6, value=act).alignment = align_left
            ws8.cell(row=r, column=7, value=asm).alignment = align_left

            ws8.cell(row=r, column=1).font = font_data
            ws8.cell(row=r, column=2).font = font_data_bold
            ws8.cell(row=r, column=3).font = font_data
            ws8.cell(row=r, column=4).font = font_data_bold
            ws8.cell(row=r, column=5).font = font_data_bold
            ws8.cell(row=r, column=6).font = font_data
            ws8.cell(row=r, column=7).font = font_data
            
            for c in range(1, 8):
                ws8.cell(row=r, column=c).border = cell_border
                if r % 2 == 1:
                    ws8.cell(row=r, column=c).fill = fill_zebra
            ws8.row_dimensions[r].height = 42
        r += 1

    auto_fit_columns(ws8, min_width=10, max_width=45)
    ws8.column_dimensions['A'].width = 16
    ws8.column_dimensions['B'].width = 12
    ws8.column_dimensions['C'].width = 38
    ws8.column_dimensions['D'].width = 14
    ws8.column_dimensions['E'].width = 28
    ws8.column_dimensions['F'].width = 36
    ws8.column_dimensions['G'].width = 32
    ws8.freeze_panes = 'A5'

    # ==========================================
    # SHEET 9: ASESMEN OTENTIK & RUBRIK
    # ==========================================
    ws9 = wb.create_sheet(title="Asesmen Otentik & Rubrik")
    apply_title_banner(
        ws9,
        "PANDUAN ASESMEN OTENTIK & RUBRIK PENILAIAN BAHASA INGGRIS SMA",
        "Prinsip Asesmen Formatif, Sumatif, Rubrik Speaking, Writing, dan Portofolio Pembelajaran Mendalam",
        max_col=6
    )

    ws9.cell(row=4, column=1, value="1. RUBRIK PENILAIAN KETERAMPILAN BERBICARA (SPEAKING RUBRIC - FASE E & F)").font = font_section_header
    ws9.merge_cells("A4:F4")
    for c in range(1, 7):
        ws9.cell(row=4, column=c).fill = fill_tbl_header
    ws9.row_dimensions[4].height = 24

    headers_rubrik_spk = ["Kriteria Penilaian", "Sangat Mahir (Skor 4 / 86-100)", "Mahir (Skor 3 / 71-85)", "Cukup Mahir (Skor 2 / 56-70)", "Perlu Bimbingan (Skor 1 / <55)", "Bobot"]
    r = 5
    for i, h in enumerate(headers_rubrik_spk, 1):
        cell = ws9.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws9.row_dimensions[5].height = 24
    r += 1

    spk_rubric = [
        ("Kelancaran (Fluency)", "Berbicara lancar alami tanpa jeda abnormal; percakapan mengalir spontan dan responsif.", "Berbicara relatif lancar dengan sesekali jeda mencari kata; tidak mengganggu komunikasi.", "Sering berhenti dan ragu-ragu dalam menyusun kalimat; kelancaran terbatas.", "Sangat terbata-bata; kesulitan meneruskan kalimat tanpa bantuan pendidik.", "25%"),
        ("Pelafalan (Pronunciation)", "Pelafalan sangat jelas, intonasi tepat sesuai makna; aksen tidak menghambat komprehensi.", "Pelafalan jelas, sebagian kecil kata keliru ucap namun mudah dipahami.", "Sering salah pelafalan sehingga pendengar harus berkonsentrasi ekstra memahami arti.", "Banyak pelafalan keliru yang mengubah makna kata; intonasi datar/monoton.", "20%"),
        ("Kosakata (Vocabulary)", "Menggunakan kosakata luas, variatif, dan presisi; mampu parafrase saat buntu kata.", "Kosakata memadai untuk topik yang dibahas; pilihan kata umum tepat.", "Kosakata terbatas; sering mengulang kata dasar yang sama secara monoton.", "Sangat miskin kosakata; kesulitan menyampaikan gagasan dasar.", "25%"),
        ("Tata Bahasa (Grammar Accuracy)", "Menggunakan struktur kalimat sederhana dan kompleks secara akurat; minim kekeliruan minor.", "Struktur kalimat umumnya tepat; ada beberapa kekeliruan tata bahasa yang tidak merusak makna.", "Sering melakukan kesalahan tata bahasa dasar (tenses, subject-verb agreement).", "Kesalahan tata bahasa sangat dominan sehingga pesan percakapan sulit dipahami.", "15%"),
        ("Penyelesaian Tugas (Task Achievement)", "Menjawab pertanyaan tuntas, mempertahankan argumen kuat didukung bukti meyakinkan.", "Menjawab sebagian besar instruksi tugas; argumen cukup jelas.", "Menjawab sebagian kecil tugas; argumen kurang berkembang.", "Tidak menjawab instruksi tugas yang diberikan.", "15%")
    ]

    for item in spk_rubric:
        krit, sm, m, cm, pb, bob = item
        ws9.cell(row=r, column=1, value=krit).alignment = align_left
        ws9.cell(row=r, column=2, value=sm).alignment = align_left
        ws9.cell(row=r, column=3, value=m).alignment = align_left
        ws9.cell(row=r, column=4, value=cm).alignment = align_left
        ws9.cell(row=r, column=5, value=pb).alignment = align_left
        ws9.cell(row=r, column=6, value=bob).alignment = align_center

        ws9.cell(row=r, column=1).font = font_data_bold
        for c in range(2, 6):
            ws9.cell(row=r, column=c).font = font_data
        ws9.cell(row=r, column=6).font = font_data_bold
        
        for c in range(1, 7):
            ws9.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws9.cell(row=r, column=c).fill = fill_zebra
        ws9.row_dimensions[r].height = 40
        r += 1

    r += 1
    # Rubrik Penulisan (Writing Rubric)
    ws9.cell(row=r, column=1, value="2. RUBRIK PENILAIAN KETERAMPILAN MENULIS ESAI (WRITING RUBRIC - FASE E & F)").font = font_section_header
    ws9.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws9.cell(row=r, column=c).fill = fill_tbl_header
    ws9.row_dimensions[r].height = 24
    r += 1

    headers_rubrik_wrt = ["Kriteria Penulisan", "Sangat Mahir (Skor 4 / 86-100)", "Mahir (Skor 3 / 71-85)", "Cukup Mahir (Skor 2 / 56-70)", "Perlu Bimbingan (Skor 1 / <55)", "Bobot"]
    for i, h in enumerate(headers_rubrik_wrt, 1):
        cell = ws9.cell(row=r, column=i, value=h)
        cell.font = font_tbl_header
        cell.fill = fill_card_header
        cell.alignment = align_center
        cell.border = cell_border
    ws9.row_dimensions[r].height = 24
    r += 1

    wrt_rubric = [
        ("Isi & Gagasan (Content & Ideas)", "Tesis jelas, argumen sangat kuat, didukung bukti data dan fakta empiris relevan dan mendalam.", "Tesis jelas, argumen logis didukung sebagian contoh relevan.", "Tesis kurang jelas; argumen masih dangkal dan minim elaborasi pendukung.", "Tidak memiliki fokus tesis; gagasan membingungkan dan tidak berdasar.", "30%"),
        ("Organisasi & Kohesi (Organization)", "Struktur teks (genre) sempurna, alur antarparagraf sangat padu (cohesive) dengan transisi mulus.", "Struktur teks lengkap, penggunaan kata hubung (connectors) tepat pada umumnya.", "Struktur teks kurang lengkap; perpindahan antarparagraf melompat-lompat.", "Tidak memiliki struktur organisasi yang jelas; tulisan tidak teratur.", "25%"),
        ("Penggunaan Kosakata (Vocabulary)", "Pilihan kata cermat, kaya, menggunakan istilah akademik (AWL) dan kolokasi alami.", "Pilihan kata tepat dan bervariasi; sedikit kekeliruan diksi minor.", "Pilihan kata sangat mendasar dan terbatas; sering mengulang kata yang sama.", "Banyak kata keliru konteks atau terjemahan harfiah bahasa Indonesia.", "20%"),
        ("Tata Bahasa (Grammar & Syntax)", "Menggunakan ragam kalimat kompleks secara akurat; kendali tenses dan klausa sempurna.", "Struktur kalimat efektif; ada beberapa kekeliruan tata bahasa yang tidak mengaburkan pesan.", "Sering salah tenses, bentuk jamak, dan preposisi yang mengganggu alur baca.", "Kesalahan tata bahasa sangat parah di hampir setiap kalimat.", "15%"),
        ("Mekanik (Spelling & Punctuation)", "Ejaan kata sempurna, penggunaan tanda baca dan huruf kapital tepat tanpa cacat.", "Ejaan dan tanda baca umumnya benar; hanya ada 1-2 kekeliruan sepele.", "Beberapa kesalahan ejaan dan tanda baca yang cukup mengganggu pembaca.", "Banyak kesalahan ejaan fatal dan tanda baca diabaikan.", "10%")
    ]

    for item in wrt_rubric:
        krit, sm, m, cm, pb, bob = item
        ws9.cell(row=r, column=1, value=krit).alignment = align_left
        ws9.cell(row=r, column=2, value=sm).alignment = align_left
        ws9.cell(row=r, column=3, value=m).alignment = align_left
        ws9.cell(row=r, column=4, value=cm).alignment = align_left
        ws9.cell(row=r, column=5, value=pb).alignment = align_left
        ws9.cell(row=r, column=6, value=bob).alignment = align_center

        ws9.cell(row=r, column=1).font = font_data_bold
        for c in range(2, 6):
            ws9.cell(row=r, column=c).font = font_data
        ws9.cell(row=r, column=6).font = font_data_bold
        
        for c in range(1, 7):
            ws9.cell(row=r, column=c).border = cell_border
            if r % 2 == 1:
                ws9.cell(row=r, column=c).fill = fill_zebra
        ws9.row_dimensions[r].height = 40
        r += 1

    r += 1
    # Rekomendasi Portofolio Pembelajaran Bahasa Inggris
    ws9.cell(row=r, column=1, value="3. PRINSIP ASESMEN FORMATIF BERKELANJUTAN & PORTOFOLIO MURID").font = font_section_header
    ws9.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    for c in range(1, 7):
        ws9.cell(row=r, column=c).fill = fill_tbl_header
    ws9.row_dimensions[r].height = 24
    r += 1

    port_info = [
        ("Penilaian Formatif Berbasis Proses", "Asesmen tidak hanya menilai hasil akhir (Assessment OF Learning), melainkan terutama untuk memandu proses belajar (Assessment FOR & AS Learning). Guru memberikan umpan balik deskriptif (descriptive feedback) segera setelah tahap Modelling atau Joint Construction."),
        ("Dokumentasi Portofolio Berkesadaran", "Siswa menyimpan seluruh draf tulisan (draf awal, draf hasil revisi peer-review, draf final) dalam Portofolio Bahasa Inggris Digital/Fisik. Siswa menuliskan refleksi metakognisi tentang peningkatan kosakata dan kemahiran mereka dari waktu ke waktu."),
        ("Asesmen Otentik Menggembirakan", "Penilaian sumatif diarahkan pada unjuk kerja di dunia nyata (authentic performance assessment), misalnya pameran karya poster digital di media sosial sekolah, penampilan drama panggung komedi bahasa Inggris, atau penulisan surat pembaca untuk majalah lokal.")
    ]
    for p_title, p_desc in port_info:
        ws9.cell(row=r, column=1, value=p_title).font = font_data_bold
        ws9.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ws9.cell(row=r, column=1).alignment = align_left
        ws9.cell(row=r, column=1).fill = fill_group

        ws9.cell(row=r, column=3, value=p_desc).font = font_data
        ws9.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
        ws9.cell(row=r, column=3).alignment = align_left

        for c in range(1, 7):
            ws9.cell(row=r, column=c).border = cell_border
            if c >= 3 and r % 2 == 1:
                ws9.cell(row=r, column=c).fill = fill_zebra
        ws9.row_dimensions[r].height = 34
        r += 1

    auto_fit_columns(ws9, min_width=10, max_width=45)
    ws9.column_dimensions['A'].width = 24
    ws9.column_dimensions['B'].width = 30
    ws9.column_dimensions['C'].width = 30
    ws9.column_dimensions['D'].width = 30
    ws9.column_dimensions['E'].width = 30
    ws9.column_dimensions['F'].width = 12
    ws9.freeze_panes = 'A5'

    # Save workbook
    wb.save(output_path)
    print(f"English SMA Workbook successfully created and saved to: {output_path}")

if __name__ == "__main__":
    out_file = r"C:\Users\KOMPUTERKU2\Documents\English SMA\Peraturan\Capaian_Pembelajaran_dan_Kurikulum_Bahasa_Inggris_SMA.xlsx"
    create_english_sma_excel(out_file)
