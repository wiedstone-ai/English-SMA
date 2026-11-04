# -*- coding: utf-8 -*-
"""
Generator script for Modul Ajar (RPP Kurikulum Merdeka) in 5-Modul Ajar XII SMA TL.
Covers TP 12.TL.1 through TP 12.TL.8 with complete pedagogical components:
Informasi Umum, Tujuan Pembelajaran, Capaian & IKTP, Pemahaman Bermakna, Pertanyaan Pemantik,
Langkah-langkah Pembelajaran Rinci (GBA: BKoF, MoT, JCoT, ICoT),
Asesmen, Pengayaan, Remedial, dan Refleksi Guru & Siswa.
"""

import os
import sys
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import docx_helper

TARGET_DIR = os.path.dirname(os.path.abspath(__file__))
os.makedirs(TARGET_DIR, exist_ok=True)

COLOR_NAVY = RGBColor(15, 41, 66)
COLOR_ROYAL = RGBColor(30, 64, 175)
COLOR_BODY = RGBColor(30, 41, 59)

def add_section_header(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(title)
    run.font.name = "Segoe UI"
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = COLOR_NAVY

def add_sub_header(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(title)
    run.font.name = "Segoe UI"
    run.font.size = Pt(10.5)
    run.font.bold = True
    run.font.color.rgb = COLOR_ROYAL

def build_meeting_scenarios(doc, scenarios):
    for meeting_title, phases in scenarios:
        add_sub_header(doc, meeting_title)
        for phase_name, steps in phases:
            p_ph = doc.add_paragraph()
            p_ph.paragraph_format.space_before = Pt(4)
            p_ph.paragraph_format.space_after = Pt(2)
            r_ph = p_ph.add_run(phase_name)
            r_ph.font.name = "Segoe UI"
            r_ph.font.size = Pt(9.5)
            r_ph.font.bold = True
            r_ph.font.color.rgb = COLOR_ROYAL

            for step in steps:
                p_st = doc.add_paragraph()
                p_st.paragraph_format.left_indent = Inches(0.2)
                p_st.paragraph_format.space_before = Pt(0)
                p_st.paragraph_format.space_after = Pt(2)
                r_bullet = p_st.add_run("• ")
                r_bullet.font.bold = True
                r_txt = p_st.add_run(step)
                r_txt.font.name = "Segoe UI"
                r_txt.font.size = Pt(9)

def add_common_sections(doc, as_rows, enrich_text, remedi_text, ref_rows):
    add_section_header(doc, "III. RENCANA ASESMEN PEMBELAJARAN (ASSESSMENT PLAN)")
    as_headers = ["Jenis Asesmen", "Teknik Penilaian", "Bentuk Instrumen", "Waktu Pelaksanaan"]
    docx_helper.add_styled_table(doc, as_headers, as_rows, [Inches(1.5), Inches(1.5), Inches(2.3), Inches(1.2)])

    add_section_header(doc, "IV. PROGRAM PENGAYAAN DAN REMEDIAL")
    docx_helper.add_callout(
        doc,
        "Program Pengayaan (Bagi Siswa yang Telah Mencapai Ketuntasan):",
        enrich_text,
        bg="DCFCE7", border_color="16A34A"
    )
    docx_helper.add_callout(
        doc,
        "Program Remedial (Bagi Siswa yang Membutuhkan Bimbingan Khusus):",
        remedi_text,
        bg="FEF3C7", border_color="D97706"
    )

    add_section_header(doc, "V. REFLEKSI GURU DAN PESERTA DIDIK")
    ref_headers = ["Aspek Refleksi", "Pertanyaan Refleksi Diri Siswa", "Pertanyaan Refleksi Guru"]
    docx_helper.add_styled_table(doc, ref_headers, ref_rows, [Inches(1.5), Inches(2.5), Inches(2.5)])


# ==============================================================================
# MODUL 1: TP 12.TL.1 - GLOBAL BIOETHICS DISCUSSION
# ==============================================================================
def create_modul_1():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "MODUL AJAR KURIKULUM MERDEKA - BAHASA INGGRIS TINGKAT LANJUT KELAS XII",
        "Modul 1: Global Academic Discussion on Bioethics & Frontier Technologies | TP 12.TL.1"
    )
    info = {
        "Satuan Pendidikan": "Sekolah Menengah Atas (SMA)",
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XII / Semester 5 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 12.TL.1 - Menganalisis wacana debat akademik tingkat tinggi (Advanced Discussion) mengenai bioetika medis, rekayasa genetika, eksplorasi antariksa, dan merumuskan sikap etis ilmiah secara komprehensif.",
        "Alokasi Waktu": "25 Jam Pelajaran (5 Minggu x 5 JP @ 45 menit)",
        "Target Peserta Didik": "Peserta Didik Peminatan / Tingkat Lanjut (32 – 36 Siswa)",
        "Model Pembelajaran": "Genre-Based Approach (GBA) & Problem-Based Learning (PBL)",
        "Profil Pelajar Pancasila": "Bernalar Kritis, Beriman & Bertakwa, Berwawasan Global",
        "Sarana & Prasarana": "LCD Proyektor, Laptop, Teks Otentik (Nature, The Lancet), Video Dokumenter CRISPR"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_section_header(doc, "I. KOMPONEN INTI (CORE COMPONENTS)")
    add_sub_header(doc, "A. Capaian Pembelajaran & Indikator Ketercapaian (IKTP)")
    p_cp = doc.add_paragraph()
    p_cp.paragraph_format.space_after = Pt(4)
    p_cp.add_run("Elemen: Menyimak - Berbicara & Membaca - Memirsa\n").bold = True
    p_cp.add_run("Pada akhir Fase F Tingkat Lanjut, peserta didik mengevaluasi wacana akademik kompleks multi-sudut pandang dan memproduksi teks lisan/tulis tingkat mahir untuk mengadvokasi sikap etis berbasis data ilmiah.")

    iktp_headers = ["Kode IKTP", "Indikator Ketercapaian Tujuan Pembelajaran", "Ranah Kognitif"]
    iktp_rows = [
        ["IKTP 1.1", "Mengidentifikasi argumen filosofis dan implikasi moral dari rekayasa genetika CRISPR dan bioetika medis.", "C4 (Analisis)"],
        ["IKTP 1.2", "Menerapkan penanda sikap akademik (stance markers) dan klausa konsesif dalam wacana sains frontier.", "C3 (Aplikasi)"],
        ["IKTP 1.3", "Mengevaluasi objektivitas dan risiko ketimpangan biologis akibat komersialisasi teknologi mutakhir.", "C5 (Evaluasi)"],
        ["IKTP 1.4", "Menyusun Position Paper bioetika 500 kata dan memimpin dialog meja bundar kebijakan ilmiah.", "C6 (Kreasi)"]
    ]
    docx_helper.add_styled_table(doc, iktp_headers, iktp_rows, [Inches(1.2), Inches(4.3), Inches(1.0)])

    add_sub_header(doc, "B. Pemahaman Bermakna (Meaningful Understanding)")
    docx_helper.add_callout(
        doc,
        "Pemahaman Bermakna (Big Ideas):",
        "1. Kemajuan bioteknologi frontier (CRISPR germline modification, AI clinical triage) menawarkan peluang penyembuhan penyakit heriditer, namun berisiko memicu jurang ketimpangan genetik dan dilema eugenika baru bila tanpa regulasi etis global.\n"
        "2. Wacana akademik tingkat tinggi menuntut kemampuan menimbang bukti ilmiah dengan integritas moral, menggunakan bahasa berpagar (hedging), serta menghargai martabat martabat kemanusiaan.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_sub_header(doc, "C. Pertanyaan Pemantik (Essential Questions)")
    p_pq = doc.add_paragraph()
    p_pq.paragraph_format.left_indent = Inches(0.2)
    p_pq.add_run("1. Should humanity place an absolute moratorium on modifying the human germline genome, even if it cures fatal diseases?\n")
    p_pq.add_run("2. Can an artificial intelligence algorithm truly make compassionate end-of-life medical decisions without violating human dignity?\n")
    p_pq.add_run("3. How do stance markers and concessive clauses transform a polarizing debate into a scholarly, balanced dialogue?")

    add_section_header(doc, "II. LANGKAH-LANGKAH PEMBELAJARAN (5 MINGGU @ 5 JP = 225 MENIT/MINGGU)")

    scenarios = [
        ("MINGGU 1 (5 JP): Building Knowledge of the Field (BKoF) - Bioetika & CRISPR", [
            ("Pendahuluan (20 Menit)", [
                "Guru membuka pembelajaran dengan salam, doa, dan presensi.",
                "Apersepsi: Menayangkan video liputan eksperimen pengeditan gen embrio manusia pertama di dunia.",
                "Penyampaian tujuan TP 12.TL.1 dan asesmen meja bundar kebijakan bioetika."
            ]),
            ("Kegiatan Inti (190 Menit - BKoF)", [
                "Eksplorasi Konseptual: Siswa mendiskusikan perbedaan esensial antara somatik (terapi individu) vs germline (warisan lintas generasi).",
                "Bedah Kasus Bioetika: Siswa menganalisis studi kasus 'The He Jiankui Affair' dan implikasinya pada konsensus sains internasional.",
                "Pemerolehan Kosakata Akademik: Membahas 8 istilah kunci (pathology, beneficence, stratification, moratorium, egalitarian, post-scarcity, dll).",
                "Speed Debating: Sesi latihan lisan singkat 1 menit per siswa: 'Apakah orang tua berhak mendesain genetik anak mereka?'"
            ]),
            ("Penutup (15 Menit)", [
                "Refleksi 1 menit: Siswa mencatat dilema moral terbesar yang mereka pelajari hari ini.",
                "Penugasan membaca mandiri teks otentik 1 (Nature Biotechnology) untuk pertemuan berikutnya."
            ])
        ]),
        ("MINGGU 2 (5 JP): Modeling of the Text (MoT) - Bedah Teks Diskusi Akademik", [
            ("Pendahuluan (15 Menit)", [
                "Review istilah akademik melalui kuis tebak kata interaktif di LCD.",
                "Orientasi target: Membedah struktur dan fitur bahasa teks model dari Nature Biotechnology."
            ]),
            ("Kegiatan Inti (195 Menit - MoT)", [
                "Close Reading: Siswa membaca intensif Teks Otentik 1: 'CRISPR and Human Germline Modification'.",
                "Dekonstruksi Struktur Teks: Memetakan Issue, Arguments For (Utilitarian beneficence), Arguments Against (Egalitarian rights, eugenics), dan Balanced Recommendation.",
                "Eksplorasi Bahasa: Mengidentifikasi stance markers ('proponents passionately articulate', 'critics caution', 'arguably') dan klausa konsesif ('Notwithstanding', 'While somatic therapies...').",
                "Latihan Tata Bahasa: Siswa mengubah kalimat argumentatif subjektif menjadi kalimat akademik berbobot dengan hedging."
            ]),
            ("Penutup (15 Menit)", [
                "Siswa merangkum bagan struktur teks diskusi akademik di lembar kerja portofolio.",
                "Guru memberikan umpan balik dan membagi kelompok meja bundar untuk Minggu 3."
            ])
        ]),
        ("MINGGU 3 (5 JP): Joint Construction of the Text (JCoT) - Analisis AI Triage & Ruang Angkasa", [
            ("Pendahuluan (15 Menit)", [
                "Guru mengondisikan 4 kelompok kerja meja bundar (Roundtable Task Forces).",
                "Penjelasan panduan LKPD Aktivitas 3: Matriks Analisis Multi-Sudut Pandang."
            ]),
            ("Kegiatan Inti (195 Menit - JCoT)", [
                "Studi Dokumen 2 & 3: Kelompok 1 & 2 menganalisis teks 'Algorithmic Triage' (The Lancet), kelompok 3 & 4 menganalisis 'Outer Space Colonization' (Harvard ILR).",
                "Penyusunan Matriks Kebijakan: Masing-masing kelompok mengisi lembar telaah pro-kontra dengan bukti data tekstual.",
                "Simulasi Meja Bundar Terbimbing: Dua perwakilan kelompok memimpin dialog meja bundar di depan kelas, saling menanggapi dengan etika kesantunan akademik.",
                "Peer Assessment: Kelompok audiens menilai kohesi argumen dan penggunaan stance markers menggunakan lembar rubrik."
            ]),
            ("Penutup (15 Menit)", [
                "Guru menyimpulkan sintesis dialog dan memberikan koreksi pengucapan serta artikulasi bahasa lisan.",
                "Pengumuman tema penulisan Position Paper mandiri untuk Minggu 4."
            ])
        ]),
        ("MINGGU 4 (5 JP): Independent Construction of the Text (ICoT) - Drafting Position Paper", [
            ("Pendahuluan (15 Menit)", [
                "Pemberian instruksi penulisan Position Paper akademik (400-500 kata).",
                "Review rubrik penilaian menulis analitik Kurikulum Merdeka (Content, Organization, Vocabulary, Language Use)."
            ]),
            ("Kegiatan Inti (195 Menit - ICoT)", [
                "Outlining: Siswa menyusun kerangka tesis pribadi dan memetakan 2 argumen tandingan yang akan disanggah.",
                "Drafting Mandiri: Siswa menulis draf lengkap esai posisi bioetika secara hening dan fokus di kelas.",
                "Guru melaksanakan konferensi meja individual (writing conference) memberikan scaffolding bagi siswa yang memerlukan dukungan tata bahasa.",
                "Self-Checklist: Siswa memeriksa keberadaan stance markers dan klausa konsesif menggunakan panduan ceklis mandiri."
            ]),
            ("Penutup (15 Menit)", [
                "Siswa mengunggah draf ke folder portofolio kelas dan menetapkan pasangan peer-review."
            ])
        ]),
        ("MINGGU 5 (5 JP): Peer-Review, Finalisasi & Simulasi Meja Bundar Resmi", [
            ("Pendahuluan (15 Menit)", [
                "Penjelasan rubrik penyuntingan antarteman (Peer-Review Protocol).",
                "Setting tata ruang kelas menjadi Konferensi Meja Bundar Bioetika Internasional."
            ]),
            ("Kegiatan Inti (195 Menit)", [
                "Peer-Editing Session (45 Menit): Siswa saling menyunting draf teman dengan fokus pada kekokohan bukti dan koherensi transisi.",
                "Revisi Naskah Akhir (45 Menit): Finalisasi naskah Position Paper 500 kata siap terbit.",
                "Simulasi Meja Bundar Resmi (90 Menit): Pelaksanaan Sidang Meja Bundar Etika Sains dengan moderator bergantian, setiap siswa menyampaikan pidato posisi 2-3 menit.",
                "Penilaian Asesmen Sumatif Unjuk Kerja oleh guru."
            ]),
            ("Penutup (15 Menit)", [
                "Refleksi pencapaian kompetensi TP 12.TL.1.",
                "Apresiasi guru atas kedalaman pemikiran etis peserta didik, diakhiri doa bersama."
            ])
        ])
    ]
    build_meeting_scenarios(doc, scenarios)

    as_rows = [
        ["1. Diagnostik", "Angket & Tes Prasyarat", "Angket Minat Bioetika & Kuis Struktur Teks Diskusi", "Awal Minggu 1"],
        ["2. Formatif", "Observasi & Portofolio", "Lembar Pengamatan Meja Bundar & Ceklis Draf Esai", "Minggu 2, 3, dan 4"],
        ["3. Sumatif", "Kinerja & Tes Tulis", "Position Paper 500 Kata & Unjuk Kerja Meja Bundar Resmi", "Akhir Minggu 5"]
    ]
    enrich_text = "Siswa menyusun policy brief 2 halaman ditujukan kepada Komite Bioetika Nasional mengenai rancangan undang-undang pengawasan terapi genetik CRISPR di Indonesia."
    remedi_text = "Bimbingan terfokus menyusun argumen berimbang menggunakan template pemandu (concessive clause sentence starters) dan latihan identifikasi pro-kontra terarah."
    ref_rows = [
        ["Pemahaman Etis & Konsep", "Apakah saya mampu membedakan isu terapi genetik somatik dan modifikasi germline?", "Apakah siswa berhasil menginternalisasi dilema moral sains frontier?"],
        ["Keterampilan Bahasa", "Apakah saya sudah mahir menggunakan stance markers dan klausa konsesif?", "Bagian mana dari tata bahasa akademik yang masih membutuhkan penguatan?"],
        ["Kolaborasi & Sikap", "Bagaimana saya mempertahankan sikap santun dan objektif dalam perdebatan?", "Apakah seluruh siswa merasa aman mengemukakan argumen kritisnya?"]
    ]
    add_common_sections(doc, as_rows, enrich_text, remedi_text, ref_rows)

    fpath = os.path.join(TARGET_DIR, "Modul_Ajar_TP_12.TL.1_Global_Bioethics_Discussion.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# MODUL 2: TP 12.TL.2 - ACADEMIC SYNTHESIS ESSAY
# ==============================================================================
def create_modul_2():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "MODUL AJAR KURIKULUM MERDEKA - BAHASA INGGRIS TINGKAT LANJUT KELAS XII",
        "Modul 2: Academic Synthesis Essay & Research Integration | TP 12.TL.2"
    )
    info = {
        "Satuan Pendidikan": "Sekolah Menengah Atas (SMA)",
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XII / Semester 5 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 12.TL.2 - Menyusun esai penelitian akademik (Academic Synthesis Essay) 1.000 kata dengan memadukan multi-referensi, kutipan langsung/tidak langsung, dan daftar pustaka standar internasional (APA/MLA).",
        "Alokasi Waktu": "25 Jam Pelajaran (5 Minggu x 5 JP @ 45 menit)",
        "Target Peserta Didik": "Peserta Didik Peminatan / Tingkat Lanjut (32 – 36 Siswa)",
        "Model Pembelajaran": "Genre-Based Approach (GBA) & Project-Based Learning (PjBL)",
        "Profil Pelajar Pancasila": "Mandiri, Bernalar Kritis, Berintegritas Akademik",
        "Sarana & Prasarana": "Laptop, Akses Database Jurnal Ilmiah (Google Scholar, ScienceDirect), Manual APA Style 7th"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_section_header(doc, "I. KOMPONEN INTI (CORE COMPONENTS)")
    add_sub_header(doc, "A. Capaian Pembelajaran & Indikator Ketercapaian (IKTP)")
    p_cp = doc.add_paragraph()
    p_cp.paragraph_format.space_after = Pt(4)
    p_cp.add_run("Elemen: Membaca - Memirsa & Menulis - Mempresentasikan\n").bold = True
    p_cp.add_run("Peserta didik menyintesis temuan dari beragam artikel jurnal penelitian internasional, menerapkan konvensi sitasi akademik yang ketat, dan menghasilkan esai riset akademik yang koheren, berbobot, dan orisinal.")

    iktp_headers = ["Kode IKTP", "Indikator Ketercapaian Tujuan Pembelajaran", "Ranah Kognitif"]
    iktp_rows = [
        ["IKTP 2.1", "Merumuskan working thesis statement dan menyintesis minimal 3 sumber jurnal ilmiah terakreditasi.", "C4 (Analisis)"],
        ["IKTP 2.2", "Menerapkan teknik sitasi dalam teks (in-text citation) gaya APA/MLA dan teknik parafrasa bebas plagiarisme.", "C3 (Aplikasi)"],
        ["IKTP 2.3", "Mengevaluasi keselarasan bukti literatur dan mengidentifikasi celah riset (research gaps) yang belum terjawab.", "C5 (Evaluasi)"],
        ["IKTP 2.4", "Menyusun makalah sintesis akademik 1.000 kata lengkap dengan daftar pustaka yang akurat dan terstandar.", "C6 (Kreasi)"]
    ]
    docx_helper.add_styled_table(doc, iktp_headers, iktp_rows, [Inches(1.2), Inches(4.3), Inches(1.0)])

    add_sub_header(doc, "B. Pemahaman Bermakna (Meaningful Understanding)")
    docx_helper.add_callout(
        doc,
        "Pemahaman Bermakna (Big Ideas):",
        "1. Menulis esai akademik bukanlah sekadar merangkum artikel satu per satu secara terpisah, melainkan menyatukan berbagai suara ilmuwan ke dalam sebuah percakapan ilmiah yang menyatu untuk membuktikan sebuah tesis orisinal.\n"
        "2. Integritas akademik melalui sitasi yang presisi (APA 7th Edition) dan teknik parafrasa yang etis merupakan fondasi utama kesiapan studi lanjut di perguruan tinggi bereputasi global.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_sub_header(doc, "C. Pertanyaan Pemantik (Essential Questions)")
    p_pq = doc.add_paragraph()
    p_pq.paragraph_format.left_indent = Inches(0.2)
    p_pq.add_run("1. Why is literature synthesis considered fundamentally different from a simple summary or book report?\n")
    p_pq.add_run("2. How does proper in-text citation protect an author's credibility while honoring intellectual property rights?\n")
    p_pq.add_run("3. What strategies prevent accidental plagiarism when incorporating complex scientific studies into your own writing?")

    add_section_header(doc, "II. LANGKAH-LANGKAH PEMBELAJARAN (5 MINGGU @ 5 JP = 225 MENIT/MINGGU)")

    scenarios = [
        ("MINGGU 1 (5 JP): BKoF - Penelusuran Jurnal Ilmiah & Plagiarisme", [
            ("Pendahuluan (20 Menit)", [
                "Salam, doa bersama, dan apersepsi seputar tantangan menulis makalah di perguruan tinggi.",
                "Penyampaian tujuan TP 12.TL.2 dan produk akhir makalah sintesis 1.000 kata."
            ]),
            ("Kegiatan Inti (190 Menit - BKoF)", [
                "Workshop Penelusuran Literatur: Siswa berlatih menggunakan boolean operators (AND, OR, NOT) di Google Scholar untuk mencari artikel penelitian peer-reviewed.",
                "Bedah Etika Akademik: Menelaah jenis-jenis plagiarisme (patchwriting, direct copy, self-plagiarism) dan batas ambang kemiripan (similarity threshold).",
                "Pengenalan Format APA 7th Edition: Membahas aturan sitasi satu penulis, dua penulis, et al., serta daftar pustaka jurnal digital (DOI).",
                "Latihan Parafrasa: Siswa mengubah penggalan teks jurnal teknis menjadi kalimat baru dengan makna yang sama persis tanpa menjiplak struktur asli."
            ]),
            ("Penutup (15 Menit)", [
                "Siswa mengumpulkan 3 tautan artikel jurnal terakreditasi sesuai topik riset pilihannya.",
                "Refleksi singkat dan doa penutup."
            ])
        ]),
        ("MINGGU 2 (5 JP): MoT - Bedah Makalah Model & Sintesis Tematik", [
            ("Pendahuluan (15 Menit)", [
                "Review kaidah sitasi APA melalui kuis identifikasi kesalahan format sitasi.",
                "Orientasi kegiatan: Membedah teks model sintesis dari Environmental Science & Technology."
            ]),
            ("Kegiatan Inti (195 Menit - MoT)", [
                "Membaca Teks Model 1: Menganalisis Teks 'Microplastic Infiltration in Terrestrial Ecosystems and Human Tissues'.",
                "Dekonstruksi Sintesis Tematik: Mengamati bagaimana penulis memadukan temuan Geyer (2017), Thompson (2022), Leslie (2022), dan Prata (2023) dalam alur paragraf yang mengalir.",
                "Analisis Reporting Verbs: Mengkaji nuansa makna kata kerja pelapor ('established', 'hypothesized', 'revealed', 'demonstrated').",
                "Latihan Matriks Sintesis: Siswa menyusun Literature Synthesis Matrix 3 kolom (Penulis & Tahun, Temuan Kunci, Analisis Sintesis/Hubungan dengan Topik)."
            ]),
            ("Penutup (15 Menit)", [
                "Siswa menyelesaikan lembar matriks sintesis untuk 3 sumber pilihan mereka.",
                "Pengarahan untuk perumusan tesis pada Minggu 3."
            ])
        ]),
        ("MINGGU 3 (5 JP): JCoT - Perumusan Working Thesis & Penyusunan Outline", [
            ("Pendahuluan (15 Menit)", [
                "Diskusi kelas terbuka: Apa yang membedakan tesis yang kuat (debatable & specific) dengan pernyataan fakta umum?",
                "Penjelasan panduan LKPD Aktivitas 3."
            ]),
            ("Kegiatan Inti (195 Menit - JCoT)", [
                "Praktik Kolaboratif Merumuskan Tesis: Berpasangan saling menguji kekuatan tesis esai teman.",
                "Penyusunan Kerangka 5 Paragraf: Introduction (Hook, Literature Background, Thesis), Body 1 (Theme A Synthesis), Body 2 (Theme B Synthesis), Body 3 (Counter-evidence & Research Gaps), Conclusion (Synthesis Restatement, Academic Significance).",
                "Bedah Draf Bersama (Group Modelling): Guru menampilkan contoh paragraf sintesis di proyektor dan kelas bersama-sama menyempurnakan kohesi transisinya.",
                "Penyusunan Draf Daftar Pustaka (References List) sesuai standar APA 7th Edition."
            ]),
            ("Penutup (15 Menit)", [
                "Review kesiapan outline oleh guru.",
                "Pengingat membawa laptop untuk sesi drafting intensif di Minggu 4."
            ])
        ]),
        ("MINGGU 4 (5 JP): ICoT - Penulisan Draf Mandiri 1.000 Kata", [
            ("Pendahuluan (15 Menit)", [
                "Instruksi penulisan esai mandiri, target minimal 800 kata pada akhir pertemuan.",
                "Pemberian rubrik esai akademik internasional."
            ]),
            ("Kegiatan Inti (195 Menit - ICoT)", [
                "Writing Marathon Mandiri: Siswa menulis draf lengkap esai akademik dengan mengintegrasikan sitasi langsung dan tidak langsung.",
                "Scaffolding Guru: Guru berkeliling memberikan umpan balik langsung pada penggunaan penanda transisi (Concurrently, In contrast, Consequently).",
                "Self-Checklist Plagiarisme: Siswa memeriksa apakah seluruh gagasan pihak ketiga telah disertai atribusi kutipan yang sah."
            ]),
            ("Penutup (15 Menit)", [
                "Siswa mengunggah draf ke sistem manajemen pembelajaran kelas.",
                "Penetapan pembagian peer-review untuk Minggu 5."
            ])
        ]),
        ("MINGGU 5 (5 JP): Peer-Review, Finalisasi & Presentasi Sidang Mini", [
            ("Pendahuluan (15 Menit)", [
                "Penjelasan pedoman penilaian antarteman (Academic Peer-Review Protocol).",
                "Fokus review: Kekuatan sintesis, kepatuhan APA 7th, dan alur penalaran logis."
            ]),
            ("Kegiatan Inti (195 Menit)", [
                "Sesi Peer-Review Kemitraan (45 Menit): Membaca draf rekan dan memberikan catatan perbaikan substantive.",
                "Revisi Naskah Akhir (60 Menit): Finalisasi makalah akademik 1.000 kata lengkap dengan daftar pustaka yang sempurna.",
                "Mini Academic Presentation (75 Menit): Setiap siswa mempresentasikan tesis dan temuan sintesisnya dalam waktu 2 menit di hadapan kelompok meja.",
                "Penilaian Sumatif Portofolio Makalah oleh guru."
            ]),
            ("Penutup (15 Menit)", [
                "Refleksi penguasaan literasi akademik tingkat lanjut.",
                "Apresiasi guru dan doa penutup."
            ])
        ])
    ]
    build_meeting_scenarios(doc, scenarios)

    as_rows = [
        ["1. Diagnostik", "Kuis Prasyarat", "Uji Pemahaman Sitasi APA & Deteksi Plagiarisme", "Awal Minggu 1"],
        ["2. Formatif", "Penugasan & Matriks", "Matriks Sintesis Literatur & Lembar Outline Esai", "Minggu 2, 3, dan 4"],
        ["3. Sumatif", "Portofolio Riset", "Makalah Sintesis Akademik 1.000 Kata & Daftar Pustaka APA", "Akhir Minggu 5"]
    ]
    enrich_text = "Siswa mengirimkan makalah sintesis terbaik mereka ke kompetisi karya tulis ilmiah siswa nasional atau mempublikasikannya di repositori sekolah ber-ISSN."
    remedi_text = "Bimbingan terstruktur mengubah draf ringkasan bertingkat menjadi matriks sintesis dua sumber dengan template kalimat pemandu (sentence stems)."
    ref_rows = [
        ["Pemahaman Konvensi Sitasi", "Apakah saya sudah konsisten menerapkan format kutipan dalam teks APA 7th Edition?", "Apakah siswa memahami batasan tegas antara parafrasa dan plagiarisme?"],
        ["Kemampuan Sintesis", "Apakah esai saya berhasil mempertemukan minimal 3 artikel ilmiah secara padu?", "Berapa persentase siswa yang berhasil melampaui level ringkasan biasa?"],
        ["Kemandirian Akademik", "Sejauh mana saya mampu menuntaskan makalah 1.000 kata tepat waktu?", "Bagaimana ketahanan belajar dan disiplin riset siswa selama proses penulisan?"]
    ]
    add_common_sections(doc, as_rows, enrich_text, remedi_text, ref_rows)

    fpath = os.path.join(TARGET_DIR, "Modul_Ajar_TP_12.TL.2_Academic_Synthesis_Essay.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# MODUL 3: TP 12.TL.3 - DRAMATIC SCRIPT & PERFORMANCE
# ==============================================================================
def create_modul_3():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "MODUL AJAR KURIKULUM MERDEKA - BAHASA INGGRIS TINGKAT LANJUT KELAS XII",
        "Modul 3: Dramatic Literature, Theatrical Scripts & Performance | TP 12.TL.3"
    )
    info = {
        "Satuan Pendidikan": "Sekolah Menengah Atas (SMA)",
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XII / Semester 5 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 12.TL.3 - Menganalisis teks naskah drama klasik/modern, menafsirkan subteks psikologis, serta memperagakan adegan monolog/dialog dramatik dengan intonasi dan pelafalan standar profesional.",
        "Alokasi Waktu": "20 Jam Pelajaran (4 Minggu x 5 JP @ 45 menit)",
        "Target Peserta Didik": "Peserta Didik Peminatan / Tingkat Lanjut (32 – 36 Siswa)",
        "Model Pembelajaran": "Genre-Based Approach (GBA) & Roleplay / Theatrical Performance",
        "Profil Pelajar Pancasila": "Kreatif, Mandiri, Berkebinekaan Global",
        "Sarana & Prasarana": "Naskah Drama Otentik (The Crucible, A Raisin in the Sun, Hamlet), Panggung Mini / Kostum Sederhana"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_section_header(doc, "I. KOMPONEN INTI (CORE COMPONENTS)")
    add_sub_header(doc, "A. Capaian Pembelajaran & Indikator Ketercapaian (IKTP)")
    p_cp = doc.add_paragraph()
    p_cp.paragraph_format.space_after = Pt(4)
    p_cp.add_run("Elemen: Menyimak - Berbicara & Membaca - Memirsa\n").bold = True
    p_cp.add_run("Peserta didik menganalisis struktur dramatik, pesan moral, dan subteks emosional dari naskah lakon drama sastra dunia, serta mengekspresikannya melalui penampilan teater atau rekaman monolog dramatik.")

    iktp_headers = ["Kode IKTP", "Indikator Ketercapaian Tujuan Pembelajaran", "Ranah Kognitif"]
    iktp_rows = [
        ["IKTP 3.1", "Menganalisis petunjuk panggung (stage directions), ketegangan dramatik, dan subteks psikologis tokoh.", "C4 (Analisis)"],
        ["IKTP 3.2", "Menginterpretasikan konflik moral tokoh ke dalam modulasi vokal, jeda dramatis, dan pelafalan presisi.", "C3 (Aplikasi)"],
        ["IKTP 3.3", "Mengevaluasi efektivitas pesan kritik sosial naskah drama terhadap realitas kemanusiaan kontemporer.", "C5 (Evaluasi)"],
        ["IKTP 3.4", "Mementaskan monolog dramatik atau adegan sandiwara berdurasi 3-4 menit dengan penghayatan artistik.", "C6 (Kreasi)"]
    ]
    docx_helper.add_styled_table(doc, iktp_headers, iktp_rows, [Inches(1.2), Inches(4.3), Inches(1.0)])

    add_sub_header(doc, "B. Pemahaman Bermakna (Meaningful Understanding)")
    docx_helper.add_callout(
        doc,
        "Pemahaman Bermakna (Big Ideas):",
        "1. Teks naskah sandiwara adalah karya sastra hidup; makna sejati dari dialog drama tersimpan di dalam subteks (emosi, motivasi tersembunyi, dan konflik batin yang tidak diucapkan secara harfiah).\n"
        "2. Seni peran teatrikal melatih kepekaan vokal, ekspresi tubuh, dan empati mendalam terhadap perjuangan moral manusia lintas peradaban dan zaman.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_sub_header(doc, "C. Pertanyaan Pemantik (Essential Questions)")
    p_pq = doc.add_paragraph()
    p_pq.paragraph_format.left_indent = Inches(0.2)
    p_pq.add_run("1. Why does John Proctor choose the gallows rather than signing away his good name in Arthur Miller's The Crucible?\n")
    p_pq.add_run("2. How do stage directions guide an actor's emotional tempo and spatial movement on stage?\n")
    p_pq.add_run("3. What vocal techniques allow a speaker to project vulnerability, rage, and moral triumph without shouting?")

    add_section_header(doc, "II. LANGKAH-LANGKAH PEMBELAJARAN (4 MINGGU @ 5 JP = 225 MENIT/MINGGU)")

    scenarios = [
        ("MINGGU 1 (5 JP): BKoF - Eksplorasi Teater Klasik & Subteks Drama", [
            ("Pendahuluan (20 Menit)", [
                "Salam, doa bersama, dan penayangan klip pementasan panggung Broadway 'The Crucible' (adegan persidangan).",
                "Penyampaian target performa pementasan drama / monolog TP 12.TL.3."
            ]),
            ("Kegiatan Inti (190 Menit - BKoF)", [
                "Pengenalan Elemen Dramatik: Plot arc (Exposition, Inciting Incident, Rising Action, Climax, Falling Action, Resolution).",
                "Konsep Subteks Teatrikal: Guru memperagakan kalimat 'I am fine' dengan 3 nada emosi berbeda (marah, putus asa, sarkasme) untuk menjelaskan subteks.",
                "Bedah Istilah Teater: Membahas 8 istilah (subtext, indomitable, soliloquy, consummation, contumely, quietus, modulation, stage direction).",
                "Latihan Olah Vokal & Pernapasan Diaphragma: Pemanasan vokal artikulasi konsonan bahasa Inggris (plosives, fricatives) dan proyeksi resonansi suara."
            ]),
            ("Penutup (15 Menit)", [
                "Siswa menuliskan satu kalimat dialog pendek dengan catatan subteks emosinya.",
                "Doa penutup."
            ])
        ]),
        ("MINGGU 2 (5 JP): MoT - Dekonstruksi Naskah Proctor, Hansberry & Hamlet", [
            ("Pendahuluan (15 Menit)", [
                "Pemanasan vokal dinamika intonasi (crescendo - decrescendo).",
                "Orientasi: Membedah naskah model otentik Arthur Miller, Lorraine Hansberry, dan Shakespeare."
            ]),
            ("Kegiatan Inti (195 Menit - MoT)", [
                "Bedah Naskah 1 (Arthur Miller): Siswa membaca dialog Proctor vs Danforth; mengidentifikasi anaphora ('Because it is my name!').",
                "Bedah Naskah 2 (Lorraine Hansberry): Menganalisis jeda elipsis (...) dalam monolog Walter Lee saat menolak suap Lindner.",
                "Bedah Naskah 3 (Hamlet): Mengkaji soliloquy 'To be or not to be' dan konflik eksistensial manusia.",
                "Anotasi Naskah: Siswa menandai naskah dengan simbol teatrikal (// = jeda panjang, CAPITALS = tekanan suara, ^ = intonasi naik)."
            ]),
            ("Penutup (15 Menit)", [
                "Siswa memilih salah satu naskah peran untuk dipraktikkan bersama kelompok di Minggu 3."
            ])
        ]),
        ("MINGGU 3 (5 JP): JCoT - Latihan Blocking Panggung & Table Reading", [
            ("Pendahuluan (15 Menit)", [
                "Briefing etika panggung dan keselamatan gerak (blocking safety).",
                "Pembagian kelompok adegan sandiwara atau mitra latihan monolog."
            ]),
            ("Kegiatan Inti (195 Menit - JCoT)", [
                "Table Reading Session (45 Menit): Membaca naskah secara berhadapan dengan fokus murni pada intonasi dan pelafalan.",
                "Latihan Blocking Gerak Panggung (60 Menit): Menata posisi tubuh, level ketinggian (standing vs kneeling), dan kontak visual tokoh.",
                "Rehearsal Terbimbing: Guru berkeliling memberikan masukan teknis pada modulasi suara dan penghayatan emosi tokoh.",
                "Peer Feedback Rehearsal: Pasangan kelompok saling menampilkan draf adegan dan mengisi checklist evaluasi."
            ]),
            ("Penutup (15 Menit)", [
                "Evaluasi kemajuan latihan dan pengumuman jadwal pementasan resmi untuk Minggu 4."
            ])
        ]),
        ("MINGGU 4 (5 JP): ICoT - Theatrical Performance Showcase & Evaluasi", [
            ("Pendahuluan (15 Menit)", [
                "Persiapan panggung mini kelas, tata lampu sederhana, dan nomor urut tampil pementasan.",
                "Doa bersama menenangkan kegugupan panggung."
            ]),
            ("Kegiatan Inti (195 Menit - ICoT Showcase)", [
                "Pementasan Resmi (150 Menit): Setiap siswa/kelompok menampilkan monolog atau adegan drama berdurasi 3-4 menit di hadapan dewan penonton.",
                "Perekaman Video Portofolio: Setiap penampilan direkam untuk dokumentasi asesmen mandiri.",
                "Penilaian Asesmen Sumatif Performa menggunakan Rubrik Teatrikal Analitik (Vocal Modulation, Stage Presence, Subtext Interpretation, Pronunciation).",
                "Curtain Call & Sesi Apresiasi Kelas (45 Menit): Seluruh pemeran naik ke panggung menerima tepuk tangan apresiasi."
            ]),
            ("Penutup (15 Menit)", [
                "Refleksi menyeluruh terhadap transformasi rasa percaya diri siswa.",
                "Doa penutup."
            ])
        ])
    ]
    build_meeting_scenarios(doc, scenarios)

    as_rows = [
        ["1. Diagnostik", "Tes Lisan Membaca", "Uji Ekspresi Dialog Singkat dengan Variasi Emosi", "Awal Minggu 1"],
        ["2. Formatif", "Observasi Latihan", "Lembar Pengamatan Table Reading & Blocking Rehearsal", "Minggu 2 dan 3"],
        ["3. Sumatif", "Unjuk Kerja Teatrikal", "Pementasan Monolog / Adegan Drama 3-4 Menit (Rubrik Seni Peran)", "Akhir Minggu 4"]
    ]
    enrich_text = "Siswa memproduksi video monolog teatrikal sinematik dengan latar musik dramatik dan mengikutsertakannya dalam festival seni drama remaja tingkat provinsi."
    remedi_text = "Latihan pengulangan terbimbing membaca teks monolog pendek (1 menit) dengan fokus perbaikan artikulasi fonetik dan intonasi dasar kalimat."
    ref_rows = [
        ["Penghayatan Seni Peran", "Apakah saya berhasil merasakan dan menyampaikan pergulatan batin tokoh saya?", "Bagaimana pementasan drama membantu siswa yang pendiam mengekspresikan diri?"],
        ["Vokal & Pelafalan", "Apakah artikulasi bahasa Inggris saya terdengar jelas oleh penonton di baris belakang?", "Aspek pengucapan (pronunciation) apa yang paling menonjol peningkatannya?"],
        ["Keberanian Tampil", "Bagaimana saya mengatasi kecemasan panggung (stage fright) saat tampil tadi?", "Apakah suasana panggung kelas mendukung rasa percaya diri seluruh siswa?"]
    ]
    add_common_sections(doc, as_rows, enrich_text, remedi_text, ref_rows)

    fpath = os.path.join(TARGET_DIR, "Modul_Ajar_TP_12.TL.3_Dramatic_Script_and_Performance.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# MODUL 4: TP 12.TL.4 - PERSUASIVE ORATORY & PARLIAMENTARY DEBATE
# ==============================================================================
def create_modul_4():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "MODUL AJAR KURIKULUM MERDEKA - BAHASA INGGRIS TINGKAT LANJUT KELAS XII",
        "Modul 4: Persuasive Oratory & Parliamentary Debate | TP 12.TL.4"
    )
    info = {
        "Satuan Pendidikan": "Sekolah Menengah Atas (SMA)",
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XII / Semester 5 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 12.TL.4 - Merancang dan membawakan pidato persuasif tingkat tinggi (Persuasive Oratory) dan mempraktikkan debat formal gaya parlemen dengan retorika persuasif yang meyakinkan.",
        "Alokasi Waktu": "10 Jam Pelajaran (2 Minggu x 5 JP @ 45 menit)",
        "Target Peserta Didik": "Peserta Didik Peminatan / Tingkat Lanjut (32 – 36 Siswa)",
        "Model Pembelajaran": "Genre-Based Approach (GBA) & Parliamentary Debate Simulation",
        "Profil Pelajar Pancasila": "Mandiri, Bernalar Kritis, Komunikatif",
        "Sarana & Prasarana": "Podium Pidato, Stop-watch/Timer Debat, Teks Pidato PBB, Format Asian/British Parliamentary"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_section_header(doc, "I. KOMPONEN INTI (CORE COMPONENTS)")
    add_sub_header(doc, "A. Capaian Pembelajaran & Indikator Ketercapaian (IKTP)")
    p_cp = doc.add_paragraph()
    p_cp.paragraph_format.space_after = Pt(4)
    p_cp.add_run("Elemen: Menyimak - Berbicara & Menulis - Mempresentasikan\n").bold = True
    p_cp.add_run("Peserta didik menerapkan seni retorika klasik (Ethos, Pathos, Logos) dan perangkat gaya bahasa untuk menyusun pidato persuasif yang memukau serta membangun kasus sanggahan logis dalam simulasi debat formal.")

    iktp_headers = ["Kode IKTP", "Indikator Ketercapaian Tujuan Pembelajaran", "Ranah Kognitif"]
    iktp_rows = [
        ["IKTP 4.1", "Mengaplikasikan retorika klasik (Ethos, Pathos, Logos) dalam menyusun argumen kebijakan publik.", "C3 (Aplikasi)"],
        ["IKTP 4.2", "Menggunakan perangkat retoris (parallelism, tricolon, antithesis) guna memperkuat daya persuasi pidato.", "C4 (Analisis)"],
        ["IKTP 4.3", "Mengevaluasi falasi logika (logical fallacies) dan merumuskan sanggahan taktis (rebuttal) dalam debat.", "C5 (Evaluasi)"],
        ["IKTP 4.4", "Menyampaikan pidato oratori persuasif 3-4 menit atau mempraktikkan debat parlemen resmi secara meyakinkan.", "C6 (Kreasi)"]
    ]
    docx_helper.add_styled_table(doc, iktp_headers, iktp_rows, [Inches(1.2), Inches(4.3), Inches(1.0)])

    add_sub_header(doc, "B. Pemahaman Bermakna (Meaningful Understanding)")
    docx_helper.add_callout(
        doc,
        "Pemahaman Bermakna (Big Ideas):",
        "1. Seni berpidato persuasif (oratory) bukanlah sekadar kepiawaian bersilat lidah, melainkan keharmonisan antara integritas moral pembicara (Ethos), ketajaman logika bukti (Logos), dan ketulusan empati kemanusiaan (Pathos).\n"
        "2. Debat parlementer melatih pemikiran multidimensi, kemampuan berpikir cepat di bawah tekanan waktu, dan rasa hormat terhadap perbedaan sudut pandang lawan.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_sub_header(doc, "C. Pertanyaan Pemantik (Essential Questions)")
    p_pq = doc.add_paragraph()
    p_pq.paragraph_format.left_indent = Inches(0.2)
    p_pq.add_run("1. How do great historical leaders combine rhythm, repetition, and contrast to make an argument unforgettable?\n")
    p_pq.add_run("2. In parliamentary debate, why is a well-structured rebuttal more influential than simply presenting new points?\n")
    p_pq.add_run("3. What distinguishes an ethical appeal to emotion (Pathos) from manipulative demagoguery?")

    add_section_header(doc, "II. LANGKAH-LANGKAH PEMBELAJARAN (2 MINGGU @ 5 JP = 225 MENIT/MINGGU)")

    scenarios = [
        ("MINGGU 1 (5 JP): BKoF & MoT - Tiga Pilar Retorika & Konstruksi Kasus Debat", [
            ("Pendahuluan (20 Menit)", [
                "Salam, doa bersama, dan penayangan cuplikan pidato pemuda di Sidang Umum PBB (United Nations).",
                "Penyampaian tujuan TP 12.TL.4 dan alur asesmen pidato & debat."
            ]),
            ("Kegiatan Inti (190 Menit - BKoF & MoT)", [
                "Eksplorasi Tiga Pilar Aristoteles: Membedah Ethos (kredibilitas), Logos (data kausal), dan Pathos (daya gugah nurani) pada teks pidato iklim 'The Arithmetic of Survival'.",
                "Analisis Perangkat Gaya Bahasa Oratori: Mengidentifikasi Tricolon ('do not compromise, does not wait, do not retreat'), Parallelism, dan Antithesis.",
                "Bedah Kasus Debat Parlemen: Menganalisis Teks Model 2 (Mosi larangan senjata otonom AI), mencermati teknik signposting ('Firstly, Secondly, Thirdly') dan antisipasi sanggahan.",
                "Latihan Sanggahan Berpasangan (Sparks Rebuttal): Siswa saling melontarkan argumen dan rekan sebangku memberikan respon sanggahan 'Even if... however...' dalam 30 detik."
            ]),
            ("Penutup (15 Menit)", [
                "Siswa merumuskan topik pidato oratori mandiri atau memilih peran dalam tim debat parlemen.",
                "Doa penutup."
            ])
        ]),
        ("MINGGU 2 (5 JP): JCoT & ICoT - Penulisan Naskah & Turnamen Oratori Kelas", [
            ("Pendahuluan (15 Menit)", [
                "Pemanasan artikulasi oratori dan penjelasan rubrik penilaian unjuk kerja pidato lisan.",
                "Penataan podium pidato di depan kelas."
            ]),
            ("Kegiatan Inti (195 Menit - JCoT & ICoT)", [
                "Drafting & Polishing (45 Menit): Siswa menyempurnakan naskah pidato oratori 300-400 kata atau kasus debat dengan memasukkan minimal satu tricolon dan satu antithesis.",
                "Turnamen Pidato Oratori & Debat Parlemen (135 Menit): Setiap siswa membawakan pidato persuasif 3-4 menit di podium kelas atau bertanding dalam putaran simulasi debat parlemen.",
                "Adjudication & Feedback: Dewan juri kelas (guru dan peer adjudicators) memberikan evaluasi berbasis rubrik (Ethos/Pathos/Logos balance, Rhetorical delivery, Eye contact, Logical reasoning).",
                "Penilaian Sumatif Unjuk Kerja Pidato Persuasif."
            ]),
            ("Penutup (15 Menit)", [
                "Refleksi keberanian berbicara di depan publik dan pemikiran kritis.",
                "Apresiasi guru kepada orator-orator terbaik dan doa penutup."
            ])
        ])
    ]
    build_meeting_scenarios(doc, scenarios)

    as_rows = [
        ["1. Diagnostik", "Identifikasi Teks Pidato", "Uji Pemahaman Perbedaan Argumen Fakta vs Emosi", "Awal Minggu 1"],
        ["2. Formatif", "Latihan Lisan & Draft", "Lembar Kerja Sanggahan Debat & Ceklis Naskah Pidato", "Minggu 1 dan 2"],
        ["3. Sumatif", "Unjuk Kerja Lisan", "Pidato Oratori Formal 3-4 Menit di Podium / Penampilan Debat", "Akhir Minggu 2"]
    ]
    enrich_text = "Siswa didaftarkan mewakili sekolah dalam kompetisi debat parlemen bahasa Inggris tingkat nasional (National Schools Debating Championship / NSDC)."
    remedi_text = "Latihan berulang membawakan satu paragraf pidato model dengan bimbingan jeda pernapasan dan penekanan kata kunci di hadapan cermin."
    ref_rows = [
        ["Penerapan Retorika", "Apakah pidato saya sudah memadukan bukti data (Logos) dengan daya gugah nurani (Pathos)?", "Sejauh mana siswa mampu keluar dari gaya membaca teks datar ke gaya oratori hidup?"],
        ["Kemampuan Sanggahan", "Apakah saya mampu menemukan kelemahan logika lawan dan menyanggahnya secara santun?", "Apakah siswa memahami konsep falasi logika dalam perdebatan?"],
        ["Wibawa Panggung", "Bagaimana ketenangan dan kontak visual saya saat berdiri di podium pidato?", "Bagaimana peningkatan kepercayaan diri siswa berbicara bahasa Inggris di ruang publik?"]
    ]
    add_common_sections(doc, as_rows, enrich_text, remedi_text, ref_rows)

    fpath = os.path.join(TARGET_DIR, "Modul_Ajar_TP_12.TL.4_Persuasive_Oratory_and_Debate.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# MODUL 5: TP 12.TL.5 - ADMISSIONS & CAREER DOSSIER
# ==============================================================================
def create_modul_5():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "MODUL AJAR KURIKULUM MERDEKA - BAHASA INGGRIS TINGKAT LANJUT KELAS XII",
        "Modul 5: Admissions & Career Dossier | TP 12.TL.5"
    )
    info = {
        "Satuan Pendidikan": "Sekolah Menengah Atas (SMA)",
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XII / Semester 6 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 12.TL.5 - Menyusun korespondensi kerja profesional dan akademik internasional: Surat Lamaran Kerja Eksekutif, Curriculum Vitae (CV) ATS-compliant, dan Pernyataan Motivasi Beasiswa/Kuliah (Personal Statement).",
        "Alokasi Waktu": "20 Jam Pelajaran (4 Minggu x 5 JP @ 45 menit)",
        "Target Peserta Didik": "Peserta Didik Peminatan / Tingkat Lanjut (32 – 36 Siswa)",
        "Model Pembelajaran": "Genre-Based Approach (GBA) & Portfolio-Based Learning",
        "Profil Pelajar Pancasila": "Mandiri, Bernalar Kritis, Berwawasan Global",
        "Sarana & Prasarana": "Laptop, Contoh Portofolio Beasiswa Dunia (Fulbright, LPDP, Oxford), Software ATS Checker"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_section_header(doc, "I. KOMPONEN INTI (CORE COMPONENTS)")
    add_sub_header(doc, "A. Capaian Pembelajaran & Indikator Ketercapaian (IKTP)")
    p_cp = doc.add_paragraph()
    p_cp.paragraph_format.space_after = Pt(4)
    p_cp.add_run("Elemen: Membaca - Memirsa & Menulis - Mempresentasikan\n").bold = True
    p_cp.add_run("Peserta didik menganalisis dan memproduksi genre korespondensi profesional dan dossier admisi akademik internasional dengan register bahasa bisnis formal, presisi leksikal tinggi, dan keselarasan profil diri.")

    iktp_headers = ["Kode IKTP", "Indikator Ketercapaian Tujuan Pembelajaran", "Ranah Kognitif"]
    iktp_rows = [
        ["IKTP 5.1", "Menyusun Executive Cover Letter dengan register bisnis formal dan orientasi kontribusi institusional.", "C6 (Kreasi)"],
        ["IKTP 5.2", "Merancang Curriculum Vitae berformat ATS-compliant menggunakan Google X-Y-Z formula dan kata kerja aksi.", "C6 (Kreasi)"],
        ["IKTP 5.3", "Menulis Personal Statement akademik 500-600 kata dengan narrative hook yang memikat dan institutional fit.", "C6 (Kreasi)"],
        ["IKTP 5.4", "Mengevaluasi keselarasan dossier lamaran terhadap kualifikasi target beasiswa atau profesi global.", "C5 (Evaluasi)"]
    ]
    docx_helper.add_styled_table(doc, iktp_headers, iktp_rows, [Inches(1.2), Inches(4.3), Inches(1.0)])

    add_sub_header(doc, "B. Pemahaman Bermakna (Meaningful Understanding)")
    docx_helper.add_callout(
        doc,
        "Pemahaman Bermakna (Big Ideas):",
        "1. Berkas lamaran kerja dan admisi perguruan tinggi kelas dunia bukanlah katalog prestasi pasif, melainkan sebuah narasi kepemimpinan strategis yang membuktikan bagaimana keunikan pengalaman kita mampu menjawab tantangan organisasi.\n"
        "2. Di era rekrutmen digital modern, pemahaman terhadap algoritma Applicant Tracking System (ATS) dan kemampuan kuantifikasi capaian (Google X-Y-Z formula) sangat menentukan lolosnya berkas ke meja wawancara.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_sub_header(doc, "C. Pertanyaan Pemantik (Essential Questions)")
    p_pq = doc.add_paragraph()
    p_pq.paragraph_format.left_indent = Inches(0.2)
    p_pq.add_run("1. What makes a personal statement opening hook memorable to an admissions officer reading thousands of essays?\n")
    p_pq.add_run("2. How does transforming a generic job duty into an action-oriented 'Google X-Y-Z' bullet enhance professional credibility?\n")
    p_pq.add_run("3. Why do complex graphics, multi-column tables, and icons cause ATS resume parsers to fail?")

    add_section_header(doc, "II. LANGKAH-LANGKAH PEMBELAJARAN (4 MINGGU @ 5 JP = 225 MENIT/MINGGU)")

    scenarios = [
        ("MINGGU 1 (5 JP): BKoF & MoT - Bedah Personal Statement Beasiswa", [
            ("Pendahuluan (20 Menit)", [
                "Salam, doa, dan apersepsi mengenai impian universitas dunia dan beasiswa bergengsi (LPDP, Chevening, Fulbright).",
                "Penyampaian tujuan TP 12.TL.5 dan portofolio dossier lengkap."
            ]),
            ("Kegiatan Inti (190 Menit - BKoF & MoT)", [
                "Eksplorasi Esai Admisi: Menganalisis Teks Model 1 (Model Personal Statement Oxford Rhodes Scholarship).",
                "Bedah Narrative Arc: Mengkaji bagaimana penulis mengaitkan kenangan masa kecil (pemadaman listrik pesisir) dengan prototipe energi ombak di SMA, lalu menyambungkannya dengan target studi kebijakan Oxford.",
                "Pemerolehan Kata Kerja Aksi: Membahas kata kerja berdampak tinggi ('spearheaded', 'orchestrated', 'quantified', 'innovated').",
                "Latihan Narrative Hook: Siswa menuliskan 3 alternatif kalimat pembuka esai diri yang orisinal dan menggugah rasa penasaran pembaca."
            ]),
            ("Penutup (15 Menit)", [
                "Siswa mengunggah 1 draft opening hook ke papan kolaborasi kelas.",
                "Doa penutup."
            ])
        ]),
        ("MINGGU 2 (5 JP): MoT & JCoT - Executive Cover Letter & Etika Bisnis Formal", [
            ("Pendahuluan (15 Menit)", [
                "Review kaidah salam pembuka dan penutup korespondensi formal bahasa Inggris (Salutation & Valediction).",
                "Orientasi: Membedah naskah cover letter analis kebijakan junior PBB (UNDP)."
            ]),
            ("Kegiatan Inti (195 Menit - MoT & JCoT)", [
                "Menganalisis Teks Model 2 (Executive Cover Letter): Mencermati struktur 4 paragraf (Tujuan & Keselarasan Nilai, Bukti Capaian Kuantitatif, Keterampilan Diplomasi Interkultural, Call to Action).",
                "De-generalizing Language: Siswa berlatih mengubah kalimat klise ('I am a hardworking person') menjadi bukti konkret ('Co-authored a 20-page policy brief presented at COP28').",
                "Joint Construction Berpasangan: Siswa saling bertukar target lowongan (magang internasional atau beasiswa) dan menyusun draf cover letter bersama rekannya.",
                "Penyuntingan Tata Bahasa & Register: Memastikan penggunaan modalitas sopan ('I welcome the opportunity to discuss...') dan format surat bisnis standar."
            ]),
            ("Penutup (15 Menit)", [
                "Refleksi kemajuan penulisan surat lamaran dan penugasan persiapan resume ATS."
            ])
        ]),
        ("MINGGU 3 (5 JP): MoT & ICoT - ATS-Compliant Resume Architecture", [
            ("Pendahuluan (15 Menit)", [
                "Demonstrasi langsung: Guru mendemonstrasikan bagaimana software ATS memindai resume satu kolom vs resume berkolom ganda yang rusak.",
                "Penjelasan prinsip Google X-Y-Z formula."
            ]),
            ("Kegiatan Inti (195 Menit - MoT & ICoT)", [
                "Bedah Formula Google X-Y-Z: 'Accomplished [X] measured by [Y] by doing [Z]'. Contoh: 'Increased student club participation by 45% across 2 semesters by orchestrating 12 interactive STEM workshops.'",
                "Penyusunan Resume Mandiri (ICoT): Siswa merancang 1 halaman Resume ATS-compliant lengkap (Contact, Objective, Education, Leadership & Projects, Skills & Certifications).",
                "Keyword Matching: Siswa menyelaraskan kata kunci pada resume dengan deskripsi kualifikasi lowongan target.",
                "Peer Scanning Checklist: Teman sebangku memposisikan diri sebagai sistem ATS, memeriksa keterbacaan font (Segoe UI / Calibri), ketiadaan tabel rumit, dan format tanggal konsisten."
            ]),
            ("Penutup (15 Menit)", [
                "Siswa mengunggah draf resume ke folder portofolio.",
                "Doa penutup."
            ])
        ]),
        ("MINGGU 4 (5 JP): Finalisasi Dossier Lengkap & Career Portfolio Review", [
            ("Pendahuluan (15 Menit)", [
                "Instruksi penyatuan berkas Dossier Admisi & Karier menjadi 1 file terpadu.",
                "Pemberian rubrik portofolio admisi internasional."
            ]),
            ("Kegiatan Inti (195 Menit)", [
                "Writing Conference & Final Polishing (90 Menit): Siswa menuntaskan revisi akhir ketiga dokumen (Cover Letter, ATS Resume, Personal Statement).",
                "Dossier Defense Walk (60 Menit): Setiap siswa memajang ringkasan berkasnya di meja; siswa lain berkeliling membaca dan memberikan umpan balik menggunakan sticky notes profesional.",
                "Penilaian Sumatif Portofolio Dossier oleh guru.",
                "Penyusunan Refleksi Karier (45 Menit): Siswa menuliskan peta jalan karier dan rencana studi 5 tahun ke depan."
            ]),
            ("Penutup (15 Menit)", [
                "Refleksi pencapaian kompetensi korespondensi profesional dunia nyata.",
                "Doa penutup."
            ])
        ])
    ]
    build_meeting_scenarios(doc, scenarios)

    as_rows = [
        ["1. Diagnostik", "Analisis Berkas Contoh", "Identifikasi Kesalahan pada Contoh Resume & Surat Lamaran", "Awal Minggu 1"],
        ["2. Formatif", "Ceklis & Writing Conference", "Lembar Kerja Google X-Y-Z & Draf Personal Statement", "Minggu 2 dan 3"],
        ["3. Sumatif", "Portofolio Lengkap", "Berkas Dossier Terpadu (Cover Letter, ATS CV, Personal Statement)", "Akhir Minggu 4"]
    ]
    enrich_text = "Siswa mendaftarkan berkas dossier mereka secara nyata pada program beasiswa pertukaran pelajar internasional (KL-YES, Sakura Science Program, atau beasiswa sarjana luar negeri)."
    remedi_text = "Pendampingan khusus mengonversi resume bertabel menjadi format polos satu kolom dengan template kata kerja aksi yang telah disediakan."
    ref_rows = [
        ["Orisinalitas Narasi Diri", "Apakah personal statement saya mencerminkan suara otentik dan visi unik saya?", "Apakah siswa berhasil menghindari klise dalam menceritakan riwayat hidupnya?"],
        ["Presisi Leksikal Profesional", "Apakah saya sudah konsisten menggunakan formula Google X-Y-Z dalam resume?", "Bagaimana ketepatan siswa dalam menggunakan register formal bahasa Inggris bisnis?"],
        ["Kesiapan Menghadapi Seleksi", "Seberapa siap saya menggunakan berkas ini untuk pendaftaran universitas nyata?", "Apakah siswa merasa lebih berdaya saing setelah menyusun berkas standar global?"]
    ]
    add_common_sections(doc, as_rows, enrich_text, remedi_text, ref_rows)

    fpath = os.path.join(TARGET_DIR, "Modul_Ajar_TP_12.TL.5_Admissions_and_Career_Dossier.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# MODUL 6: TP 12.TL.6 - JOB INTERVIEW & ORAL DEFENSE
# ==============================================================================
def create_modul_6():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "MODUL AJAR KURIKULUM MERDEKA - BAHASA INGGRIS TINGKAT LANJUT KELAS XII",
        "Modul 6: Executive Interview & Academic Oral Defense | TP 12.TL.6"
    )
    info = {
        "Satuan Pendidikan": "Sekolah Menengah Atas (SMA)",
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XII / Semester 6 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 12.TL.6 - Mempraktikkan wawancara kerja profesional (Job Interview Simulation) dan pertahanan karya ilmiah lisan (Oral Defense) dengan respons spontan, taktis, dan register bahasa formal.",
        "Alokasi Waktu": "15 Jam Pelajaran (3 Minggu x 5 JP @ 45 menit)",
        "Target Peserta Didik": "Peserta Didik Peminatan / Tingkat Lanjut (32 – 36 Siswa)",
        "Model Pembelajaran": "Genre-Based Approach (GBA) & High-Stakes Spoken Simulation",
        "Profil Pelajar Pancasila": "Mandiri, Bernalar Kritis, Berintegritas",
        "Sarana & Prasarana": "Meja Sidang Penguji, Kursi Wawancara Eksekutif, Bank Pertanyaan Behavioral STAR, Lembar Rubrik Lisan"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_section_header(doc, "I. KOMPONEN INTI (CORE COMPONENTS)")
    add_sub_header(doc, "A. Capaian Pembelajaran & Indikator Ketercapaian (IKTP)")
    p_cp = doc.add_paragraph()
    p_cp.paragraph_format.space_after = Pt(4)
    p_cp.add_run("Elemen: Menyimak - Berbicara\n").bold = True
    p_cp.add_run("Peserta didik merespons secara spontan, runtut, dan persuasif terhadap pertanyaan berbasis perilaku dalam wawancara kerja eksekutif serta mempertahankan validitas temuan riset dalam ujian sidang karya ilmiah lisan.")

    iktp_headers = ["Kode IKTP", "Indikator Ketercapaian Tujuan Pembelajaran", "Ranah Kognitif"]
    iktp_rows = [
        ["IKTP 6.1", "Menerapkan metode STAR (Situation, Task, Action, Result) dalam menjawab pertanyaan perilaku wawancara.", "C3 (Aplikasi)"],
        ["IKTP 6.2", "Mempertahankan kesimpulan riset secara diplomatis di hadapan kritik metodologis dewan penguji.", "C5 (Evaluasi)"],
        ["IKTP 6.3", "Menggunakan jeda taktis (tactical fillers) dan ekspresi formal untuk mengelola ketegangan berpikir lisan.", "C3 (Aplikasi)"],
        ["IKTP 6.4", "Menampilkan simulasi wawancara kerja atau sidang oral defense mandiri dengan wibawa eksekutif.", "C6 (Kreasi)"]
    ]
    docx_helper.add_styled_table(doc, iktp_headers, iktp_rows, [Inches(1.2), Inches(4.3), Inches(1.0)])

    add_sub_header(doc, "B. Pemahaman Bermakna (Meaningful Understanding)")
    docx_helper.add_callout(
        doc,
        "Pemahaman Bermakna (Big Ideas):",
        "1. Dalam wawancara tingkat tinggi, pewawancara tidak mencari jawaban teoritis hafalan, melainkan bukti ketahanan mental dan ketangkasan adaptasi masa lalu yang diceritakan secara terstruktur melalui kerangka STAR.\n"
        "2. Sidang pertahanan karya ilmiah lisan (viva voce) menguji integritas intelektual: kemampuan mengakui keterbatasan riset secara elegan justru memperkuat kredibilitas peneliti di hadapan para pakar.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_sub_header(doc, "C. Pertanyaan Pemantik (Essential Questions)")
    p_pq = doc.add_paragraph()
    p_pq.paragraph_format.left_indent = Inches(0.2)
    p_pq.add_run("1. Why is describing specific individual actions ('I convened...') far more convincing in an interview than saying 'We did...'?\n")
    p_pq.add_run("2. How does a professional researcher respond to a harsh methodological critique without becoming defensive?\n")
    p_pq.add_run("3. What is the tactical purpose of a 2-second deliberate pause before answering a high-stakes question?")

    add_section_header(doc, "II. LANGKAH-LANGKAH PEMBELAJARAN (3 MINGGU @ 5 JP = 225 MENIT/MINGGU)")

    scenarios = [
        ("MINGGU 1 (5 JP): BKoF & MoT - Metode STAR & Bahasa Diplomatis Ujian Lisan", [
            ("Pendahuluan (20 Menit)", [
                "Salam, doa bersama, dan penayangan rekaman video simulasi wawancara kerja McKinsey & sidang tesis Cambridge.",
                "Penyampaian target performa lisan TP 12.TL.6."
            ]),
            ("Kegiatan Inti (190 Menit - BKoF & MoT)", [
                "Bedah Kerangka STAR: Mengkaji alokasi waktu ideal (15% Situation & Task, 60% Specific Actions, 25% Quantified Results & Lessons).",
                "Analisis Transkrip Model 1: Menganalisis respon kandidat saat sensor robotik rusak 48 jam sebelum perlombaan.",
                "Bedah Transkrip Model 2 (Oral Defense): Menganalisis bagaimana kandidat menanggapi kritik ukuran sampel ('That is an incisive critique... however...').",
                "Penguasaan Tactical Fillers: Berlatih menggunakan delay formulae ('That is a perceptive question', 'If I understand your premise correctly...') untuk menggantikan 'uhm' atau 'like'."
            ]),
            ("Penutup (15 Menit)", [
                "Siswa menuliskan 1 pengalaman krisis nyata di sekolah yang siap diformulasikan ke format STAR.",
                "Doa penutup."
            ])
        ]),
        ("MINGGU 2 (5 JP): JCoT - Mock Interview Panels & Sidang Mini Terbimbing", [
            ("Pendahuluan (15 Menit)", [
                "Setting kelas menjadi 4 stasiun panel (2 Panel Wawancara Karier & 2 Panel Sidang Karya Ilmiah).",
                "Penjelasan panduan LKPD Aktivitas 3."
            ]),
            ("Kegiatan Inti (195 Menit - JCoT)", [
                "Simulasi Panel Berpasangan/Kelompok: Siswa bergantian berperan sebagai: (A) Kandidat Wawancara, (B) Lead Interviewer, (C) Rigorous Examiner, (D) Observer Pengisi Rubrik.",
                "Uji Coba Pertanyaan Behavioral Acak: Siswa mencabut kartu pertanyaan krisis (konflik tim, kegagalan proyek, kepemimpinan di bawah tekanan) dan merespons dalam format STAR.",
                "Simulasi Ujian Skripsi: Kandidat memaparkan metodologi esai mereka dan menjawab 2 pertanyaan kritis penguji.",
                "Debriefing & Umpan Balik Kelas: Membahas kontak mata, postur tubuh tegak, dan artikulasi vokal."
            ]),
            ("Penutup (15 Menit)", [
                "Refleksi kemajuan latihan dan pengundian jadwal ujian performa mandiri di Minggu 3."
            ])
        ]),
        ("MINGGU 3 (5 JP): ICoT - Ujian Resmi Simulasi Wawancara & Sidang Oral Defense", [
            ("Pendahuluan (15 Menit)", [
                "Pengkondisian ruang sidang resmi; tata meja formal dengan dewan penguji.",
                "Doa bersama menenangkan fokus mental."
            ]),
            ("Kegiatan Inti (195 Menit - ICoT Showcase)", [
                "Pelaksanaan Ujian Performa Mandiri (165 Menit): Setiap siswa maju ke meja panel resmi, menghadapi sesi wawancara/sidang lisan berdurasi 4-5 menit di hadapan guru dan panelis tamu.",
                "Penilaian Sumatif Berbasis Rubrik Analitik Lisan (STAR/Tactical Framework, Executive Presence & Composure, Vocabulary Precision, Spoken Fluency).",
                "Sesi Evaluasi & Feedback Langsung (30 Menit): Guru memberikan catatan apresiasi dan rekomendasi pengembangan kesiapan wawancara kerja nyata."
            ]),
            ("Penutup (15 Menit)", [
                "Refleksi akhir pencapaian kompetensi lisan tingkat tinggi.",
                "Doa penutup."
            ])
        ])
    ]
    build_meeting_scenarios(doc, scenarios)

    as_rows = [
        ["1. Diagnostik", "Tes Respons Spontan", "Uji Ketangkasan Menjawab Pertanyaan Spontan dalam 1 Menit", "Awal Minggu 1"],
        ["2. Formatif", "Mock Panel Observation", "Lembar Pengamatan STAR Practice & Ceklis Bahasa Diplomatis", "Minggu 2"],
        ["3. Sumatif", "Ujian Kinerja Lisan", "Simulasi Wawancara Eksekutif & Sidang Oral Defense Mandiri", "Akhir Minggu 3"]
    ]
    enrich_text = "Siswa berperan sebagai pewawancara senior melatih adik kelas tingkat X-XI dalam seleksi kepengurusan OSIS atau tim lomba sekolah."
    remedi_text = "Latihan pengulangan terstruktur menjawab 1 pertanyaan STAR dengan kartu pengingat alur (visual cue cards) di hadapan guru."
    ref_rows = [
        ["Struktur Jawaban STAR", "Apakah saya berhasil menjelaskan tindakan spesifik saya tanpa bertele-tele?", "Apakah siswa mampu membedakan tugas tim dengan aksi kepemimpinan pribadinya?"],
        ["Wibawa & Pengendalian Diri", "Apakah saya mampu mengendalikan jeda hening tanpa panik mengeluarkan kata 'uhm'?", "Bagaimana ketenangan emosional siswa menghadapi pertanyaan sulit?"],
        ["Artikulasi Bahasa Formal", "Apakah diksi bahasa Inggris yang saya pilih mencerminkan register profesional?", "Sejauh mana perbaikan kelancaran dan kepercayaan diri berbicara siswa?"]
    ]
    add_common_sections(doc, as_rows, enrich_text, remedi_text, ref_rows)

    fpath = os.path.join(TARGET_DIR, "Modul_Ajar_TP_12.TL.6_Job_Interview_and_Oral_Defense.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# MODUL 7: TP 12.TL.7 - RESEARCH ABSTRACT & INFOGRAPHIC
# ==============================================================================
def create_modul_7():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "MODUL AJAR KURIKULUM MERDEKA - BAHASA INGGRIS TINGKAT LANJUT KELAS XII",
        "Modul 7: Scientific Abstract & Infographic Symposium | TP 12.TL.7"
    )
    info = {
        "Satuan Pendidikan": "Sekolah Menengah Atas (SMA)",
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XII / Semester 6 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 12.TL.7 - Menyusun abstrak laporan penelitian ilmiah (Research Abstract) 250 kata dan mendesain presentasi infografis saintifik untuk diseminasi publik dalam simposium ilmiah siswa.",
        "Alokasi Waktu": "15 Jam Pelajaran (3 Minggu x 5 JP @ 45 menit)",
        "Target Peserta Didik": "Peserta Didik Peminatan / Tingkat Lanjut (32 – 36 Siswa)",
        "Model Pembelajaran": "Genre-Based Approach (GBA) & Student Scientific Symposium",
        "Profil Pelajar Pancasila": "Bernalar Kritis, Kreatif, Literasi Saintifik",
        "Sarana & Prasarana": "Papan Display Poster Ilmiah, Laptop/Software Desain Grafis (Canva), Jurnal Terakreditasi IMRaD"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_section_header(doc, "I. KOMPONEN INTI (CORE COMPONENTS)")
    add_sub_header(doc, "A. Capaian Pembelajaran & Indikator Ketercapaian (IKTP)")
    p_cp = doc.add_paragraph()
    p_cp.paragraph_format.space_after = Pt(4)
    p_cp.add_run("Elemen: Membaca - Memirsa & Menulis - Mempresentasikan\n").bold = True
    p_cp.add_run("Peserta didik mengabstraksikan temuan riset ilmiah ke dalam naskah terstruktur IMRaD maksimal 250 kata dan mendesain media komunikasi multimodal (poster infografis) untuk dipresentasikan dalam simposium ilmiah.")

    iktp_headers = ["Kode IKTP", "Indikator Ketercapaian Tujuan Pembelajaran", "Ranah Kognitif"]
    iktp_rows = [
        ["IKTP 7.1", "Merangkum latar belakang, metode, hasil, dan simpulan riset ke dalam struktur IMRaD (maks. 250 kata).", "C5 (Evaluasi)"],
        ["IKTP 7.2", "Menggunakan nominalisasi ilmiah (scientific nominalization) dan kalimat pasif objektif secara presisi.", "C3 (Aplikasi)"],
        ["IKTP 7.3", "Mendesain poster infografis saintifik berprinsip 'Ten-Second Rule' dengan hierarki visual tiga kolom.", "C6 (Kreasi)"],
        ["IKTP 7.4", "Menyajikan presentasi flash talk 2 menit dan mempertahankan poster dalam ajang simposium ilmiah siswa.", "C6 (Kreasi)"]
    ]
    docx_helper.add_styled_table(doc, iktp_headers, iktp_rows, [Inches(1.2), Inches(4.3), Inches(1.0)])

    add_sub_header(doc, "B. Pemahaman Bermakna (Meaningful Understanding)")
    docx_helper.add_callout(
        doc,
        "Pemahaman Bermakna (Big Ideas):",
        "1. Abstrak ilmiah adalah etalase utama karya riset: kemampuan merangkum berbulan-bulan eksperimen laboratorium ke dalam 250 kata yang padat dan menarik adalah puncak kemahiran menulis ilmiah.\n"
        "2. Di era banjir informasi, literasi multimodal melalui desain poster saintifik yang estetik dan berbasis data memungkinkan diseminasi sains menjangkau masyarakat luas secara inklusif.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_sub_header(doc, "C. Pertanyaan Pemantik (Essential Questions)")
    p_pq = doc.add_paragraph()
    p_pq.paragraph_format.left_indent = Inches(0.2)
    p_pq.add_run("1. How do scientific writers condense complex causal processes into tight grammatical nouns using nominalization?\n")
    p_pq.add_run("2. What is the 'Ten-Second Rule' in an academic conference poster hall, and how do visual organizers achieve it?\n")
    p_pq.add_run("3. Why is an active, conclusion-driven headline far more impactful on a research poster than a passive topic title?")

    add_section_header(doc, "II. LANGKAH-LANGKAH PEMBELAJARAN (3 MINGGU @ 5 JP = 225 MENIT/MINGGU)")

    scenarios = [
        ("MINGGU 1 (5 JP): BKoF & MoT - Struktur IMRaD & Bedah Abstrak Scopus", [
            ("Pendahuluan (20 Menit)", [
                "Salam, doa bersama, dan pameran galeri foto poster ilmiah pemenang International Science & Engineering Fair (ISEF).",
                "Penyampaian tujuan TP 12.TL.7 dan agenda Simposium Ilmiah Kelas."
            ]),
            ("Kegiatan Inti (190 Menit - BKoF & MoT)", [
                "Dekonstruksi Struktur IMRaD: Membedah Teks Model 1 (Model Abstract Green Nanotechnology 245 kata) memetakan Introduction, Methodology, Results, Discussion.",
                "Eksplorasi Nominalisasi: Siswa berlatih mengubah kalimat verbal bertele-tele menjadi frasa nomina padat ('Rapid precipitation occurred via an endothermic reaction').",
                "Kalkulasi Alokasi Kata: Menganalisis proporsi kata ideal (15% Background, 25% Methods, 45% Results & Data, 15% Significance).",
                "Latihan Penulisan Bagian Results: Siswa merangkum data tabel eksperimen fiktif menjadi 2 kalimat saintifik berbobot."
            ]),
            ("Penutup (15 Menit)", [
                "Siswa menetapkan topik penelitian yang akan diabstraksikan (bisa dari riset Modul 2 atau proyek sains mandiri).",
                "Doa penutup."
            ])
        ]),
        ("MINGGU 2 (5 JP): MoT & JCoT - Desain Poster Infografis Tiga Kolom", [
            ("Pendahuluan (15 Menit)", [
                "Review aturan 'The Ten-Second Rule' dan etika visual data (menghindari grafik menyesatkan).",
                "Orientasi panduan tata letak poster ilmiah di software desain."
            ]),
            ("Kegiatan Inti (195 Menit - MoT & JCoT)", [
                "Analisis Teks Model 2 (Poster Architecture): Mempelajari hierarki tiga kolom (Kolom 1: Problem & Objectives, Kolom 2: Methods & Workflow Diagram, Kolom 3: Key Results & Practical Impact).",
                "Formulasi Active Headline: Siswa mengubah judul pasif ('A Study on Bio-filters') menjadi headline bertenaga ('Zeolite Bio-filters Remove 95% of Toxic Bacteria in Rural Wells').",
                "Praktik Desain Kolaboratif (JCoT): Berpasangan saling membantu menata proporsi ruang kosong (white space), pemilihan palet warna kontras, dan diagram alur penelitian.",
                "Peer Feedback Desain: Memeriksa apakah teks poster tidak melebihi 400 kata agar tidak membebani mata pembaca."
            ]),
            ("Penutup (15 Menit)", [
                "Siswa merampungkan draf poster digital siap cetak/tayang untuk Minggu 3.",
                "Doa penutup."
            ])
        ]),
        ("MINGGU 3 (5 JP): ICoT - Student Scientific Poster Symposium Showcase", [
            ("Pendahuluan (15 Menit)", [
                "Setting ruang kelas menjadi Hall Simposium Ilmiah Internasional (papan stan poster berderet rapi).",
                "Pembagian ID Card Presenter & Visitor Log."
            ]),
            ("Kegiatan Inti (195 Menit - ICoT Showcase)", [
                "Simposium Sesi 1 (60 Menit): Separuh siswa berdiri di stan poster masing-masing menyampaikan flash talk 2 menit dan melayani tanya jawab dari pengunjung.",
                "Simposium Sesi 2 (60 Menit): Bertukar peran; separuh siswa berikutnya menjadi presenter di stannya.",
                "Penilaian Asesmen Sumatif Produk & Presentasi Multimodal oleh guru menggunakan rubrik analitik poster simposium.",
                "Sesi Pleno Penghargaan (45 Menit): Pengumuman poster terfavorit (Best Visual Design) dan presenter terbaik (Best Flash Talk)."
            ]),
            ("Penutup (15 Menit)", [
                "Refleksi keberhasilan mengomunikasikan sains secara visual dan verbal.",
                "Apresiasi guru atas dedikasi literasi saintifik siswa, diakhiri doa bersama."
            ])
        ])
    ]
    build_meeting_scenarios(doc, scenarios)

    as_rows = [
        ["1. Diagnostik", "Identifikasi Bagian Abstrak", "Uji Pemetaan 4 Elemen IMRaD pada Abstrak Acak", "Awal Minggu 1"],
        ["2. Formatif", "Ceklis Desain & Draf", "Lembar Telaah Nominalisasi & Pratinjau Desain Poster Tiga Kolom", "Minggu 2"],
        ["3. Sumatif", "Produk Multimodal & Kinerja", "Abstrak Ilmiah 250 Kata & Gelar Simposium Poster Ilmiah", "Akhir Minggu 3"]
    ]
    enrich_text = "Siswa mengirimkan naskah abstrak dan posternya ke ajang National Young Inventors Award (NYIA) atau Olimpiade Penelitian Siswa Indonesia (OPSI)."
    remedi_text = "Bimbingan terarah mengisi template abstrak 4 kotak (Background, Methods, Results, Conclusion) dengan panduan jumlah kata per kotak."
    ref_rows = [
        ["Keringkasan Abstrak", "Apakah naskah abstrak saya mampu memuat seluruh esensi riset di bawah 250 kata?", "Berapa banyak siswa yang berhasil menguasai teknik nominalisasi?"],
        ["Komunikasi Visual Poster", "Apakah pengunjung simposium mampu memahami temuan kunci saya dalam 10 detik?", "Bagaimana efektivitas tata letak visual infografis yang dirancang siswa?"],
        ["Daya Papar Lisan (Flash Talk)", "Seberapa lancar saya menyampaikan penjelasan riset 2 menit tanpa melihat teks?", "Apakah suasana simposium berhasil menumbuhkan atmosfer ilmiah yang positif?"]
    ]
    add_common_sections(doc, as_rows, enrich_text, remedi_text, ref_rows)

    fpath = os.path.join(TARGET_DIR, "Modul_Ajar_TP_12.TL.7_Research_Abstract_and_Infographic.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# MODUL 8: TP 12.TL.8 - STANDARDIZED READINESS (B2+)
# ==============================================================================
def create_modul_8():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "MODUL AJAR KURIKULUM MERDEKA - BAHASA INGGRIS TINGKAT LANJUT KELAS XII",
        "Modul 8: Standardized English Proficiency (IELTS/TOEFL/CEFR B2+) | TP 12.TL.8"
    )
    info = {
        "Satuan Pendidikan": "Sekolah Menengah Atas (SMA)",
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XII / Semester 6 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 12.TL.8 - Mengembangkan kesiapan literasi bahasa Inggris standar internasional (International Academic Literacy: IELTS/TOEFL/CEFR B2+ reading & critical analytical writing).",
        "Alokasi Waktu": "10 Jam Pelajaran (2 Minggu x 5 JP @ 45 menit)",
        "Target Peserta Didik": "Peserta Didik Peminatan / Tingkat Lanjut (32 – 36 Siswa)",
        "Model Pembelajaran": "Genre-Based Approach (GBA) & Timed Benchmark Assessment",
        "Profil Pelajar Pancasila": "Mandiri, Bernalar Kritis, Berdaya Saing Global",
        "Sarana & Prasarana": "Paket Soal Simulasi Standar (IELTS/TOEFL Format), Grafik Tren Energi Bersih, Band Descriptors 9 Skala"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_section_header(doc, "I. KOMPONEN INTI (CORE COMPONENTS)")
    add_sub_header(doc, "A. Capaian Pembelajaran & Indikator Ketercapaian (IKTP)")
    p_cp = doc.add_paragraph()
    p_cp.paragraph_format.space_after = Pt(4)
    p_cp.add_run("Elemen: Menyimak, Membaca, Menulis & Berbicara\n").bold = True
    p_cp.add_run("Peserta didik mendemonstrasikan kemahiran membaca analitis teks saintifik panjang dan menulis laporan deskripsi data grafik matematis di bawah batasan waktu ketat berstandar internasional (CEFR B2+).")

    iktp_headers = ["Kode IKTP", "Indikator Ketercapaian Tujuan Pembelajaran", "Ranah Kognitif"]
    iktp_rows = [
        ["IKTP 8.1", "Menerapkan teknik membaca cepat (skimming, scanning, inferencing) pada teks akademik kutub/paleoklimatologi.", "C4 (Analisis)"],
        ["IKTP 8.2", "Menyusun laporan analisis tren grafik (IELTS Academic Task 1) 150-180 kata dengan presisi leksikal.", "C6 (Kreasi)"],
        ["IKTP 8.3", "Menggunakan variasi kohesi dan struktur tata bahasa kompleks (inversion, participial clauses) secara akurat.", "C3 (Aplikasi)"],
        ["IKTP 8.4", "Menyelesaikan tes simulasi kemahiran terstandar terpadu sebagai tolak ukur kesiapan perguruan tinggi.", "C5 (Evaluasi)"]
    ]
    docx_helper.add_styled_table(doc, iktp_headers, iktp_rows, [Inches(1.2), Inches(4.3), Inches(1.0)])

    add_sub_header(doc, "B. Pemahaman Bermakna (Meaningful Understanding)")
    docx_helper.add_callout(
        doc,
        "Pemahaman Bermakna (Big Ideas):",
        "1. Ujian kemahiran bahasa Inggris terstandar internasional (IELTS Band 7.5+ / TOEFL 100+) bukan sekadar menguji perbendaharaan kata, melainkan menguji kecepatan pemrosesan kognitif, ketajaman inferensi logis, dan disiplin manajemen waktu.\n"
        "2. Penguasaan deskripsi data visual (angka, persentase, tren fluktuasi) adalah kompetensi universal yang mutlak dibutuhkan dalam studi sains, rekayasa, ekonomi, dan kebijakan publik di tingkat universitas.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_sub_header(doc, "C. Pertanyaan Pemantik (Essential Questions)")
    p_pq = doc.add_paragraph()
    p_pq.paragraph_format.left_indent = Inches(0.2)
    p_pq.add_run("1. How do skimming for main ideas and scanning for specific numerical data allow readers to beat the exam clock?\n")
    p_pq.add_run("2. What specific elements differentiate a mediocre graph description from an official Band 8+ academic report?\n")
    p_pq.add_run("3. Why does deploying varied grammatical structures (such as inversion and participial phrases) boost writing band scores?")

    add_section_header(doc, "II. LANGKAH-LANGKAH PEMBELAJARAN (2 MINGGU @ 5 JP = 225 MENIT/MINGGU)")

    scenarios = [
        ("MINGGU 1 (5 JP): BKoF & MoT - Strategi Membaca Akademik & Deskripsi Grafik", [
            ("Pendahuluan (20 Menit)", [
                "Salam, doa bersama, dan pengenalan format tes kemahiran global (IELTS Academic vs TOEFL iBT vs CEFR B2+).",
                "Penyampaian tujuan benchmark akhir kelulusan TP 12.TL.8."
            ]),
            ("Kegiatan Inti (190 Menit - BKoF & MoT)", [
                "Workshop Reading Strategies: Berlatih teknik skimming (30 detik per paragraf) dan scanning kata kunci pada Teks Otentik 1 (Paleoclimatology & Ice Cores).",
                "Inferencing Deep-Dive: Membedah pertanyaan tipe 'True/False/Not Given' dan 'Author's Implicit Meaning'.",
                "Bedah Model Writing Task 1: Menganalisis Teks Model 2 (Investasi Energi Bersih 2015-2025). Mengamati 4 bagian vital (Introduction Paraphrase, Overall Overview, Body 1 Trend, Body 2 Divergence).",
                "Perbendaharaan Kosakata Tren: Membahas kata kerja tren ('plummeted', 'soared', 'plateaued', 'reached parity', 'outstripped') dan rasio matematis."
            ]),
            ("Penutup (15 Menit)", [
                "Siswa merangkum rumus kalimat overview grafik.",
                "Doa penutup."
            ])
        ]),
        ("MINGGU 2 (5 JP): JCoT & ICoT - Latihan Terbatas & Diagnostic Benchmark Test", [
            ("Pendahuluan (15 Menit)", [
                "Review kriteria Band Descriptors (Task Achievement, Coherence & Cohesion, Lexical Resource, Grammatical Range).",
                "Instruksi pelaksanaan simulasi tes terstandar di bawah batasan waktu ketat."
            ]),
            ("Kegiatan Inti (195 Menit - JCoT & ICoT)", [
                "Timed Writing Task 1 Practice (45 Menit): Siswa menulis laporan analisis grafik 150 kata dalam waktu tepat 20 menit, dilanjutkan koreksi silang kemitraan (peer correction).",
                "Diagnostic Benchmark Test Mandiri (90 Menit - ICoT):",
                "  • Sesi 1: Academic Reading Comprehension (40 Menit - Teks saintifik kutub & 10 butir soal inferensial HOTS).",
                "  • Sesi 2: Academic Analytical Writing Task (40 Menit - Penulisan laporan data grafik terstandar).",
                "  • Sesi 3: Lembar Evaluasi Diri (10 Menit).",
                "Pembahasan Kunci & Analisis Skor (60 Menit): Guru membedah kunci jawaban, membahas pola jebakan soal inferensi, dan memetakan prediksi capaian CEFR kelas."
            ]),
            ("Penutup (15 Menit)", [
                "Refleksi paripurna atas seluruh perjalanan belajar Bahasa Inggris Tingkat Lanjut Fase F.",
                "Pesan motivasi guru menyongsong gerbang perguruan tinggi dan doa penutup."
            ])
        ])
    ]
    build_meeting_scenarios(doc, scenarios)

    as_rows = [
        ["1. Diagnostik", "Diagnostic Pre-Test", "Pemetaan Level Awal Membaca Cepat & Kosakata Standar", "Awal Minggu 1"],
        ["2. Formatif", "Timed Drills & Peer Review", "Latihan Menulis Grafik 20 Menit & Ceklis Band Descriptors", "Minggu 1 dan 2"],
        ["3. Sumatif", "Standardized Benchmark Test", "Simulasi Resmi IELTS Academic Reading & Writing Task 1", "Akhir Minggu 2"]
    ]
    enrich_text = "Siswa difasilitasi mengikuti tes kemahiran resmi bersertifikasi internasional (IELTS / TOEFL ITP / Duolingo English Test) untuk persyaratan beasiswa universitas luar negeri."
    remedi_text = "Bimbingan terfokus mengenali kata kunci penanda waktu (time indicators) dan formula kalimat pembuka deskripsi grafik yang baku."
    ref_rows = [
        ["Kecepatan Membaca & Inferensi", "Apakah saya mampu menyelesaikan teks bacaan panjang tanpa kehabisan waktu?", "Bagaimana persentase akurasi siswa dalam menjawab butir inferensial tersirat?"],
        ["Akurasi Deskripsi Data", "Apakah laporan grafik saya sudah memuat gambaran umum (overview) yang jelas?", "Apakah siswa mampu menghindari spekulasi pribadi saat mendeskripsikan data grafik?"],
        ["Kesiapan Menghadapi Ujian Global", "Seberapa percaya diri saya menghadapi tes kemahiran standar universitas?", "Bagaimana capaian kompetensi rata-rata siswa di akhir Fase F Tingkat Lanjut?"]
    ]
    add_common_sections(doc, as_rows, enrich_text, remedi_text, ref_rows)

    fpath = os.path.join(TARGET_DIR, "Modul_Ajar_TP_12.TL.8_Standardized_English_Proficiency_Readiness.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


def main():
    print("Generating comprehensive Modul Ajar for 5-Modul Ajar XII SMA TL...")
    create_modul_1()
    create_modul_2()
    create_modul_3()
    create_modul_4()
    create_modul_5()
    create_modul_6()
    create_modul_7()
    create_modul_8()
    print("All 8 Modul Ajar documents successfully generated!")

if __name__ == "__main__":
    main()

