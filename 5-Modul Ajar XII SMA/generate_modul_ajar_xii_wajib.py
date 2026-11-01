# -*- coding: utf-8 -*-
"""
Generator script for Modul Ajar (RPP Merdeka) in 5-Modul Ajar XII SMA (Wajib).
Covers TP 12.W.1 through TP 12.W.6 with complete pedagogical components:
Informasi Umum, Tujuan Pembelajaran, Langkah-langkah Pembelajaran Rinci (GBA: BKoF, MoT, JCoT, ICoT),
Asesmen, Pengayaan, Remedial, Refleksi Guru & Siswa, serta Lampiran.
"""

import os
import sys
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import docx_helper

TARGET_DIR = r"C:\Users\KOMPUTERKU2\Documents\English SMA\5-Modul Ajar XII SMA"
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

# ==============================================================================
# MODUL 1: TP 12.W.1 - DISCUSSION TEXT (FUTURE OF WORK & AUTOMATION)
# ==============================================================================
def create_modul_1():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "MODUL AJAR KURIKULUM MERDEKA - BAHASA INGGRIS KELAS XII SMA",
        "Modul 1: Discussion Text (The Future of Work, Automation & AI in the Workplace)"
    )

    info = {
        "Satuan Pendidikan": "Sekolah Menengah Atas (SMA)",
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XII / Semester 5 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 12.W.1 - Menganalisis teks diskusi mengenai disrupsi dunia kerja masa depan, etika otomasi AI, dan keberlanjutan lingkungan, serta menyusun teks diskusi berimbang.",
        "Alokasi Waktu": "10 Jam Pelajaran (5 Pertemuan @ 2 JP = 45 menit/JP)",
        "Target Peserta Didik": "Peserta Didik Reguler (32 – 36 Siswa)",
        "Model Pembelajaran": "Genre-Based Approach (GBA) & Problem-Based Learning (PBL)",
        "Profil Pelajar Pancasila": "Bernalar Kritis, Mandiri, Gotong Royong",
        "Sarana & Prasarana": "LCD Proyektor, Laptop, Teks Otentik (World Economic Forum, BBC), Lembar Kerja Siswa"
    }
    docx_helper.add_tp_info_box(doc, info)

    # I. KOMPONEN INTI
    add_section_header(doc, "I. KOMPONEN INTI (CORE COMPONENTS)")
    
    add_sub_header(doc, "A. Capaian Pembelajaran & Indikator Ketercapaian (IKTP)")
    p_cp = doc.add_paragraph()
    p_cp.paragraph_format.space_after = Pt(4)
    p_cp.add_run("Elemen: Menyimak - Berbicara & Membaca - Memirsa\n").bold = True
    p_cp.add_run("Pada akhir Fase F, peserta didik menggunakan teks lisan, tulisan, dan visual dalam bahasa Inggris untuk berkomunikasi sesuai dengan situasi, tujuan, dan pemirsa/pembaca. Peserta didik menganalisis dan mengevaluasi wacana teks diskusi multi-sudut pandang mengenai isu-isu kontemporer dunia kerja dan teknologi.")

    iktp_headers = ["Kode IKTP", "Indikator Ketercapaian Tujuan Pembelajaran", "Ranah Kognitif"]
    iktp_rows = [
        ["IKTP 1.1", "Mengidentifikasi isu kontroversial, argumen pendukung (pros), dan argumen penentang (cons) dalam teks diskusi mengenai otomasi tempat kerja.", "C4 (Analisis)"],
        ["IKTP 1.2", "Menemukan fungsi dan menggunakan konjungsi konsesif (although, despite, in spite of) serta penanda transisi kontras (conversely, on the other hand).", "C3 (Aplikasi)"],
        ["IKTP 1.3", "Mengevaluasi objektivitas dan teknik bahasa berpagar (hedging) dalam menyajikan argumen netral.", "C5 (Evaluasi)"],
        ["IKTP 1.4", "Menyusun draf teks diskusi tertulis 250-300 kata yang berimbang disertai rekomendasi jalan tengah yang solutif.", "C6 (Kreasi)"]
    ]
    docx_helper.add_styled_table(doc, iktp_headers, iktp_rows, [Inches(1.2), Inches(4.3), Inches(1.0)])

    add_sub_header(doc, "B. Pemahaman Bermakna (Meaningful Understanding)")
    docx_helper.add_callout(
        doc,
        "Pemahaman Bermakna (Big Ideas):",
        "1. Perkembangan kecerdasan buatan (AI) dan otomasi kerja bukanlah ancaman mutlak ataupun anugerah tanpa cela; keduanya membawa peluang efisiensi sekaligus tantangan disrupsi tenaga kerja.\n"
        "2. Kemampuan menimbang dua sudut pandang secara adil, objektif, dan berbasis data empiris adalah keterampilan abad ke-21 yang sangat krusial bagi calon lulusan SMA yang akan memasuki dunia perkuliahan dan kerja.",
        bg="EFF6FF", border_color="1E40AF"
    )

    add_sub_header(doc, "C. Pertanyaan Pemantik (Essential Questions)")
    p_pq = doc.add_paragraph()
    p_pq.paragraph_format.left_indent = Inches(0.2)
    p_pq.add_run("1. Will artificial intelligence and automated robotics take away jobs, or will they create better opportunities for young people?\n")
    p_pq.add_run("2. What unique human qualities (such as empathy, ethics, and creativity) can never be fully replaced by computer algorithms?\n")
    p_pq.add_run("3. How do we write a fair and balanced essay when addressing a deeply controversial issue without taking sides too early?")

    # II. KEGIATAN PEMBELAJARAN
    add_section_header(doc, "II. LANGKAH-LANGKAH PEMBELAJARAN LENGKAP (5 PERTEMUAN)")

    scenarios = [
        ("PERTEMUAN 1 (2 JP @ 45 Menit): Building Knowledge of the Field (BKoF)", [
            ("Pendahuluan (15 Menit)", [
                "Guru membuka pembelajaran dengan salam, doa bersama, dan presensi kehadiran siswa.",
                "Apersepsi & Motivasi: Menayangkan video singkat (2 menit) tentang robot humanoid yang melayani di kafetaria atau gudang otomasi logistik modern.",
                "Guru menyampaikan Tujuan Pembelajaran (TP 12.W.1) dan mengaitkannya dengan kesiapan karier siswa pasca-kelulusan SMA."
            ]),
            ("Kegiatan Inti (65 Menit - Sintaks BKoF)", [
                "Eksplorasi Konteks: Siswa berpartisipasi dalam polling interaktif kelas: 'Profesi apa saja yang menurutmu akan tergantikan oleh AI dalam 5 tahun ke depan?'",
                "Diskusi Kelas Terbuka: Memetakan profesi yang rentan (kasir, data entry, customer service) versus profesi yang aman (seniman, konselor psikologi, perawat, peneliti etika).",
                "Pengenalan Kosakata Kunci: Membahas 8 istilah otentik ketenagakerjaan (displacement, productivity, commute, reskilling, hybrid model, co-pilot) menggunakan flashcard kontekstual."
            ]),
            ("Penutup (10 Menit)", [
                "Siswa menuliskan 1 kalimat refleksi pada kertas sticky note: 'Satu hal yang paling membuat saya penasaran mengenai masa depan dunia kerja.'",
                "Guru memberikan gambaran materi pertemuan berikutnya (membaca teks model) dan menutup kelas dengan doa."
            ])
        ]),
        ("PERTEMUAN 2 (2 JP @ 45 Menit): Modeling of the Text (MoT) - Bedah Struktur Diskusi", [
            ("Pendahuluan (15 Menit)", [
                "Review singkat kosakata pertemuan sebelumnya melalui kuis tebak kata berpasangan (Word Pairing).",
                "Penyampaian target: Menganalisis teks model otentik dari World Economic Forum."
            ]),
            ("Kegiatan Inti (65 Menit - Sintaks MoT)", [
                "Membaca Teks Model: Siswa membaca teks otentik 1: 'Will Artificial Intelligence Replace Human Workers or Empower Them?' (World Economic Forum Report).",
                "Bedah Struktur Teks: Guru memandu dekonstruksi teks di papan tulis: (1) Issue, (2) Arguments For (Pros), (3) Arguments Against (Cons), (4) Conclusion / Balanced Recommendation.",
                "Analisis Bahasa: Mengidentifikasi penanda wacana kontrastif (conversely, on the other hand) dan klausa konsesif (although, despite, in spite of). Siswa berlatih menggabungkan dua kalimat kontras menjadi satu kalimat majemuk berimbang."
            ]),
            ("Penutup (10 Menit)", [
                "Siswa merangkum rumus struktur teks diskusi di buku catatan.",
                "Guru memberikan umpan balik dan mengumumkan pembagian kelompok untuk pertemuan 3."
            ])
        ]),
        ("PERTEMUAN 3 (2 JP @ 45 Menit): Joint Construction of the Text (JCoT) - Analisis Kasus Kelompok", [
            ("Pendahuluan (10 Menit)", [
                "Orientasi tugas kelompok: Siswa berkumpul dalam kelompok beranggotakan 4 orang.",
                "Pemberian LKPD Aktivitas 3 (Matriks Analisis Pro-Kontra Sistem Kerja Jarak Jauh / Remote Work)."
            ]),
            ("Kegiatan Inti (70 Menit - Sintaks JCoT)", [
                "Eksplorasi Teks 2: Setiap kelompok membaca teks otentik 2: 'Remote Work vs. Office Collaboration: The Great Workplace Divide' (BBC Worklife).",
                "Penyusunan Matriks: Kelompok mengisi tabel analisis pro dan kontra mengenai kerja dari rumah vs kerja di kantor.",
                "Perumusan Rekomendasi Solutif: Setiap kelompok menyusun 1 paragraf rekomendasi jalan tengah (hybrid working model) menggunakan bahasa berpagar (hedging).",
                "Gallery Walk: Setiap kelompok menempelkan hasil matriksnya di dinding kelas; perwakilan kelompok berkeliling memberikan komentar dan bintang apresiasi."
            ]),
            ("Penutup (10 Menit)", [
                "Guru memberikan evaluasi terhadap hasil kerja kelompok, menyoroti kekuatan argumen dan ketepatan penggunaan konjungsi kontras.",
                "Pengingat untuk persiapan penulisan mandiri pada pertemuan 4."
            ])
        ]),
        ("PERTEMUAN 4 (2 JP @ 45 Menit): Independent Construction of the Text (ICoT) - Penulisan Draf Mandiri", [
            ("Pendahuluan (10 Menit)", [
                "Pengkondisian kelas mandiri dan penjelasan instruksi tugas penulisan esai teks diskusi (250-300 kata).",
                "Guru membagikan LKPD Aktivitas 4 dan lembar rubrik penulisan."
            ]),
            ("Kegiatan Inti (70 Menit - Sintaks ICoT)", [
                "Pemilihan Topik: Siswa memilih salah satu topik: (A) Otomasi Layanan Pelanggan dengan Chatbot AI, atau (B) Kewajiban Magang Teknologi bagi Siswa SMA.",
                "Drafting Mandiri: Siswa menyusun draf esai lengkap 4 paragraf (Issue -> Pro -> Con -> Recommendation).",
                "Guru berkeliling memberikan bimbingan individual (scaffolding) bagi siswa yang mengalami kendala struktur kalimat atau pemilihan kosakata."
            ]),
            ("Penutup (10 Menit)", [
                "Siswa menyimpan draf penulisan dan menandai bagian yang ingin dikonsultasikan pada sesi peer-review pertemuan 5."
            ])
        ]),
        ("PERTEMUAN 5 (2 JP @ 45 Menit): Peer-Review, Finalisasi & Asesmen Sumatif", [
            ("Pendahuluan (10 Menit)", [
                "Penjelasan kriteria penyuntingan antarteman (Peer-Editing Checklist).",
                "Siswa bertukar draf dengan teman sebangku."
            ]),
            ("Kegiatan Inti (70 Menit)", [
                "Peer-Editing Session (25 Menit): Memeriksa keselarasan struktur, keberimbangan argumen, penggunaan konjungsi kontras, serta ejaan dan tanda baca.",
                "Revisi & Finalisasi (35 Menit): Siswa memperbaiki draf berdasarkan masukan teman dan menuliskan naskah akhir yang rapi.",
                "Pengumpulan Portofolio & Mini-Presentation (10 Menit): Dua perwakilan siswa membacakan rekomendasi solutif terbaik di depan kelas."
            ]),
            ("Penutup (10 Menit)", [
                "Refleksi menyeluruh terhadap pencapaian TP 12.W.1.",
                "Apresiasi guru atas dedikasi dan penalaran kritis siswa, diakhiri doa penutup."
            ])
        ])
    ]

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

    # III. RENCANA ASESMEN
    add_section_header(doc, "III. RENCANA ASESMEN PEMBELAJARAN (ASSESSMENT PLAN)")
    as_headers = ["Jenis Asesmen", "Teknik Penilaian", "Bentuk Instrumen", "Waktu Pelaksanaan"]
    as_rows = [
        ["1. Diagnostik Awal", "Tes Tulis Singkat & Angket", "Soal Pemetaan Pengetahuan Awal & Angket Minat", "Awal Pertemuan 1"],
        ["2. Formatif (Proses)", "Observasi & Kinerja", "Lembar Pengamatan Diskusi Kelompok & Ceklis LKPD", "Pertemuan 2, 3, dan 4"],
        ["3. Sumatif (Akhir)", "Tes Tertulis / Proyek", "Esai Teks Diskusi Berimbang (Dinilai dengan Rubrik)", "Akhir Pertemuan 5"]
    ]
    docx_helper.add_styled_table(doc, as_headers, as_rows, [Inches(1.5), Inches(1.5), Inches(2.3), Inches(1.2)])

    # IV. PENGAYAAN & REMEDIAL
    add_section_header(doc, "IV. PROGRAM PENGAYAAN DAN REMEDIAL")
    docx_helper.add_callout(
        doc,
        "Program Pengayaan (Bagi Siswa yang Telah Mencapai Ketuntasan):",
        "Siswa diminta menuliskan esai komparasi kebijakan ketenagakerjaan di 2 negara maju (misal: Jepang dan Jerman) dalam merespons otomasi AI, serta menyajikannya dalam bentuk infografis digital.",
        bg="DCFCE7", border_color="16A34A"
    )
    docx_helper.add_callout(
        doc,
        "Program Remedial (Bagi Siswa yang Membutuhkan Bimbingan):",
        "Bimbingan terfokus secara perorangan atau kelompok kecil untuk mengidentifikasi argumen pro dan kontra menggunakan template pemandu (guided writing template) serta latihan khusus konjungsi kontras sederhana.",
        bg="FEF3C7", border_color="D97706"
    )

    # V. REFLEKSI
    add_section_header(doc, "V. REFLEKSI GURU DAN PESERTA DIDIK")
    ref_headers = ["Aspek Refleksi", "Pertanyaan Refleksi Diri Siswa", "Pertanyaan Refleksi Guru"]
    ref_rows = [
        ["Pemahaman Konsep", "Apakah saya sudah memahami perbedaan argumen pendukung dan penentang?", "Apakah alur pembelajaran GBA berhasil membantu siswa menyusun argumen objektif?"],
        ["Keterampilan Bahasa", "Apakah saya mampu menggunakan kata penghubung kontras dengan tepat?", "Bagian tata bahasa mana yang masih sering menjadi kendala bagi siswa?"],
        ["Kolaborasi & Sikap", "Bagaimana kontribusi saya dalam diskusi kelompok tadi?", "Apakah seluruh siswa terlibat aktif dan merasa aman mengemukakan pendapatnya?"]
    ]
    docx_helper.add_styled_table(doc, ref_headers, ref_rows, [Inches(1.5), Inches(2.5), Inches(2.5)])

    fpath = os.path.join(TARGET_DIR, "Modul_Ajar_TP_12.W.1_Discussion_Text.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# MODUL 2: TP 12.W.2 - NEWS ITEM TEXT & MEDIA LITERACY
# ==============================================================================
def create_modul_2():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "MODUL AJAR KURIKULUM MERDEKA - BAHASA INGGRIS KELAS XII SMA",
        "Modul 2: News Item Text & Media Literacy (News Broadcasting & Journalism)"
    )

    info = {
        "Satuan Pendidikan": "Sekolah Menengah Atas (SMA)",
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XII / Semester 5 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 12.W.2 - Mengidentifikasi fakta utama, latar belakang peristiwa, dan kutipan sumber dalam teks berita aktual (News Item), serta mempraktikkan pembacaan berita (News Anchor).",
        "Alokasi Waktu": "10 Jam Pelajaran (5 Pertemuan @ 2 JP = 45 menit/JP)",
        "Target Peserta Didik": "Peserta Didik Reguler (32 – 36 Siswa)",
        "Model Pembelajaran": "Genre-Based Approach (GBA) & Project-Based Learning (News Casting Simulation)",
        "Profil Pelajar Pancasila": "Bernalar Kritis, Mandiri, Berkebinekaan Global",
        "Sarana & Prasarana": "Kamera Ponsel, Mikrofon Sederhana, Naskah Berita Otentik Antara/BBC, LCD Proyektor"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_section_header(doc, "I. KOMPONEN INTI")
    add_sub_header(doc, "A. Capaian Pembelajaran & Indikator Ketercapaian (IKTP)")
    p_cp = doc.add_paragraph()
    p_cp.paragraph_format.space_after = Pt(4)
    p_cp.add_run("Elemen: Membaca - Memirsa & Menulis - Mempresentasikan\n").bold = True
    p_cp.add_run("Peserta didik menganalisis struktur teks berita jurnalistik aktual, memverifikasi kredibilitas fakta versus opini, dan menyajikan naskah siaran berita secara lisan di depan kamera.")

    iktp_headers = ["Kode IKTP", "Indikator Ketercapaian Tujuan Pembelajaran", "Ranah Kognitif"]
    iktp_rows = [
        ["IKTP 2.1", "Mengidentifikasi unsur 5W+1H pada bagian Newsworthy Event (Lead) dalam teks berita cetak dan daring.", "C2 (Pemahaman)"],
        ["IKTP 2.2", "Menganalisis fungsi kutipan narasumber (quotes) dan kata kerja pelapor (reporting verbs) dalam memperkuat objektivitas berita.", "C4 (Analisis)"],
        ["IKTP 2.3", "Menulis naskah berita liputan sekolah/daerah sepanjang 120-150 kata dengan tata bahasa baku.", "C6 (Kreasi)"],
        ["IKTP 2.4", "Memperagakan peran pembaca berita TV (News Anchor) dengan intonasi, artikulasi, dan kontak mata yang profesional.", "P3 (Kinerja Lisan)"]
    ]
    docx_helper.add_styled_table(doc, iktp_headers, iktp_rows, [Inches(1.2), Inches(4.3), Inches(1.0)])

    add_sub_header(doc, "B. Pemahaman Bermakna & Pertanyaan Pemantik")
    docx_helper.add_callout(
        doc,
        "Pemahaman Bermakna:",
        "Berita jurnalistik yang kredibel dibangun di atas fakta empiris yang terverifikasi, bukan sensasi atau spekulasi. Kemampuan menyaring dan menyampaikan berita akurat adalah tanggung jawab moral warga digital cerdas.",
        bg="EFF6FF", border_color="1E40AF"
    )
    p_pq = doc.add_paragraph()
    p_pq.paragraph_format.left_indent = Inches(0.2)
    p_pq.add_run("1. How do professional news journalists distinguish between verified facts and personal opinions?\n")
    p_pq.add_run("2. What makes a television news anchor sound confident, trustworthy, and authoritative?\n")
    p_pq.add_run("3. Why are direct quotes from official eyewitnesses or experts crucial in a news report?")

    add_section_header(doc, "II. LANGKAH-LANGKAH PEMBELAJARAN LENGKAP (5 PERTEMUAN)")
    scenarios_2 = [
        ("PERTEMUAN 1 (2 JP): BKoF - Eksplorasi Jurnalisme & Breaking News", [
            ("Langkah Inti", [
                "Menonton video pendek siaran berita bahasa Inggris aktual (BBC World News / CNA / Antara News).",
                "Brainstorming unsur berita: What, Who, When, Where, Why, How.",
                "Eksplorasi kosakata berita: headline, inauguration, milestone, congestion, official statement."
            ])
        ]),
        ("PERTEMUAN 2 (2 JP): MoT - Dekonstruksi Teks Berita 'Whoosh Bullet Train'", [
            ("Langkah Inti", [
                "Membaca teks otentik: 'Indonesia Launches Southeast Asia's First High-Speed Railway: Whoosh Begins Operations'.",
                "Membedah struktur berita: Headline -> Newsworthy Event (Lead) -> Elaboration / Background Events -> Sources (Kutipan Presiden).",
                "Latihan mengubah kalimat langsung (Direct Speech) menjadi kalimat tidak langsung (Reported Speech) menggunakan reporting verbs (stated, explained, announced)."
            ])
        ]),
        ("PERTEMUAN 3 (2 JP): JCoT - Menulis Naskah Berita Kelompok", [
            ("Langkah Inti", [
                "Bekerja berpasangan memilih satu peristiwa aktual di sekolah atau daerah (misal: kemenangan olimpiade sains, festival seni, pameran inovasi hijau).",
                "Menyusun naskah berita siaran 120-150 kata menggunakan format piramida terbalik (Inverted Pyramid).",
                "Saling menyunting draf antarkelompok (Peer-editing fokus pada akurasi 5W+1H)."
            ])
        ]),
        ("PERTEMUAN 4 (2 JP): Latihan News Casting & Modulasi Suara", [
            ("Langkah Inti", [
                "Workshop teknik pembaca berita TV: postur tubuh tegak, kontak mata ke kamera, intonasi jeda berita, modulasi nada suara.",
                "Latihan membaca naskah berita di depan cermin atau kamera ponsel berpasangan (drill simulasi studio TV)."
            ])
        ]),
        ("PERTEMUAN 5 (2 JP): Asesmen Sumatif Unjuk Kerja - News Anchor Studio Showcase", [
            ("Langkah Inti", [
                "Penampilan rekaman siaran berita TV (Live News Anchor Showcase) di depan kelas.",
                "Penilaian kinerja menggunakan Rubrik Pembaca Berita Formal (Fluency, Pronunciation, Professional Presence, Content Accuracy).",
                "Refleksi akhir bab dan umpan balik konstruktif dari guru dan rekan sekelas."
            ])
        ])
    ]
    for m_title, phases in scenarios_2:
        add_sub_header(doc, m_title)
        for p_name, steps in phases:
            for s in steps:
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Inches(0.2)
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(2)
                p.add_run("• ").bold = True
                p.add_run(s).font.name = "Segoe UI"

    add_section_header(doc, "III. RENCANA ASESMEN")
    as_rows_2 = [
        ["1. Diagnostik", "Tes Menjodohkan", "Menjodohkan Headline dengan Lead Berita yang Tepat", "Pertemuan 1"],
        ["2. Formatif", "Lembar Kerja Siswa", "Analisis 5W+1H & Latihan Reported Speech", "Pertemuan 2 & 3"],
        ["3. Sumatif", "Kinerja Lisan (Performance)", "Simulasi News Anchor 1-2 Menit (Dinilai Rubrik)", "Pertemuan 5"]
    ]
    docx_helper.add_styled_table(doc, iktp_headers[:2] + ["Waktu Pelaksanaan"], [[r[0], r[1] + " - " + r[2], r[3]] for r in as_rows_2], [Inches(1.5), Inches(3.8), Inches(1.2)])

    fpath = os.path.join(TARGET_DIR, "Modul_Ajar_TP_12.W.2_News_Item_Text.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# MODUL 3: TP 12.W.3 - REVIEW TEXT (CREATIVE PRODUCTS & HERITAGE)
# ==============================================================================
def create_modul_3():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "MODUL AJAR KURIKULUM MERDEKA - BAHASA INGGRIS KELAS XII SMA",
        "Modul 3: Review Text (Promoting Indonesian Creative Products & Cultural Heritage)"
    )
    info = {
        "Satuan Pendidikan": "Sekolah Menengah Atas (SMA)",
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XII / Semester 5 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 12.W.3 - Mengevaluasi dan menyusun ulasan kritis (Review Text) terhadap produk kreatif nusantara, kriya lokal, dan kuliner khas daerah secara multimodal.",
        "Alokasi Waktu": "10 Jam Pelajaran (5 Pertemuan @ 2 JP = 45 menit/JP)",
        "Target Peserta Didik": "Peserta Didik Reguler (32 – 36 Siswa)",
        "Model Pembelajaran": "Genre-Based Approach (GBA) & Digital Publishing Project",
        "Profil Pelajar Pancasila": "Berkebinekaan Global, Kreatif, Cinta Tanah Air",
        "Sarana & Prasarana": "Katalog Produk Kemenparekraf, Foto Produk Kriya, Blog/Media Sosial Sekolah"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_section_header(doc, "I. KOMPONEN INTI")
    iktp_rows = [
        ["IKTP 3.1", "Menganalisis struktur teks ulasan produk (Orientation, Interpretative Recount, Evaluation, Evaluative Summary).", "C4 (Analisis)"],
        ["IKTP 3.2", "Mengidentifikasi kosakata evaluasi estetika bernada positif dan konstruktif (intricate, breathtaking, painstaking).", "C2 (Pemahaman)"],
        ["IKTP 3.3", "Menulis artikel ulasan produk kreatif lokal nusantara (200-250 kata) disertai penilaian bintang dan tips perawatan.", "C6 (Kreasi)"]
    ]
    docx_helper.add_styled_table(doc, ["Kode IKTP", "Indikator Ketercapaian Tujuan Pembelajaran", "Ranah"], iktp_rows, [Inches(1.2), Inches(4.3), Inches(1.0)])

    add_section_header(doc, "II. LANGKAH-LANGKAH PEMBELAJARAN (5 PERTEMUAN)")
    steps_3 = [
        "Pertemuan 1 (BKoF): Apresiasi Produk Ekonomi Kreatif Nusantara (Batik Tulis, Tenun Ikat, Kriya Rotan, Kopi Spesialti). Brainstorming produk unggulan daerah siswa.",
        "Pertemuan 2 (MoT): Membedah Teks Resensi Otentik Batik Tulis Indonesia (Standar UNESCO) & Mempelajari Pola Rating Bintang serta bahasa evaluasi santun.",
        "Pertemuan 3 (JCoT): Berkelompok menyusun draf ulasan produk mebel rotan ramah lingkungan atau kopi khas daerah menggunakan matriks evaluasi.",
        "Pertemuan 4 (ICoT): Menulis mandiri draf artikel ulasan produk kreatif lokal pilihan siswa dilengkapi foto estetika pendukung.",
        "Pertemuan 5 (Finalisasi & Sumatif): Penyuntingan antarteman, pemberian skor bintang, dan publikasi artikel di blog atau mading sekolah."
    ]
    for s in steps_3:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(3)
        p.add_run("• ").bold = True
        p.add_run(s).font.name = "Segoe UI"

    fpath = os.path.join(TARGET_DIR, "Modul_Ajar_TP_12.W.3_Review_Text.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# MODUL 4: TP 12.W.4 - JOB APPLICATION LETTER & ATS CV
# ==============================================================================
def create_modul_4():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "MODUL AJAR KURIKULUM MERDEKA - BAHASA INGGRIS KELAS XII SMA",
        "Modul 4: Job Application Letter & Modern ATS-Compliant Curriculum Vitae (CV)"
    )
    info = {
        "Satuan Pendidikan": "Sekolah Menengah Atas (SMA)",
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XII / Semester 6 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 12.W.4 - Menyusun surat lamaran pekerjaan formal (Job Application Letter) dan daftar riwayat hidup modern (Curriculum Vitae / Resume) sesuai standar profesional dunia kerja modern.",
        "Alokasi Waktu": "10 Jam Pelajaran (5 Pertemuan @ 2 JP = 45 menit/JP)",
        "Target Peserta Didik": "Peserta Didik Reguler (32 – 36 Siswa)",
        "Model Pembelajaran": "Genre-Based Approach & Career Portfolio Workshop",
        "Profil Pelajar Pancasila": "Mandiri, Bernalar Kritis, Kreatif",
        "Sarana & Prasarana": "Iklan Lowongan Kerja Otentik, Templat CV ATS-Friendly, Komputer/Laptop Lab"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_section_header(doc, "I. KOMPONEN INTI")
    iktp_rows = [
        ["IKTP 4.1", "Menganalisis kualifikasi dan persyaratan pada iklan lowongan pekerjaan (job vacancy) otentik.", "C4 (Analisis)"],
        ["IKTP 4.2", "Menyusun surat lamaran pekerjaan formal (Cover Letter) berformat bisnis baku bebas kesalahan tata bahasa.", "C6 (Kreasi)"],
        ["IKTP 4.3", "Merancang Curriculum Vitae (CV) 1 halaman yang terstruktur dan ramah sistem pemindaian komputer (ATS-friendly).", "C6 (Kreasi)"]
    ]
    docx_helper.add_styled_table(doc, ["Kode IKTP", "Indikator Ketercapaian Tujuan Pembelajaran", "Ranah"], iktp_rows, [Inches(1.2), Inches(4.3), Inches(1.0)])

    add_section_header(doc, "II. LANGKAH-LANGKAH PEMBELAJARAN (5 PERTEMUAN)")
    steps_4 = [
        "Pertemuan 1 (BKoF): Menganalisis tren pasar kerja dan membedah iklan lowongan kerja digital (Junior Digital Content Creator).",
        "Pertemuan 2 (MoT): Membedah anatomi surat lamaran kerja formal (Letterhead, Salutation, 3 Body Paragraphs, Professional Sign-off) & Formula Action Verbs.",
        "Pertemuan 3 (MoT CV): Membedah format Curriculum Vitae modern ATS-friendly (tanpa tabel rumit, fokus data kuantitatif terukur).",
        "Pertemuan 4 (JCoT & ICoT): Menyusun draf berkas dossier lamaran kerja lengkap disesuaikan dengan minat dan keahlian siswa.",
        "Pertemuan 5 (Asesmen Sumatif): Peer-editing dossier lamaran kerja menggunakan panduan ceklis profesionalitas & penilaian portofolio karier."
    ]
    for s in steps_4:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(3)
        p.add_run("• ").bold = True
        p.add_run(s).font.name = "Segoe UI"

    fpath = os.path.join(TARGET_DIR, "Modul_Ajar_TP_12.W.4_Job_Application_and_CV.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# MODUL 5: TP 12.W.5 - JOB INTERVIEW SIMULATION (STAR METHOD)
# ==============================================================================
def create_modul_5():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "MODUL AJAR KURIKULUM MERDEKA - BAHASA INGGRIS KELAS XII SMA",
        "Modul 5: Job Interview Simulation & Professional Oral Presentation (STAR Method)"
    )
    info = {
        "Satuan Pendidikan": "Sekolah Menengah Atas (SMA)",
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XII / Semester 6 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 12.W.5 - Mempraktikkan percakapan wawancara kerja formal (Job Interview Simulation) dan keterampilan presentasi diri (Elevator Pitch) secara percaya diri, santun, dan runtut menggunakan metode STAR.",
        "Alokasi Waktu": "8 Jam Pelajaran (4 Pertemuan @ 2 JP = 45 menit/JP)",
        "Target Peserta Didik": "Peserta Didik Reguler (32 – 36 Siswa)",
        "Model Pembelajaran": "Role-Play Simulation & Behavioral Interview Drill",
        "Profil Pelajar Pancasila": "Mandiri, Bernalar Kritis, Berakhlak Mulia",
        "Sarana & Prasarana": "Daftar Pertanyaan Wawancara Standar, Ruang Simulasi Meja Wawancara, Lembar Observasi"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_section_header(doc, "I. KOMPONEN INTI")
    iktp_rows = [
        ["IKTP 5.1", "Menyusun teks perkenalan diri profesional 60 detik (Elevator Pitch) yang memikat dan terstruktur.", "C6 (Kreasi)"],
        ["IKTP 5.2", "Menerapkan kerangka kerja STAR (Situation, Task, Action, Result) dalam merespons pertanyaan situasional wawancara.", "C3 (Aplikasi)"],
        ["IKTP 5.3", "Mendemonstrasikan simulasi wawancara kerja formal dengan bahasa tubuh profesional dan modulasi vokal santun.", "P3 (Kinerja Lisan)"]
    ]
    docx_helper.add_styled_table(doc, ["Kode IKTP", "Indikator Ketercapaian Tujuan Pembelajaran", "Ranah"], iktp_rows, [Inches(1.2), Inches(4.3), Inches(1.0)])

    add_section_header(doc, "II. LANGKAH-LANGKAH PEMBELAJARAN (4 PERTEMUAN)")
    steps_5 = [
        "Pertemuan 1: Bedah Strategi Wawancara Kerja & Penyusunan 60-Second Elevator Pitch (Who, Accomplishment, Direction).",
        "Pertemuan 2: Pendalaman Metode STAR (Situation, Task, Action, Result) untuk menjawab pertanyaan perilaku kepemimpinan & problem-solving.",
        "Pertemuan 3: Mock Interview Drill berpasangan (Siswa bergantian menjadi pewawancara dan pelamar kerja menggunakan checklist observasi).",
        "Pertemuan 4 (Asesmen Sumatif): Ujian Praktik Simulasi Wawancara Kerja Lisan di hadapan guru penguji dengan penilaian rubrik profesional."
    ]
    for s in steps_5:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(3)
        p.add_run("• ").bold = True
        p.add_run(s).font.name = "Segoe UI"

    fpath = os.path.join(TARGET_DIR, "Modul_Ajar_TP_12.W.5_Job_Interview_Simulation.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# MODUL 6: TP 12.W.6 - CAPTIONS, DATA GRAPHS & MULTIMODAL INFOGRAPHICS
# ==============================================================================
def create_modul_6():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "MODUL AJAR KURIKULUM MERDEKA - BAHASA INGGRIS KELAS XII SMA",
        "Modul 6: Photo Captions, Data Graphs & Multimodal Infographics"
    )
    info = {
        "Satuan Pendidikan": "Sekolah Menengah Atas (SMA)",
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XII / Semester 6 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 12.W.6 - Menafsirkan makna dan pesan tersirat dalam takarir (Caption), grafik statistik, dan infografis multimodal, serta menyajikan ringkasan visual ke dalam bahasa Inggris komunikatif.",
        "Alokasi Waktu": "6 Jam Pelajaran (3 Pertemuan @ 2 JP = 45 menit/JP)",
        "Target Peserta Didik": "Peserta Didik Reguler (32 – 36 Siswa)",
        "Model Pembelajaran": "Multimodal Literacy Project & Visual Data Analysis",
        "Profil Pelajar Pancasila": "Bernalar Kritis, Kreatif, Mandiri",
        "Sarana & Prasarana": "Foto Jurnalistik World Press Photo, Grafik Tren BPS/World Bank, Poster Infografis WHO"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_section_header(doc, "I. KOMPONEN INTI")
    iktp_rows = [
        ["IKTP 6.1", "Menganalisis kaidah penulisan takarir foto jurnalistik (Sentence 1: Present, Sentence 2: Past/Context, Photo Credit).", "C4 (Analisis)"],
        ["IKTP 6.2", "Membaca dan mendeskripsikan tren kenaikan, penurunan, dan fluktuasi pada grafik batang dan diagram garis.", "C3 (Aplikasi)"],
        ["IKTP 6.3", "Menyusun takarir foto informatif dan ringkasan visual infografis kesehatan/sosial secara multimodal.", "C6 (Kreasi)"]
    ]
    docx_helper.add_styled_table(doc, ["Kode IKTP", "Indikator Ketercapaian Tujuan Pembelajaran", "Ranah"], iktp_rows, [Inches(1.2), Inches(4.3), Inches(1.0)])

    add_section_header(doc, "II. LANGKAH-LANGKAH PEMBELAJARAN (3 PERTEMUAN)")
    steps_6 = [
        "Pertemuan 1: Bedah Takarir Foto Jurnalistik (Cutline Guidelines World Press Photo) & Latihan menulis caption foto aksi sosial sekolah.",
        "Pertemuan 2: Membaca Grafik Data Statistik (Trends Vocabulary: rose sharply, declined, leveled off, accounted for) & Menulis ringkasan data.",
        "Pertemuan 3 (Asesmen Sumatif): Analisis Poster Infografis Multimodal WHO (5 Pilar Kesehatan Mental Remaja) & Portofolio Akhir Jenjang."
    ]
    for s in steps_6:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.2)
        p.paragraph_format.space_after = Pt(3)
        p.add_run("• ").bold = True
        p.add_run(s).font.name = "Segoe UI"

    fpath = os.path.join(TARGET_DIR, "Modul_Ajar_TP_12.W.6_Captions_and_Infographics.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


if __name__ == "__main__":
    print("Generating complete Modul Ajar files for 5-Modul Ajar XII SMA (Wajib)...")
    create_modul_1()
    create_modul_2()
    create_modul_3()
    create_modul_4()
    create_modul_5()
    create_modul_6()
    print("All 6 Modul Ajar (.docx) for Kelas XII SMA Wajib successfully created!")
