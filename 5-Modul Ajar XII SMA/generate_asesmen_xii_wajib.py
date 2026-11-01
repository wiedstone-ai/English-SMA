# -*- coding: utf-8 -*-
"""
Generator script for complete Assessment Instruments (.docx) in 5-Modul Ajar XII SMA (Wajib).
Covers TP 12.W.1 through TP 12.W.6 with:
- Constructive Alignment Matrix
- Bagian I: Asesmen Diagnostik (Non-kognitif & Kognitif)
- Bagian II: Asesmen Formatif (Lembar Observasi & Rubrik Proses)
- Bagian III: Asesmen Sumatif (Kisi-kisi, 10 Soal PG HOTS, 5 Soal Esai, Performance Task, Rubrik Analitik, Kunci Jawaban & Pedoman Penskoran)
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

def add_as_section(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(title)
    run.font.name = "Segoe UI"
    run.font.size = Pt(11.5)
    run.font.bold = True
    run.font.color.rgb = COLOR_NAVY

def add_as_sub(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(title)
    run.font.name = "Segoe UI"
    run.font.size = Pt(10)
    run.font.bold = True
    run.font.color.rgb = COLOR_ROYAL

# ==============================================================================
# ASESMEN 1: TP 12.W.1 - DISCUSSION TEXT
# ==============================================================================
def create_asesmen_1():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "PERANGKAT INSTRUMEN ASESMEN LENGKAP - BAHASA INGGRIS KELAS XII",
        "Modul 1: Discussion Text (Workplace Automation & AI Ethics) | TP 12.W.1"
    )

    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XII / Semester 5 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 12.W.1 - Menganalisis wacana teks diskusi berimbang mengenai disrupsi otomasi kerja masa depan, etika AI, dan keberlanjutan lingkungan.",
        "Prinsip Utama": "Constructive Alignment (IKTP <-> Teknik & Bentuk Asesmen)",
        "Target Kurikulum": "Kurikulum Merdeka 2025 (Puskurjar BSKAP)"
    }
    docx_helper.add_tp_info_box(doc, info)

    # MATRIKS ALIGNMENT
    add_as_section(doc, "I. MATRIKS KESELARASAN INDIKATOR DENGAN TEKNIK ASESMEN")
    al_headers = ["Kode IKTP", "Indikator Ketercapaian", "Teknik Asesmen", "Bentuk Instrumen"]
    al_rows = [
        ["IKTP 1.1", "Mengidentifikasi isu, argumen pro, dan kontra otomasi kerja.", "Tes Tertulis", "Pilihan Ganda & Isian LKPD"],
        ["IKTP 1.2", "Menerapkan konjungsi konsesif & penanda wacana kontras.", "Tes Tertulis / Kinerja", "Soal Analisis Kalimat & Rubrik Esai"],
        ["IKTP 1.3", "Mengevaluasi objektivitas dan bahasa berpagar (hedging).", "Tes Tertulis HOTS", "Soal Uraian Penalaran Kritis"],
        ["IKTP 1.4", "Menyusun draf teks diskusi berimbang dan rekomendasi solutif.", "Penilaian Kinerja Produk", "Esai Teks Diskusi (Rubrik Analitik)"]
    ]
    docx_helper.add_styled_table(doc, al_headers, al_rows, [Inches(1.2), Inches(3.0), Inches(1.3), Inches(1.5)])

    # BAGIAN 1: ASESMEN DIAGNOSTIK
    add_as_section(doc, "II. INSTRUMEN ASESMEN DIAGNOSTIK (AWAL PEMBELAJARAN)")
    add_as_sub(doc, "A. Angket Diagnostik Non-Kognitif (Kesiapan Belajar & Minat Karier)")
    p_d = doc.add_paragraph()
    p_d.paragraph_format.left_indent = Inches(0.2)
    p_d.add_run("1. Bagaimana perasaanmu ketika mendengar bahwa kecerdasan buatan (AI) dapat menggantikan beberapa profesi manusia? (Sangat Cemas / Waspada / Sangat Antusias)\n")
    p_d.add_run("2. Karier atau bidang studi apa yang kamu rencanakan setelah lulus SMA?\n")
    p_d.add_run("3. Gaya belajar apa yang paling membantumu memahami materi bahasa Inggris? (Membaca artikel / Diskusi kelompok / Menonton video edukasi)")

    add_as_sub(doc, "B. Tes Diagnostik Kognitif Awal (5 Soal Prasyarat)")
    p_cog = doc.add_paragraph()
    p_cog.paragraph_format.left_indent = Inches(0.2)
    p_cog.add_run("1. What is the main purpose of a Discussion Text compared to an Analytical Exposition Text?\n")
    p_cog.add_run("2. Combine these two sentences using 'Although': (A) AI tools can write essays quickly. (B) They lack human empathy.\n")
    p_cog.add_run("3. Which connector shows contrast: 'Furthermore', 'Consequently', or 'Conversely'?\n")
    p_cog.add_run("4. Identify the subject and verb in: 'Workplace automation eliminates routine repetitive tasks.'\n")
    p_cog.add_run("5. What does the word 'displacement' mean in the context of employment?")

    # BAGIAN 2: ASESMEN FORMATIF
    add_as_section(doc, "III. INSTRUMEN ASESMEN FORMATIF (PROSES PEMBELAJARAN)")
    add_as_sub(doc, "Lembar Observasi Kinerja Diskusi Kelompok & Kolaborasi (Skala 1 - 4)")
    obs_headers = ["No", "Nama Peserta Didik", "Keaktifan Diskusi", "Menghargai Pendapat", "Penggunaan Bahasa", "Total Skor", "Predikat"]
    obs_rows = [
        ["1", "Aditya Pratama", "4", "4", "3", "11", "Sangat Baik"],
        ["2", "Devina Anindya", "4", "4", "4", "12", "Sangat Baik"],
        ["3", "Farhan Ramadhan", "3", "3", "3", "9", "Baik"],
        ["4", "................................................", "....", "....", "....", "....", "..............."]
    ]
    docx_helper.add_styled_table(doc, obs_headers, obs_rows, [Inches(0.5), Inches(2.2), Inches(1.1), Inches(1.1), Inches(1.1), Inches(0.8), Inches(1.0)])

    # BAGIAN 3: ASESMEN SUMATIF
    add_as_section(doc, "IV. INSTRUMEN ASESMEN SUMATIF (AKHIR MODUL)")
    add_as_sub(doc, "A. Kisi-Kisi Soal Asesmen Sumatif")
    kisi_headers = ["No", "Materi Pokok / Indikator Soal", "Level Kognitif", "Bentuk Soal", "Nomor Soal"]
    kisi_rows = [
        ["1", "Menemukan ide pokok isu disrupsi otomasi kerja pada wacana teks.", "C2 (Pemahaman)", "Pilihan Ganda", "1, 2"],
        ["2", "Menganalisis argumen pendukung (pros) peningkatan produktivitas AI.", "C4 (Analisis)", "Pilihan Ganda", "3, 4"],
        ["3", "Menganalisis argumen penentang (cons) ancaman pengangguran.", "C4 (Analisis)", "Pilihan Ganda", "5, 6"],
        ["4", "Menentukan makna konjungsi konsesif dan transisi kontras (conversely).", "C3 (Aplikasi)", "Pilihan Ganda", "7, 8"],
        ["5", "Menyimpulkan rekomendasi jalan tengah yang seimbang (compromise).", "C5 (Evaluasi)", "Pilihan Ganda", "9, 10"],
        ["6", "Menguraikan dampak sosial jangka panjang otomasi terhadap lulusan SMA.", "C5 (Evaluasi)", "Uraian / Esai", "1, 2, 3"],
        ["7", "Menyusun draf teks diskusi orisinal 250-300 kata secara berimbang.", "C6 (Kreasi)", "Tugas Kinerja", "Tugas Produk"]
    ]
    docx_helper.add_styled_table(doc, kisi_headers, kisi_rows, [Inches(0.6), Inches(3.2), Inches(1.2), Inches(1.2), Inches(0.8)])

    add_as_sub(doc, "B. Naskah Soal Pilihan Ganda (10 Butir Soal HOTS)")
    pg_text = (
        "Stimulus Teks: Read the following passage to answer questions 1 – 5:\n"
        "'The rapid acceleration of artificial intelligence and robotic automation in modern workplaces has triggered a global debate about the future of employment. Supporters of workplace automation point out that technology relieves humans from repetitive and dangerous labor, freeing workers to focus on creative problem-solving. Furthermore, history shows that while technology eliminates some older jobs, it creates brand new industries. Conversely, opponents emphasize that the speed of disruption threatens millions of routine livelihoods before society can adapt. Older workers without specialized technical skills face long-term unemployment. Therefore, governments and corporations must collaborate to fund nationwide reskilling programs, while schools prepare students with human-centric skills that machines cannot easily replicate.'\n\n"
        "1. What is the primary issue addressed in the passage?\n"
        "   A. The history of the Industrial Revolution.\n"
        "   B. The controversy regarding the impact of AI automation on human employment.\n"
        "   C. Why robotic arms are cheaper to maintain than human factory workers.\n"
        "   D. The technical programming steps required to build generative chatbots.\n"
        "   E. Why schools should eliminate humanities and arts subjects.\n\n"
        "2. Which of the following is an argument in FAVOR of workplace automation?\n"
        "   A. It creates immediate long-term unemployment for older workers.\n"
        "   B. It relieves humans from hazardous and repetitive physical labor.\n"
        "   C. It reduces corporate investment in employee healthcare and training.\n"
        "   D. It guarantees that no older industries will ever disappear.\n"
        "   E. It makes human empathy and ethical judgment obsolete.\n\n"
        "3. What risk is highlighted by opponents regarding technological disruption?\n"
        "   A. Machines take too long to process office data and financial invoices.\n"
        "   B. Robotic machinery is completely impossible to operate in tropical climates.\n"
        "   C. Millions of routine jobs may disappear before society has time to adapt.\n"
        "   D. Computers will eventually run out of electrical battery power.\n"
        "   E. Corporations refuse to use cloud engineering and digital networks.\n\n"
        "4. The word 'Conversely' in the third sentence is used to indicate...\n"
        "   A. an additional supporting example.\n"
        "   B. a clear shift to an opposing point of view.\n"
        "   C. the chronological time of a past historical event.\n"
        "   D. an author's emotional apology for an error.\n"
        "   E. a numerical percentage calculation.\n\n"
        "5. What balanced recommendation is proposed in the conclusion?\n"
        "   A. An outright worldwide ban on all computers and automated robotic machinery.\n"
        "   B. Nationwide reskilling programs and cultivating uniquely human skills in education.\n"
        "   C. Allowing private tech corporations to eliminate employee benefits freely.\n"
        "   D. Requiring all high school graduates to become robotic assembly technicians.\n"
        "   E. Postponing high school graduation until machines take over all office duties.\n\n"
        "6. Which sentence uses a concessive clause correctly?\n"
        "   A. Although automation boosts productivity, it can cause short-term job displacement.\n"
        "   B. Despite it can cause job displacement, but automation boosts productivity.\n"
        "   C. Automation boosts productivity because it is despite displacement.\n"
        "   D. Although but automation causes displacement, productivity rises.\n"
        "   E. In spite of automation is fast, workers are still needed.\n\n"
        "7. In a hybrid working model, what is the main advantage of working from home?\n"
        "   A. Completely eliminating communication with team members.\n"
        "   B. Saving time and energy by cutting out stressful daily traffic commutes.\n"
        "   C. Allowing employees to ignore company work deadlines.\n"
        "   D. Replacing all collaborative office meetings with automated robots.\n"
        "   E. Ensuring that employees never have to learn digital computer tools.\n\n"
        "8. In contrast, what is a potential drawback of prolonged remote work?\n"
        "   A. Employees spend too much money on office parking tickets.\n"
        "   B. Increased opportunities for casual spontaneous hallway conversations.\n"
        "   C. Social isolation, loneliness, and blurred boundaries between work and rest.\n"
        "   D. Inability to use modern video conferencing platforms like Zoom.\n"
        "   E. Having too many healthy home-cooked meals every week.\n\n"
        "9. What makes human skills like 'empathy' and 'ethical judgment' irreplaceable by AI?\n"
        "   A. Machines can process emotional feelings much faster than humans.\n"
        "   B. AI algorithms lack authentic human consciousness, lived experience, and moral compassion.\n"
        "   C. Computer software is not allowed to connect to electrical power grids.\n"
        "   D. Humans are always mathematically more accurate than computer calculations.\n"
        "   E. Ethical laws strictly prohibit computers from calculating numbers.\n\n"
        "10. Which characteristic is ESSENTIAL for the conclusion of a formal Discussion Text?\n"
        "   A. Choosing one extreme side and mocking the opposing viewpoint.\n"
        "   B. Providing a balanced summary and proposing a fair, constructive compromise.\n"
        "   C. Introducing five brand new controversial topics that were never discussed.\n"
        "   D. Copying and pasting the entire opening paragraph without any changes.\n"
        "   E. Concluding with a fictional fairy-tale happy ending."
    )
    p_pg = doc.add_paragraph()
    p_pg.paragraph_format.left_indent = Inches(0.2)
    p_pg.add_run(pg_text)

    add_as_sub(doc, "C. Naskah Soal Uraian / Esai (5 Butir Soal HOTS)")
    p_es = doc.add_paragraph()
    p_es.paragraph_format.left_indent = Inches(0.2)
    p_es.add_run("1. Explain why a Discussion Text must maintain an objective and neutral tone rather than taking sides aggressively like a persuasive speech.\n")
    p_es.add_run("2. Based on your understanding, how can governments and schools prevent young graduates from being left behind by rapid AI automation?\n")
    p_es.add_run("3. Analyze the statement: 'Technology is neither inherently good nor bad; it depends on human stewardship.' Do you agree? Why or why not?\n")
    p_es.add_run("4. Differentiate between 'displacement of labor' and 'creation of new industries' during major technological revolutions.\n")
    p_es.add_run("5. Construct two complex sentences using 'In spite of' and 'Conversely' to discuss the pros and cons of using AI in education.")

    add_as_sub(doc, "D. Rubrik Penilaian Analitik Esai Teks Diskusi (Skala 1 - 4)")
    rub_headers = ["Kriteria Penilaian", "Skor 4 (Sangat Baik)", "Skor 3 (Baik)", "Skor 2 (Cukup)", "Skor 1 (Kurang)"]
    rub_rows = [
        ["1. Content & Arguments (Bobot 30%)", "Isu jelas; argumen pro dan kontra disajikan sangat seimbang didukung data nyata.", "Isu jelas; argumen pro dan kontra cukup seimbang dengan contoh relevan.", "Argumen kurang berimbang; salah satu sisi jauh lebih dominan.", "Tidak tampak keberimbangan isu; argumen tidak relevan."],
        ["2. Organization & Structure (Bobot 25%)", "Struktur sempurna (Issue -> Pro -> Con -> Solutive Recommendation); kohesif.", "Struktur lengkap; transisi antarbargumen mengalir cukup lancar.", "Ada bagian struktur yang terlewati atau penempatan argumen acak.", "Struktur tidak teratur; paragraf tidak terhubung dengan baik."],
        ["3. Language & Hedging (Bobot 25%)", "Pilihan kosakata formal kaya; konjungsi kontras dan hedging digunakan presisi.", "Kosakata tepat; konjungsi kontras digunakan dengan benar.", "Ada beberapa kesalahan penggunaan konjungsi kontras/tata bahasa.", "Banyak kesalahan tata bahasa yang mengaburkan makna wacana."],
        ["4. Mechanics (Bobot 20%)", "Bebas dari kesalahan ejaan, tanda baca, dan kapitalisasi.", "Hanya sedikit kesalahan ejaan minor yang tidak mengganggu pemahaman.", "Cukup banyak kesalahan ejaan dan tanda baca.", "Banyak kesalahan ejaan mendasar di setiap paragraf."]
    ]
    docx_helper.add_styled_table(doc, rub_headers, rub_rows, [Inches(1.5), Inches(1.5), Inches(1.5), Inches(1.5), Inches(1.5)])

    add_as_sub(doc, "E. Kunci Jawaban & Pedoman Penskoran")
    docx_helper.add_callout(
        doc,
        "Kunci Jawaban Pilihan Ganda & Konversi Nilai:",
        "Kunci Jawaban PG:\n"
        "1. B | 2. B | 3. C | 4. B | 5. B | 6. A | 7. B | 8. C | 9. B | 10. B\n\n"
        "Pedoman Penskoran:\n"
        "• Nilai Pilihan Ganda = (Jumlah Benar / 10) x 40 poin (Maksimal: 40)\n"
        "• Nilai Soal Esai = (Skor Perolehan / 20) x 20 poin (Maksimal: 20)\n"
        "• Nilai Tugas Kinerja Esai = (Skor Rubrik / 16) x 40 poin (Maksimal: 40)\n"
        "• NILAI AKHIR SUMATIF = Nilai PG + Nilai Esai + Nilai Tugas Kinerja (Total Maksimal: 100)",
        bg="DCFCE7", border_color="16A34A"
    )

    fpath = os.path.join(TARGET_DIR, "Instrumen_Asesmen_TP_12.W.1_Discussion_Text.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# ASESMEN 2: TP 12.W.2 - NEWS ITEM TEXT
# ==============================================================================
def create_asesmen_2():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "PERANGKAT INSTRUMEN ASESMEN LENGKAP - BAHASA INGGRIS KELAS XII",
        "Modul 2: News Item Text & Media Literacy | TP 12.W.2"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XII / Semester 5 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 12.W.2 - Mengidentifikasi fakta utama, latar belakang peristiwa, dan kutipan sumber berita dalam teks berita aktual (News Item), serta mempraktikkan pembacaan berita (News Anchor).",
        "Target Kurikulum": "Kurikulum Merdeka 2025"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_as_section(doc, "I. MATRIKS KESELARASAN ASESMEN")
    al_rows = [
        ["IKTP 2.1", "Mengidentifikasi 5W+1H pada Lead berita aktual.", "Tes Tertulis", "Pilihan Ganda & Analisis LKPD"],
        ["IKTP 2.2", "Menganalisis kutipan narasumber & reporting verbs.", "Tes Tertulis", "Soal Analisis Kalimat Tidak Langsung"],
        ["IKTP 2.3", "Menulis naskah berita siaran TV (120-150 kata).", "Penilaian Produk", "Naskah Berita Berformat Piramida Terbalik"],
        ["IKTP 2.4", "Menampilkan simulasi pembaca berita TV (News Anchor).", "Unjuk Kerja Lisan", "Rubrik News Casting Showcase"]
    ]
    docx_helper.add_styled_table(doc, ["Kode IKTP", "Indikator Ketercapaian", "Teknik Asesmen", "Bentuk Instrumen"], al_rows, [Inches(1.2), Inches(3.0), Inches(1.3), Inches(1.5)])

    add_as_section(doc, "II. KISI-KISI & SOAL ASESMEN SUMATIF")
    pg_news = (
        "Stimulus Teks Berita: 'JAKARTA — Indonesia officially launched Southeast Asia's first high-speed bullet train on Monday, marking a historic technological milestone. The bullet train, named Whoosh, connects Jakarta with Bandung. Operating at speeds of up to 350 km/h, the train cuts travel time from three hours to just 45 minutes. Powered by clean electricity, the railway features earthquake detection systems. \"This high-speed train is a symbol of Indonesia's modernization and a giant leap forward,\" said President Joko Widodo during the inauguration.'\n\n"
        "1. What is the newsworthy event reported in the passage?\n"
        "   A. The closing of traditional bus terminals in Bandung.\n"
        "   B. The official launch of Southeast Asia's first high-speed bullet train, Whoosh.\n"
        "   C. An earthquake that damaged electrical power grids in West Java.\n"
        "   D. A traffic strike along the Jakarta highway.\n"
        "   E. The invention of a new diesel locomotive.\n\n"
        "2. What significant travel time improvement is achieved by the Whoosh bullet train?\n"
        "   A. Travel time increased from 45 minutes to three hours.\n"
        "   B. Travel time was reduced from over three hours to just 45 minutes.\n"
        "   C. Trains will only operate once every three days.\n"
        "   D. Travel time remains completely unchanged.\n"
        "   E. Passengers must spend two days waiting at the station.\n\n"
        "3. In the quote by President Joko Widodo, what does he emphasize regarding the train?\n"
        "   A. It is too expensive for ordinary citizens to afford.\n"
        "   B. It is a symbol of modernization and a giant leap forward in mass transit.\n"
        "   C. It should only be used by foreign tourists visiting Jakarta.\n"
        "   D. It relies exclusively on burning heavy coal fuel.\n"
        "   E. It will replace all domestic airplane flights in Indonesia.\n\n"
        "4. Which of the following is a reporting verb commonly used to cite expert sources?\n"
        "   A. sprinted | B. stated | C. danced | D. smiled | E. kicked\n\n"
        "5. The inverted pyramid structure in news writing means that...\n"
        "   A. the most important facts (Lead) are presented at the beginning, followed by supporting details.\n"
        "   B. the main point of the article is kept secret until the very last sentence.\n"
        "   C. only fictional imaginative stories can be published.\n"
        "   D. articles must be printed in the shape of a triangle.\n"
        "   E. no quotes from eyewitnesses are allowed."
    )
    p_pn = doc.add_paragraph()
    p_pn.paragraph_format.left_indent = Inches(0.2)
    p_pn.add_run(pg_news)

    add_as_section(doc, "III. RUBRIK PENILAIAN PRAKTIK PEMBACA BERITA (NEWS ANCHOR)")
    rub_news = [
        ["1. Articulation & Pronunciation (30%)", "Pengucapan kosakata bahasa Inggris sangat jelas, fasih, dan akurat.", "Pengucapan jelas dengan sedikit kekeliruan minor yang tidak mengganggu.", "Beberapa kesalahan artikulasi yang membuat penonton harus menyimak ulang.", "Banyak kata diucapkan keliru sehingga sulit dipahami."],
        ["2. Voice Modulation & Pacing (25%)", "Intonasi dan tempo berita sangat tepat; jeda lugas layaknya presenter TV profesional.", "Intonasi berita tepat dan tempo pembacaan stabil.", "Intonasi cenderung monoton atau tempo membaca terlalu cepat/lambat.", "Membaca datar tanpa jeda dan terdengar gugup."],
        ["3. Eye Contact & Presence (25%)", "Kontak mata ke kamera sangat konsisten (85%+); postur tubuh tegap dan percaya diri.", "Kontak mata ke kamera cukup baik; postur tubuh tegak.", "Terlalu sering menunduk membaca naskah; kontak mata jarang.", "Hanya menatap naskah sepanjang pembacaan berita."],
        ["4. Script Quality & Accuracy (20%)", "Naskah memuat 5W+1H lengkap, kutipan narasumber, dan bahasa baku jurnalistik.", "Naskah memenuhi unsur 5W+1H dengan tata bahasa yang baik.", "Naskah kekurangan beberapa detail fakta penting.", "Naskah tidak memenuhi kaidah penulisan berita jurnalistik."]
    ]
    docx_helper.add_styled_table(doc, ["Kriteria Penilaian", "Skor 4", "Skor 3", "Skor 2", "Skor 1"], rub_news, [Inches(1.8), Inches(1.4), Inches(1.4), Inches(1.4), Inches(1.5)])

    fpath = os.path.join(TARGET_DIR, "Instrumen_Asesmen_TP_12.W.2_News_Item_Text.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# ASESMEN 3: TP 12.W.3 - REVIEW TEXT
# ==============================================================================
def create_asesmen_3():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "PERANGKAT INSTRUMEN ASESMEN LENGKAP - BAHASA INGGRIS KELAS XII",
        "Modul 3: Review Text (Indonesian Creative Heritage) | TP 12.W.3"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XII / Semester 5 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 12.W.3 - Mengevaluasi dan menyusun ulasan kritis (Review Text) terhadap produk kreatif nusantara, kriya lokal, dan warisan budaya untuk promosi ekonomi kreatif.",
        "Target Kurikulum": "Kurikulum Merdeka 2025"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_as_section(doc, "I. KISI-KISI & ASESMEN SUMATIF REVIEW TEXT")
    pg_rev = (
        "Stimulus: 'Indonesian Batik Tulis is a profound cultural masterpiece inscribed on UNESCO's Representative List. The artistic quality lies in its meticulous process. Artisans use a small brass pen called a canting filled with hot beeswax to trace intricate motifs onto fine cotton or silk. The fabric is repeatedly dipped into botanical dyes before the wax is gently melted away in boiling water. While genuine hand-drawn batik commands a premium price, its exceptional durability and ancestral dignity make it an extraordinary heirloom investment. Overall, it earns a flawless 5 out of 5 stars.'\n\n"
        "1. What is the primary purpose of the Review Text above?\n"
        "   A. To teach readers how to build a commercial fabric manufacturing factory.\n"
        "   B. To critically evaluate and promote the artistic quality and heritage value of Batik Tulis.\n"
        "   C. To describe the historical military conflicts of the kingdoms in Java.\n"
        "   D. To warn international tourists against buying Indonesian handicrafts.\n"
        "   E. To advertise modern chemical textile dyes to batik workshops.\n\n"
        "2. According to the review, what traditional tool is used by batik artisans to draw wax patterns?\n"
        "   A. A paint roller | B. A canting | C. An ink printer | D. An iron hammer | E. A wooden paddle\n\n"
        "3. How does the author balance the review by mentioning a minor limitation?\n"
        "   A. By stating that the batik dissolves immediately when washed.\n"
        "   B. By acknowledging that genuine hand-drawn batik commands a premium price.\n"
        "   C. By claiming that Indonesian batik is not recognized internationally.\n"
        "   D. By suggesting that buyers should only purchase cheap printed cloth.\n"
        "   E. By complaining about the natural aroma of botanical dyes.\n\n"
        "4. Which of the following words is an evaluative adjective expressing high aesthetic quality?\n"
        "   A. rectangular | B. intricate | C. wooden | D. synthetic | E. daily\n\n"
        "5. The final paragraph of a Review Text typically serves to...\n"
        "   A. summarize the overall evaluation and provide a clear recommendation or star rating.\n"
        "   B. list the names and telephone numbers of all factory employees.\n"
        "   C. apologize for writing about cultural topics.\n"
        "   D. tell a long personal childhood fairy tale.\n"
        "   E. argue against the importance of international tourism."
    )
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    p.add_run(pg_rev)

    add_as_section(doc, "II. RUBRIK PENILAIAN MENULIS RESENSI PRODUK KREATIF")
    rub_rev = [
        ["1. Content & Depth of Review (30%)", "Ulasan sangat mendalam; membedah keunikan bahan, estetika, fungsionalitas, dan nilai budaya.", "Ulasan baik; menjelaskan kelebihan produk dan proses pembuatannya.", "Ulasan cukup; hanya mendeskripsikan produk secara umum tanpa evaluasi mendalam.", "Ulasan dangkal; tidak memuat kriteria penilaian yang jelas."],
        ["2. Text Structure (25%)", "Struktur lengkap: Orientation, Interpretative Recount, Evaluation, Evaluative Summary & Star Rating.", "Struktur lengkap dan runtut.", "Ada bagian struktur yang hilang (misal: tidak ada rating atau rekomendasi).", "Struktur acak dan tidak mengikuti kaidah review text."],
        ["3. Evaluative Vocabulary (25%)", "Menggunakan variasi kosakata estetika dan apresiasi budaya yang kaya dan tepat.", "Menggunakan kosakata evaluasi yang sesuai konteks.", "Pengulangan kosakata evaluasi sederhana (hanya good, nice, beautiful).", "Pilihan kosakata sangat terbatas dan banyak kekeliruan."],
        ["4. Mechanics (20%)", "Tata bahasa baku, ejaan dan tanda baca tertata sangat rapi.", "Tata bahasa baik dengan sedikit kesalahan minor.", "Cukup banyak kesalahan ejaan dan tata bahasa.", "Banyak kesalahan tata bahasa yang mengganggu keterbacaan."]
    ]
    docx_helper.add_styled_table(doc, ["Aspek Penilaian", "Skor 4", "Skor 3", "Skor 2", "Skor 1"], rub_rev, [Inches(1.8), Inches(1.4), Inches(1.4), Inches(1.4), Inches(1.5)])

    fpath = os.path.join(TARGET_DIR, "Instrumen_Asesmen_TP_12.W.3_Review_Text.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# ASESMEN 4: TP 12.W.4 - JOB APPLICATION & CV
# ==============================================================================
def create_asesmen_4():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "PERANGKAT INSTRUMEN ASESMEN LENGKAP - BAHASA INGGRIS KELAS XII",
        "Modul 4: Job Application Letter & Modern ATS CV | TP 12.W.4"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XII / Semester 6 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 12.W.4 - Menyusun surat lamaran pekerjaan formal (Job Application Letter) dan daftar riwayat hidup modern (Curriculum Vitae / Resume) sesuai standar profesional dunia kerja.",
        "Target Kurikulum": "Kurikulum Merdeka 2025"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_as_section(doc, "I. KISI-KISI & ASESMEN SUMATIF COVER LETTER & CV")
    pg_job = (
        "Stimulus: 'March 15, 2025. Ms. Sarah Jenkins, HR Director, Nexus Media Group. Dear Ms. Jenkins, I am writing to express my enthusiastic interest in the Junior Digital Marketing Associate position advertised on your career portal. As a motivated high school graduate with a strong academic foundation in English communication and digital design, I am confident in making a valuable contribution. During my final year, I served as Head of Publications for the Student Council, managing our school Instagram and increasing student engagement by 40%. In addition, I possess strong collaborative skills and adapt quickly to fast-paced deadlines. Thank you for your consideration. I have enclosed my CV for your review. Sincerely, Devina Anindya.'\n\n"
        "1. What is the primary purpose of an opening paragraph in a formal cover letter?\n"
        "   A. To complain about the applicant's current high school teachers.\n"
        "   B. To state the specific job position being applied for and where the opening was found.\n"
        "   C. To demand an immediate high salary and company car.\n"
        "   D. To list all the applicant's childhood hobbies and sports trophies.\n"
        "   E. To ask the company for personal financial loans.\n\n"
        "2. What quantifiable achievement does Devina present to prove her marketing competence?\n"
        "   A. She won an international cooking tournament.\n"
        "   B. She increased student social media engagement by 40% as Head of Publications.\n"
        "   C. She memorized all English dictionary words.\n"
        "   D. She worked for twenty years in commercial television advertising.\n"
        "   E. She built a private website for selling shoes.\n\n"
        "3. Which salutation is most appropriate when writing to a female hiring manager named Sarah Jenkins?\n"
        "   A. Hey Sarah! | B. Dear Ms. Jenkins, | C. What's up Sarah? | D. To Whom It May Concern Sarah, | E. Hi Buddy,\n\n"
        "4. Why should a modern Curriculum Vitae (CV) be 'ATS-friendly'?\n"
        "   A. So that automated Applicant Tracking Systems can parse and index the candidate's qualifications smoothly.\n"
        "   B. So that the CV can be colored with neon paints.\n"
        "   C. Because employers only hire candidates who use comic book fonts.\n"
        "   D. To hide the applicant's real name from the employer.\n"
        "   E. So that the file size is larger than ten gigabytes.\n\n"
        "5. Which bullet point uses an impactful, professional Action Verb?\n"
        "   A. I was responsible for looking at pictures.\n"
        "   B. Coordinated a 5-member creative committee and designed 60+ promotional event banners.\n"
        "   C. We sometimes did some social media stuff on weekends.\n"
        "   D. Making flyers was something I liked doing.\n"
        "   E. My teacher told me to help the school council."
    )
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    p.add_run(pg_job)

    add_as_section(doc, "II. RUBRIK PENILAIAN DOSSIER LAMARAN KERJA (COVER LETTER + CV)")
    rub_job = [
        ["1. Format & Professional Layout (25%)", "Format bisnis resmi lengkap; CV 1 halaman terstruktur bersih, ATS-compliant, dan proporsional.", "Format surat dan CV lengkap dengan tata letak rapi.", "Ada bagian format penting yang terlewat (misal: tanpa kontak atau salutation).", "Format tidak teratur; menyerupai pesan obrolan informal."],
        ["2. Content Alignment & Relevance (30%)", "Kualifikasi diri sangat relevan dengan iklan lowongan kerja; didukung bukti prestasi kuantitatif.", "Kualifikasi diri relevan dan didukung contoh pengalaman nyata.", "Kualifikasi diri kurang spesifik dengan posisi yang dilamar.", "Pengalaman dan kualifikasi tidak relevan dengan lowongan."],
        ["3. Action Verbs & Tone (25%)", "Menggunakan action verbs terukur di setiap poin; nada surat percaya diri, santun, dan antusias.", "Menggunakan action verbs dengan nada profesional.", "Terlalu banyak kalimat pasif atau bahasa yang kurang percaya diri.", "Gaya bahasa terlalu santai atau tidak sopan."],
        ["4. Language Conventions (20%)", "Bebas dari kesalahan tata bahasa (zero grammatical/spelling errors).", "Hanya 1-2 kesalahan kecil yang tidak mengganggu.", "Beberapa kesalahan tata bahasa yang terlihat jelas.", "Banyak kesalahan tata bahasa dan ejaan yang menurunkan citra profesional."]
    ]
    docx_helper.add_styled_table(doc, ["Kriteria Penilaian", "Skor 4", "Skor 3", "Skor 2", "Skor 1"], rub_job, [Inches(1.8), Inches(1.4), Inches(1.4), Inches(1.4), Inches(1.5)])

    fpath = os.path.join(TARGET_DIR, "Instrumen_Asesmen_TP_12.W.4_Job_Application_and_CV.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# ASESMEN 5: TP 12.W.5 - JOB INTERVIEW SIMULATION
# ==============================================================================
def create_asesmen_5():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "PERANGKAT INSTRUMEN ASESMEN LENGKAP - BAHASA INGGRIS KELAS XII",
        "Modul 5: Job Interview Simulation (STAR Method) | TP 12.W.5"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XII / Semester 6 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 12.W.5 - Mempraktikkan percakapan wawancara kerja formal (Job Interview Simulation) dan keterampilan presentasi diri (Elevator Pitch) secara percaya diri, santun, dan runtut menggunakan metode STAR.",
        "Target Kurikulum": "Kurikulum Merdeka 2025"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_as_section(doc, "I. KISI-KISI & ASESMEN SUMATIF WAWANCARA KERJA")
    pg_int = (
        "1. In behavioral job interviews, what does the 'STAR' acronym stand for?\n"
        "   A. Strategy, Timing, Argument, Review\n"
        "   B. Situation, Task, Action, Result\n"
        "   C. Skills, Talents, Aspirations, Rewards\n"
        "   D. Salary, Training, Attendance, Retirement\n"
        "   E. Speaking, Testing, Analyzing, Recording\n\n"
        "2. When responding to the opening prompt 'Tell me about yourself,' what should a candidate focus on?\n"
        "   A. Their complete family tree and childhood medical history.\n"
        "   B. A concise 60-second summary of their identity, relevant achievements, and career aspirations.\n"
        "   C. A list of all their favorite pop music songs and video games.\n"
        "   D. Complaining about how difficult their previous school exams were.\n"
        "   E. Asking the interviewer for a personal loan.\n\n"
        "3. When asked: 'What is your greatest weakness?', what is the most constructive strategy?\n"
        "   A. Claiming that you are completely perfect and never make any mistakes.\n"
        "   B. Admitting an honest area for growth and explaining the proactive steps you take to improve it.\n"
        "   C. Telling the interviewer that you easily get angry at teammates.\n"
        "   D. Refusing to answer the question because it is too personal.\n"
        "   E. Stating that you hate waking up in the morning.\n\n"
        "4. Which non-verbal behavior demonstrates confidence and professionalism during an interview?\n"
        "   A. Slouching low in the chair and looking at your smartphone screen.\n"
        "   B. Maintaining steady eye contact, sitting upright, and smiling warmly.\n"
        "   C. Tapping your feet loudly and crossing your arms defensively.\n"
        "   D. Whispering in a very quiet voice so the panel cannot hear.\n"
        "   E. Chewing bubblegum and laughing at the interviewer.\n\n"
        "5. Why should a job candidate always prepare 1–2 thoughtful questions to ask the interviewer at the end?\n"
        "   A. To show genuine enthusiasm, curiosity, and serious interest in the role and company culture.\n"
        "   B. Because candidates are legally required to interrogate the company director.\n"
        "   C. To test whether the hiring manager knows how to speak English.\n"
        "   D. To delay the interview so the next candidate cannot enter.\n"
        "   E. Because the interview cannot finish without an argument."
    )
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    p.add_run(pg_int)

    add_as_section(doc, "II. RUBRIK PENILAIAN UJIAN PRAKTIK WAWANCARA KERJA FORMAL")
    rub_int = [
        ["1. Professional Executive Presence (25%)", "Postur tubuh tegak, kontak mata stabil, busana formal rapi, dan senyum bersahabat.", "Bahasa tubuh baik; kontak mata cukup stabil dan sopan.", "Tampak gugup; kontak mata sering teralihkan; postur tubuh kaku.", "Bahasa tubuh tidak sesuai standar formal (gelisah/bungkuk)."],
        ["2. Delivery of 60-Second Elevator Pitch (25%)", "Pengenalan diri sangat runtut (Who-What-Where), memikat, lancar, dan tepat waktu (60 dtk).", "Pengenalan diri lancar dan informatif.", "Pengenalan diri terlalu bertele-tele atau melebihi waktu 2 menit.", "Pengenalan diri sangat singkat (hanya nama) atau gugup terbata-bata."],
        ["3. Application of STAR Method (30%)", "Menjawab pertanyaan situasional dengan alur S-T-A-R sangat jelas didukung bukti hasil terukur.", "Menerapkan alur STAR dengan baik dalam menjawab pengalaman nyata.", "Hanya menjelaskan situasi tanpa menguraikan tindakan nyata (Action) atau hasil (Result).", "Jawaban abstrak tanpa contoh konkret (hanya teori)."],
        ["4. Inquiring Engagement (20%)", "Mengajukan pertanyaan cerdas dan relevan mengenai budaya kerja atau pengembangan diri tim.", "Mengajukan pertanyaan yang baik kepada pewawancara.", "Pertanyaan yang diajukan kurang relevan atau sekadar menanyakan gaji.", "Tidak memiliki pertanyaan sama sekali untuk pewawancara."]
    ]
    docx_helper.add_styled_table(doc, ["Aspek Penilaian", "Skor 4", "Skor 3", "Skor 2", "Skor 1"], rub_int, [Inches(1.8), Inches(1.4), Inches(1.4), Inches(1.4), Inches(1.5)])

    fpath = os.path.join(TARGET_DIR, "Instrumen_Asesmen_TP_12.W.5_Job_Interview_Simulation.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# ASESMEN 6: TP 12.W.6 - CAPTIONS & INFOGRAPHICS
# ==============================================================================
def create_asesmen_6():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "PERANGKAT INSTRUMEN ASESMEN LENGKAP - BAHASA INGGRIS KELAS XII",
        "Modul 6: Photo Captions & Multimodal Infographics | TP 12.W.6"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XII / Semester 6 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 12.W.6 - Menafsirkan makna dan pesan tersirat dalam takarir foto jurnalistik (Caption), membaca tren grafik data statistik, dan menyajikan ringkasan infografis multimodal secara komunikatif.",
        "Target Kurikulum": "Kurikulum Merdeka 2025"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_as_section(doc, "I. KISI-KISI & ASESMEN SUMATIF CAPTIONS & DATA GRAPHS")
    pg_cap = (
        "Stimulus: 'Indonesian environmental volunteers plant red mangrove seedlings along the eroding coastline of Muara Gembong, West Java, on Saturday, February 15, 2025. The community reforestation initiative has planted over 50,000 mangrove trees since 2022 to protect local fishing villages from rising tidal sea floods. (Photo: Antara News / Ahmad Subaidi)'\n\n"
        "1. In professional photojournalism, what is the grammatical tense used in the first sentence of a caption?\n"
        "   A. Simple Present Tense (to describe the immediate action visible in the frame)\n"
        "   B. Future Perfect Continuous Tense\n"
        "   C. Past Continuous Tense exclusively\n"
        "   D. Passive Subjunctive Tense\n"
        "   E. Imperative Command Tense\n\n"
        "2. What essential context does the second sentence of the caption provide?\n"
        "   A. The price of the digital camera used by the photographer.\n"
        "   B. The historical background and ecological purpose of planting 50,000 mangroves.\n"
        "   C. A list of all fish species living in the Pacific Ocean.\n"
        "   D. An advertisement for buying commercial beach resorts.\n"
        "   E. The weather forecast for next month.\n\n"
        "3. Which phrase best describes a statistical trend that rises very fast and dramatically?\n"
        "   A. dropped slightly | B. surged sharply | C. leveled off | D. remained completely stable | E. declined gradually\n\n"
        "4. When describing a pie chart where clean solar power represents 75% of the circle, we write:\n"
        "   A. Solar power accounted for the vast majority (three-quarters) of total energy.\n"
        "   B. Solar power was completely negligible.\n"
        "   C. Only one single person used solar electricity.\n"
        "   D. Coal power was three times higher than solar energy.\n"
        "   E. The chart shows no data at all.\n\n"
        "5. What makes an infographic 'multimodal'?\n"
        "   A. It uses only black-and-white text without any spaces.\n"
        "   B. It integrates visual illustrations, numerical data, icons, and concise textual explanations.\n"
        "   C. It can only be listened to using radio headphones.\n"
        "   D. It requires an electric battery to be read.\n"
        "   E. It has no title or author credit."
    )
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    p.add_run(pg_cap)

    add_as_section(doc, "II. RUBRIK PENILAIAN LITERASI MULTIMODAL (CAPTION & GRAFIK)")
    rub_cap = [
        ["1. Caption Structure & Grammar (30%)", "Struktur 2 kalimat sempurna (Kalimat 1: Present 5W1H; Kalimat 2: Past/Konteks; Credit Line).", "Struktur caption lengkap dengan tata bahasa baik.", "Caption hanya 1 kalimat atau tenses terbalik.", "Caption tidak memuat unsur fakta 5W1H yang jelas."],
        ["2. Data Trends Accuracy (30%)", "Menafsirkan pergerakan angka, kenaikan/penurunan grafik dengan kosakata tren yang presisi.", "Menafsirkan tren data statistik dengan benar.", "Ada kesalahan pembacaan persentase atau tren data.", "Salah menafsirkan arah tren data grafik."],
        ["3. Multimodal Synthesis (25%)", "Mampu mengaitkan pesan visual ikonik dengan ringkasan teks secara ringkas dan padat.", "Mampu merangkum pesan visual infografis dengan baik.", "Ringkasan terlalu panjang atau mengabaikan unsur visual.", "Tidak mampu menyintesis pesan visual infografis."],
        ["4. Mechanics (15%)", "Bebas dari kesalahan ejaan, tanda kurung kredit foto, dan tanda baca.", "Hanya sedikit kesalahan tanda baca minor.", "Beberapa kesalahan ejaan.", "Banyak kesalahan mekanika penulisan."]
    ]
    docx_helper.add_styled_table(doc, ["Kriteria Penilaian", "Skor 4", "Skor 3", "Skor 2", "Skor 1"], rub_cap, [Inches(1.8), Inches(1.4), Inches(1.4), Inches(1.4), Inches(1.5)])

    fpath = os.path.join(TARGET_DIR, "Instrumen_Asesmen_TP_12.W.6_Captions_and_Infographics.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


if __name__ == "__main__":
    print("Generating complete Assessment Instruments (.docx) for 5-Modul Ajar XII SMA (Wajib)...")
    create_asesmen_1()
    create_asesmen_2()
    create_asesmen_3()
    create_asesmen_4()
    create_asesmen_5()
    create_asesmen_6()
    print("All 6 Assessment Instruments (.docx) for Kelas XII SMA Wajib successfully created!")
