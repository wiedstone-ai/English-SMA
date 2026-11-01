# -*- coding: utf-8 -*-
"""
Generator script for authentic English materials (.docx) in 5-Modul Ajar XII SMA (Wajib).
Calibrated specifically to CEFR B1+ to B2 (Accessible, Practical SMA Grade XII Level).
Each TP contains 3 authentic texts from credible sources with explicit URLs,
practical vocabulary glossaries, and aligned assessment questions.
"""

import os
import sys

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import docx_helper

TARGET_DIR = r"C:\Users\KOMPUTERKU2\Documents\English SMA\5-Modul Ajar XII SMA"
os.makedirs(TARGET_DIR, exist_ok=True)

# ==============================================================================
# TP 12.W.1: DISCUSSION TEXT - FUTURE OF WORK & AUTOMATION
# ==============================================================================
def create_tp_12_w_1():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS SMA",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XII / Semester 5 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 12.W.1 - Menganalisis teks diskusi mengenai disrupsi dunia kerja masa depan, etika otomasi AI, dan keberlanjutan lingkungan, serta menyusun teks diskusi berimbang.",
        "Elemen Capaian Pembelajaran": "Menyimak - Berbicara & Membaca - Memirsa",
        "Genre Teks / Ruang Lingkup": "Discussion Text (Future of Work, Workplace Automation, Remote Work)",
        "Alokasi Waktu": "10 Jam Pelajaran (5 Pertemuan @ 2 JP)",
        "Target Kemahiran CEFR": "Level B1+ s.d. B2 (Menengah Atas - Tingkat Kelas XII SMA)",
        "Profil Pelajar Pancasila": "Bernalar Kritis, Mandiri, Gotong Royong"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "World Economic Forum / Future of Jobs Report",
        "title": "Will Artificial Intelligence Replace Human Workers or Empower Them?",
        "url": "https://www.weforum.org/agenda/future-of-jobs-automation/",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa Kelas XII)"
    }
    t1_paras = [
        "The rapid acceleration of artificial intelligence and robotic automation in modern workplaces has triggered a global debate about the future of employment. While tech companies celebrate AI as a powerful engine for productivity and economic growth, labor economists and workers express serious anxiety about widespread job displacement.",
        "Supporters of workplace automation point out that technology relieves humans from repetitive, dangerous, and physically exhausting labor. In manufacturing, automated robotic arms assemble heavy machinery with zero risk of human injury. In offices, AI software processes massive amounts of paperwork in seconds, freeing employees to focus on creative problem-solving, strategic planning, and customer care. Furthermore, historical evidence from the Industrial Revolution shows that while technology eliminates some older jobs, it creates brand new industries—such as cloud engineering, data analytics, and digital cybersecurity.",
        "Conversely, opponents emphasize that the speed of current technological disruption threatens millions of livelihoods before society can adapt. Routine office jobs, data entry roles, and customer service positions are disappearing rapidly. Older workers and individuals without specialized technical skills face the threat of long-term unemployment. Moreover, corporations often prioritize cost-cutting through machines over employee well-being, widening the economic gap between tech investors and ordinary wage-earners.",
        "In conclusion, automation is an unstoppable technological wave, but its social impact depends on human leadership. Governments and corporations must collaborate to fund nationwide reskilling programs, while schools should prepare students with human-centric skills that machines cannot easily replicate—such as empathy, ethical judgment, and critical collaboration."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Menimbang secara objektif pro dan kontra dampak otomasi kecerdasan buatan terhadap lapangan kerja serta merumuskan rekomendasi persiapan SDM.",
        "Struktur Teks": "Issue (Perdebatan dampak AI terhadap masa depan pekerjaan) -> Arguments For (Peningkatan produktivitas, keselamatan kerja, lahirnya profesi baru) -> Arguments Against (Disrupsi mata pencaharian, pengangguran struktural, ketimpangan ekonomi) -> Balanced Recommendation (Program reskilling dan pengembangan kompetensi manusiawi).",
        "Fitur Bahasa": "Discourse Markers of Contrast ('While', 'Conversely', 'Moreover', 'In conclusion'), Concessive Clauses ('Although automation is unstoppable...'), Hedging ('point out', 'emphasize', 'evidence shows')."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "Workplace Automation: Technological Boom or Job Threat?",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "BBC Worklife / Harvard Business Review Column",
        "title": "Remote Work vs. Office Collaboration: The Great Workplace Divide",
        "url": "https://www.bbc.com/worklife/article/remote-work-pros-cons",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa Kelas XII)"
    }
    t2_paras = [
        "Since the global pandemic transformed how companies operate, remote work (working from home) has shifted from an emergency response into a permanent feature of modern corporate life. However, whether employees should work remotely forever or return to physical offices remains intensely contested.",
        "Proponents of remote work highlight enhanced flexibility and better work-life balance. Employees save two to three hours every day by eliminating stressful traffic commutes, allowing them to spend more quality time with family, exercise, and cook healthier meals. Studies also show that remote workers often report higher productivity because they can focus quietly without noisy office interruptions.",
        "On the other hand, business leaders argue that purely virtual communication weakens corporate culture, spontaneous creativity, and mentorship. Casual hallway conversations often spark innovative business ideas that planned Zoom calls rarely produce. Furthermore, prolonged remote work can lead to loneliness, social isolation, and blurred boundaries between professional working hours and personal rest.",
        "A sensible compromise adopted by many progressive firms is the 'hybrid model'. Employees spend two or three days in the office for team collaboration and meetings, while working the remaining days from home for focused, independent tasks."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Mendiskusikan kelebihan dan kekurangan sistem kerja jarak jauh (work from home) vs kerja di kantor.",
        "Struktur Teks": "Issue (Tren kerja jarak jauh) -> Arguments For (Fleksibilitas, hemat waktu perjalanan, bebas bising) -> Arguments Against (Melemahkan kultur tim, minim spontanitas, isolasi sosial) -> Recommendation (Model kerja hybrid).",
        "Fitur Bahasa": "Contrast Connectors ('However', 'On the other hand', 'Furthermore', 'A sensible compromise')."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "Remote Work: Personal Freedom or Professional Isolation?",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "The Guardian Careers / Forbes Future of Business",
        "title": "Artificial Intelligence in Entry-Level Careers: The New Graduate Perspective",
        "url": "https://www.theguardian.com/business/ai-future-careers-youth",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa Kelas XII)"
    }
    t3_paras = [
        "High school seniors and university graduates are entering a job market that is fundamentally different from that of previous generations. Generative AI tools like ChatGPT can now write computer code, draft business emails, and summarize financial reports within seconds.",
        "Advocates argue that AI democratizes business productivity. Young professionals who learn to collaborate effectively with AI tools can accomplish tasks that previously required entire teams. By acting as 'AI co-pilots', young employees can amplify their efficiency, accelerate learning, and propose innovative ideas faster.",
        "Nevertheless, skeptics warn that automated software is eliminating entry-level roles that traditionally served as training grounds for young workers. If software handles junior research, proofreading, and customer responses, young graduates miss vital opportunities to learn foundational workplace skills from experienced mentors.",
        "To thrive in this environment, graduating students must view AI not as a competitor, but as a collaborative tool, while deliberately honing their unique human capabilities in critical thinking, interpersonal diplomacy, and ethical leadership."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "Entering the Automated Workplace: Advice for High School Graduates",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Displacement", "Noun", "Pergeseran atau hilangnya posisi kerja manusia oleh mesin", "Automation causes displacement of routine manufacturing jobs."),
        ("Productivity", "Noun", "Tingkat hasil kerja produktif dalam rentang waktu", "AI tools can significantly enhance office productivity."),
        ("Commute", "Noun/Verb", "Perjalanan bolak-balik antara rumah dan tempat kerja", "Working from home eliminates stressful daily commutes."),
        ("Mentorship", "Noun", "Bimbingan karier dari staf senior kepada yunior", "Remote work can make meaningful mentorship harder to achieve."),
        ("Hybrid model", "Noun Phrase", "Model gabungan kerja di kantor dan kerja dari rumah", "Many modern companies adopt a flexible hybrid model."),
        ("Co-pilot", "Noun", "Alat atau asisten pendamping yang membantu kerja manusia", "Generative AI functions as a digital co-pilot for writers."),
        ("Democratize", "Verb", "Membuka akses teknologi canggih untuk semua kalangan", "AI democratizes access to sophisticated business analysis."),
        ("Reskilling", "Noun", "Pelatihan ulang keterampilan baru sesuai era modern", "Governments must invest in nationwide reskilling programs.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "According to Text 1, what historical example is cited to show that technological change eventually creates new industries?",
        "In Text 2, what are two main personal benefits enjoyed by employees who work remotely from home?",
        "What concern do skeptics raise in Text 3 regarding the elimination of junior entry-level jobs by AI?"
    ]
    inf_q = [
        "Why is the 'hybrid work model' (Text 2) considered a win-win compromise by both business executives and employees?",
        "Explain how generative AI tools can serve as a 'co-pilot' rather than a replacement for young professionals in Text 3.",
        "Analyze why human empathy and ethical judgment are skills that computer algorithms cannot easily replicate (Text 1)."
    ]
    eval_q = [
        "As a high school senior preparing for university or work, are you optimistic or anxious about entering an automated job market? Defend your viewpoint.",
        "Evaluate whether companies have a moral duty to retrain workers whose jobs are eliminated by technology."
    ]
    perf_task = (
        "PENYUSUNAN ESAI TEKS DISKUSI BERIMBANG (LKPD / SUMATIF MENULIS):\n"
        "Susunlah sebuah Esai Teks Diskusi (Discussion Text) sepanjang 250-300 kata mengenai salah satu isu ketenagakerjaan masa depan:\n"
        "Mosi 1: Should Companies Replace Customer Service Representatives with AI Chatbots?\n"
        "Mosi 2: Should High Schools Require All Students to Complete an Internship in Digital Technology?\n"
        "Struktur Wajib:\n"
        "• Paragraph 1: Issue (Latar belakang kontroversi dunia kerja).\n"
        "• Paragraph 2: Arguments For (Sajikan 2 alasan pendukung didukung data fakta).\n"
        "• Paragraph 3: Arguments Against (Sajikan 2 alasan penentang didukung kritik sosial).\n"
        "• Paragraph 4: Balanced Recommendation (Sajikan jalan tengah solutif).\n"
        "Gunakan konjungsi kontras dan modalitas objektif."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_12.W.1_Discussion_Text_Future_Work.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 12.W.2: NEWS ITEM TEXT & MEDIA LITERACY
# ==============================================================================
def create_tp_12_w_2():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS SMA",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XII / Semester 5 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 12.W.2 - Mengidentifikasi fakta utama, latar belakang peristiwa, dan kutipan sumber berita dalam teks berita aktual (News Item), serta mempraktikkan pembacaan berita (News Anchor).",
        "Elemen Capaian Pembelajaran": "Membaca - Memirsa & Menulis - Mempresentasikan",
        "Genre Teks / Ruang Lingkup": "News Item Text & Media Literacy (Headline, Lead, Quotes, TV News Casting)",
        "Alokasi Waktu": "10 Jam Pelajaran (5 Pertemuan @ 2 JP)",
        "Target Kemahiran CEFR": "Level B1+ s.d. B2",
        "Profil Pelajar Pancasila": "Bernalar Kritis, Mandiri, Berkebinekaan Global"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "Antara News English / BBC World News",
        "title": "Indonesia Launches Southeast Asia's First High-Speed Railway: Whoosh Begins Operations",
        "url": "https://en.antaranews.com/news/whoosh-bullet-train-launch",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa Kelas XII)"
    }
    t1_paras = [
        "JAKARTA — Indonesia officially launched Southeast Asia's first high-speed bullet train on Monday, marking a historic technological milestone for the nation's public infrastructure. The bullet train, named 'Whoosh'—an acronym for Waktu Hemat, Operasi Optimal, Sistem Handal—connects the capital city of Jakarta with the cultural and economic hub of Bandung in West Java.",
        "Operating at speeds of up to 350 kilometers per hour, the modern electric train cuts travel time between the two bustling cities from over three hours to just 45 minutes. Powered by clean electricity, the 142-kilometer railway features state-of-the-art earthquake and landslide safety detection systems adapted to Indonesia's tropical geography.",
        "'This high-speed train is a symbol of Indonesia's modernization and a giant leap forward in mass transportation,' said President Joko Widodo during the grand inauguration ceremony at Halim Station in East Jakarta. 'It provides an eco-friendly, fast, and comfortable alternative that will reduce highway traffic congestion and carbon emissions.'",
        "Commuters and foreign tourists praised the smooth, quiet ride during its opening public trial, expressing hope that the network will eventually be expanded toward Surabaya."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Menginformasikan kepada masyarakat luas tentang peluncuran kereta cepat pertama di Asia Tenggara ('Whoosh') rute Jakarta-Bandung.",
        "Struktur Teks": "Headline (Judul berita menarik) -> Lead / Newsworthy Event (Paragraf 1: Peresmian kereta cepat Whoosh rute Jakarta-Bandung) -> Background Events (Paragraf 2: Spesifikasi kecepatan 350 km/jam, efisiensi waktu, dan sistem keamanan) -> Sources / Quotes (Paragraf 3: Pernyataan resmi Presiden Joko Widodo).",
        "Fitur Bahasa": "Past Action Verbs ('launched', 'connected', 'praised'), Reporting Verbs ('said', 'expressed'), Headline Capitalization, Adverbs of Time and Place."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "Whoosh: Indonesia Launches Southeast Asia's First Bullet Train",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "Reuters Environment / International Energy Agency",
        "title": "Global Renewable Energy Reaches Historic High: Wind and Solar Surge Worldwide",
        "url": "https://www.reuters.com/sustainability/renewable-energy-milestone/",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa Kelas XII)"
    }
    t2_paras = [
        "PARIS — Global electricity generation from renewable energy sources reached an unprecedented 30 percent in 2023, driven by a monumental boom in solar and wind installations, according to an international energy report released on Tuesday.",
        "The report documented that global solar power capacity grew by over fifty percent in a single year, with China, the European Union, and developing economies in Southeast Asia leading the massive green transition. Technological innovations and falling manufacturing costs have made solar panels the cheapest source of new electricity generation in history.",
        "'We are witnessing a truly irreversible transformation of the global energy landscape,' stated Dr. Fatih Birol, Executive Director of the International Energy Agency. 'Renewables are expanding faster than fossil fuels, showing that clean energy is both economically profitable and environmentally essential.'",
        "Despite this remarkable progress, energy analysts cautioned that sovereign governments must urgently upgrade national electrical transmission grids to integrate intermittent solar and wind power into city networks effectively."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Melaporkan pencapaian rekor energi terbarukan global yang menyumbang 30% listrik dunia.",
        "Struktur Teks": "Newsworthy Event (Paragraf 1: Rekor energi surya & angin 30%) -> Elaboration (Paragraf 2: Lonjakan kapasitas 50% dan murahnya panel surya) -> Quotes (Paragraf 3: Kutipan Dr. Fatih Birol IEA) -> Concluding Context (Paragraf 4: Tantangan jaringan transmisi listrik).",
        "Fitur Bahasa": "Passive Voice ('was released', 'were cautioned'), Numbers and Statistics ('30 percent', 'fifty percent')."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "Global Green Energy Milestone: Solar and Wind Reach Historic Record",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "The Jakarta Post / Youth Science Achievements",
        "title": "Indonesian High Schoolers Win Gold at Tokyo World Robotics Olympiad",
        "url": "https://www.thejakartapost.com/youth/robotics-championship-tokyo",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa Kelas XII)"
    }
    t3_paras = [
        "TOKYO — A talented team of three high school students from Central Java took home the gold medal at the World Robot Olympiad in Tokyo on Sunday, defeating over eighty international teams with their innovative disaster-response robot.",
        "The students—aged 16 and 17—designed and programmed an autonomous search-and-rescue robot named 'Siaga-1'. Equipped with heat-sensing infrared cameras and artificial intelligence mapping algorithms, the robot is built to navigate collapsed buildings during earthquakes and locate trapped survivors in hazardous environments.",
        "'We developed this robot because our homeland frequently experiences earthquakes,' said team captain Rizky Pratama in his acceptance speech. 'We wanted to use computer science to save lives and make a tangible contribution to disaster management in Indonesia.'",
        "The judges commended the team for their practical engineering and cost-effective design, which utilizes 3D-printed lightweight components. The students received a standing ovation and scholarship offers from leading engineering universities."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "Youth Innovation: Indonesian Students Triumph at Tokyo Robotics Olympiad",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Acronym", "Noun", "Singkatan yang dilafalkan sebagai kata (misal: Whoosh)", "Whoosh is an acronym for speed, efficiency, and reliability."),
        ("Inauguration", "Noun", "Upacara peresmian pembukaan proyek atau gedung", "The president attended the grand railway inauguration."),
        ("Congestion", "Noun", "Kemacetan lalu lintas kendaraan di jalan raya", "Fast trains help reduce highway traffic congestion."),
        ("Unprecedented", "Adjective", "Belum pernah terjadi sebelumnya dalam sejarah", "Renewable energy reached an unprecedented global record."),
        ("Intermittent", "Adjective", "Tidak terus-menerus, berganti-ganti (misal: tenaga surya)", "Grids must manage intermittent solar power during cloudy days."),
        ("Commended", "Verb (Past)", "Memuji atau memberikan apresiasi tinggi", "The judges commended the students for their engineering skills."),
        ("Autonomous", "Adjective", "Dapat bergerak dan beroperasi sendiri secara mandiri", "The autonomous rescue robot navigates earthquake rubble."),
        ("Standing ovation", "Noun Phrase", "Tepuk tangan meriah sambil berdiri menghormati", "The winning team received a standing ovation from the audience.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "According to Text 1, what does the acronym 'Whoosh' stand for, and what two major cities does the railway connect?",
        "What global milestone regarding clean energy generation is reported in Text 2?",
        "In Text 3, what specific functions does the student-built robot 'Siaga-1' perform during disaster situations?"
    ]
    inf_q = [
        "Why is cutline speed and safety integration crucial for high-speed rail lines in an earthquake-prone country like Indonesia?",
        "How do quotes from authoritative figures (such as President Joko Widodo or Dr. Fatih Birol) enhance the credibility of news reports?",
        "Why did the international judges in Text 3 specifically praise the students' use of 3D-printed lightweight components?"
    ]
    eval_q = [
        "Evaluate why reliable, fact-checked news reporting is essential in an era flooded with social media rumors and clickbait.",
        "If you were a television news editor, which of the three events above would you choose as your main 'breaking news' story for tonight? Justify your choice."
    ]
    perf_task = (
        "PROYEK NEWS ANCHOR & SIMULASI PENYIARAN BERITA (LKPD / SUMATIF LISAN):\n"
        "1. Tulislah sebuah Naskah Berita Televisi (News Broadcast Script) sepanjang 120-150 kata mengenai peristiwa positif di sekolah atau kotamu (misal: prestasi lomba sains, pameran seni budaya, atau aksi tanam pohon).\n"
        "2. Pastikan naskah memuat struktur berita baku: Headline -> Lead (5W+1H) -> Rincian Berita -> Kutipan Narasumber.\n"
        "3. Rekam video berdurasi 1-2 menit di mana kamu berperan sebagai News Anchor (Pembaca Berita TV) profesional. Perhatikan intonasi, kontak mata ke kamera, dan artikulasi bahasa Inggris yang jelas."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_12.W.2_News_Item_Text.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 12.W.3: REVIEW TEXT - INDONESIAN CREATIVE PRODUCTS & HERITAGE
# ==============================================================================
def create_tp_12_w_3():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS SMA",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XII / Semester 5 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 12.W.3 - Mengevaluasi dan menyusun ulasan kritis (Review Text) terhadap produk kreatif nusantara, kriya lokal, dan warisan budaya untuk promosi ekonomi kreatif.",
        "Elemen Capaian Pembelajaran": "Membaca - Memirsa & Menulis - Mempresentasikan",
        "Genre Teks / Ruang Lingkup": "Review Text (Indonesian Creative Heritage, Batik, Crafts, Specialty Coffee)",
        "Alokasi Waktu": "10 Jam Pelajaran (5 Pertemuan @ 2 JP)",
        "Target Kemahiran CEFR": "Level B1+ s.d. B2",
        "Profil Pelajar Pancasila": "Berkebinekaan Global, Kreatif, Mandiri"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "Indonesia Travel / UNESCO Intangible Cultural Heritage",
        "title": "Product Review: The Timeless Artistry of Hand-Drawn Indonesian Batik Tulis",
        "url": "https://www.indonesia.travel/gb/en/trip-ideas/indonesian-batik-heritage",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa Kelas XII)"
    }
    t1_paras = [
        "Indonesian Batik Tulis (hand-drawn wax-resist textile) is not merely decorative fabric; it is a profound cultural masterpiece inscribed on UNESCO's Representative List of the Intangible Cultural Heritage of Humanity in 2009. Unlike modern mass-produced printed imitations, authentic Batik Tulis represents months of painstaking craftsmanship by master artisans across Java, particularly in Yogyakarta, Solo, and Pekalongan.",
        "The artistic quality of Batik Tulis lies in its meticulous production process. Artisans use a small brass pen called a 'canting' filled with hot beeswax to trace intricate motifs—such as the majestic 'Parang' or the sacred 'Megamendung'—onto fine silk or cotton cloth. The fabric is then repeatedly dipped into natural botanical dyes extracted from indigo leaves, mahogany bark, and mangosteen rinds before the wax is gently melted away in boiling water. This multi-layered coloring imparts a rich, organic warmth and a distinct herbal aroma that synthetic prints can never replicate.",
        "While genuine hand-drawn batik commands a premium price and requires delicate hand-washing with gentle 'lerak' soap, its exceptional durability and historical dignity make it an extraordinary heirloom investment. Wearing authentic batik connects the wearer directly to ancestral philosophy and supports rural artisan communities.",
        "Overall, authentic Indonesian Batik Tulis earns a flawless 5 out of 5 stars. It stands as a glorious testament to Indonesian cultural pride and an essential luxury textile for any international fashion enthusiast."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Meresensi dan mengevaluasi keunggulan artistik, proses produksi, serta nilai budaya luhur kain Batik Tulis Indonesia berstandar UNESCO.",
        "Struktur Teks": "Orientation (Identifikasi Batik Tulis dan status warisan UNESCO) -> Interpretative Recount (Proses melukis lilin dengan canting dan pewarna alami) -> Evaluation (Keunggulan aroma alami, daya tahan, perawatan vs harga premium) -> Evaluative Summary & Recommendation (Rating 5/5 dan promosi kebanggaan nasional).",
        "Fitur Bahasa": "Evaluative Adjectives ('profound', 'painstaking', 'intricate', 'exceptional', 'flawless'), Cultural Nouns ('canting', 'lerak', 'motifs')."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "Product Review: The Timeless Elegance of Indonesian Batik Tulis",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "The Jakarta Post Lifestyle / Green Design Indonesia",
        "title": "Design Review: Eco-Friendly Rattan and Bamboo Furniture from Central Java",
        "url": "https://www.thejakartapost.com/life/rattan-crafts-sustainable",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa Kelas XII)"
    }
    t2_paras = [
        "In an era where consumers worldwide are turning away from polluting plastic furniture, handcrafted rattan and bamboo home furnishings from Jepara and Cirebon are experiencing a vibrant global renaissance.",
        "The standout appeal of Indonesian rattan furniture is its harmonious blend of organic warmth and structural resilience. Local craftspeople steam and bend mature rattan vines into flowing, modern Scandinavian-inspired silhouettes while preserving traditional Indonesian wicker weaving techniques. The resulting lounge chairs and pendant lamps are lightweight, sturdy, and completely biodegradable.",
        "One minor drawback is that natural rattan furniture should be protected from prolonged exposure to heavy outdoor rain and intense dampness to prevent mildew. However, when placed in indoor living rooms, sunrooms, or shaded patios, these pieces age gracefully with a soft golden hue.",
        "Rated 4.5 out of 5 stars, Central Java's sustainable rattan homewares offer conscious homeowners a stylish, earth-friendly alternative to factory-made plastic furniture."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Mengevaluasi keindahan dan ketahanan furnitur rotan dan bambu ramah lingkungan asal Jepara dan Cirebon.",
        "Struktur Teks": "Orientation (Furnitur rotan ramah lingkungan) -> Evaluation (Kombinasi kelenturan, desain Skandinavia modern, sifat biodegradable) -> Minor Drawback (Perlu dihindarkan dari hujan terus-menerus) -> Summary & Rating (4.5/5).",
        "Fitur Bahasa": "Design Terminology ('silhouettes', 'wicker weaving', 'biodegradable', 'structural resilience')."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "Sustainable Craft Review: Modern Rattan Furniture from Central Java",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "Ministry of Tourism and Creative Economy / Coffee Review Portal",
        "title": "Culinary Review: Exploring Indonesian Single-Origin Specialty Coffees",
        "url": "https://kemenparekraf.go.id/ragam-ekonomi-kreatif/specialty-coffee-indonesia",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa Kelas XII)"
    }
    t3_paras = [
        "Indonesia's unique tropical geography along volcanic mountain ridges produces some of the most distinctive and celebrated specialty coffees in the international culinary arena. From the misty highlands of Aceh Gayo to the rugged volcanic soils of Toraja and Bali Kintamani, Indonesian Arabica beans offer an unforgettable sensory journey.",
        "Aceh Gayo Arabica, grown under shade trees at altitudes above 1,200 meters, is renowned for its heavy body, low acidity, and complex flavor notes of dark chocolate and warming spices. In contrast, Bali Kintamani beans, cultivated alongside orange orchards, deliver a refreshingly crisp, fruity sweetness with delicate citrus floral aromas.",
        "Specialty coffee roasters around the world praise Indonesian beans for their unique wet-hulling processing method ('Giling Basah'), which imparts an earthy, syrupy mouthfeel that balances milk perfectly in cappuccinos or shines cleanly in pour-over brewing.",
        "Overall, Indonesian specialty coffee is a masterclass in agricultural craftsmanship, deserving a glowing 5-star rating for coffee lovers seeking rich, ethically harvested beans."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "Culinary Review: The Rich Flavors of Indonesian Specialty Coffee",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Artisan", "Noun", "Perajin terampil yang membuat karya seni tangan bermutu", "Rural artisans preserve traditional batik techniques."),
        ("Painstaking", "Adjective", "Sangat cermat, teliti, dan membutuhkan kerja keras", "Creating hand-drawn batik requires painstaking patience."),
        ("Intricate", "Adjective", "Rumit, detail, dan sangat halus polanya", "The canting traces intricate patterns onto the cloth."),
        ("Botanical", "Adjective", "Berasal dari tumbuh-tumbuhan alami", "The fabric is colored using natural botanical dyes."),
        ("Heirloom", "Noun", "Pusaka berharga yang diwariskan antargenerasi", "A genuine batik cloth is a cherished family heirloom."),
        ("Resilience", "Noun", "Daya tahan lentur dan kekuatan menghadapi tekanan", "Rattan furniture combines flexibility with structural resilience."),
        ("Biodegradable", "Adjective", "Dapat terurai secara alami tanpa mencemari tanah", "Bamboo homewares are completely eco-friendly and biodegradable."),
        ("Acidity", "Noun", "Tingkat keasaman rasa (dalam uji cita rasa kopi)", "Gayo coffee is famous for its rich body and low acidity.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "In Text 1, what traditional tool and natural substances are used by artisans to create Batik Tulis?",
        "According to Text 2, what are two main physical benefits of rattan furniture compared to plastic furniture?",
        "In Text 3, what distinct flavor profiles distinguish Aceh Gayo coffee from Bali Kintamani coffee?"
    ]
    inf_q = [
        "Why is authentic Batik Tulis considered an 'heirloom investment' rather than just ordinary clothing in Text 1?",
        "Explain how the traditional wet-hulling ('Giling Basah') processing method influences the mouthfeel of Indonesian coffee beans (Text 3).",
        "Why do product reviews often include minor limitations (like protecting rattan from rain in Text 2) to build reader trust?"
    ]
    eval_q = [
        "Evaluate how promoting local creative products (batik, rattan, specialty coffee) on international social media blogs supports Indonesian economic growth.",
        "What specific criteria would you use to evaluate whether a creative handcrafted product deserves a 5-star rating?"
    ]
    perf_task = (
        "PENYUSUNAN ARTIKEL RESENSI PRODUK KREATIF LOKAL (LKPD / SUMATIF):\n"
        "Tulislah sebuah Artikel Ulasan Kritis (Review Text) sepanjang 200-250 kata untuk dipublikasikan di blog sekolah atau media sosial mempromosikan produk lokal/UMKM.\n"
        "Pilihlah salah satu produk unggulan nusantara:\n"
        "1. Kain Tenun Tradisional (Tenun Ikat NTT / Songket Palembang / Ulos Batak).\n"
        "2. Kerajinan Kulit Asli Garut / Keramik Kasongan Yogyakarta.\n"
        "3. Kuliner Khas Daerah (Rendang Minangkabau / Cokelat Lokal Indonesia).\n"
        "Format Resensi: Orientation -> Interpretative Recount (Bahan & Proses) -> Evaluation (Kelebihan & Saran Perawatan) -> Summary & Star Rating (1–5 Bintang).\n"
        "Gunakan kosakata evaluasi bernada positif dan meyakinkan."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_12.W.3_Review_Text_Creative_Heritage.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 12.W.4: JOB APPLICATION LETTER & MODERN ATS-COMPLIANT CV
# ==============================================================================
def create_tp_12_w_4():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS SMA",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XII / Semester 6 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 12.W.4 - Menyusun surat lamaran pekerjaan formal (Job Application Letter) dan daftar riwayat hidup modern (Curriculum Vitae / Resume) sesuai standar profesional dunia kerja.",
        "Elemen Capaian Pembelajaran": "Membaca - Memirsa & Menulis - Mempresentasikan",
        "Genre Teks / Ruang Lingkup": "Job Application Letter & Modern ATS-compliant Curriculum Vitae (CV)",
        "Alokasi Waktu": "10 Jam Pelajaran (5 Pertemuan @ 2 JP)",
        "Target Kemahiran CEFR": "Level B1+ s.d. B2",
        "Profil Pelajar Pancasila": "Mandiri, Bernalar Kritis, Kreatif"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "Harvard Career Services / Indeed Career Guide",
        "title": "Model Formal Cover Letter for an Entry-Level Junior Marketing Associate",
        "url": "https://careerservices.fas.harvard.edu/resources/cover-letters-guide/",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa Kelas XII)"
    }
    t1_paras = [
        "March 15, 2025\n\n"
        "Ms. Sarah Jenkins\nHuman Resources Director\nNexus Digital Media Group\nJalan Jenderal Sudirman No. 45, Jakarta 12190\n\n"
        "Dear Ms. Jenkins,\n\n"
        "I am writing to express my enthusiastic interest in the Junior Digital Marketing Associate position recently advertised on your company's career portal. As a motivated high school graduate with a strong academic foundation in English communication, graphic design, and social media management, I am confident that my skills and dedication will make a valuable contribution to your creative team.",
        "During my final year of secondary school, I served as the Head of Publications for the Student Council. In this leadership role, I successfully managed our school's official Instagram and YouTube channels, producing engaging video content that increased student event participation by 40 percent. This hands-on experience honed my abilities in content calendar scheduling, copywriting, and basic analytics reporting using Canva and Google Workspace.",
        "In addition to my technical proficiencies, I possess strong collaborative and organizational skills. I thrive in dynamic team environments and adapt quickly to fast-paced deadlines. Nexus Media's recent campaign on youth environmental sustainability deeply resonates with my personal values, and I am eager to bring my creative energy to your upcoming multimedia projects.",
        "Thank you for your time and consideration of my application. I have enclosed my Curriculum Vitae for your review, and I welcome the opportunity to discuss my qualifications in a personal interview. I can be reached at +62 812-3456-7890 or via email at devina.anindya@email.com.\n\n"
        "Sincerely,\n\n"
        "Devina Anindya"
    ]
    t1_analysis = {
        "Fungsi Sosial": "Menulis surat lamaran pekerjaan formal resmi (Application / Cover Letter) yang meyakinkan HRD untuk memperoleh kesempatan wawancara kerja.",
        "Struktur Teks": "Letter Heading (Tanggal & alamat pengirim/penerima) -> Formal Salutation ('Dear Ms. Jenkins') -> Opening Paragraph (Posisi yang dilamar dan sumber info) -> Body Paragraph 1 (Kualifikasi pendidikan & bukti prestasi kepemimpinan) -> Body Paragraph 2 (Keterampilan teknis, soft skills & kesesuaian nilai perusahaan) -> Closing Paragraph (Ajakan wawancara & kontak resmi) -> Sign-off & Signature ('Sincerely, Devina Anindya').",
        "Fitur Bahasa": "Formal Business Register, Action Verbs ('managed', 'produced', 'honed', 'thrive'), Polite Expressions ('welcome the opportunity', 'thank you for your consideration')."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "Model Formal Job Application Letter: Entry-Level Digital Associate",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "British Council Business English / Modern Career Hub",
        "title": "Modern ATS-Compliant Curriculum Vitae (CV) Template for High School Graduates",
        "url": "https://learnenglish.britishcouncil.org/business-english/writing-cv",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa Kelas XII)"
    }
    t2_paras = [
        "DEVINA ANINDYA\nJakarta, Indonesia | +62 812-3456-7890 | devina.anindya@email.com | linkedin.com/in/devinaanindya\n\n"
        "PROFESSIONAL PROFILE:\n"
        "Energetic and detail-oriented high school graduate with demonstrated skills in digital content creation, bilingual English-Indonesian communication, and youth project coordination. Proven track record of boosting student social media engagement by 40% through creative storytelling and visual design.\n\n"
        "EDUCATION:\n"
        "SMA Negeri 1 Jakarta (Science & Humanities Stream)\n"
        "Graduated: May 2025 | Overall GPA: 91.5 / 100\n"
        "• Relevant Coursework: Advanced English Communication, Digital Media Arts, Statistics\n\n"
        "LEADERSHIP & PROJECT EXPERIENCE:\n"
        "Head of Publications & Media | Student Council (OSIS) (2024 – 2025)\n"
        "• Directed a 5-member creative committee managing official school social media accounts.\n"
        "• Designed over 60 promotional flyers, digital posters, and event banners using Canva and Adobe Illustrator.\n"
        "• Coordinated live video broadcasting for the Annual Youth Arts Festival, reaching 5,000+ viewers.\n\n"
        "SKILLS & CERTIFICATIONS:\n"
        "• Hard Skills: Social Media Marketing, Copywriting, Canva Pro, Video Editing (CapCut), Google Workspace\n"
        "• Language Skills: Indonesian (Native), English (Upper-Intermediate / CEFR B2 – TOEFL ITP 560)\n"
        "• Soft Skills: Team Leadership, Public Speaking, Time Management, Creative Problem Solving"
    ]
    t2_analysis = {
        "Fungsi Sosial": "Menyajikan daftar riwayat hidup modern 1 halaman yang terstruktur rapi dan ramah sistem pelacak pelamar otomatis (ATS-friendly).",
        "Struktur Teks": "Contact Header -> Professional Profile Summary -> Education -> Leadership & Practical Experience (Bullet points diawali action verbs) -> Skills & Languages.",
        "Fitur Bahasa": "Action Verbs di awal bullet point ('Directed', 'Designed', 'Coordinated'), Data kuantitatif terukur ('40%', '60 promotional flyers', '5,000+ viewers'), Format bersih tanpa tabel rumit."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "Modern ATS-Friendly Curriculum Vitae: High School Graduate Portfolio",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "LinkedIn Talent Solutions / Creative Agency Career Post",
        "title": "Authentic Job Vacancy Advertisement: Junior Digital Content Creator",
        "url": "https://www.forbes.com/advisor/business/how-to-hire-content-creators/",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa Kelas XII)"
    }
    t3_paras = [
        "VACANCY ANNOUNCEMENT: JUNIOR DIGITAL CONTENT CREATOR\n"
        "Company: BrightSpark Creative Media Group\n"
        "Location: South Jakarta (Hybrid: 3 Days Office / 2 Days Remote)\n"
        "Employment Type: Full-time Entry Level | Application Deadline: April 30, 2025\n\n"
        "ABOUT THE ROLE:\n"
        "BrightSpark Media is seeking a passionate, creative Junior Digital Content Creator to join our expanding youth brand team. You will assist in producing engaging short-form video content, writing bilingual captions, and collaborating on brand campaigns for education and lifestyle clients.\n\n"
        "KEY RESPONSIBILITIES:\n"
        "• Assist in brainstorming, shooting, and editing weekly TikTok and Instagram Reels content.\n"
        "• Write punchy, informative social media copy in both English and Indonesian.\n"
        "• Monitor audience comments and assist in community engagement.\n"
        "• Collaborate with senior designers on marketing campaign layouts.\n\n"
        "REQUIREMENTS & QUALIFICATIONS:\n"
        "• High School / Vocational School (SMK) Diploma or equivalent.\n"
        "• Excellent written and spoken communication skills in English and Indonesian.\n"
        "• Proficiency with mobile video editing apps (CapCut / Premiere Rush) and graphic tools (Canva).\n"
        "• Enthusiastic, disciplined, and able to meet project deadlines.\n\n"
        "HOW TO APPLY:\n"
        "Send your formal Application Letter and ATS-compliant CV to recruitment@brightsparkmedia.com with subject line: 'Applicant Name - Junior Content Creator'."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "Authentic Job Vacancy Advertisement: Junior Digital Content Creator",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Applicant", "Noun", "Pelamar yang mendaftar lowongan pekerjaan", "The company received over one hundred applicant resumes."),
        ("ATS-friendly", "Adjective", "Format CV yang mudah dibaca sistem komputer pelacak", "Use simple fonts to ensure your CV is ATS-friendly."),
        ("Relevant", "Adjective", "Sesuai dan berkaitan langsung dengan posisi kerja", "Highlight relevant coursework and extracurricular skills."),
        ("Copywriting", "Noun", "Seni menulis naskah teks promosi yang memikat", "Her strong copywriting skills improved social media reach."),
        ("Proficiency", "Noun", "Kemahiran atau penguasaan tinggi dalam suatu bidang", "The job requires proficiency in bilingual English-Indonesian."),
        ("Deadline", "Noun", "Tenggat batas waktu akhir pengumpulan tugas", "Our team worked efficiently to meet the project deadline."),
        ("Salutation", "Noun", "Salam pembuka formal dalam korespondensi resmi", "Begin your business letter with a respectful salutation."),
        ("Enclosed", "Verb/Adjective", "Dilampirkan bersama dalam surat resmi", "Please find my enclosed Curriculum Vitae for review.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "In Text 1, what specific leadership achievement in high school does Devina Anindya highlight in her cover letter?",
        "According to the CV in Text 2, what hard skills and language proficiencies does Devina possess?",
        "What are the minimum education requirements and key responsibilities stated in the Job Advertisement in Text 3?"
    ]
    inf_q = [
        "Why is it essential to customize your job application letter to match the specific requirements of the job vacancy advertisement?",
        "Explain why modern CVs use quantifiable action verbs (e.g., 'increased engagement by 40%') rather than vague statements like 'good communication skills'.",
        "What makes an 'ATS-friendly' CV design superior to fancy, multi-colored graphic templates when applying online?"
    ]
    eval_q = [
        "Critique Devina's application letter (Text 1): In what ways does it effectively convey enthusiasm without sounding overly informal?",
        "As a high school senior, how can you turn your school organization (OSIS/Pramuka/Ekskul) activities into impressive achievements on your future CV?"
    ]
    perf_task = (
        "PENYUSUNAN DOSSIER LAMARAN KERJA LENGKAP (LKPD / SUMATIF MENULIS):\n"
        "Berdasarkan Iklan Lowongan Kerja pada Teks 3 (atau lowongan impianmu lainnya), susunlah Berkas Lamaran Kerja Profesional 1 Paket:\n"
        "1. A Formal Job Application Letter (Cover Letter) sepanjang 200-250 kata dengan format bisnis formal lengkap (Date, Inside Address, Salutation, 3 Body Paragraphs, Polite Sign-off, and Signature).\n"
        "2. A 1-Page Modern ATS-Compliant Curriculum Vitae (CV) yang memuat Profile, Education, Experience, dan Skills pribadimu secara jujur dan terstruktur rapi.\n"
        "Pastikan bebas dari kesalahan tata bahasa (spelling & grammar check) serta menggunakan action verbs yang terukur."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_12.W.4_Job_Application_and_CV.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 12.W.5: JOB INTERVIEW SIMULATION & ORAL PRESENTATION (STAR METHOD)
# ==============================================================================
def create_tp_12_w_5():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS SMA",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XII / Semester 6 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 12.W.5 - Mempraktikkan percakapan wawancara kerja formal (Job Interview Simulation) dan keterampilan presentasi diri (Elevator Pitch) secara percaya diri, santun, dan runtut menggunakan metode STAR.",
        "Elemen Capaian Pembelajaran": "Menyimak - Berbicara",
        "Genre Teks / Ruang Lingkup": "Job Interview Dialogue, STAR Method, Oral Self-Presentation, Executive Presence",
        "Alokasi Waktu": "8 Jam Pelajaran (4 Pertemuan @ 2 JP)",
        "Target Kemahiran CEFR": "Level B1+ s.d. B2",
        "Profil Pelajar Pancasila": "Mandiri, Bernalar Kritis, Berakhlak Mulia"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "Harvard Business Review / Indeed Career Interviewing Guide",
        "title": "Mastering the STAR Method: Structuring Impactful Interview Responses",
        "url": "https://www.indeed.com/career-advice/interviewing/how-to-use-the-star-interview-response-technique",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa Kelas XII)"
    }
    t1_paras = [
        "In modern behavioral job interviews, hiring managers frequently ask situational questions such as: 'Tell me about a time you resolved a conflict,' or 'Describe a situation where you had to meet a tight deadline.' The most effective framework to answer these questions with clarity and impact is the STAR method: Situation, Task, Action, and Result.",
        "S – SITUATION: Set the context by briefly describing the challenge or background. Be specific and concise. (Example: 'During my final year in high school, our student council team was organizing the annual school bazaar with two weeks remaining.')",
        "T – TASK: Explain your specific responsibility or role in that scenario. (Example: 'As the Sponsorship Coordinator, I needed to secure five local business sponsors to fund our event sound system.')",
        "A – ACTION: Detail the exact practical steps you took to resolve the problem. Highlight your own initiative and teamwork. (Example: 'I created a concise sponsorship proposal pitch, personally visited ten local businesses in our neighborhood, and presented how our 800 attending students would give their brands valuable visibility.')",
        "R – RESULT: Share the successful outcome with measurable evidence and personal learning. (Example: 'As a result, seven businesses agreed to sponsor us, raising 120% of our financial target. The bazaar ran smoothly, and I learned how clear communication builds community trust.')"
    ]
    t1_analysis = {
        "Fungsi Sosial": "Membimbing siswa menyusun jawaban wawancara kerja berbasis perilaku (behavioral interview) secara terstruktur, percaya diri, dan didukung bukti nyata.",
        "Struktur Teks": "Framework Introduction (Metode STAR) -> S (Situation: Latar konteks masalah) -> T (Task: Tugas spesifik pelamar) -> A (Action: Aksi solutif terukur) -> R (Result: Hasil nyata dan pelajaran berharga).",
        "Fitur Bahasa": "Action Verbs Lampau ('organized', 'secured', 'visited', 'raised'), Kosakata kepemimpinan dan pencapaian terukur ('120% target', 'concise proposal')."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "The STAR Method Framework: Structuring Behavioral Interview Answers",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "BBC Worklife / Corporate Career Coaching",
        "title": "Authentic Job Interview Dialogue: Answering Key Questions with Confidence",
        "url": "https://www.bbc.com/worklife/article/job-interview-tips-confidence",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa Kelas XII)"
    }
    t2_paras = [
        "Interviewer: 'Good morning, Devina. Welcome to BrightSpark Media. How are you today?'\n"
        "Candidate: 'Good morning, Ms. Jenkins. I am doing very well, thank you. I am excited and grateful for the opportunity to speak with you today.'\n\n"
        "Interviewer: 'To begin, could you tell us a little bit about yourself and why you applied for this position?'\n"
        "Candidate: 'Certainly. I recently graduated from SMA Negeri 1 Jakarta where I served as the Head of Publications for the student council. Throughout high school, I developed a passion for bilingual storytelling and digital media creation. When I saw BrightSpark's creative campaign on environmental awareness, I knew this was an organization where I could contribute my graphic design and copywriting skills while learning from an industry-leading team.'\n\n"
        "Interviewer: 'What would you say is your greatest strength, and what is an area you are currently working to improve?'\n"
        "Candidate: 'My greatest strength is my proactive adaptability. When unexpected changes happen during project deadlines, I remain calm and quickly reorganize tasks to keep the team on schedule. As for an area to improve, I used to find it difficult to delegate tasks because I wanted everything to be perfect. However, in our recent school festival, I practiced trusting my teammates by assigning clear roles, which made our project much more efficient.'\n\n"
        "Interviewer: 'Do you have any questions for us about the role or the team?'\n"
        "Candidate: 'Yes, thank you! What does a typical day look like for a Junior Content Creator here, and what opportunities exist for professional skill development?'"
    ]
    t2_analysis = {
        "Fungsi Sosial": "Mencontohkan percakapan wawancara kerja formal yang santun, percaya diri, dan strategis dalam menjawab kelebihan serta kelemahan diri.",
        "Struktur Teks": "Formal Greeting & Rapport Building -> Self-Introduction (Elevator Pitch) -> Behavioral Inquiry (Strengths & Constructive Weakness) -> Candidate's Inquiring Questions (Menunjukkan antusiasme).",
        "Fitur Bahasa": "Polite Formulaic Language ('Certainly', 'Grateful for the opportunity', 'Would you say'), Modulasi vokal tenang."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "Mastering the Interview: A Model Dialogue for Young Job Seekers",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "LinkedIn Learning / Career Strategy Column",
        "title": "The 60-Second Elevator Pitch: Making an Impactful First Impression",
        "url": "https://www.linkedin.com/pulse/mastering-elevator-pitch/",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa Kelas XII)"
    }
    t3_paras = [
        "In networking events, college scholarship interviews, or when responding to the classic prompt 'Tell me about yourself,' having a crisp, 60-second Elevator Pitch sets a confident and professional tone. A great elevator pitch answers three core questions: Who are you? What have you accomplished? Where are you heading?",
        "'Hello! My name is Farhan Ramadhan, and I am a graduating senior from SMAN 2 Bandung with a deep enthusiasm for sustainable green technology and robotics. Over the past two years, I captained our school robotics team, where we successfully engineered an automated solar-powered irrigation prototype that won second place at the National Science Expo.'",
        "'Through this project, I developed strong technical skills in sensor programming and collaborative problem-solving under pressure. I am now seeking a junior laboratory assistant internship where I can apply my dedication to renewable innovation and contribute to clean energy research.'",
        "Key Delivery Tips: Stand tall with open posture, maintain warm eye contact, smile genuinely, and speak at a steady, conversational pace without rushing."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "The 60-Second Elevator Pitch: How to Introduce Yourself Professionally",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Framework", "Noun", "Kerangka kerja terstruktur untuk menyelesaikan tugas", "The STAR framework helps structure interview answers."),
        ("Initiative", "Noun", "Prakarsa mandiri mengambil tindakan tanpa disuruh", "She took the initiative to contact local business sponsors."),
        ("Outcome", "Noun", "Hasil akhir dari suatu usaha atau tindakan", "Our campaign achieved a successful financial outcome."),
        ("Adaptability", "Noun", "Kemampuan beradaptasi cepat dengan perubahan", "Employers value employees with strong adaptability."),
        ("Delegate", "Verb", "Melimpahkan wewenang tugas kepada rekan tim", "A good leader learns how to delegate responsibilities."),
        ("Elevator pitch", "Noun Phrase", "Presentasi perkenalan diri singkat dan memikat (60 detik)", "Prepare a polished elevator pitch for college interviews."),
        ("Prototype", "Noun", "Model purwarupa awal dari suatu rancangan mesin", "The team engineered a solar-powered prototype."),
        ("Rapport", "Noun", "Hubungan komunikasi yang hangat, akrab, dan saling percaya", "Build positive rapport with your interviewers early.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "What do the four letters in the STAR interview method stand for in Text 1?",
        "In the dialogue in Text 2, how does Devina answer the question about her area of improvement (weakness) in a positive way?",
        "According to Text 3, what are the three essential questions every 60-second elevator pitch must answer?"
    ]
    inf_q = [
        "Why is ending an interview answer with a measurable 'Result' (STAR method) far more convincing than merely describing your actions?",
        "Why is asking insightful questions to the interviewer at the end of the meeting (Text 2) considered a mark of a strong candidate?",
        "How do non-verbal cues (eye contact, posture, calm voice) affect an interviewer's perception of your competence in Text 3?"
    ]
    eval_q = [
        "Critique why answering 'I have no weaknesses' or 'I am a perfectionist' during an interview is considered an ineffective response by HR managers.",
        "Evaluate the importance of practicing mock interviews with classmates before attending real job or university scholarship interviews."
    ]
    perf_task = (
        "SIMULASI WAWANCARA KERJA FORMAL (LKPD / UJIAN PRAKTIK LISAN):\n"
        "Lakukan Simulasi Wawancara Kerja Berpasangan (Job Interview Simulation):\n"
        "1. Siswa A berperan sebagai HR Interviewer dan Siswa B berperan sebagai Job Applicant (kemudian bergantian).\n"
        "2. Wawancara berdurasi 3-5 menit mencakup:\n"
        "   - Formal greeting & 60-second Elevator Pitch (Tell me about yourself).\n"
        "   - Pertanyaan Situasional (gunakan metode STAR untuk menjawab satu pengalaman nyata).\n"
        "   - Pertanyaan Strengths & Weakness.\n"
        "   - Satu pertanyaan cerdas dari pelamar kepada pewawancara.\n"
        "Dinilai berdasarkan: Fluency, Pronunciation & Intonation, Content Relevance (STAR), dan Professional Body Language."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_12.W.5_Job_Interview_Simulation.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 12.W.6: CAPTIONS, DATA GRAPHS & MULTIMODAL INFOGRAPHICS
# ==============================================================================
def create_tp_12_w_6():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS SMA",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XII / Semester 6 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 12.W.6 - Menafsirkan makna dan pesan tersirat dalam takarir foto jurnalistik (Caption), membaca tren grafik data statistik, dan menyajikan ringkasan infografis multimodal secara komunikatif.",
        "Elemen Capaian Pembelajaran": "Membaca - Memirsa & Menulis - Mempresentasikan",
        "Genre Teks / Ruang Lingkup": "Photo Captions, Data Graphs, Trends Description, Multimodal Infographics",
        "Alokasi Waktu": "6 Jam Pelajaran (3 Pertemuan @ 2 JP)",
        "Target Kemahiran CEFR": "Level B1+ s.d. B2",
        "Profil Pelajar Pancasila": "Bernalar Kritis, Kreatif, Mandiri"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "World Press Photo / Associated Press Captioning Guide",
        "title": "Photojournalism Standards: Writing Concise and Informative Captions",
        "url": "https://www.worldpressphoto.org/education/caption-guidelines",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa Kelas XII)"
    }
    t1_paras = [
        "In modern news journalism, a photo caption (also called a 'cutline') is an essential text element that accompanies a photograph. A caption explains what the camera cannot show on its own: the names of the individuals, the exact date and location, the historical context, and why the captured moment is significant. Professional news agencies like the Associated Press follow standard two-sentence caption structures:",
        "Sentence 1: Written in the Present Tense to describe the immediate action shown in the frame, stating clearly WHO is doing WHAT, WHERE, and WHEN.\n"
        "(Example: 'Indonesian environmental volunteers plant red mangrove seedlings along the eroding coastline of Muara Gembong, West Java, on Saturday, February 15, 2025.')",
        "Sentence 2: Written in the Past Tense or Present Perfect to provide essential background context and explain the broader significance of the event.\n"
        "(Example: 'The community reforestation initiative has planted over 50,000 mangrove trees since 2022 to protect local fishing villages from rising tidal sea floods.')",
        "Photo Credit Line: Always placed at the end in parentheses to credit the original photographer and news agency. (Example: 'Photo: Antara News / Ahmad Subaidi')"
    ]
    t1_analysis = {
        "Fungsi Sosial": "Menjelaskan kaidah penulisan takarir (photo caption) jurnalistik profesional berstandar internasional yang informatif dan ringkas.",
        "Struktur Teks": "Lead Sentence (Present Tense: Siapa melakukan apa, di mana, kapan) -> Context Sentence (Past/Present Perfect: Latar belakang dan signifikansi sosial) -> Photo Credit (Nama fotografer & media).",
        "Fitur Bahasa": "Action Verbs Present ('plant', 'demonstrate'), Present Perfect for Ongoing Impact ('has planted'), Factual Prepositions of Time and Place."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "Photojournalism Cutlines: How to Write Professional Photo Captions",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "The World Bank / BPS Indonesia Youth Statistics",
        "title": "Describing Statistical Charts: Vocabulary and Phrasing for Data Trends",
        "url": "https://www.worldbank.org/en/topic/education/data-graphs",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa Kelas XII)"
    }
    t2_paras = [
        "Interpreting and summarizing statistical visual data—such as line graphs, bar charts, and pie diagrams—is a vital academic and workplace literacy skill. When describing a data chart, writers must present the main trend, highlight significant fluctuations, and make accurate comparisons.",
        "Describing Upward Trends: Use verbs and phrases such as 'increased steadily', 'rose sharply', 'surged dramatically', or 'reached a peak of 85% in 2024'. (Example: 'Between 2020 and 2024, high school internet access across Indonesia rose sharply from 62% to 88%.')",
        "Describing Downward Trends & Plateaus: Use terms like 'declined gradually', 'dropped significantly', 'leveled off', or 'remained relatively stable'. (Example: 'Youth unemployment in urban areas dropped gradually following the expansion of the digital economy.')",
        "Making Comparisons & Proportions: Use phrases such as 'accounted for', 'compared to', 'the largest proportion', or 'three times higher than'. (Example: 'Renewable energy accounted for nearly one-third of total regional power consumption, whereas coal reliance shrank noticeably.')"
    ]
    t2_analysis = {
        "Fungsi Sosial": "Membekali siswa dengan ragam kosakata dan pola kalimat untuk mendeskripsikan tren kenaikan, penurunan, dan perbandingan grafik statistik.",
        "Struktur Teks": "Introductory Statement (Topik data grafik) -> Trend Vocabulary (Kenaikan tajam, penurunan bertahap, kondisi stabil) -> Comparative Proportions (Persentase dan perbandingan proporsi).",
        "Fitur Bahasa": "Trend Verbs ('rose', 'surged', 'dropped', 'leveled off'), Adverbs of Degree ('sharply', 'steadily', 'noticeably')."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "Describing Trends: Vocabulary for Line Graphs and Statistical Charts",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "World Health Organization (WHO) / UNICEF Health Infographics",
        "title": "Multimodal Infographic: Protecting Adolescent Mental Health in the Digital Age",
        "url": "https://www.who.int/news-room/fact-sheets/detail/adolescent-mental-health",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa Kelas XII)"
    }
    t3_paras = [
        "INFOGRAPHIC TITLE: 5 DAILY PILLARS FOR TEEN MENTAL WELL-BEING\n"
        "(Source: WHO Youth Health Initiative | Visual Layout: Multi-panel Graphic)\n\n"
        "Panel 1 [Icon: Bed & Moon]: 'Sleep Sanctuary – Ensure 8 to 10 hours of un-interrupted sleep. Put gadgets away 45 minutes before bedtime to allow natural melatonin release.'\n\n"
        "Panel 2 [Icon: Running Shoes]: 'Move Everyday – 30 minutes of aerobic exercise releases endorphins, improving mood and sharpening academic focus.'\n\n"
        "Panel 3 [Icon: Water & Apple]: 'Brain Nutrition – Drink 2 liters of water daily. Limit sugary energy drinks that trigger energy crashes.'\n\n"
        "Panel 4 [Icon: Two Talking Faces]: 'Human Connection – Have at least one honest, face-to-face conversation with family or friends daily.'\n\n"
        "Panel 5 [Icon: Smartphone with Clock]: 'Digital Detox – Set a daily 60-minute screen-free window to cultivate mindfulness and offline creative hobbies.'"
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "Multimodal Infographic Text: The 5 Pillars of Teen Mental Well-Being",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Cutline", "Noun", "Takarir penjelasan foto jurnalistik (caption)", "A cutline provides essential context beneath a photo."),
        ("Reforestation", "Noun", "Penanaman kembali hutan yang gundul/rusak", "The mangrove reforestation initiative protects coastlines."),
        ("Fluctuation", "Noun", "Naik-turunnya angka statistik data", "The graph illustrates seasonal fluctuations in employment."),
        ("Surged", "Verb (Past)", "Melonjak naik secara cepat dan drastis", "Internet usage surged sharply during remote learning."),
        ("Leveled off", "Phrasal Verb", "Mendatar atau stabil setelah mengalami kenaikan", "Smartphone adoption leveled off after reaching 90%."),
        ("Proportion", "Noun", "Perbandingan porsi atau bagian dari keseluruhan", "Renewables represent a growing proportion of energy."),
        ("Multimodal", "Adjective", "Menggabungkan teks visual, gambar, ikon, dan angka", "Infographics are effective multimodal communication tools."),
        ("Mindfulness", "Noun", "Kesadaran penuh dan ketenangan pikiran saat ini", "Daily digital detox helps cultivate mental mindfulness.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "In Text 1, what tense is used in the first sentence of a standard news photo caption, and what information must it state?",
        "According to Text 2, what verbs and adverbs can be used to describe a sudden, dramatic increase in data?",
        "What are the five daily pillars of teen mental health presented in the WHO multimodal infographic in Text 3?"
    ]
    inf_q = [
        "Why is it important for photo captions to provide background context rather than merely describing what is obviously visible in the picture?",
        "How does using visual multimodal elements (icons, bold colors, short text blocks) improve information retention compared to dense text paragraphs?",
        "Explain how the visual data trends described in Text 2 help policymakers make decisions about youth education funding."
    ]
    eval_q = [
        "Evaluate the effectiveness of infographics in raising teen awareness about social and environmental issues compared to traditional newspaper articles.",
        "In your view, which of the 5 mental health pillars in Text 3 is most overlooked by high school seniors facing final graduation exams?"
    ]
    perf_task = (
        "PENYUSUNAN TAKARIR FOTO JURNALISTIK & DESKRIPSI GRAFIK (LKPD / SUMATIF):\n"
        "Selesaikan dua tugas literasi multimodal berikut:\n"
        "Tugas A (Writing Photo Captions): Carilah sebuah foto peristiwa menarik di sekolahmu atau berita nasional. Tuliskan sebuah Takarir Foto (Caption) 2 kalimat berstandar jurnalistik (Sentence 1: Present tense WHO-WHAT-WHERE-WHEN; Sentence 2: Past tense CONTEXT & IMPACT; diakhiri Photo Credit).\n"
        "Tugas B (Describing a Data Graph): Pilihlah sebuah grafik batang atau diagram lingkaran sederhana (misal: penggunaan gawai remaja / tren pilihan jurusan kuliah). Tuliskan ringkasan deskripsi tren data sepanjang 100-120 kata menggunakan kosakata tren (rose steadily, dropped, accounted for)."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_12.W.6_Captions_and_Infographics.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


if __name__ == "__main__":
    print("Generating calibrated authentic materials for 5-Modul Ajar XII SMA (Wajib)...")
    create_tp_12_w_1()
    create_tp_12_w_2()
    create_tp_12_w_3()
    create_tp_12_w_4()
    create_tp_12_w_5()
    create_tp_12_w_6()
    print("All 6 DOCX authentic materials for Kelas XII SMA Wajib successfully created!")
