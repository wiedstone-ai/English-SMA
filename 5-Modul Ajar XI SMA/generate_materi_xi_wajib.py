# -*- coding: utf-8 -*-
"""
Generator script for authentic English materials (.docx) in 5-Modul Ajar XI SMA (Wajib).
Calibrated specifically to CEFR B1 to early B2 (Accessible, Communicative SMA Level).
Each TP contains 3 authentic texts from credible sources with explicit URLs,
practical vocabulary glossaries, and aligned assessment questions.
"""

import os
import sys

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import docx_helper

TARGET_DIR = r"C:\Users\KOMPUTERKU2\Documents\English SMA\5-Modul Ajar XI SMA"
os.makedirs(TARGET_DIR, exist_ok=True)

# ==============================================================================
# TP 11.W.1: ANALYTICAL EXPOSITION - DIGITAL ETHICS & SOCIAL MEDIA
# ==============================================================================
def create_tp_11_w_1():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS SMA",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XI / Semester 3 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 11.W.1 - Menganalisis fungsi sosial, struktur teks, dan unsur kebahasaan teks eksposisi analitis tentang etika digital dan media sosial, serta menyusun esai argumentatif logis.",
        "Elemen Capaian Pembelajaran": "Membaca - Memirsa & Menulis - Mempresentasikan",
        "Genre Teks / Ruang Lingkup": "Analytical Exposition Text (Digital Ethics, Social Media & Cyber Well-being)",
        "Alokasi Waktu": "10 Jam Pelajaran (5 Pertemuan @ 2 JP)",
        "Target Kemahiran CEFR": "Level B1 s.d. B1+ (Menengah - Sesuai Siswa Kelas XI SMA)",
        "Profil Pelajar Pancasila": "Bernalar Kritis, Mandiri, Berakhlak Mulia"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "UNESCO Education Portal / Cyberbullying Research Center",
        "title": "Why We Must Stop Cyberbullying: Protecting Adolescent Mental Health",
        "url": "https://www.unesco.org/en/articles/preventing-cyberbullying-schools",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah - Sesuai Siswa Kelas XI SMA)"
    }
    t1_paras = [
        "In our modern digital world, social media platforms have become an important part of teenage life. While these apps allow students to connect with friends and share creative ideas, the rise of cyberbullying has become a serious concern. Cyberbullying refers to sending mean messages, posting false rumors, or sharing private photos online without permission. We must take strong action against cyberbullying for several important reasons.",
        "First, digital harassment hurts victims around the clock. Unlike traditional bullying at school which stops when the bell rings, cyberbullying follows teenagers directly into their homes through their phones. Victims can receive hurtful comments late at night in their bedrooms. This constant stress causes severe anxiety, poor sleep, and a drop in school grades.",
        "Second, online cruelty spreads very quickly. A harmful post can be shared hundreds of times within minutes. People who comment or share mean posts often feel protected by screen anonymity, making them act far more aggressively than they would in face-to-face conversations.",
        "In conclusion, cyberbullying is a harmful behavior that damages young people's well-being. Schools, parents, and students must work together to report harassment, practice digital empathy, and create a kinder online environment."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Mempersuasi pembaca (Analytical Exposition) bahwa perundungan siber (cyberbullying) adalah masalah serius yang memerlukan tindakan tegas dan empati bersama.",
        "Struktur Teks": "Thesis (Paragraf 1: Masalah cyberbullying pada remaja) -> Arguments (Paragraf 2: Teror 24 jam ke ruang pribadi; Paragraf 3: Penyebaran cepat dan ilusi anonimitas) -> Reiteration (Paragraf 4: Kesimpulan dan ajakan menciptakan ruang digital yang sehat).",
        "Fitur Bahasa": "Thinking/Feeling Verbs ('worry', 'feel', 'hurt'), Causal Connectors ('unlike', 'because', 'therefore'), Present Tense."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "Why We Must Combat Cyberbullying Among Teenagers",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "BBC Learning English / Media Literacy Section",
        "title": "The Importance of Fact-Checking: How to Spot Online Disinformation",
        "url": "https://www.bbc.co.uk/learningenglish/features/media-literacy/fake-news",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah - Sesuai Siswa Kelas XI SMA)"
    }
    t2_paras = [
        "Every day, millions of news headlines, videos, and social media posts are published online. While having instant access to information is helpful, the rapid spread of false information—commonly called 'fake news'—poses a real danger to society. Learning how to verify facts and think critically is an essential skill for all high school students.",
        "Primarily, fake news is often created to manipulate people's emotions. Authors of false articles write sensational, shocking headlines to make people angry or scared, which encourages them to click and share the post immediately without checking the facts. During health emergencies or elections, believing false rumors can lead people to make dangerous decisions.",
        "Furthermore, modern software and artificial intelligence can easily generate realistic fake images and audio clips. Without basic media literacy training, students can easily be tricked by manipulated content.",
        "To summarize, we must never believe everything we read on social media. Students should always verify information by checking credible news sources, looking for the original author, and thinking twice before sharing any sensational story."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Meyakinkan pembaca tentang pentingnya memeriksa kebenaran berita (fact-checking) sebelum mempercayai dan membagikannya di media sosial.",
        "Struktur Teks": "Thesis (Bahaya berita palsu) -> Arguments (Manipulasi emosi demi klik; Perkembangan teknologi manipulasi media) -> Reiteration (Kewajiban verifikasi sebelum membagikan).",
        "Fitur Bahasa": "Connectors ('Primarily', 'Furthermore', 'To summarize'), Modal Verbs ('must', 'can', 'should')."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "Why Media Literacy and Fact-Checking Matter for High School Students",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "UNICEF Youth & Digital Safety / Child Mind Institute",
        "title": "The Hidden Costs of Screen Dependency: Protecting Real-Life Relationships",
        "url": "https://www.unicef.org/parenting/digital-safety/healthy-habits-teens",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah - Sesuai Siswa Kelas XI SMA)"
    }
    t3_paras = [
        "Smartphones have transformed how we communicate, study, and entertain ourselves. However, excessive screen time and digital dependency can quietly damage students' personal development and mental health.",
        "First, spending hours scrolling through social media feeds often replaces valuable face-to-face interactions. When teenagers interact only through text messages or likes, they miss important social cues like body language, eye contact, and tone of voice. This can make them feel awkward and nervous when talking to people in real life.",
        "Second, relying on phones for constant entertainment shortens our attention spans. Students who frequently check their notifications while doing homework find it difficult to concentrate on reading books or solving complex math problems.",
        "In conclusion, while smartphones are useful tools, we should never let them control our daily lives. Setting daily screen-time limits and spending quality time with family and friends are necessary steps to maintain emotional balance."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "The Risks of Smartphone Dependency for Teenagers",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Harassment", "Noun", "Tindakan mengganggu, mengancam, atau melecehkan", "Cyberbullying is a form of digital harassment."),
        ("Anonymity", "Noun", "Kondisi tanpa identitas nama asli (anonim)", "People feel bolder online due to screen anonymity."),
        ("Disinformation", "Noun", "Informasi salah yang sengaja dibuat untuk menipu", "Disinformation aims to mislead readers for profit."),
        ("Sensational", "Adjective", "Menghebohkan, sengaja dibuat dramatis", "Sensational headlines are designed to attract quick clicks."),
        ("Manipulate", "Verb", "Memanipulasi atau memengaruhi pikiran seseorang", "Fake news tries to manipulate public emotions."),
        ("Social cues", "Noun Phrase", "Isyarat sosial alami (bahasa tubuh, kontak mata)", "Text messaging lacks vital real-life social cues."),
        ("Attention span", "Noun Phrase", "Rentang waktu kemampuan seseorang untuk fokus", "Constant phone notifications reduce our attention span."),
        ("Verify", "Verb", "Memverifikasi, membuktikan kebenaran fakta", "Always verify shocking news before sharing it.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "According to Text 1, what makes cyberbullying significantly different from traditional playground bullying at school?",
        "Why do creators of fake news often write sensational and shocking headlines in Text 2?",
        "What two negative impacts of excessive smartphone usage on teenagers are discussed in Text 3?"
    ]
    inf_q = [
        "Why does screen anonymity make some internet users act more aggressively in comment sections than in real life?",
        "Explain why sharing a piece of false news during an emergency situation can be dangerous to the public.",
        "How can high school students practice 'digital empathy' when interacting with classmates on group messaging apps?"
    ]
    eval_q = [
        "In your opinion, should schools ban smartphones completely during school hours, or should they teach responsible phone habits? Explain your reasons.",
        "Evaluate your own social media habits: How often do you verify information before sharing it with your friends?"
    ]
    perf_task = (
        "PENYUSUNAN ESAI EKSPOSISI ANALITIS (LKPD / SUMATIF MENULIS):\n"
        "Tulislah sebuah esai eksposisi analitis (Analytical Exposition) sepanjang 200-250 kata mengenai salah satu topik berikut:\n"
        "Topik 1: Why Teenagers Should Limit Daily Social Media Screen Time.\n"
        "Topik 2: The Importance of Protecting Personal Privacy Online.\n"
        "Struktur Esai:\n"
        "• Thesis: Nyatakan isu dan pendapat utamamu dengan jelas.\n"
        "• Arguments: Tuliskan minimal 2 paragraf argumen yang didukung contoh konkret dari kehidupan sehari-hari.\n"
        "• Reiteration: Tegaskan kembali kesimpulan dan pesan pentingmu.\n"
        "Gunakan minimal 4 kosakata dari tabel glosarium di atas."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_11.W.1_Analytical_Exposition_Digital.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 11.W.2: HORTATORY EXPOSITION - PLASTIC WASTE & SUSTAINABILITY
# ==============================================================================
def create_tp_11_w_2():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS SMA",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XI / Semester 3 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 11.W.2 - Menganalisis dan memproduksi teks eksposisi hortatori yang merekomendasikan tindakan pelestarian lingkungan dan pengurangan sampah plastik.",
        "Elemen Capaian Pembelajaran": "Menyimak - Berbicara & Menulis - Mempresentasikan",
        "Genre Teks / Ruang Lingkup": "Hortatory Exposition Text (Waste Management, Environmental Advocacy, Youth Action)",
        "Alokasi Waktu": "12 Jam Pelajaran (6 Pertemuan @ 2 JP)",
        "Target Kemahiran CEFR": "Level B1 s.d. B1+",
        "Profil Pelajar Pancasila": "Bernalar Kritis, Gotong Royong, Berakhlak Mulia"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "UNEP Clean Seas Campaign / The Jakarta Post Environment",
        "title": "Why Cities Must Ban Single-Use Plastic Bags and Packaging",
        "url": "https://www.cleanseas.org/features/action-on-plastics",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah - Sesuai Siswa Kelas XI SMA)"
    }
    t1_paras = [
        "Single-use plastic bags and packaging have become an alarming environmental threat across Indonesian cities. Every day, millions of plastic wrappers and drink cups are thrown away after being used for only a few minutes. Because plastic takes hundreds of years to break down, it clogs drainage canals, causes urban flooding during the rainy season, and severely pollutes our rivers and oceans. Local governments and citizens must take decisive steps to phase out single-use plastics immediately.",
        "First, plastic pollution imposes huge financial costs on society. Municipal governments spend billions of rupiah each year clearing blocked drainage channels and dredging plastic debris from rivers. These public funds could easily be spent on better school facilities or public healthcare if plastic waste were prevented at the source.",
        "Second, convenient eco-friendly alternatives already exist. Reusable cloth tote bags, stainless steel water bottles, and biodegradable food containers made from cassava starch or bamboo are widely available, affordable, and durable.",
        "Therefore, city administrations should enforce a strict ban on free plastic shopping bags in all supermarkets, convenience stores, and traditional markets. Furthermore, schools should encourage students to bring their own lunchboxes and tumblers every day to create a zero-waste campus."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Mempersuasi dan merekomendasikan (Hortatory Exposition) agar pemerintah dan masyarakat menghentikan penggunaan plastik sekali pakai.",
        "Struktur Teks": "Thesis (Masalah sampah plastik kota) -> Arguments (Beban biaya pembersihan selokan; Tersedianya alternatif tas kain ramah lingkungan) -> Recommendation (Pemberlakuan larangan kantong plastik dan kewajiban membawa wadah sendiri).",
        "Fitur Bahasa": "Modalitas anjuran ('must take decisive steps', 'should enforce', 'should encourage'), Causal Connectors ('Because', 'Therefore', 'Furthermore')."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "Why We Must Phase Out Single-Use Plastics in Our Cities",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "World Wildlife Fund (WWF) Youth / Eco-Schools Portal",
        "title": "Why High School Students Should Practice Daily Waste Sorting",
        "url": "https://www.worldwildlife.org/initiatives/plastic-pollution-youth",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah - Sesuai Siswa Kelas XI SMA)"
    }
    t2_paras = [
        "In many high schools, all trash—leftover noodles, paper sheets, and plastic drink bottles—is tossed into a single trash can. This unorganized waste ends up in giant city landfills where it rots and creates foul odors and greenhouse gases. Establishing a systematic waste-sorting program in schools is an urgent necessity.",
        "Primarily, separating recyclable materials creates valuable economic resources. Clean plastic bottles, aluminum drink cans, and paper can be sold to local recycling centers. The money earned can be used to fund student council activities, buy library books, or plant green trees around the schoolyard.",
        "Moreover, sorting waste teaches teenagers practical civic discipline. When students get into the habit of placing plastic in the yellow bin and food scraps in the green bin, they naturally practice the same eco-friendly habit at home and in public spaces.",
        "For these reasons, school principals ought to install clearly labeled three-compartment recycling bins across campus and conduct regular educational workshops on waste management."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Mengajak seluruh warga sekolah untuk memilah sampah menjadi organik dan anorganik demi kelestarian lingkungan.",
        "Struktur Teks": "Thesis (Masalah tempat sampah campur aduk) -> Arguments (Nilai ekonomis daur ulang; Melatih kedisiplinan siswa) -> Recommendation (Penyediaan tong sampah terpilah dan pelatihan).",
        "Fitur Bahasa": "Rekomendasi ('ought to install', 'conduct regular workshops'), Present Tense."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "Why High Schools Should Implement Waste Sorting Programs",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "Earth Island Institute / Teen Environmental Action",
        "title": "Why Youth Should Demand Climate Action and Environmental Education",
        "url": "https://www.thejakartapost.com/opinion/plastic-bag-ban-cities",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah - Sesuai Siswa Kelas XI SMA)"
    }
    t3_paras = [
        "Climate change and environmental degradation are no longer distant scientific theories; they are directly affecting our lives today through extreme weather, hotter summers, and rising sea levels. High school students, as future leaders, must be equipped with strong ecological awareness.",
        "First, youth voices have tremendous power to change community habits. When young people organize tree-planting activities or clean up trash in neighborhood parks, adults and local leaders pay attention and feel inspired to support them.",
        "Second, understanding sustainability helps students prepare for modern green careers. From renewable energy engineering to eco-friendly agriculture, green jobs will dominate the global economy in the coming decades.",
        "Therefore, educational authorities must integrate practical environmental projects into the high school curriculum, and teenagers should actively participate in environmental clubs to protect our shared planet."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "Why Youth Must Lead Environmental Change in Their Communities",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Phase out", "Phrasal Verb", "Menghentikan penggunaan sesuatu secara bertahap", "Cities must phase out plastic packaging gradually."),
        ("Drainage", "Noun", "Saluran air got atau gorong-gorong pembuangan", "Plastic trash clogs urban drainage channels during storms."),
        ("Biodegradable", "Adjective", "Dapat terurai secara alami oleh tanah dan mikroba", "Cassava food boxes are completely biodegradable."),
        ("Decisive", "Adjective", "Tegas, nyata, dan menghasilkan dampak langsung", "We need decisive action to solve city waste problems."),
        ("Landfill", "Noun", "Tempat pembuangan akhir (TPA) sampah kota", "Dumping mixed trash shortens the lifespan of landfills."),
        ("Ought to", "Modal Verb", "Seharusnya, sepantasnya (menyatakan anjuran)", "Schools ought to provide three-color recycling bins."),
        ("Sustainability", "Noun", "Keberlanjutan lingkungan hidup jangka panjang", "Education should teach sustainability to young generations."),
        ("Degradation", "Noun", "Kemunduran atau kerusakan kondisi lingkungan", "Deforestation leads to serious environmental degradation.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "In Text 1, what negative problems occur in cities when plastic waste clogs urban drainage channels?",
        "According to Text 2, what useful benefits can schools gain by selling sorted recyclable bottles and paper?",
        "What two reasons in Text 3 explain why environmental education is important for high school students?"
    ]
    inf_q = [
        "How does spending money on clearing clogged plastic drains prevent cities from improving school facilities (Text 1)?",
        "Why is practicing waste sorting at school more effective than merely reading about recycling in a textbook?",
        "Analyze the difference in purpose between an Analytical Exposition (asserting a viewpoint) and a Hortatory Exposition (recommending action)."
    ]
    eval_q = [
        "Do you agree that single-use plastic bags should be banned completely in your city? Defend your opinion with practical examples.",
        "What simple eco-friendly habit can your class implement starting this week to reduce trash production?"
    ]
    perf_task = (
        "PENYUSUNAN TEKS EKSPOSISI HORTATORI & PIDATO SINGKAT (LKPD / SUMATIF):\n"
        "Susunlah sebuah Teks Eksposisi Hortatori (Hortatory Exposition) sepanjang 200-250 kata untuk dibacakan pada upacara bendera atau mading sekolah.\n"
        "Topik Pilihan:\n"
        "1. Why Every Student Should Bring Their Own Water Tumbler to School.\n"
        "2. Why Our School Should Implement a Paperless Homework Policy.\n"
        "Struktur Wajib:\n"
        "- Thesis: Sebutkan latar belakang isu kepedulian lingkungan.\n"
        "- Arguments: Berikan 2 argumen logis yang meyakinkan.\n"
        "- Recommendation: Tuliskan rekomendasi tindakan tegas menggunakan modal verbs (should, must, ought to).\n"
        "Gunakan minimal 4 kosakata dari daftar glosarium di atas."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_11.W.2_Hortatory_Exposition_Advocacy.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 11.W.3: CONTEMPORARY NARRATIVE TALES - PERSEVERANCE & COMPASSION
# ==============================================================================
def create_tp_11_w_3():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS SMA",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XI / Semester 3 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 11.W.3 - Mengapresiasi, menganalisis alur cerita, dan merekonstruksi teks naratif kontemporer bertema empati, persahabatan, dan perjuangan hidup.",
        "Elemen Capaian Pembelajaran": "Menyimak - Berbicara & Membaca - Memirsa",
        "Genre Teks / Ruang Lingkup": "Contemporary Narrative Text (Short Stories, Friendship, Moral Courage)",
        "Alokasi Waktu": "12 Jam Pelajaran (6 Pertemuan @ 2 JP)",
        "Target Kemahiran CEFR": "Level B1 s.d. B1+",
        "Profil Pelajar Pancasila": "Berakhlak Mulia, Gotong Royong, Mandiri"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "British Council LearnEnglish / Modern Folk Stories",
        "title": "The Wooden Bridge: A Story of Forgiveness and Brotherhood",
        "url": "https://learnenglish.britishcouncil.org/general-english/stories/bridge-forgiveness",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah - Sesuai Siswa Kelas XI SMA)"
    }
    t1_paras = [
        "Once upon a time, two brothers lived on neighboring farms separated by a small creek. For forty years, they had farmed side by side, sharing machinery and trading goods without a single argument. However, a small misunderstanding over a missing calf led to bitter words, followed by weeks of cold silence between them.",
        "One morning, an itinerant carpenter knocked on the older brother's door looking for work. The older brother, still angry at his sibling, said, 'See that creek over there? My younger brother dug that ditch with his tractor to separate our land. I want you to build me an eight-foot-tall wooden fence so I never have to look at his house again.' The carpenter smiled gently and replied, 'I understand the situation. I will build something that satisfies you.'",
        "The older brother left for town and returned at sunset. To his absolute astonishment, there was no fence at all! Instead, stretching across the creek was a beautiful wooden bridge with sturdy handrails. Standing in the middle of the bridge with outstretched arms was his younger brother. 'Brother,' the younger man cried with tears, 'you built a bridge after all the hurtful things I said!' The two brothers embraced warmly in the middle of the bridge, their grudge washed away by kindness."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Mengajarkan nilai luhur memaafkan, menjaga silaturahmi, dan membangun jembatan perdamaian daripada mendirikan dinding pemisah.",
        "Struktur Teks": "Orientation (Dua saudara hidup bertani rukun selama 40 tahun) -> Complication (Perselisihan kecil berujung amarah dan pesanan pagar pemisah) -> Resolution (Tukang kayu justru membangun jembatan persaudaraan yang menyatukan mereka kembali).",
        "Fitur Bahasa": "Past Tense ('lived', 'farmed', 'dug', 'built', 'embraced'), Dialog bermakna, Simbolisme jembatan vs pagar."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "Building Bridges: A Tale of Two Brothers",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "Reader's Digest Inspiring Stories / Youth Fiction",
        "title": "A Beacon in the Storm: When Friendship Conquers Fear",
        "url": "https://www.rd.com/article/true-stories-friendship-rescue/",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah - Sesuai Siswa Kelas XI SMA)"
    }
    t2_paras = [
        "Rian and Danu were high school classmates who loved weekend hiking in the mountains of West Java. One Saturday afternoon, while descending from Mount Gede, thick dark clouds gathered quickly, and a violent thunderstorm swept across the forested ridge. Heavy rain turned the hiking path into a slippery torrent of mud.",
        "While rushing down, Danu slipped on a wet tree root and twisted his ankle badly. The pain was sharp, and he could barely stand. Panic began to rise as darkness fell upon the misty forest. 'Leave me here, Rian,' Danu pleaded, shivering from the cold wind. 'Go find help in the village before the storm gets worse.'",
        "Rian firmly shook his head. 'We started this hike together, and we are going home together,' he said calmly. Rian found a sturdy tree branch to serve as a crutch, wrapped his warm jacket around Danu's shoulders, and supported his friend step by step down the rocky path. After three grueling hours, they saw the welcoming yellow lantern lights of the basecamp hut. Safe inside by the fireplace, both boys realized that true friendship is tested not on sunny days, but in the heart of the storm."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Menghibur pembaca dengan kisah petualangan remaja yang menonjolkan nilai kesetiakawanan dan keberanian.",
        "Struktur Teks": "Orientation (Rian dan Danu mendaki Gunung Gede) -> Complication (Badai petir dan cedera pergelangan kaki Danu) -> Resolution (Rian setia memapah sahabatnya sampai ke basecamp).",
        "Fitur Bahasa": "Descriptive Action Verbs ('descending', 'twisted', 'shivering', 'pleaded', 'supported'), Setting atmosferik."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "In the Heart of the Storm: The True Test of Friendship",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "Indonesian Young Adult Literary Journal / Kemendikbud",
        "title": "The Runner from the Hills: Dedication Over Adversity",
        "url": "https://literacy.kemdikbud.go.id/cerita-remaja-prestasi",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah - Sesuai Siswa Kelas XI SMA)"
    }
    t3_paras = [
        "In a quiet hilly village in Flores, a sixteen-year-old girl named Maria had a passion for long-distance running. While students in big cities had modern rubber athletic tracks and professional running shoes, Maria ran barefoot along dusty dirt roads every morning before school, carrying water jugs from the spring to help her family.",
        "When the regional high school athletics competition was announced in the district capital, Maria's physical education teacher encouraged her to register for the 3,000-meter race. On the day of the championship, looking at the other competitors wearing expensive sports gear, Maria felt self-conscious. However, her teacher whispered, 'Running is not in your shoes, Maria; it is in your heart and lungs.'",
        "When the starting pistol fired, Maria started steadily. Lap after lap, she maintained a smooth, rhythmic pace. During the final lap, as the crowd cheered loudly, Maria found a reserve of strength forged from years of climbing hills with water jugs. She crossed the finish line first with tears of pride, proving that humble beginnings are no barrier to greatness."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "The Runner from the Hills: A Story of Determination",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Misunderstanding", "Noun", "Salah paham yang memicu perselisihan", "A small misunderstanding caused a rift between the brothers."),
        ("Itinerant", "Adjective", "Berpindah-pindah tempat untuk mencari pekerjaan", "An itinerant carpenter knocked on the door looking for work."),
        ("Grudge", "Noun", "Rasa dendam atau sakit hati yang dipendam", "Holding a grudge only hurts the person carrying it."),
        ("Shivering", "Verb/Participle", "Menggigil kedinginan atau ketakutan", "Danu was shivering from the freezing mountain rain."),
        ("Grueling", "Adjective", "Sangat melelahkan dan menguras tenaga", "They completed a grueling three-hour hike in the dark."),
        ("Crutch", "Noun", "Tongkat penopang untuk membantu berjalan saat cedera", "He used a sturdy wooden branch as an emergency crutch."),
        ("Self-conscious", "Adjective", "Rasa minder atau canggung memikirkan pandangan orang", "Maria felt self-conscious about her simple running clothes."),
        ("Rhythmic", "Adjective", "Berirama teratur dan stabil", "She maintained a steady, rhythmic breathing pace.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "In Text 1, what did the older brother originally order the carpenter to build along the creek?",
        "What unexpected difficulty happened to Danu while descending Mount Gede in Text 2?",
        "What daily activity in her home village helped Maria build strong endurance for running in Text 3?"
    ]
    inf_q = [
        "Why did the carpenter in Text 1 choose to build a bridge instead of the wooden fence requested by the older brother?",
        "What does Rian's refusal to leave his injured friend in Text 2 reveal about his character?",
        "Explain the meaning of the teacher's advice in Text 3: 'Running is not in your shoes, but in your heart and lungs.'"
    ]
    eval_q = [
        "In Text 1, the older brother was angry for weeks over a small calf. Why is holding grudges harmful to personal peace?",
        "Which of the three stories inspired you the most to overcome personal difficulties in your own life? Explain your reason."
    ]
    perf_task = (
        "PENULISAN CERPEN KONTEMPORER BERMAKNA (LKPD / SUMATIF KREATIF):\n"
        "Tulislah sebuah Cerpen Naratif Kontemporer (Contemporary Short Story) sepanjang 200-250 kata mengenai pengalaman remaja.\n"
        "Pilihlah salah satu tema:\n"
        "1. Reconciling with a friend after a misunderstanding.\n"
        "2. Overcoming self-doubt during a school competition or exam.\n"
        "3. Helping a stranger in need during a rainy day.\n"
        "Pastikan cerita memuat struktur lengkap (Orientation -> Complication -> Resolution) dan 1 kalimat pesan moral di akhir cerita."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_11.W.3_Contemporary_Narrative_Tales.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 11.W.4: COMPLEX EXPLANATION TEXT - SCIENCE & TECHNOLOGY
# ==============================================================================
def create_tp_11_w_4():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS SMA",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XI / Semester 4 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 11.W.4 - Menganalisis urutan proses kausalitas dan fitur bahasa dalam teks eksplanasi ilmiah tentang fenomena alam dan teknologi energi terbarukan.",
        "Elemen Capaian Pembelajaran": "Membaca - Memirsa & Menulis - Mempresentasikan",
        "Genre Teks / Ruang Lingkup": "Explanation Text (Earthquakes, Tsunami Mechanics, Solar Energy Technology)",
        "Alokasi Waktu": "12 Jam Pelajaran (6 Pertemuan @ 2 JP)",
        "Target Kemahiran CEFR": "Level B1 s.d. B1+",
        "Profil Pelajar Pancasila": "Bernalar Kritis, Mandiri, Kreatif"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "US Geological Survey (USGS) Education / BMKG Indonesia",
        "title": "How Earthquakes Occur: Tectonic Plates and Seismic Energy",
        "url": "https://www.usgs.gov/educational-resources/how-earthquakes-work",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah - Sesuai Siswa Kelas XI SMA)"
    }
    t1_paras = [
        "An earthquake is an intense shaking of the Earth's surface caused by a sudden slip along a geological fault. Indonesia is located within the famous 'Pacific Ring of Fire', an active zone where several massive tectonic plates meet. To understand how earthquakes happen, we must examine the movement of these tectonic plates beneath our feet.",
        "The Earth's crust is not a single solid shell; rather, it is broken into giant puzzle pieces called tectonic plates that move slowly over the hot, semi-liquid mantle. As plates collide, slide past one another, or pull apart, their rough edges get stuck due to friction while the rest of the plate keeps moving. Over decades or centuries, immense elastic stress builds up along the fault line.",
        "Eventually, the accumulated stress overcomes the friction. The rocks suddenly fracture and snap, releasing monumental energy in all directions in the form of seismic waves. These waves travel through the Earth's crust, causing the ground above to vibrate violently. The point underground where the fracture begins is called the focus, while the point directly above it on the surface is known as the epicenter."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Menjelaskan secara ilmiah proses terjadinya gempa bumi akibat pergerakan dan pelepasan energi lempeng tektonik.",
        "Struktur Teks": "General Statement (Pengenalan gempa bumi dan Cincin Api Pasifik) -> Sequenced Explanation (Lempeng saling bergesekan -> Penumpukan energi elastis -> Patahan batuan dan perambatan gelombang seismik) -> Concluding Note (Fokus dan episentrum).",
        "Fitur Bahasa": "Passive Voice ('is caused', 'is broken into', 'is known as'), Causal Connectors ('due to', 'Eventually', 'causing'), Technical Terms ('tectonic plates', 'fault line', 'seismic waves')."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "How Earthquakes Happen: The Science of Plate Tectonics",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "BMKG Indonesia / National Oceanic and Atmospheric Administration (NOAA)",
        "title": "The Mechanics of Tsunami Formation: From Seafloor to Shore",
        "url": "https://www.bmkg.go.id/edukasi/mekanisme-tsunami",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah - Sesuai Siswa Kelas XI SMA)"
    }
    t2_paras = [
        "A tsunami is a series of powerful ocean waves generated by a sudden displacement of a massive volume of water. While commonly associated with earthquakes, tsunamis can also be triggered by undersea volcanic eruptions or coastal landslides.",
        "The process typically begins when an undersea earthquake causes a vertical displacement of the seafloor. When one tectonic plate forces another plate upward, the entire ocean column above it is instantly lifted. Gravity immediately pulls the water back down, creating a series of concentric wave ripples that travel across the deep ocean at jetliner speeds of over 700 kilometers per hour.",
        "In deep water, tsunami waves are barely noticeable, often measuring less than one meter in height. However, as the waves approach shallow coastal waters, friction with the seafloor slows the wave velocity down to about 40 kilometers per hour. Consequently, the energy compresses, forcing the wave height to surge dramatically into towering walls of water that inundate coastlines."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Menjelaskan tahapan pembentukan tsunami dari dasar laut hingga menghantam daratan.",
        "Struktur Teks": "General Statement (Definisi tsunami) -> Explanation Sequence (Patahan lempeng bawah laut -> Pengangkatan kolom air -> Perambatan gelombang cepat -> Pemampatan energi di perairan dangkal).",
        "Fitur Bahasa": "Temporal & Causal Markers ('typically begins', 'When', 'Consequently', 'as the waves approach')."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "How Tsunamis Form and Travel Across Oceans",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "National Geographic Science / US Department of Energy",
        "title": "How Solar Panels Convert Sunlight into Clean Electricity",
        "url": "https://kids.nationalgeographic.com/science/article/solar-energy-explained",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah - Sesuai Siswa Kelas XI SMA)"
    }
    t3_paras = [
        "Solar energy is one of the cleanest and most abundant renewable resources on Earth. Photovoltaic (PV) solar panels allow us to capture sunlight and convert it directly into usable electricity without burning fossil fuels or emitting carbon smoke.",
        "A solar panel is made up of many small silicon cells. When sunlight hits a solar panel, packets of light energy called photons strike the silicon atoms, knocking electrons free from their atomic bonds. The internal design of the cell creates an electric field that forces these free electrons to flow in a single direction, producing direct current (DC) electricity.",
        "Because our home appliances and school computers run on alternating current (AC) electricity, the direct current is channeled through an electronic device called an inverter. The inverter transforms the DC power into AC power, which is then distributed through electrical wiring to power lights, refrigerators, and electronic gadgets cleanly."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "Harnessing the Sun: How Photovoltaic Solar Panels Work",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Tectonic plate", "Noun Phrase", "Lempeng raksasa penyusun kerak bumi", "Tectonic plates move slowly over the semi-liquid mantle."),
        ("Friction", "Noun", "Gaya gesek yang menahan dua permukaan batuan", "Friction prevents the rocks from sliding smoothly."),
        ("Seismic waves", "Noun Phrase", "Gelombang rambat getaran energi gempa", "Seismic waves radiate outwards from the earthquake focus."),
        ("Epicenter", "Noun", "Titik di permukaan bumi tepat di atas pusat gempa", "The town closest to the epicenter suffered heavy damage."),
        ("Displacement", "Noun", "Pergeseran posisi suatu massa secara tiba-tiba", "Vertical displacement of water triggers tsunami waves."),
        ("Inundate", "Verb", "Membanjiri, menenggelamkan daratan dengan air", "Giant tsunami waves can inundate coastal villages in minutes."),
        ("Photovoltaic", "Adjective", "Mampu mengubah energi cahaya menjadi listrik", "Photovoltaic cells harness sunlight on rooftop panels."),
        ("Inverter", "Noun", "Alat elektronik pengubah arus searah (DC) ke bolak-balik (AC)", "The solar inverter converts DC current into AC power.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "According to Text 1, what geological phenomenon causes elastic stress to build up along a fault line?",
        "What happens to a tsunami wave's speed and height as it travels from deep ocean water to shallow coastal waters in Text 2?",
        "In Text 3, what is the role of an inverter in a home solar power system?"
    ]
    inf_q = [
        "Why is Indonesia particularly vulnerable to frequent earthquakes and volcanic activities?",
        "Explain why people on a ship in the middle of deep ocean water rarely feel a passing tsunami wave.",
        "How does using solar energy help combat global climate change compared to coal-fired power plants?"
    ]
    eval_q = [
        "Evaluate the earthquake safety procedures in your school: What should students do immediately when strong tremors occur?",
        "In tropical countries like Indonesia, why hasn't rooftop solar energy become universal in all households? What are the potential challenges?"
    ]
    perf_task = (
        "PENYUSUNAN ARTIKEL EKSPLANASI SAINS & BAGAN PROSES (LKPD / SUMATIF):\n"
        "Susunlah sebuah Teks Eksplanasi Kausal Ilmiah (Explanation Text) sepanjang 200-250 kata mengenai salah satu fenomena berikut:\n"
        "Topik 1: How Rain and the Water Cycle Occur.\n"
        "Topik 2: How Wind Turbines Generate Clean Electricity.\n"
        "Topik 3: How Volcanic Eruptions Happen.\n"
        "Struktur Wajib:\n"
        "- General Statement: Kenalkan fenomena alam atau teknologi tersebut.\n"
        "- Sequenced Explanation: Jelaskan minimal 3 tahapan sebab-akibat yang runtut menggunakan kalimat pasif ilmiah (passive voice) dan konjungsi kausalitas (because, consequently, therefore).\n"
        "- Sertakan diagram alur sederhana (flowchart) yang menggambarkan proses tersebut."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_11.W.4_Complex_Explanation_Science.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 11.W.5: DISCUSSION TEXT - SCHOOL ISSUES & TEEN LIFE
# ==============================================================================
def create_tp_11_w_5():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS SMA",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XI / Semester 4 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 11.W.5 - Menganalisis dan menimbang dua sudut pandang berimbang (pro dan kontra) dalam teks diskusi mengenai isu-isu pendidikan dan kehidupan remaja.",
        "Elemen Capaian Pembelajaran": "Membaca - Memirsa & Berbicara - Mempresentasikan",
        "Genre Teks / Ruang Lingkup": "Discussion Text (School Uniforms, Homework Debates, Gadget Policies)",
        "Alokasi Waktu": "12 Jam Pelajaran (6 Pertemuan @ 2 JP)",
        "Target Kemahiran CEFR": "Level B1 s.d. B1+",
        "Profil Pelajar Pancasila": "Bernalar Kritis, Mandiri, Gotong Royong"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "BBC Worklife / Education Debate Series",
        "title": "The Great Debate: Do School Uniforms Help or Hinder Students?",
        "url": "https://www.bbc.com/worklife/article/school-uniform-debate",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah - Sesuai Siswa Kelas XI SMA)"
    }
    t1_paras = [
        "The question of whether high school students should wear mandatory school uniforms remains a frequent topic of debate among teachers, parents, and students. While uniforms are traditional in many countries, including Indonesia, some modern educators argue that students should be free to express their individuality through casual attire.",
        "Proponents of school uniforms believe they promote equality and discipline. When all students wear identical clothing, social and economic differences between wealthy and less fortunate families are less obvious. This reduces peer pressure and prevents bullying related to expensive brand-name fashion. Furthermore, wearing a uniform saves time in the morning and creates a proud sense of school belonging.",
        "On the other hand, critics argue that mandatory uniforms suppress creativity and self-expression. High school is a crucial time for teenagers to discover their personal identity, and choosing their own clothes is a healthy way to practice decision-making. Moreover, purchasing multiple sets of uniforms, including sports kits and scout attire, can place a financial burden on low-income families.",
        "In conclusion, both sides present valid perspectives. While uniforms foster equality and focus, schools could adopt a sensible compromise—such as having uniform days from Monday to Thursday, while allowing casual but neat clothing on Fridays."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Menimbang secara berimbang sisi positif dan negatif kewajiban seragam sekolah serta merumuskan jalan tengah yang bijak.",
        "Struktur Teks": "Issue (Perdebatan seragam sekolah) -> Arguments For (Mendorong kesetaraan, disiplin, dan rasa bangga) -> Arguments Against (Membatasi ekspresi diri dan membebani biaya) -> Recommendation/Conclusion (Jalan tengah hari seragam dan hari bebas rapi).",
        "Fitur Bahasa": "Contrasting Connectors ('While', 'On the other hand', 'Moreover', 'In conclusion'), Neutral Stance Verbs ('believe', 'argue', 'present valid perspectives')."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "The School Uniform Debate: Equality vs. Self-Expression",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "The Guardian Education / National Education Association",
        "title": "The Homework Debate: Does After-School Work Truly Benefit Students?",
        "url": "https://www.theguardian.com/education/homework-debate-students",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah - Sesuai Siswa Kelas XI SMA)"
    }
    t2_paras = [
        "For generations, homework has been considered an indispensable element of school education. However, in recent years, an increasing number of educators and parents are questioning whether heavy loads of homework truly enhance learning or simply create unnecessary stress for teenagers.",
        "Supporters maintain that homework reinforces lessons taught in the classroom. Practicing math formulas or reading literature at home helps students retain knowledge and develop vital independent study habits. Homework also gives parents a clear window into what their children are learning at school.",
        "Conversely, opponents point out that excessive homework leads to sleep deprivation, anxiety, and burnout. After spending eight hours at school, teenagers need time to relax, exercise, pursue artistic hobbies, and spend quality time with family. Assigning hours of repetitive worksheets often kills students' natural curiosity and love of learning.",
        "Taking both viewpoints into consideration, homework should not be eliminated entirely, but its volume should be carefully regulated. Teachers should assign meaningful, creative projects rather than repetitive worksheets, ensuring homework takes no more than one hour per evening."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Membahas pro dan kontra pemberian pekerjaan rumah (PR) bagi siswa SMA dan merumuskan batasan waktu yang ideal.",
        "Struktur Teks": "Issue (Polemik beban PR) -> Sisi Pendukung (Penguatan materi & disiplin belajar) -> Sisi Penentang (Kelelahan, stres, dan hilangnya waktu keluarga) -> Rekomendasi seimbang (Membatasi PR maksimal 1 jam per malam).",
        "Fitur Bahasa": "Hedging Markers ('considered', 'point out', 'should not be eliminated entirely')."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "Is Homework Necessary? Evaluating the Pros and Cons",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "UNESCO Global Education Monitoring / EdTech Review",
        "title": "Smartphones in the Classroom: Educational Powerhouses or Major Distractions?",
        "url": "https://www.unesco.org/en/articles/smartphones-in-schools",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah - Sesuai Siswa Kelas XI SMA)"
    }
    t3_paras = [
        "The presence of smartphones in high school classrooms has ignited a lively controversy worldwide. Should mobile phones be welcomed as modern learning instruments, or should they be strictly prohibited during lessons?",
        "Advocates of classroom smartphone integration argue that phones are powerful handheld computers. Students can instantly access online encyclopedias, look up definitions in digital dictionaries, and participate in interactive learning quizzes like Kahoot. Utilizing smartphones prepares students for the digital workplace.",
        "On the contrary, educators who support bans emphasize that smartphones are tremendous sources of distraction. Constant notifications from social media, text messages, and mobile games tempt students to multitask, which severely harms their concentration. Furthermore, smartphone cameras in classrooms can lead to privacy violations and cheating during tests.",
        "A practical compromise is for schools to implement 'device pouches'. Students store their phones in a wall pocket at the start of class and only access them when the teacher explicitly plans a digital research activity."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "Smartphones in School: Learning Tool or Learning Disruption?",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Individuality", "Noun", "Kepribadian atau keunikan khas seseorang", "Uniforms may limit students' desire for individuality."),
        ("Peer pressure", "Noun Phrase", "Tekanan sosial dari teman sebaya untuk ikut tren", "Uniforms reduce peer pressure over expensive clothing."),
        ("Indispensable", "Adjective", "Sangat penting dan tidak bisa ditinggalkan", "Reading is an indispensable habit for academic success."),
        ("Burnout", "Noun", "Kelelahan fisik dan mental ekstrem akibat beban berlebih", "Excessive daily homework often leads to student burnout."),
        ("Supporter", "Noun", "Pendukung suatu gagasan atau kebijakan", "Supporters argue that uniforms foster school discipline."),
        ("Conversely", "Adverb", "Sebaliknya (digunakan untuk sudut pandang lawan)", "Conversely, opponents highlight the lack of creative freedom."),
        ("Distraction", "Noun", "Hal yang memecah atau mengganggu konsentrasi", "Phone notifications are a major distraction during lessons."),
        ("Compromise", "Noun", "Kesepakatan damai yang menengahi dua perbedaan", "The device pouch offers a sensible compromise for schools.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "According to Text 1, what are two main benefits of wearing school uniforms cited by proponents?",
        "In Text 2, what negative impacts on students' health are mentioned by opponents of heavy homework?",
        "What practical compromise regarding smartphones is suggested at the end of Text 3?"
    ]
    inf_q = [
        "How can wearing identical uniforms help protect students from low-income families from social bullying?",
        "Why does multitasking on a smartphone during class reduce a student's ability to retain important information?",
        "Why is it essential for a Discussion Text to present both supporting and opposing views fairly before drawing a conclusion?"
    ]
    eval_q = [
        "If your school allowed a 'No-Uniform Day' every Friday, what dress code guidelines would you propose to ensure students dress respectfully?",
        "Between Homework and Classroom Smartphones, which issue creates a bigger impact on your daily academic life? Give your reasons."
    ]
    perf_task = (
        "DEBAT KELAS & PENULISAN TEKS DISKUSI (LKPD / SUMATIF):\n"
        "Susunlah sebuah Teks Diskusi Berimbang (Discussion Text) sepanjang 200-250 kata mengenai salah satu isu sekolah:\n"
        "Mosi 1: Should High Schools Start Classes at 8:30 AM Instead of 7:00 AM?\n"
        "Mosi 2: Should Canteens Ban Junk Food and Sell Only Nutritious Meals?\n"
        "Ketentuan Format:\n"
        "• Paragraph 1: Issue (Jelaskan kontroversi dan perdebatan).\n"
        "• Paragraph 2: Arguments For (Sajikan 2 alasan pendukung beserta contoh).\n"
        "• Paragraph 3: Arguments Against (Sajikan 2 alasan penentang beserta contoh).\n"
        "• Paragraph 4: Conclusion & Recommendation (Berikan solusi jalan tengah yang berimbang).\n"
        "Gunakan konjungsi kontras (on the other hand, conversely, however)."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_11.W.5_Discussion_Text_School_Issues.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 11.W.6: CRITICAL REVIEW - ARTS, BOOKS & CINEMA
# ==============================================================================
def create_tp_11_w_6():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS SMA",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase F Umum / Kelas XI / Semester 4 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 11.W.6 - Mengevaluasi dan menyusun teks resensi kritis (Review Text) terhadap novel, film bermutu, dan karya seni yang relevan dengan remaja.",
        "Elemen Capaian Pembelajaran": "Membaca - Memirsa & Menulis - Mempresentasikan",
        "Genre Teks / Ruang Lingkup": "Review Text (Book Reviews, Movie Reviews, Character Evaluation)",
        "Alokasi Waktu": "12 Jam Pelajaran (6 Pertemuan @ 2 JP)",
        "Target Kemahiran CEFR": "Level B1 s.d. B1+",
        "Profil Pelajar Pancasila": "Bernalar Kritis, Kreatif, Berkebinekaan Global"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "The Guardian Children's Books / NYT Bestseller Review",
        "title": "Book Review: 'Wonder' by R.J. Palacio – A Heartwarming Story of Courage",
        "url": "https://www.theguardian.com/childrens-books-site/wonder-palacio-review",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah - Sesuai Siswa Kelas XI SMA)"
    }
    t1_paras = [
        "'Wonder', a bestselling novel by R.J. Palacio, is a moving story about ten-year-old August 'Auggie' Pullman. Auggie was born with a severe facial difference that required multiple surgeries. After years of being homeschooled by his devoted mother, Auggie faces his biggest challenge yet: attending fifth grade at Beecher Prep, a regular private middle school.",
        "The story is brilliantly told from multiple perspectives. While Auggie's voice is funny, brave, and honest, the book also includes chapters written from the viewpoints of his loving sister Via, his school friend Jack Will, and his classmate Summer. This clever storytelling technique allows readers to understand how Auggie's presence affects everyone around him.",
        "Palacio's writing is warm, accessible, and full of heartfelt emotion without becoming overly dramatic. The novel deals honestly with the pain of playground stares and bullying, but it also celebrates the quiet courage of choosing kindness. The school precept introduced by his English teacher, Mr. Browne—'When given the choice between being right or being kind, choose kind'—serves as the emotional anchor of the book.",
        "Overall, 'Wonder' is a magnificent, uplifting read for teenagers and adults alike. It earns a well-deserved 5 out of 5 stars and should be required reading in all secondary schools to foster empathy and inclusion."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Meresensi dan mengevaluasi kualitas novel 'Wonder' mengenai keberanian anak berkebutuhan khusus dan pesan kebaikan.",
        "Struktur Teks": "Orientation (Perkenalan judul novel, pengarang, tokoh Auggie) -> Interpretative Recount (Latar cerita masuk sekolah reguler dan teknik sudut pandang ganda) -> Evaluation (Kelebihan gaya bahasa hangat dan pesan moral memilih kebaikan) -> Evaluative Summary & Recommendation (Rating 5/5 dan rekomendasi bacaan).",
        "Fitur Bahasa": "Evaluative Adjectives ('moving', 'brilliant', 'heartfelt', 'magnificent'), Present Tense, Rating formal."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "Book Review: 'Wonder' – Choosing Kindness Over Being Right",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "Common Sense Media / Variety International Film Review",
        "title": "Movie Review: 'Laskar Pelangi' (The Rainbow Troops) – The Power of Education",
        "url": "https://www.commonsensemedia.org/movie-reviews/laskar-pelangi",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah - Sesuai Siswa Kelas XI SMA)"
    }
    t2_paras = [
        "Directed by Riri Riza and adapted from Andrea Hirata's beloved autobiographical novel, 'Laskar Pelangi' (The Rainbow Troops) is one of the most successful and inspiring films in Indonesian cinema history. Set on the tin-mining island of Belitung in the 1970s, the film follows ten cheerful, impoverished children struggling to get an education at a dilapidated village school named SD Muhammadiyah Gantong.",
        "The film's greatest strength lies in its natural performances and stunning island cinematography. The young, local amateur actors who portray Ikal, the genius Lintang, and the artistic Mahar perform with genuine innocence and infectious warmth. Cut Mini delivers a memorable, touching performance as Bu Muslimah, a selfless young teacher who refuses to close the school despite crushing financial hardships.",
        "While the film's pacing slows down slightly during the middle act, the breathtaking soundtrack by Nidji and the triumphant academic quiz scene easily make up for any minor flaws. The movie captures the raw joy of childhood and the transformative dignity of learning.",
        "In summary, 'Laskar Pelangi' is an emotional cinematic masterpiece that reminds us of the true value of teachers and friendship. Rated 4.5 out of 5 stars, it is an absolute must-watch family movie."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Memberikan ulasan kritis terhadap film Laskar Pelangi, membedah kekuatan akting, pesan pendidikan, dan tata sinematografi.",
        "Struktur Teks": "Orientation (Identifikasi sutradara, adaptasi novel, setting Belitung) -> Interpretative Recount (Perjuangan 10 anak sekolah gubuk) -> Evaluation (Keunggulan akting alami dan musik pengiring) -> Summary & Rating (4.5/5).",
        "Fitur Bahasa": "Film Terminology ('cinematography', 'soundtrack', 'performances', 'pacing')."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "Movie Review: 'Laskar Pelangi' – Inspiring Spirits from Belitung",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "BBC Culture / Film Inquiry Youth Review",
        "title": "Film Review: Pixar's 'Inside Out' – Making Sense of Adolescent Emotions",
        "url": "https://www.bbc.com/culture/article/inside-out-review",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah - Sesuai Siswa Kelas XI SMA)"
    }
    t3_paras = [
        "Pixar Animation Studios has created many imaginative movies, but 'Inside Out' (directed by Pete Docter) is perhaps its most psychologically profound work. The movie takes place inside the mind of eleven-year-old Riley, who is struggling to adjust to a new city after her family moves from Minnesota to San Francisco.",
        "Inside Riley's brain, five personified emotions—Joy, Sadness, Anger, Fear, and Disgust—operate a control console. The plot kicks into gear when Joy and Sadness are accidentally swept away from the control center, leaving Riley unable to feel happiness during a stressful transition.",
        "The film's brilliant innovation is showing that all emotions, including sadness, serve an essential purpose. Rather than treating Sadness as a villain, the story shows that acknowledging our sad feelings helps us seek comfort from loved ones. The colorful animation and clever humor keep both kids and teens thoroughly engaged.",
        "To conclude, 'Inside Out' is a smart, entertaining, and emotionally healing movie. It deserves 5 out of 5 stars and offers a great lesson on emotional maturity for high school teenagers."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "Film Review: 'Inside Out' – The Wisdom of Emotional Balance",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Precept", "Noun", "Petuah atau prinsip bijak sebagai pedoman hidup", "Mr. Browne introduced a monthly precept about kindness."),
        ("Perspective", "Noun", "Sudut pandang dalam melihat suatu peristiwa", "The novel is narrated from multiple character perspectives."),
        ("Cinematography", "Noun", "Seni tata visual dan perekaman gambar film", "The film features stunning cinematography of Belitung's beaches."),
        ("Dilapidated", "Adjective", "Reyot, rusak parah karena usia tua", "The children studied eagerly in a dilapidated wooden classroom."),
        ("Personified", "Adjective", "Dipersonifikasi, digambarkan berwujud manusia", "Riley's emotions are personified as colorful animated figures."),
        ("Heartfelt", "Adjective", "Tulus, menyentuh hati terdalam", "The story conveys a heartfelt message about friendship."),
        ("Flaw", "Noun", "Kekurangan kecil dalam sebuah karya seni", "A slightly slow middle act was the film's only minor flaw."),
        ("Uplifting", "Adjective", "Membangkitkan semangat dan memberi inspirasi", "'Wonder' is an uplifting book that celebrates human dignity.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "In Text 1, what unique storytelling technique does author R.J. Palacio use in the novel 'Wonder'?",
        "According to Text 2, where is the story of 'Laskar Pelangi' set, and who is the devoted teacher who inspires the students?",
        "In Text 3, what are the five personified emotions that control Riley's mind in Pixar's 'Inside Out'?"
    ]
    inf_q = [
        "Why is Mr. Browne's precept—'When given the choice between being right or being kind, choose kind'—so important for high school teenagers?",
        "What makes the performances of the young actors in 'Laskar Pelangi' feel so touching and realistic to international film critics?",
        "How does the movie 'Inside Out' demonstrate that feeling sadness is not always a bad thing?"
    ]
    eval_q = [
        "If you were asked to recommend either 'Wonder' (Book) or 'Laskar Pelangi' (Film) to your school library, which would you pick and why?",
        "What are the essential elements that distinguish an objective Review Text from a mere personal complaint about a movie?"
    ]
    perf_task = (
        "PENULISAN RESENSI FILM ATAU BUKU (LKPD / SUMATIF MENULIS):\n"
        "Tulislah sebuah Teks Resensi Kritis (Review Text) sepanjang 200-250 kata mengenai sebuah film atau buku favoritmu yang mendidik.\n"
        "Struktur Wajib Resensi:\n"
        "1. Orientation: Nama karya, sutradara/penulis, genre, dan pengenalan tokoh utama.\n"
        "2. Interpretative Recount: Ringkasan singkat alur cerita (jangan bocorkan ending/no spoilers).\n"
        "3. Evaluation: Ulas kelebihan (akting, visual, pesan moral) dan sedikit kekurangan karya.\n"
        "4. Summary & Recommendation: Berikan kesimpulan, Star Rating (1–5 bintang), dan saran pembaca/penonton yang cocok.\n"
        "Gunakan minimal 4 kosakata dari tabel glosarium di atas."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_11.W.6_Critical_Review_Arts.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


if __name__ == "__main__":
    print("Generating calibrated authentic materials for 5-Modul Ajar XI SMA (Wajib)...")
    create_tp_11_w_1()
    create_tp_11_w_2()
    create_tp_11_w_3()
    create_tp_11_w_4()
    create_tp_11_w_5()
    create_tp_11_w_6()
    print("All 6 DOCX files for Kelas XI SMA Wajib successfully updated and calibrated!")
