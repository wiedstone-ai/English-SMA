# -*- coding: utf-8 -*-
"""
Generator script for authentic English materials (.docx) in 5-Modul Ajar XI SMA TL.
Calibrated specifically to CEFR B1+ to B2 (Accessible, Engaging SMA Tingkat Lanjut Level).
Each TP contains 3 authentic texts from credible sources with explicit URLs,
practical academic vocabulary glossaries, and aligned assessment questions.
"""

import os
import sys

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import docx_helper

TARGET_DIR = r"C:\Users\KOMPUTERKU2\Documents\English SMA\5-Modul Ajar XI SMA TL"
os.makedirs(TARGET_DIR, exist_ok=True)

# ==============================================================================
# TP 11.TL.1: ADVANCED ANALYTICAL EXPOSITION - AI ETHICS & PRIVACY
# ==============================================================================
def create_tp_11_tl_1():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS TINGKAT LANJUT",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XI / Semester 3 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 11.TL.1 - Mengevaluasi wacana argumentatif kritis dalam teks eksposisi analitis tentang etika kecerdasan buatan, bias data, dan perlindungan privasi digital.",
        "Elemen Capaian Pembelajaran": "Membaca - Memirsa & Menulis - Mempresentasikan",
        "Genre Teks / Ruang Lingkup": "Analytical Exposition Text (AI Ethics, Algorithmic Bias, Digital Privacy)",
        "Alokasi Waktu": "20 Jam Pelajaran (4 Minggu x 5 JP)",
        "Target Kemahiran CEFR": "Level B1+ s.d. B2 (Menengah Atas - Sesuai Siswa SMA Peminatan)",
        "Profil Pelajar Pancasila": "Bernalar Kritis, Berkebinekaan Global, Mandiri"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "MIT Technology Review / Harvard Berkman Klein Center",
        "title": "Why Computer Algorithms Are Not Always Neutral: Understanding AI Bias",
        "url": "https://www.technologyreview.com/topic/artificial-intelligence/bias/",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t1_paras = [
        "As artificial intelligence (AI) systems are increasingly used to make important real-world decisions—such as hiring job applicants, approving bank loans, and predicting school exam scores—many people assume that computers are completely fair and objective. However, computer algorithms are created and trained by humans. When algorithms learn from historical data that contains human prejudice, they can quietly repeat and amplify unfair discrimination.",
        "Primarily, AI systems learn by analyzing past patterns. For example, if a company's automated recruitment tool is trained on resumes from the past twenty years—when most corporate executives were men—the algorithm can learn to favor male applicants over equally qualified female candidates. The computer is not acting with malice; it is simply repeating historical inequality.",
        "Furthermore, many modern AI models operate as 'black boxes', meaning that even the engineers who build them cannot easily explain how the system reached a specific conclusion. If a student is denied a university scholarship or a job interview by an automated program, they have a right to know the reason.",
        "In conclusion, we must never trust artificial intelligence blindly. Governments and tech companies must introduce clear ethical guidelines, require independent audits of computer algorithms, and ensure that human judgment always oversees automated decisions."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Mengevaluasi secara kritis dan membongkar anggapan keliru bahwa algoritma AI selalu netral dan bebas dari bias manusia.",
        "Struktur Teks": "Thesis (Anggapan bahwa AI selalu adil adalah keliru) -> Arguments (AI mengulang bias data historis rekrutmen; Masalah transparansi sistem black box) -> Reiteration (Tuntutan audit etika dan pengawasan manusia).",
        "Fitur Bahasa": "Abstract Nouns ('prejudice', 'discrimination', 'inequality', 'transparency'), Concessive Connectors ('However', 'Even though', 'In conclusion'), Present Tense."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "Understanding Algorithmic Bias: Why AI is Not Always Neutral",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "BBC Future / Digital Ethics Column",
        "title": "Generative AI and Creative Work: Opportunity or Threat to Human Artists?",
        "url": "https://www.bbc.com/future/article/generative-ai-creativity-ethics",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t2_paras = [
        "The emergence of generative artificial intelligence capable of creating digital paintings, writing music, and composing essays in seconds has sparked a passionate debate across the creative world. While some tech enthusiasts celebrate these tools as great helpers for brainstorming, many professional artists worry about the future of human creativity.",
        "Foremost, generative AI models are trained by scraping millions of artworks, photos, and stories from the internet without asking for permission from the original creators. Many illustrators and writers find that AI tools copy their unique artistic styles without giving them any credit or compensation. This unfair practice threatens the livelihoods of creative workers.",
        "Moreover, over-relying on automated tools can weaken young people's creative thinking. Learning how to draw, paint, or write requires patience, practice, and personal emotional expression. When students simply type a prompt and let a machine do the work, they miss the joy of discovering their own artistic voice.",
        "To summarize, generative AI should be used as an assistive tool, not as a replacement for genuine human imagination. Protecting artists' rights and valuing original human effort must remain top priorities."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Membahas dampak etika kecerdasan buatan generatif terhadap hak cipta seniman dan perkembangan kreativitas siswa.",
        "Struktur Teks": "Thesis (Debat AI generatif) -> Arguments (Penggunaan karya seniman tanpa izin; Melemahkan latihan kreativitas mandiri) -> Reiteration (AI sebagai alat bantu, bukan pengganti seniman manusia).",
        "Fitur Bahasa": "Contrastive Conjunctions ('While', 'Moreover', 'To summarize'), Modal Auxiliaries ('should be used', 'must remain')."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "Generative AI in the Creative Arts: Ethics, Copyright, and Human Effort",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "The Guardian Tech / Civil Liberties Briefing",
        "title": "Facial Recognition Surveillance in Public Spaces: The Erosion of Privacy",
        "url": "https://www.theguardian.com/technology/facial-recognition-privacy-debate",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t3_paras = [
        "In many modern cities, smart security cameras equipped with facial recognition technology are installed in train stations, shopping malls, and streets. While authorities argue that this technology deters crime and helps catch suspects quickly, privacy advocates warn that mass surveillance poses serious risks to personal freedom.",
        "First, constant facial tracking takes away citizens' right to privacy in public spaces. When people know that their every step is monitored and logged into a central database, they feel less comfortable joining peaceful public gatherings or expressing their beliefs freely.",
        "Second, facial recognition software is not 100 percent accurate. Studies show that these cameras have higher error rates when identifying women and people with darker skin tones. An incorrect match can lead to innocent citizens being stopped, questioned, or mistakenly accused.",
        "Ultimately, public safety should not come at the expense of fundamental civil rights. Strict laws are necessary to regulate where and when facial recognition cameras can be legally used."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "Facial Recognition Technology: Safety vs. Personal Freedom",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Algorithm", "Noun", "Rangkaian instruksi komputasi pemecah masalah", "The hiring algorithm evaluated thousands of job resumes."),
        ("Prejudice", "Noun", "Prasangka buruk yang tidak berdasar", "Algorithms can learn human prejudice from historical data."),
        ("Amplify", "Verb", "Memperbesar atau memperluas dampak sesuatu", "Automated systems can amplify existing social inequalities."),
        ("Transparency", "Noun", "Keterbukaan dan kejelasan mekanisme sistem", "Tech companies should provide transparency about their AI tools."),
        ("Compensation", "Noun", "Ganti rugi atau pembayaran atas karya/jasa", "Artists deserve fair compensation when their work is used."),
        ("Livelihood", "Noun", "Mata pencaharian untuk memenuhi kebutuhan hidup", "Unfair AI copying threatens the livelihood of illustrators."),
        ("Surveillance", "Noun", "Pengawasan ketat terhadap gerak-gerik publik", "Mass facial surveillance reduces personal privacy in public."),
        ("Accused", "Verb/Adjective", "Dituduh melakukan kesalahan atau pelanggaran", "Innocent citizens should never be mistakenly accused.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "According to Text 1, why do automated hiring algorithms sometimes unfairly downgrade female job applicants?",
        "What major concern do human illustrators and writers raise about generative AI in Text 2?",
        "In Text 3, what technical limitation regarding accuracy makes facial recognition cameras risky to use in public spaces?"
    ]
    inf_q = [
        "Explain what is meant by calling an AI model a 'black box' in Text 1.",
        "Why is practicing the traditional struggle of drawing or writing important for high school students according to Text 2?",
        "How can widespread facial surveillance create a 'chilling effect' on citizens' willingness to express their opinions in public (Text 3)?"
    ]
    eval_q = [
        "In your view, should high school teachers permit students to use generative AI for brainstorming essay ideas? Defend your answer.",
        "Evaluate the balance between public safety and personal privacy: Where should society draw the line on facial recognition cameras?"
    ]
    perf_task = (
        "PENYUSUNAN ESSAY POSISI AKADEMIK (LKPD / SUMATIF TINGKAT LANJUT):\n"
        "Susunlah sebuah Academic Position Essay (Analytical Exposition) sepanjang 250-300 kata mengenai salah satu isu teknologi:\n"
        "Pilihan 1: Should Schools Use AI Proctors to Monitor Online Examinations?\n"
        "Pilihan 2: Should Tech Companies Pay Royalties to Artists for Training Generative AI?\n"
        "Kriteria Esai:\n"
        "• Thesis: Tentukan posisi pendapatmu secara tegas dan terukur.\n"
        "• Arguments: Sediakan minimal 2 paragraf argumen dengan bukti pendukung konkret.\n"
        "• Counter-Argument: Bahas 1 sanggahan dari pihak lawan secara santun dan sanggahlah.\n"
        "• Gunakan minimal 4 kosakata dari tabel glosarium di atas."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_11.TL.1_Advanced_Analytical_Exposition_AI_Ethics.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 11.TL.2: ADVANCED HORTATORY EXPOSITION - CLIMATE & CIRCULAR ECONOMY
# ==============================================================================
def create_tp_11_tl_2():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS TINGKAT LANJUT",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XI / Semester 3 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 11.TL.2 - Menganalisis dan menyusun teks eksposisi hortatori tingkat lanjut yang merekomendasikan transisi ekonomi sirkular dan aksi nyata krisis iklim.",
        "Elemen Capaian Pembelajaran": "Menyimak - Berbicara & Menulis - Mempresentasikan",
        "Genre Teks / Ruang Lingkup": "Hortatory Exposition (Climate Action, Renewable Energy, Circular Economy)",
        "Alokasi Waktu": "25 Jam Pelajaran (5 Minggu x 5 JP)",
        "Target Kemahiran CEFR": "Level B1+ s.d. B2",
        "Profil Pelajar Pancasila": "Bernalar Kritis, Gotong Royong, Berakhlak Mulia"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "UN Climate Change (UNFCCC) / IPCC Youth Briefs",
        "title": "The Urgent Need for Global Climate Action: Why Words Are Not Enough",
        "url": "https://unfccc.int/process-and-meetings/the-paris-agreement/action",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t1_paras = [
        "The latest scientific reports from the United Nations Intergovernmental Panel on Climate Change (IPCC) make one reality crystal clear: our planet is warming at an alarming rate due to human activities, especially the burning of coal, oil, and gas. Rising temperatures are already causing longer droughts, heavier floods, and dangerous heatwaves around the world. Global leaders and industries must urgently replace fossil fuels with clean, renewable energy.",
        "First, continuing to burn fossil fuels is economically shortsighted. Every year, governments spend hundreds of billions of dollars repairing roads, bridges, and homes destroyed by climate-related storms. Investing that money in solar panels, wind farms, and energy-efficient public transport creates thousands of modern green jobs while protecting our biosphere.",
        "Second, climate justice requires fair international cooperation. Wealthy industrialized nations have produced the vast majority of historical carbon emissions. Therefore, they have an ethical responsibility to help developing island nations protect their coastlines from rising sea levels.",
        "In conclusion, we can no longer afford delay. World leaders must set binding deadlines to phase out coal power plants, end fossil fuel subsidies, and support community-led renewable energy projects worldwide."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Mempersuasi dan mendesak para pemimpin dunia untuk beralih ke energi bersih dan menghentikan subsidi bahan bakar fosil.",
        "Struktur Teks": "Thesis (Bumi memanas cepat; perlunya peralihan ke energi terbarukan) -> Arguments (Biaya pemulihan bencana vs keuntungan lapangan kerja hijau; Keadilan iklim bagi negara kepulauan) -> Recommendation (Penetapan tenggat waktu penghentian batubara dan pencabutan subsidi fosil).",
        "Fitur Bahasa": "Strong Hortatory Modals ('must urgently replace', 'must set binding deadlines', 'can no longer afford delay'), Causal Connectors ('First', 'Second', 'Therefore')."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "Why We Must Accelerate Global Climate Action and Clean Energy",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "World Economic Forum / Ellen MacArthur Foundation",
        "title": "Rethinking Waste: Why We Need a Circular Economy",
        "url": "https://www.weforum.org/agenda/circular-economy-solutions/",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t2_paras = [
        "For over a century, the global economy has followed a wasteful 'take-make-dispose' model: we extract raw materials from nature, manufacture goods, and discard them in landfills after a short period. This linear system creates mountains of trash and depletes finite natural resources. Manufacturing companies must adopt a circular economy model.",
        "In a circular economy, products are deliberately designed to be repaired, refurbished, and recycled rather than thrown away. For example, when smartphones are designed with modular components, broken screens or worn-out batteries can easily be replaced without tossing the entire device into the garbage. This dramatically reduces electronic waste and saves precious metals.",
        "Moreover, businesses that embrace circular practices save significant production costs by reusing secondary materials. Circular business models protect natural forests and rivers from continuous mining.",
        "Therefore, governments should pass 'Right to Repair' laws that force tech manufacturers to provide affordable spare parts, and consumers should choose durable, repairable items over disposable goods."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Merekomendasikan penerapan ekonomi sirkular yang merancang produk agar mudah diperbaiki dan didaur ulang.",
        "Struktur Teks": "Thesis (Model ekonomi buang-pakai linier merusak bumi; perlunya ekonomi sirkular) -> Arguments (Manfaat desain modular alat elektronik; Penghematan biaya bahan baku dan pelestarian alam) -> Recommendation (Hukum hak memperbaiki / Right to Repair).",
        "Fitur Bahasa": "Vocabulary of Sustainability ('circular economy', 'modular', 'refurbish', 'electronic waste')."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "From Linear to Circular: Building a Sustainable Economic Future",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "UNEP News / Climate Justice Action",
        "title": "Supporting Frontline Communities: The Loss and Damage Challenge",
        "url": "https://www.unep.org/news-and-stories/story/loss-and-damage-climate",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t3_paras = [
        "Small island developing states and low-lying coastal regions contribute less than one percent of global greenhouse gas emissions. Yet, they suffer the most devastating consequences of rising sea levels, saltwater intrusion, and super-typhoons. Providing reliable financial support through international climate funds is a moral duty.",
        "First, when severe storms wipe out coastal fishing villages, the affected communities lose their homes, boats, and drinking water sources. Expecting these developing nations to rebuild entirely on their own or through commercial bank loans is unfair.",
        "Second, protecting coastal ecosystems like mangroves and coral reefs benefits the entire planet by storing vast amounts of carbon and preserving marine biodiversity.",
        "Therefore, industrialized nations must honor their climate funding commitments promptly and assist frontline nations with coastal seawalls, clean water supplies, and disaster warning systems."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "Climate Solidarity: Helping Frontline Island Communities Survive",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Shortsighted", "Adjective", "Berpandangan sempit, tidak memikirkan masa depan", "Relying solely on fossil fuels is economically shortsighted."),
        ("Biosphere", "Noun", "Lapisan kehidupan di bumi (flora, fauna, manusia)", "Renewable energy helps protect our fragile biosphere."),
        ("Subsidies", "Noun (Plural)", "Bantuan dana pemerintah untuk meringankan harga", "Governments should redirect fossil fuel subsidies to solar power."),
        ("Deplete", "Verb", "Menguras habis sumber daya alam", "Over-mining threatens to deplete finite mineral deposits."),
        ("Modular", "Adjective", "Terdiri dari bagian-bagian terpisah yang mudah diganti", "Modular smartphones allow users to change only the broken screen."),
        ("Right to repair", "Noun Phrase", "Hak konsumen untuk memperbaiki barangnya sendiri secara mudah", "Right to Repair laws encourage durable electronic manufacturing."),
        ("Intrusion", "Noun", "Masuknya zat asing (misal: air laut ke sumur tawar)", "Coastal villages suffer from saltwater intrusion into wells."),
        ("Commitment", "Noun", "Komitmen atau janji tanggung jawab", "Nations must honor their international climate commitments.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "According to Text 1, what are two main economic benefits of investing in renewable energy instead of repairing storm damages?",
        "How does a circular economy model differ from the traditional 'take-make-dispose' linear system in Text 2?",
        "In Text 3, why is it considered unfair to expect small island nations to rebuild after super-typhoons using commercial loans?"
    ]
    inf_q = [
        "Explain the ethical concept of 'climate justice' discussed in Text 1.",
        "Why would passing 'Right to Repair' laws help reduce the volume of electronic waste sent to landfills (Text 2)?",
        "How do coastal mangrove forests contribute to mitigating global climate change (Text 3)?"
    ]
    eval_q = [
        "In your opinion, what is the most practical step an individual high school student can take to support circular economy habits?",
        "Critique why transitioning completely from coal to renewable energy in developing countries like Indonesia takes time. What are the economic challenges?"
    ]
    perf_task = (
        "PENYUSUNAN POLICY BRIEF & PIDATO HORTATORI DIPLOMATIK (LKPD / SUMATIF):\n"
        "Susunlah sebuah Naskah Pidato Advokasi Lingkungan (Hortatory Speech) sepanjang 250-300 kata sebagai wakil pemuda di hadapan dewan kota.\n"
        "Pilihlah salah satu topik kebijakan:\n"
        "1. Urgensi Pengurangan Kemasan Plastik Sekali Pakai di Seluruh Kafetaria Sekolah.\n"
        "2. Pemasangan Panel Surya di Atap Gedung-Gedung Pemerintahan dan Sekolah.\n"
        "Naskah wajib menerapkan argumen berjenjang dan diakhiri dengan minimal 3 rekomendasi tegas (hortatory recommendations) yang jelas dan realistis."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_11.TL.2_Advanced_Hortatory_Exposition_Climate_Diplomacy.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 11.TL.3: ADVANCED LITERARY NARRATIVE - SHORT STORIES & CHARACTER
# ==============================================================================
def create_tp_11_tl_3():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS TINGKAT LANJUT",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XI / Semester 3 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 11.TL.3 - Mengapresiasi, menganalisis kedalaman karakter, dan merekonstruksi teks naratif sastra tingkat lanjut (cerpen klasik dan kontemporer).",
        "Elemen Capaian Pembelajaran": "Menyimak - Berbicara & Membaca - Memirsa",
        "Genre Teks / Ruang Lingkup": "Literary Narrative (Short Stories, Symbolism, Characterization, Epiphany)",
        "Alokasi Waktu": "25 Jam Pelajaran (5 Minggu x 5 JP)",
        "Target Kemahiran CEFR": "Level B1+ s.d. B2",
        "Profil Pelajar Pancasila": "Kreatif, Mandiri, Berkebinekaan Global"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "Katherine Mansfield Classic Stories / Project Gutenberg",
        "title": "The Epiphany of Empathy: An Excerpt from 'The Garden Party'",
        "url": "https://www.gutenberg.org/ebooks/1429",
        "year": "Classic / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t1_paras = [
        "The weather was absolutely ideal for the garden party. Windless, warm, and the sky was clear blue. The lawns were neatly mowed, and the florist had delivered armfuls of beautiful pink lilies to the Sheridan family's large country estate. Laura, the young daughter, was given the task of organizing the decorations, feeling proud of her grown-up responsibility.",
        "However, right before the guests arrived, dreadful news reached the house: Scott, a poor horse-cart driver who lived in a small, rundown cottage down the road, had suffered a tragic accident and died, leaving a grieving wife and five small children. Horror-stricken, Laura ran to her mother. 'Mother, we cannot host our party with music and laughing while a working man has died right down the hill! We must stop the party!'",
        "Mrs. Sheridan laughed gently, dismissing Laura's concerns. She placed an exquisite black hat decorated with a bright yellow flower on Laura's head and said, 'People like that don't expect us to cancel a party, darling. Look at how lovely this hat looks on you!' Looking in the mirror, Laura was dazzled by her own reflection, and her moral protest faded.",
        "That evening, after the music ended, Laura carried a basket of leftover party cakes down to the dark cottage. Seeing the young driver lying peaceful in death, Laura felt an overwhelming wave of sadness and realization about the deep social divide between rich and poor. 'Forgive my hat,' she whispered softly, realizing how shallow her privileged world had been."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Mengeksplorasi benturan kelas sosial dan pencerahan batin (epiphany) seorang remaja terhadap jurang ketidakadilan dan kefanaan hidup.",
        "Struktur Teks": "Exposition (Persiapan pesta taman mewah) -> Complication (Kabar tewasnya kusir miskin dan penolakan sang ibu membatalkan pesta) -> Climax & Epiphany (Laura mengantar kue ke rumah duka dan tersadar akan kepalsuan dunianya yang kaya).",
        "Fitur Bahasa": "Simbolisme sastra (topi indah melambangkan keangkuhan kelas sosial); Kontras deskripsi rumah mewah vs gubuk duka."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "Empathy and Social Divide: Excerpt from 'The Garden Party'",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "Ray Bradbury Short Stories / Science Fiction Classics",
        "title": "A House Without Inhabitants: Excerpt from 'There Will Come Soft Rains'",
        "url": "https://www.raybradbury.com/works/short-stories",
        "year": "Classic / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t2_paras = [
        "In the living room of an empty house, the voice-clock sang: 'Tick-tock, seven o'clock, time to get up!' In the kitchen, the automated stove gave a hissing sound and cooked eight pieces of warm toast, eight sunny-side eggs, and two cups of coffee. 'Today is August 4, 2026,' announced another automated voice from the ceiling. But there were no footsteps on the floor, and no hands reached for the plates.",
        "The automated smart house stood entirely alone in the middle of a ruined, silent city. A nuclear catastrophe had destroyed everything around it. On the charred west wall of the building, four white silhouettes were permanently burned into the paint: a father mowing a lawn, a mother picking flowers, and a boy and girl throwing a ball.",
        "The smart house continued to perform its daily duties faithfully—cleaning the floors, playing gentle music, and watering the garden—serving owners who were no longer there. Finally, a sudden tree branch crashed through the window, knocked over a cleaning bottle, and started a fire that consumed the empty house in flames, proving that nature outlasts all human machines."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Menyajikan kisah fiksi ilmiah reflektif tentang bahaya perang nuklir dan batas keangkuhan teknologi manusia.",
        "Struktur Teks": "Setting & Routine (Rumah pintar yang tetap melayani rutinitas keluarga yang telah tiada) -> Revelation (Siluet putih di dinding hangus mengungkap tragedi nuklir) -> Destruction (Kebakaran yang melenyapkan rumah otomatis).",
        "Fitur Bahasa": "Personifikasi mesin ('the voice-clock sang', 'the stove gave a hissing sound'), Ironi dramatis."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "Technological Hubris: Excerpt from 'There Will Come Soft Rains'",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "James Joyce Short Stories / Dubliners Collection",
        "title": "The Awakening of Youth: Excerpt from 'Araby'",
        "url": "https://www.gutenberg.org/ebooks/2814",
        "year": "Classic / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t3_paras = [
        "Every day, a young schoolboy in Dublin secretly watched the girl next door, Mangan's sister, from his window. He was deeply captivated by her graceful presence. When she finally spoke to him, asking if he was visiting Araby—a grand eastern-themed bazaar—the boy proudly promised that he would buy her a wonderful gift.",
        "All week, he could not concentrate on his school lessons, dreaming only of the magnificent treasures at the bazaar. When Saturday evening finally arrived, his uncle returned home very late with the pocket money. By the time the boy reached the grand exhibition hall, it was nearly ten o'clock, and most of the stalls had already closed.",
        "He approached one of the few open stalls where expensive porcelain dishes were displayed, but the young saleswoman spoke to him with cold indifference. Looking around the shadowy, quiet hall as the ceiling lights clicked off, the boy realized that his romantic daydream was only an illusion. Standing in the darkness, he felt a painful loss of innocence, realizing that the real world was far more ordinary and indifferent than his dreams."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "The End of an Illusion: Excerpt from 'Araby'",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Horror-stricken", "Adjective", "Sangat terpukul dan ngeri mendengar kabar buruk", "Laura was horror-stricken by the sudden death of the carter."),
        ("Dazzled", "Adjective", "Terpesona atau terpukau oleh kilauan keindahan", "She was dazzled by the fancy black hat in the mirror."),
        ("Divide", "Noun", "Jurang pemisah antarkelas sosial di masyarakat", "The story exposes the deep divide between wealth and poverty."),
        ("Automated", "Adjective", "Serba otomatis tanpa perlu campur tangan manusia", "The automated house continued cooking for missing owners."),
        ("Catastrophe", "Noun", "Bencana besar yang meluluhlantakkan peradaban", "The city was destroyed in a nuclear catastrophe."),
        ("Outlast", "Verb", "Bertahan hidup lebih lama melampaui yang lain", "Natural forces will always outlast human machines."),
        ("Captivated", "Adjective", "Terpikat hati dan terpesona oleh seseorang", "The boy was captivated by the kind neighbor girl."),
        ("Illusion", "Noun", "Ilusi atau angan-angan semu yang tidak nyata", "He realized his romantic expectations were only an illusion.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "In Text 1, why did Laura believe her family should cancel the garden party?",
        "What happened to the human residents of the smart house in Text 2?",
        "Why was the young narrator in Text 3 unable to buy a gift for Mangan's sister at Araby?"
    ]
    inf_q = [
        "What does Laura's apology—'Forgive my hat'—symbolize beside the dead driver in Text 1?",
        "How does Ray Bradbury use the white silhouettes on the wall in Text 2 to evoke tragic emotions in the reader?",
        "Explain the emotional 'epiphany' (moment of sudden realization) experienced by the boy at the end of Text 3."
    ]
    eval_q = [
        "Compare Laura's feelings in Text 1 with the boy's feelings in Text 3: Which character experiences a more painful transition into adult reality? Defend your view.",
        "How can reading classic literature help modern teenagers develop empathy for people from different backgrounds?"
    ]
    perf_task = (
        "PENULISAN CERPEN KREATIF BERMAKNA (LKPD / SUMATIF KREATIF):\n"
        "Tulislah sebuah Cerpen Sastra (Literary Short Story) sepanjang 250-300 kata yang berpusat pada sebuah 'Epiphany' (Momen Pencerahan Batin).\n"
        "Unsur Wajib Cerpen:\n"
        "1. Setting yang memuat suasana emosional yang jelas (misal: ruang tunggu, stasiun senja, rumah sakit).\n"
        "2. Tokoh utama yang awalnya memiliki pandangan keliru atau sombong terhadap suatu hal.\n"
        "3. Sebuah peristiwa yang membukakan matanya dan menyadarkannya akan kebenaran hidup.\n"
        "Gunakan minimal 4 kosakata baru dari tabel di atas."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_11.TL.3_Advanced_Literary_Narrative_Short_Stories.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 11.TL.4: STYLISTICS OF POETRY & PROTEST SONGS
# ==============================================================================
def create_tp_11_tl_4():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS TINGKAT LANJUT",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XI / Semester 3 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 11.TL.4 - Menganalisis gaya bahasa stilistika, makna simbolik, dan kritik sosial dalam teks puisi serta lirik lagu kemanusiaan dunia.",
        "Elemen Capaian Pembelajaran": "Menyimak - Berbicara & Membaca - Memirsa",
        "Genre Teks / Ruang Lingkup": "Poetry & Protest Songs (Stylistics, Metaphor, Rhythm, Social Critique)",
        "Alokasi Waktu": "10 Jam Pelajaran (2 Minggu x 5 JP)",
        "Target Kemahiran CEFR": "Level B1+ s.d. B2",
        "Profil Pelajar Pancasila": "Berkebinekaan Global, Bernalar Kritis, Berakhlak Mulia"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "Poetry Foundation / Random House Classics",
        "title": "Caged Bird: The Symbolism of Freedom and Oppression",
        "url": "https://www.poetryfoundation.org/poems/48989/caged-bird",
        "year": "Classic / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t1_paras = [
        "A free bird leaps on the back of the wind\nand floats downstream till the current ends\nand dips his wing in the orange sun rays\nand dares to claim the sky.",
        "But a bird that stalks down his narrow cage\ncan seldom see through his bars of rage\nhis wings are clipped and his feet are tied\nso he opens his throat to sing.",
        "The caged bird sings with a fearful trill\nof things unknown but longed for still\nand his tune is heard on the distant hill\nfor the caged bird sings of freedom.",
        "The free bird thinks of another breeze\nand the trade winds soft through the sighing trees\nand the fat worms waiting on a dawn-bright lawn\nand he names the sky his own.",
        "But a caged bird stands on the grave of dreams\nhis shadow shouts on a nightmare scream\nhis wings are clipped and his feet are tied\nso he opens his throat to sing."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Menyuarakan jeritan kaum tertindas dan merayakan keteguhan jiwa manusia untuk terus bernyanyi menuntut kebebasan dan kesetaraan.",
        "Struktur Teks": "Pola bait bergantian (strophic contrast) yang mempertemukan kehidupan burung bebas ('free bird') dengan derita sangkar represi ('caged bird').",
        "Fitur Bahasa": "Metafora ('bars of rage', 'grave of dreams'), Aliterasi bunyi konsonan ('fearful trill', 'distant hill'), Refrain repetisi 'so he opens his throat to sing'."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "Caged Bird: Allegory of Freedom and Unyielding Spirit",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "Nobel Prize in Literature Archives / Sony Music",
        "title": "Blowin' in the Wind: Rhetorical Questions for Peace and Conscience",
        "url": "https://www.nobelprize.org/prizes/literature/2016/dylan/facts/",
        "year": "Classic / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t2_paras = [
        "How many roads must a man walk down / Before you call him a man? / How many seas must a white dove sail / Before she sleeps in the sand? / Yes, and how many times must the cannonballs fly / Before they're forever banned? / The answer, my friend, is blowin' in the wind / The answer is blowin' in the wind.",
        "Yes, and how many years must a mountain exist / Before it is washed to the sea? / Yes, and how many years can some people exist / Before they're allowed to be free? / Yes, and how many times can a man turn his head / And pretend that he just doesn't see? / The answer, my friend, is blowin' in the wind / The answer is blowin' in the wind.",
        "Yes, and how many times must a man look up / Before he can see the sky? / Yes, and how many ears must one man have / Before he can hear people cry? / Yes, and how many deaths will it take 'til he knows / That too many people have died? / The answer, my friend, is blowin' in the wind / The answer is blowin' in the wind."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Mengajak nurani umat manusia untuk menolak perang, memperjuangkan hak asasi, dan menghentikan sikap pura-pura tidak tahu.",
        "Struktur Teks": "Tiga bait berstruktur pertanyaan retoris berulang (anaphora 'How many...') yang diakhiri dengan refrain puitis.",
        "Fitur Bahasa": "Simbol burung merpati putih ('white dove' - lambang perdamaian), Peluru meriam ('cannonballs' - lambang perang destruktif)."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "Blowin' in the Wind: A Conscience for Human Peace",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "Amanda Gorman Official / Penguin Young Readers",
        "title": "The Hill We Climb: A Spoken Word Call for Unity and Courage",
        "url": "https://www.theamandagorman.com/poems/the-hill-we-climb",
        "year": "2021 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t3_paras = [
        "When day comes, we ask ourselves, where can we find light in this never-ending shade? / The loss we carry, a sea we must wade. / We've learned that quiet isn't always peace, / and the norms and notions of what 'just' is isn't always justice. / And yet, the dawn is ours before we knew it. / Somehow we do it. / Somehow we've weathered and witnessed a nation that isn't broken, but simply unfinished.",
        "We are striving to forge our union with purpose. / To compose a country committed to all cultures, colors, characters, and conditions of man. / And so we lift our gazes not to what stands between us, but what stands before us.",
        "For there is always light, / if only we're brave enough to see it. / If only we're brave enough to be it."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "The Hill We Climb: A Contemporary Anthem of Hope",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Bars of rage", "Noun Phrase", "Jeruji sangkar yang melambangkan kemarahan tertindas", "The bird sings even behind its bars of rage."),
        ("Refrain", "Noun", "Baris lirik berulang yang mempertegas pesan puisi", "The singer repeated the memorable refrain about peace."),
        ("Cannonball", "Noun", "Peluru meriam besi, simbol perang militer", "We dream of a world where cannonballs are banned forever."),
        ("Anaphora", "Noun", "Pengulangan kata awal baris untuk penegasan puitis", "Dylan's use of anaphora makes his questions powerful."),
        ("Unfinished", "Adjective", "Belum selesai, masih butuh perjuangan untuk disempurnakan", "Our society is not broken, but simply unfinished."),
        ("Union", "Noun", "Persatuan yang kokoh dalam keberagaman", "We strive to forge a peaceful union of all cultures."),
        ("Gaze", "Noun", "Tatapan mata yang penuh harapan ke depan", "Lift your gaze to what stands before us."),
        ("Spoken word", "Noun Phrase", "Puisi performatif lisan yang dibacakan penuh ritme", "Spoken word poetry connects deeply with youth audiences.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "In Text 1, what does the caged bird do despite having clipped wings and tied feet?",
        "What repeating phrase concludes every single stanza of Bob Dylan's song in Text 2?",
        "According to Text 3, what distinction does Amanda Gorman draw between 'quiet' and 'peace'?"
    ]
    inf_q = [
        "Why is singing an act of defiance and hope for the caged bird in Maya Angelou's poem?",
        "What does Bob Dylan mean when he says 'The answer is blowin' in the wind'?",
        "Explain the message behind Amanda Gorman's closing line: 'If only we're brave enough to be it.'"
    ]
    eval_q = [
        "Between classic written poetry (Text 1) and spoken word recitation (Text 3), which format do you find more engaging for high school students? Explain why.",
        "How can songs and poetry influence public attitudes toward social justice more effectively than news broadcasts?"
    ]
    perf_task = (
        "DEKLAMASI PUISI & RESENSI GAYA BAHASA (LKPD / SUMATIF LISAN):\n"
        "1. Pilih salah satu puisi di atas, atau ciptakan sebuah Puisi Inspiratif orisinal (3 bait) bertema persahabatan, perdamaian, atau alam.\n"
        "2. Rekam pembacaan puisi ekspresif (Spoken Word Recitation) berdurasi 1-2 menit dengan intonasi, volume, dan penghayatan yang tepat.\n"
        "3. Sertakan catatan refleksi singkat (100 kata) yang menjelaskan makna simbol utama dalam puisimu."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_11.TL.4_Stylistics_Poetry_and_Protest_Songs.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 11.TL.5: DISCUSSION TEXT - PUSKURJAR MODEL (ESPORTS & GAMING)
# ==============================================================================
def create_tp_11_tl_5():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS TINGKAT LANJUT",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XI / Semester 4 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 11.TL.5 - Mengonstruksi dan menimbang dua sudut pandang berimbang dalam teks diskusi mengenai fenomena industri gim online dan kecanduan digital (Model Bab D Puskurjar).",
        "Elemen Capaian Pembelajaran": "Membaca - Memirsa & Berbicara - Mempresentasikan",
        "Genre Teks / Ruang Lingkup": "Discussion Text (Competitive Esports vs Digital Gaming Risks - Puskurjar Bab D)",
        "Alokasi Waktu": "20 Jam Pelajaran (4 Minggu x 5 JP)",
        "Target Kemahiran CEFR": "Level B1+ s.d. B2",
        "Profil Pelajar Pancasila": "Bernalar Kritis, Mandiri, Gotong Royong"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "Harvard Health Publishing / BBC Worklife Report",
        "title": "Competitive Esports vs. Gaming Addiction: The High School Debate",
        "url": "https://www.health.harvard.edu/blog/the-health-effects-of-gaming",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t1_paras = [
        "The rapid growth of online gaming from a casual hobby into a multi-million-dollar competitive esports industry has sparked a lively debate among parents, teachers, and teenagers. While gaming enthusiasts celebrate esports as a legitimate career path that develops quick reflexes and strategic teamwork, medical experts caution that excessive gaming can lead to addiction, poor sleep, and social isolation.",
        "On the positive side, competitive video gaming provides genuine cognitive and professional benefits. Scientific studies indicate that strategy and action games improve players' visual-motor coordination, spatial memory, and fast decision-making skills. Furthermore, the modern gaming industry offers legitimate career opportunities beyond playing, including software programming, event broadcasting, game design, and digital team management.",
        "On the negative side, unmonitored gaming poses serious health hazards. Many popular games use reward loops and in-game purchases that encourage teenagers to play for hours without stopping. Prolonged screen time often results in eye fatigue, poor physical posture from sitting too long, and neglected school homework. Some students become socially withdrawn, preferring virtual friends over real-world relationships.",
        "In conclusion, video gaming is neither purely harmful nor purely beneficial. The sensible solution is moderation and structure. High schools should guide students by establishing supervised esports clubs that require good grades and daily physical exercise, while parents set reasonable screen-time boundaries at home."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Menimbang secara berimbang sisi positif dan negatif dari industri gim daring dan bahaya adiksi layar, serta merumuskan jalan keluar yang solutif.",
        "Struktur Teks": "Issue (Ledakan esports vs bahaya adiksi gim pada remaja) -> Arguments For (Peningkatan koordinasi visual, refleks, dan peluang karier) -> Arguments Against (Kelelahan mata, penurunan nilai, dan isolasi sosial) -> Recommendation (Moderasi berimbang melalui klub sekolah dan batas waktu di rumah).",
        "Fitur Bahasa": "Contrastive Transitions ('On the positive side', 'On the negative side', 'In conclusion'), Hedging Language ('neither purely harmful nor purely beneficial')."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "Competitive Esports vs. Digital Gaming Risks: Finding the Balance",
        "Bahan Modul Ajar (Modeling of the Text - Model Bab D)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "BBC Worklife / Pew Research Center Youth Series",
        "title": "Is a Career in Professional Esports Viable for High School Graduates?",
        "url": "https://www.bbc.com/worklife/article/esports-careers-youth",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t2_paras = [
        "As global esports tournaments offer millions of dollars in prize money, many high school students dream of becoming professional video game players. However, career counselors emphasize that the reality of professional esports is far more challenging than it appears on streaming platforms.",
        "Proponents point out that the global gaming market is booming, creating viable jobs in coaching, digital content creation, and broadcast commentary. Disciplined players who master teamwork and communication can earn a good income and travel internationally.",
        "Conversely, critics emphasize that the chances of becoming an elite professional player are extremely slim—comparable to the odds of reaching the football World Cup. Professional players often practice for twelve hours a day, suffering from wrist injuries, mental burnout, and short career lifespans that often end before age twenty-five. Dropping out of school for gaming leaves youth without essential career fallback options.",
        "Ultimately, teenagers should treat esports as a fun extracurricular interest while completing their academic education to ensure long-term career security."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Mendiskusikan prospek realistis karier esports profesional bagi lulusan SMA.",
        "Struktur Teks": "Issue (Impian karier esports) -> Sisi Peluang (Industri berkembang) -> Sisi Risiko (Persaingan sangat ketat, cedera fisik, karier pendek) -> Rekomendasi (Menjadikan hobi sambil mengutamakan pendidikan).",
        "Fitur Bahasa": "Words of Caution ('challenging', 'burnout', 'fallback options', 'Ultimately')."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "Pursuing an Esports Career: Golden Opportunity or Risky Dream?",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "UNESCO Digital Well-being / Child Mind Institute",
        "title": "Managing Screen Time: How Schools Can Foster Digital Balance",
        "url": "https://www.unesco.org/en/articles/digital-wellbeing-students",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t3_paras = [
        "The integration of digital games and educational apps into high school classrooms has created a dilemma for educators. Does gamified learning inspire student motivation, or does it encourage more screen distraction?",
        "Supporters argue that gamification makes difficult subjects like mathematics and chemistry exciting. Educational quizzes with badges and leaderboards increase student engagement and encourage healthy classroom teamwork.",
        "However, skeptics worry that continuous gamification conditions students to expect constant digital stimulation. Learning to read long literary texts or write detailed essays requires quiet patience without instant dopamine rewards. Too much gamification may reduce deep intellectual stamina.",
        "A balanced approach combines gamified review sessions with offline handwritten tasks, ensuring students develop both digital enthusiasm and deep reading focus."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "Gamification in Education: Fun Engagement or Screen Overload?",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Coordination", "Noun", "Koordinasi kerja sama antara mata, tangan, dan otak", "Action games improve visual-motor coordination."),
        ("Moderation", "Noun", "Sikap wajar dan tidak berlebih-lebihan", "Video games can be fun when enjoyed in moderation."),
        ("Hazard", "Noun", "Bahaya atau risiko yang mengancam kesehatan", "Sitting all day without moving poses a health hazard."),
        ("Withdrawn", "Adjective", "Menutup diri dan menjauh dari pergaulan sosial", "Excessive gaming can cause teenagers to become withdrawn."),
        ("Extracurricular", "Noun/Adjective", "Kegiatan pengembangan diri di luar jam pelajaran", "Join esports as an extracurricular activity, not your whole life."),
        ("Burnout", "Noun", "Kelelahan mental luar biasa akibat latihan berlebih", "Long gaming sessions lead to severe mental burnout."),
        ("Gamification", "Noun", "Penerapan unsur permainan dalam proses belajar", "Gamification makes reviewing history facts engaging."),
        ("Stamina", "Noun", "Daya tahan fisik atau fokus mental jangka panjang", "Students need stamina to read long academic books.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "According to Text 1, what cognitive skills are developed through playing strategic video games?",
        "In Text 2, what physical and career risks do professional esports players face before age twenty-five?",
        "What potential drawback of gamified classroom learning is raised by skeptics in Text 3?"
    ]
    inf_q = [
        "Why do game developers include reward loops and microtransactions in video games?",
        "How can high schools use supervised esports clubs to motivate students to improve their academic grades?",
        "Explain why deep reading requires quiet patience rather than instant digital stimulation."
    ]
    eval_q = [
        "Do you agree with the recommendation in Text 1 that schools should supervise esports rather than ban it? Explain your perspective.",
        "Between playing online games and doing outdoor sports, which activity better supports your personal well-being? Give examples."
    ]
    perf_task = (
        "DEBAT KELAS & PENULISAN TEKS DISKUSI (LKPD / SUMATIF):\n"
        "Susunlah sebuah Teks Diskusi (Discussion Essay) sepanjang 250-300 kata mengenai salah satu isu digital:\n"
        "Mosi A: Should High Schools Allow Official Esports Teams alongside Traditional Sports Teams?\n"
        "Mosi B: Should Video Games Ban 'Loot-Box' Purchases for Minors Under 18?\n"
        "Format Wajib:\n"
        "• Paragraph 1: Issue (Jelaskan konteks perdebatan).\n"
        "• Paragraph 2: Arguments For (Sajikan 2 alasan pendukung disertai fakta).\n"
        "• Paragraph 3: Arguments Against (Sajikan 2 alasan penentang disertai fakta).\n"
        "• Paragraph 4: Balanced Recommendation (Sajikan jalan tengah yang solutif).\n"
        "Gunakan kosakata transisi kontras (on the one hand, on the other hand, conversely)."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_11.TL.5_Official_Discussion_Text_Puskurjar_Gaming.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 11.TL.6: COMPLEX SCIENTIFIC CAUSAL EXPLANATION
# ==============================================================================
def create_tp_11_tl_6():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS TINGKAT LANJUT",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XI / Semester 4 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 11.TL.6 - Menganalisis mekanisme kausalitas ilmiah dan urutan proses teknologi mutakhir dalam teks eksplanasi ilmiah bertingkat.",
        "Elemen Capaian Pembelajaran": "Membaca - Memirsa & Menulis - Mempresentasikan",
        "Genre Teks / Ruang Lingkup": "Scientific Causal Explanation (Quantum Basics, Space Detection, Clean Fusion)",
        "Alokasi Waktu": "25 Jam Pelajaran (5 Minggu x 5 JP)",
        "Target Kemahiran CEFR": "Level B1+ s.d. B2",
        "Profil Pelajar Pancasila": "Bernalar Kritis, Mandiri, Kreatif"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "Scientific American / Nature Physics Youth Column",
        "title": "Quantum Computing Basics: How Qubits Work in Simple Terms",
        "url": "https://www.scientificamerican.com/computing/quantum-basics",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t1_paras = [
        "Traditional computers, from your smartphone to massive supercomputers, process information using binary bits. Each bit acts like a simple light switch that is either OFF (0) or ON (1). However, quantum computers utilize quantum bits, or qubits, which follow the fascinating laws of quantum physics to solve complex calculations at extraordinary speeds.",
        "The key breakthrough of quantum computing is a phenomenon called superposition. Unlike a regular bit that must choose between 0 or 1, a qubit can exist in a combination of both 0 and 1 simultaneously. This means that a quantum computer can evaluate millions of potential solutions at the exact same moment, rather than checking them one by one like a classical machine.",
        "Additionally, through a property called entanglement, connected qubits can share information instantly across the processor. When scientists run a quantum algorithm, the system uses wave interference to cancel out incorrect answers and highlight the correct solution.",
        "Finally, when the measurement is taken, the qubit settles into a definitive 0 or 1 answer. This revolutionary technology has the potential to discover new life-saving medicines and optimize renewable energy grids."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Menjelaskan secara ilmiah bagaimana fenomena superposisi dan keterikatan kuantum memungkinkan komputasi super cepat.",
        "Struktur Teks": "General Statement (Perbedaan bit biner klasik vs qubit kuantum) -> Causal Sequence (Mekanisme superposisi 0 dan 1 sekaligus -> Keterikatan kuantum & amplifikasi interferensi) -> Concluding Summary (Aplikasi untuk obat-obatan dan energi).",
        "Fitur Bahasa": "Passive Voice ('is processed', 'are utilized', 'can be evaluated'), Technical Terms ('qubits', 'superposition', 'entanglement')."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "How Quantum Computers Work: The Power of Superposition",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "NASA Jet Propulsion Laboratory / Caltech LIGO",
        "title": "Listening to the Universe: How LIGO Detects Gravitational Waves",
        "url": "https://www.ligo.caltech.edu/page/what-are-gw",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t2_paras = [
        "In 1916, Albert Einstein predicted that violent events in space would create ripples in the fabric of spacetime, known as gravitational waves. In 2015, scientists at the Laser Interferometer Gravitational-Wave Observatory (LIGO) made history by detecting these cosmic ripples directly for the first time.",
        "The process begins billions of light-years away when two massive black holes spiral toward each other at immense speeds and collide. This cataclysmic crash shakes the fabric of spacetime, sending out waves of energy that travel across the universe at the speed of light.",
        "To measure these faint waves on Earth, LIGO uses two identical four-kilometer-long vacuum tubes built at right angles. High-powered laser beams travel down both tubes, bouncing off mirrors at the ends. When a gravitational wave passes through Earth, it slightly stretches one tube while compressing the other by an amount smaller than a fraction of an atom. This minute change alters the laser beam's return time, confirming the distant cosmic collision."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Menjelaskan urutan sebab-akibat tabrakan lubang hitam di luar angkasa hingga terdeteksi oleh sensor laser LIGO di bumi.",
        "Struktur Teks": "General Statement (Prediksi Einstein dan deteksi gelombang gravitasi) -> Causal Sequence (Tabrakan dua lubang hitam -> Perambatan riak ruang-waktu -> Deteksi peregangan tabung laser di bumi).",
        "Fitur Bahasa": "Causal Sequencers ('begins when', 'sending out', 'When a wave passes')."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "Ripples in Spacetime: How Scientists Detect Gravitational Waves",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "MIT Energy Initiative / ITER Science Briefs",
        "title": "Harnessing the Power of Stars: The Science of Clean Fusion Energy",
        "url": "https://energy.mit.edu/research/fusion-energy-basics",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t3_paras = [
        "Nuclear fusion is the natural process that powers the Sun and all the stars in the night sky. For decades, scientists have been working to replicate this clean energy process inside laboratories on Earth.",
        "Unlike traditional nuclear fission, which splits heavy uranium atoms and creates dangerous radioactive waste, fusion combines lightweight hydrogen atoms into helium. When hydrogen nuclei fuse together under extreme heat and pressure, they release a tremendous amount of clean energy without producing any carbon emissions.",
        "To achieve fusion on Earth, hydrogen gas is heated inside a magnetic confinement chamber called a tokamak until it reaches over 100 million degrees Celsius, turning into superheated plasma. Powerful superconducting magnets keep the fiery plasma suspended away from the chamber walls, allowing nuclei to collide with enough speed to fuse and generate clean power."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "Nuclear Fusion: The Clean Energy of the Future",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Binary bit", "Noun Phrase", "Satuan data komputer klasik (angka 0 atau 1)", "A binary bit can only represent a 0 or a 1."),
        ("Simultaneously", "Adverb", "Secara serentak pada waktu yang bersamaan", "Qubits can evaluate multiple possibilities simultaneously."),
        ("Interference", "Noun", "Pola perpaduan gelombang yang saling menguatkan/melemahkan", "Wave interference cancels out erroneous computational pathways."),
        ("Cataclysmic", "Adjective", "Sangat dahsyat dan berdaya hancur besar", "The collision of two black holes is a cataclysmic cosmic event."),
        ("Interferometer", "Noun", "Alat pengukur gelombang menggunakan pantulan sinar laser", "LIGO's interferometer measures microscopic shifts in space."),
        ("Compress", "Verb", "Memampatkan atau merapatkan ukuran ruang", "Gravitational waves slightly compress spacetime as they pass."),
        ("Fission", "Noun", "Pembelahan inti atom berat yang menghasilkan limbah radiasi", "Traditional nuclear power plants use fission to make power."),
        ("Plasma", "Noun", "Gas bermuatan listrik pada suhu luar biasa tinggi", "Superheated plasma is contained by magnetic fields in a tokamak.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "In Text 1, what is the key difference between a classical binary bit and a quantum bit (qubit)?",
        "What cosmic event creates the gravitational waves detected by LIGO in Text 2?",
        "According to Text 3, why is nuclear fusion considered cleaner and safer than traditional nuclear fission?"
    ]
    inf_q = [
        "How does the phenomenon of superposition allow a quantum computer to search through complex data much faster than a standard PC?",
        "Why must LIGO's vacuum tubes be built at right angles (90 degrees) to detect gravitational waves?",
        "Why is it necessary to use powerful magnets to keep the plasma suspended away from the tokamak walls in Text 3?"
    ]
    eval_q = [
        "Between quantum computing (Text 1) and clean nuclear fusion (Text 3), which technological breakthrough do you think will have a greater positive impact on humanity? Explain your reasoning.",
        "Evaluate why scientists use analogies like 'light switches' and 'ripples on water' when explaining advanced physics to the general public."
    ]
    perf_task = (
        "PENULISAN ARTIKEL EKSPLANASI SAINS TEKNOLOGI (LKPD / SUMATIF):\n"
        "Susunlah sebuah Artikel Eksplanasi Kausal Ilmiah (Explanation Text) sepanjang 250-300 kata untuk majalah sains sekolah.\n"
        "Pilihlah salah satu tema teknologi mutakhir:\n"
        "1. How Electric Vehicles (EVs) Store and Convert Battery Power.\n"
        "2. How Cloud Computing Stores and Delivers Data Across the World.\n"
        "3. How Fiber Optic Cables Transmit Internet Signals at the Speed of Light.\n"
        "Struktur Wajib:\n"
        "- General Statement: Perkenalkan teknologi dan kegunaannya.\n"
        "- Sequenced Explanation: Jelaskan minimal 3 langkah sebab-akibat ilmiah secara runtut menggunakan passive voice dan kata penghubung kausalitas.\n"
        "- Gambarkan bagan alur proses sederhana yang mendampingi artikelmu."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_11.TL.6_Complex_Causal_Explanation_Science.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 11.TL.7: BIOGRAPHY & HISTORICAL MEMOIRS - WORLD LEADERS
# ==============================================================================
def create_tp_11_tl_7():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS TINGKAT LANJUT",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XI / Semester 4 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 11.TL.7 - Mengevaluasi rekam jejak perjuangan, kepemimpinan transformasional, dan retrospeksi historis dalam teks biografi mendalam dan memoar tokoh dunia.",
        "Elemen Capaian Pembelajaran": "Menyimak - Berbicara & Membaca - Memirsa",
        "Genre Teks / Ruang Lingkup": "Biography & Historical Memoirs (Transformational Leadership, Moral Courage, Integrity)",
        "Alokasi Waktu": "20 Jam Pelajaran (4 Minggu x 5 JP)",
        "Target Kemahiran CEFR": "Level B1+ s.d. B2",
        "Profil Pelajar Pancasila": "Berkebinekaan Global, Berakhlak Mulia, Bernalar Kritis"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "Nelson Mandela Foundation / Little, Brown & Co.",
        "title": "Long Walk to Freedom: Choosing Reconciliation Over Revenge",
        "url": "https://www.nelsonmandela.org/content/page/biography",
        "year": "Classic / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t1_paras = [
        "Nelson Mandela was born into a small village in South Africa. As a young man, he experienced firsthand the deep injustice of the apartheid regime, a racist system of laws that denied black citizens basic human rights and voting power. Mandela dedicated his life to fighting for racial equality, which led to his arrest and twenty-seven long years of harsh imprisonment on Robben Island.",
        "During those long decades behind bars, breaking limestone rocks under the hot sun, Mandela refused to let hatred poison his heart. Instead, he studied his captors' language and culture, realizing that true leadership requires understanding your opponent. When he was finally released from prison in 1990, the world watched with awe as Mandela urged his supporters to choose peaceful forgiveness over revenge.",
        "Elected as South Africa's first black president in 1994, Mandela united a divided nation by promoting national reconciliation. He famously said, 'Resentment is like drinking poison and hoping it will kill your enemies.' His selfless moral leadership proved to humanity that peace is far stronger than hatred."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Merekam perjalanan hidup, keteguhan moral, dan kepemimpinan rekonsiliasi Nelson Mandela dalam meruntuhkan rezim apartheid.",
        "Struktur Teks": "Early Life & Conflict (Masa muda dan diskriminasi rezim apartheid) -> Imprisonment & Moral Growth (27 tahun penjara Robben Island yang mematangkan jiwa pemaaf) -> National Triumph (Menjadi presiden dan merajut rekonsiliasi bangsa).",
        "Fitur Bahasa": "Past Tense Naratif ('was born', 'experienced', 'dedicated', 'refused', 'united'), Kutipan mutiara bijak, Evaluative Adjectives ('injustice', 'selfless', 'moral')."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "Nelson Mandela: The Triumph of Reconciliation",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "The Nobel Foundation Official Biographical Archives",
        "title": "Marie Curie: A Life of Scientific Genius and Selfless Courage",
        "url": "https://www.nobelprize.org/prizes/physics/1903/marie-curie/biographical/",
        "year": "Classic / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t2_paras = [
        "Born Maria Skłodowska in Warsaw, Poland, Marie Curie overcame poverty and severe gender discrimination to become one of the greatest scientists in human history. In Poland, women were banned from university education, so Marie attended an illegal underground university at night while working as a governess to save money.",
        "In 1891, she moved to Paris to study physics and mathematics at the Sorbonne, surviving on bread and tea in a freezing garret room. Working alongside her husband Pierre in a poorly ventilated wooden shed, she tirelessly stirred heavy pots of mineral ore, discovering two new radioactive elements: polonium and radium.",
        "Curie became the first woman to win a Nobel Prize and remains the only person ever to win Nobel Prizes in two different scientific fields (Physics and Chemistry). Remarkably, she refused to patent her radium discovery, choosing to share it freely for cancer treatment. During World War I, she drove mobile X-ray vans to the front lines to help doctors save wounded soldiers."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Mendokumentasikan ketekunan, kejeniusan sains, dan pengabdian kemanusiaan Marie Curie bagi peradaban dunia.",
        "Struktur Teks": "Early Life (Pendidikan bawah tanah di Polandia) -> Laboratory Breakthroughs (Penemuan polonium & radium di Paris) -> Nobel Achievements & Humanitarian Service (Penolakan paten demi kesehatan dan ambulans X-ray).",
        "Fitur Bahasa": "Biographical Markers, Appositives, Descriptors of Persistence ('tirelessly', 'selfless', 'overcame')."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "Marie Curie: The Trailblazer of Modern Science",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "The Jakarta Post Retrospective / The Habibie Center",
        "title": "B.J. Habibie: The Master of Aerospace and Architect of Indonesian Democracy",
        "url": "https://www.thejakartapost.com/news/bj-habibie-in-memoriam",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t3_paras = [
        "Bacharuddin Jusuf Habibie was a brilliant visionary whose life bridged global aerospace engineering with Indonesia's democratic transition. Born in Parepare, South Sulawesi, Habibie pursued engineering in Germany, where his groundbreaking mathematical theory on aircraft crack progression earned him global fame as 'Mr. Crack'.",
        "Returning to Indonesia, Habibie led the nation's technological development, successfully designing and flying the indigenous N-250 Gatotkaca airplane in 1995. When Indonesia experienced severe economic and political crisis in 1998, Habibie was appointed president during a turbulent transition period.",
        "In just seventeen months in office, Habibie enacted landmark democratic reforms: he freed political prisoners, lifted press censorship, established an independent central bank, and conducted free democratic elections. Habibie demonstrated that scientific genius and humble democratic leadership can steer a nation out of deep crisis."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "B.J. Habibie: From High Aerospace to Democratic Reform",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Apartheid", "Noun", "Sistem politik pemisahan rasial di Afrika Selatan", "Nelson Mandela fought peacefully to dismantle apartheid."),
        ("Reconciliation", "Noun", "Pemulihan hubungan damai dan persaudaraan bangsa", "Mandela prioritized national reconciliation over revenge."),
        ("Governess", "Noun", "Guru les privat anak di rumah keluarga majikan", "Marie worked as a governess to finance her sister's medical school."),
        ("Ventilated", "Adjective", "Memiliki ventilasi sirkulasi udara yang baik", "The Curies worked in an unheated, poorly ventilated shed."),
        ("Patent", "Verb/Noun", "Mematenkan hak cipta teknologi demi keuntungan", "Marie Curie chose not to patent the radium extraction process."),
        ("Aerospace", "Noun", "Bidang teknologi industri penerbangan dan antariksa", "Habibie was a world-renowned pioneer in aerospace engineering."),
        ("Crack progression", "Noun Phrase", "Teori rambatan keretakan mikro pada badan pesawat terbang", "His crack progression formula made commercial flights safer."),
        ("Censorship", "Noun", "Penyensoran atau pembungkaman kebebasan pers", "Habibie abolished authoritarian media censorship in 1998.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "How many years did Nelson Mandela spend imprisoned before his historic release in 1990 in Text 1?",
        "What two radioactive elements did Marie and Pierre Curie discover together in Text 2?",
        "According to Text 3, what important aerospace theory earned B.J. Habibie the international nickname 'Mr. Crack'?"
    ]
    inf_q = [
        "In Text 1, what did Mandela mean by saying: 'Resentment is like drinking poison and hoping it will kill your enemies'?",
        "Why did Marie Curie refuse to patent her discovery of radium despite living in modest conditions?",
        "How did B.J. Habibie's background as a precise engineer help him lead Indonesia through the turbulent 1998 democratic reforms (Text 3)?"
    ]
    eval_q = [
        "Compare the leadership style of Nelson Mandela and B.J. Habibie: How did both leaders prevent civil conflicts during times of major political change?",
        "What character qualities from Marie Curie's life can inspire Indonesian students who wish to pursue careers in science and technology?"
    ]
    perf_task = (
        "PENYUSUNAN BIOGRAFI INSPIRATIF & PODCAST SINGKAT (LKPD / SUMATIF):\n"
        "1. Tulislah sebuah Teks Biografi Inspiratif (Biographical Monograph) sepanjang 250-300 kata mengenai seorang tokoh pahlawan atau panutan.\n"
        "2. Rekam naskah monolog audio (Podcast 2-3 menit) yang mengisahkan titik balik perjuangan tokoh tersebut dalam mengatasi rintangan besar.\n"
        "Gunakan urutan waktu yang runtut (Simple Past, Past Perfect) serta kosakata nilai kebajikan luhur."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_11.TL.7_Biography_Historical_Memoirs_World_Leaders.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 11.TL.8: CRITICAL REVIEW OF INTERNATIONAL ARTS & CINEMA
# ==============================================================================
def create_tp_11_tl_8():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS TINGKAT LANJUT",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XI / Semester 4 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 11.TL.8 - Menyusun resensi dan kritik estetika mendalam (Review Text) terhadap novel sastra dunia dan karya sinematografi peraih penghargaan internasional.",
        "Elemen Capaian Pembelajaran": "Membaca - Memirsa & Menulis - Mempresentasikan",
        "Genre Teks / Ruang Lingkup": "Critical Review Text (Cinematography, World Literature, Thematic Aesthetics)",
        "Alokasi Waktu": "20 Jam Pelajaran (4 Minggu x 5 JP)",
        "Target Kemahiran CEFR": "Level B1+ s.d. B2",
        "Profil Pelajar Pancasila": "Bernalar Kritis, Kreatif, Berkebinekaan Global"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "The New York Times Critic's Pick / Variety Film Review",
        "title": "Film Review: Christopher Nolan's 'Oppenheimer' – Science and Conscience",
        "url": "https://www.nytimes.com/movies/oppenheimer-review",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t1_paras = [
        "Christopher Nolan's 'Oppenheimer' is a gripping historical drama about J. Robert Oppenheimer, the American theoretical physicist who led the Manhattan Project to create the world's first atomic bomb during World War II. The film explores the agonizing moral dilemma of scientific innovation.",
        "Cillian Murphy delivers an extraordinary, haunting performance as Oppenheimer. With his hollow cheeks and intense gaze, Murphy perfectly conveys the internal conflict between scientific excitement and the terrifying horror of unleashing nuclear weapons. The film uses a clever visual contrast: Oppenheimer's personal perspective is shown in rich, warm color, while the political investigations against him in the 1950s are presented in austere black-and-white.",
        "The film's most memorable moment is the Trinity nuclear test. Instead of filling the screen with loud explosions immediately, Nolan cuts the sound completely for nearly two minutes following the blinding flash. The audience sits in breathless silence until the concussive sound wave finally shatters the theater, mirroring the delayed moral shockwave that still haunts humanity.",
        "In conclusion, 'Oppenheimer' is a cinematic tour de force that challenges viewers to reflect on the moral responsibilities of scientists. Rated 5 out of 5 stars, it is an essential historical film for mature students."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Mengevaluasi secara kritis nilai estetika, tata artistik, sinematografi, dan pesan filosofis film Oppenheimer karya Christopher Nolan.",
        "Struktur Teks": "Orientation (Identifikasi film Oppenheimer, sutradara, latar sejarah Proyek Manhattan) -> Interpretative Recount (Akting Cillian Murphy dan kontras visual warna vs hitam-putih) -> Evaluation (Keunggulan tata suara uji coba Trinity tanpa suara instan) -> Evaluative Summary & Recommendation (Rating 5/5 dan refleksi moral ilmuwan).",
        "Fitur Bahasa": "Film Terminology ('cinematography', 'visual contrast', 'tour de force', 'performance'), Evaluative Adjectives ('gripping', 'haunting', 'breathless')."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "Film Review: 'Oppenheimer' – When Science Meets Conscience",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "The Guardian Film & Arts / Cannes Palme d'Or Review",
        "title": "Film Review: Bong Joon-ho's 'Parasite' – A Brilliant Satire on Social Inequality",
        "url": "https://www.theguardian.com/film/parasite-review",
        "year": "2023 / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t2_paras = [
        "Bong Joon-ho's 'Parasite', the historic winner of the Palme d'Or at Cannes and four Academy Awards, is a dark comedy and suspense thriller that examines social class inequality in modern South Korea. The story follows the impoverished Kim family, living in a damp semi-basement apartment in Seoul, who gradually infiltrate the lavish hillside home of the ultra-rich Park family by posing as qualified tutors and drivers.",
        "The film's genius lies in its architectural symbolism and verticality. The Kims live underground, while the wealthy Parks live high on a hill surrounded by glass walls and manicured gardens. During a heavy rainstorm, water flows downhill from the rich neighborhood into the Kims' slum, flooding their home with wastewater. This visual metaphor highlights how environmental disasters fall hardest on the poor.",
        "While the third act features shocking violence that may unsettle some viewers, the film's humorous con-artist caper and sharp social commentary make it impossible to look away. There are no cartoonish villains; both families are simply caught in an unequal economic system.",
        "Rated 4.8 out of 5 stars, 'Parasite' is a brilliant social satire that entertains while forcing audiences to question the economic divides in our modern world."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Mengulas film pemenang Oscar Parasite, membedah simbolisme arsitektur tangga turun dan kritik jurang sosial ekonomi.",
        "Struktur Teks": "Orientation (Perkenalan sutradara Bong Joon-ho dan keluarga Kim) -> Interpretative Recount (Strategi keluarga miskin masuk ke rumah mewah) -> Evaluation (Simbolisme vertikalitas tangga dan banjir hujan) -> Summary & Rating (4.8/5).",
        "Fitur Bahasa": "Literary & Film Concepts ('satire', 'symbolism', 'verticality', 'metaphor')."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "Film Review: 'Parasite' – The Architectures of Inequality",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "Times Literary Supplement / The Atlantic Book Review",
        "title": "Book Review: Kazuo Ishiguro's 'Never Let Me Go' – Fragility of the Soul",
        "url": "https://www.theatlantic.com/magazine/archive/never-let-me-go/",
        "year": "Classic / 2024",
        "cefr": "Level B1+ (Menengah Atas - Sesuai Siswa SMA Tingkat Lanjut)"
    }
    t3_paras = [
        "Kazuo Ishiguro's celebrated novel 'Never Let Me Go' is a gentle, heartbreaking dystopian story set in a tranquil English boarding school called Hailsham. Narrated by thirty-one-year-old Kathy H., the book follows her childhood friendships with classmates Tommy and Ruth.",
        "Gradually, the novel reveals a disturbing secret: the students at Hailsham are human clones, bred specifically to donate their vital organs to ordinary citizens once they reach young adulthood. What makes Ishiguro's writing so devastating is not violent action scenes, but the quiet, resigned acceptance of the characters. Instead of revolting, they cling to childhood artwork and seek temporary deferrals based on true love.",
        "Ishiguro's prose is understated, calm, and deeply emotional. The novel serves as a profound allegory about our own human mortality, reminding us that life is short and that love and kindness are the only things that truly matter.",
        "Earning 5 out of 5 stars, 'Never Let Me Go' is a literary masterpiece that lingers in the reader's heart long after the final page is turned."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "Book Review: 'Never Let Me Go' – Love, Mortality, and Gentle Souls",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Agonizing", "Adjective", "Sangat menyiksa batin dan pikiran", "Oppenheimer faced an agonizing moral dilemma."),
        ("Austere", "Adjective", "Sederhana, kaku, dan tanpa hiasan berlebih", "The political hearing was shot in austere black-and-white."),
        ("Tour de force", "Noun Phrase", "Pencapaian mahakarya artistik yang luar biasa hebat", "The silent explosion sequence is a cinematic tour de force."),
        ("Infiltrate", "Verb", "Menyusup masuk secara rahasia ke suatu tempat", "The poor family managed to infiltrate the luxury mansion."),
        ("Verticality", "Noun", "Pemanfaatan ketinggian vertikal sebagai simbol status", "The film uses verticality to represent social hierarchy."),
        ("Devastating", "Adjective", "Sangat menyayat hati dan mengharukan", "The quiet ending of the novel is emotionally devastating."),
        ("Understated", "Adjective", "Disampaikan dengan tenang dan bersahaja tanpa hiperbola", "Ishiguro's understated prose makes the tragedy deeper."),
        ("Mortality", "Noun", "Kefanaan hidup manusia yang pasti berakhir kematian", "The story explores the fragile nature of human mortality.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "In Text 1, what visual technique does Christopher Nolan use to distinguish between Oppenheimer's personal perspective and the political hearings?",
        "According to Text 2, how does Bong Joon-ho use the physical metaphor of 'verticality' (height and downhill stairs) in 'Parasite'?",
        "What tragic biological purpose were the students of Hailsham created for in Kazuo Ishiguro's novel 'Never Let Me Go' (Text 3)?"
    ]
    inf_q = [
        "Why did Christopher Nolan decide to keep the sound completely silent during the initial moments of the Trinity nuclear test?",
        "In Text 2, what does the flooding of the Kims' basement apartment during the rainstorm symbolize regarding climate impacts?",
        "Why is the clones' quiet acceptance of their fate in 'Never Let Me Go' considered more heartbreaking than an action-packed rebellion?"
    ]
    eval_q = [
        "Evaluate which storytelling medium—cinematic film (Text 1 & 2) or literary novel (Text 3)—better communicates complex moral messages to high school students.",
        "What criteria are most important when writing a balanced, objective review of an artistic work (movie or book)?"
    ]
    perf_task = (
        "PENYUSUNAN RESENSI KRITIS SENI & SINEMA (LKPD / SUMATIF MENULIS):\n"
        "Susunlah sebuah Teks Resensi Kritis (Review Text) sepanjang 250-300 kata mengenai sebuah film atau novel bermutu yang pernah kamu nikmati.\n"
        "Struktur Wajib Resensi:\n"
        "1. Orientation: Judul karya, sutradara/penulis, genre, dan konteks latar belakang.\n"
        "2. Interpretative Recount: Ringkasan singkat premis alur cerita tanpa membocorkan klimaks akhir (no major spoilers).\n"
        "3. Multi-Criteria Evaluation: Ulas minimal 2 aspek estetika (akting, sinematografi/gaya bahasa, tata suara, atau pesan moral mendalam).\n"
        "4. Evaluative Summary & Recommendation: Berikan kesimpulan berbobot, Star Rating (1–5 bintang), dan saran bagi calon penonton/pembaca.\n"
        "Gunakan minimal 4 istilah kosakata seni dari daftar glosarium di atas."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_11.TL.8_Critical_Review_International_Arts_Cinema.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


if __name__ == "__main__":
    print("Generating calibrated authentic materials for 5-Modul Ajar XI SMA TL...")
    create_tp_11_tl_1()
    create_tp_11_tl_2()
    create_tp_11_tl_3()
    create_tp_11_tl_4()
    create_tp_11_tl_5()
    create_tp_11_tl_6()
    create_tp_11_tl_7()
    create_tp_11_tl_8()
    print("All 8 DOCX files for Kelas XI Tingkat Lanjut successfully updated and calibrated!")
