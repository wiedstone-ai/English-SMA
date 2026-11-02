import openpyxl
from build_core import build_sts_sheet, build_sas_sheet, build_gradebook_sheet

wb = openpyxl.Workbook()

# Setup sheets
ws_sts = wb.active
ws_sts.title = "STS_2_Topik"
ws_sas = wb.create_sheet(title="SAS_4_Topik")
ws_gradebook = wb.create_sheet(title="Rekap_Nilai_Siswa")

# Passages for Kelas 11
t1_passage = {
    "source": "Adapted from Contemporary Indonesian Literary Heritage Review (2025)",
    "text": [
        "In the mist-shrouded foothills of Mount Muria during the late eighteenth century, master clockmaker Ki Ageng Semito crafted bronze timepieces of exquisite precision.",
        "However, his heart hardened with pride; he treated time purely as a gold commodity and harshly expelled an impoverished, soaked traveler seeking shelter during a violent storm.",
        "Moments later, a subterranean spring burst through his floor, flooding the workshop and spinning his golden clock backwards, teaching him through sudden destitution that human compassion transcends mechanical mastery."
    ]
}

t2_passage = {
    "source": "Adapted from Global Educational Technology & AI Ethics Journal (2025)",
    "text": [
        "The proliferation of generative artificial intelligence and synthetic media has fundamentally reshaped academic research and information consumption.",
        "While AI tools offer unprecedented productivity, they expose secondary school students to automated plagiarism, biased algorithms, and cognitive passivity.",
        "Therefore, integrating critical digital literacy into the high school curriculum is an urgent imperative to ensure youth evaluate source veracity, maintain academic integrity, and become discerning digital citizens."
    ]
}

# 10 Questions for STS
sts_questions = [
    {
        "no": 1, "topic": "Narrative Text", "kktp": "Menganalisis perkembangan watak tokoh (Character Arc)",
        "stem": "How did Ki Ageng Semito's character undergo a transformative arc from the beginning to the end of the narrative?",
        "options": ["A. He started as a compassionate hermit and became an arrogant merchant.", "B. He evolved from an arrogant, materialistic craftsman into a humble, compassionate human being after suffering devastating loss.", "C. He abandoned clockmaking to become a full-time sailor.", "D. He remained persistently hostile to travelers despite the flood.", "E. He gave up human relationships to live underground."],
        "key": "B", "explanation": "Alur karakter (character arc) Ki Ageng bergerak dari keangkuhan intelektual/harta menjadi kerendahan hati sejati setelah tertimpa musibah air bah."
    },
    {
        "no": 2, "topic": "Narrative Text", "kktp": "Menafsirkan simbolisme sastra (Symbolism)",
        "stem": "What does the clock's hands spinning wildly in reverse symbolize in the climax of the story?",
        "options": ["A. A faulty gear mechanism caused by cheap bronze.", "B. That time can literally be purchased with gold coins.", "C. The collapse of his illusion of controlling time and an opportunity for moral rebirth.", "D. A warning that an earthquake was approaching Central Java.", "E. That the traveler was a secret mechanical engineer."],
        "key": "C", "explanation": "Jarum jam yang berputar balik melambangkan runtuhnya ilusi arogansi manusia atas waktu, sekaligus simbol kesempatan kedua untuk bertobat."
    },
    {
        "no": 3, "topic": "Narrative Text", "kktp": "Menganalisis struktur komplikasi (Complication)",
        "stem": "Which event serves as the central moral conflict (Complication) that triggers the story's downfall?",
        "options": ["A. The government ordering twenty clocks.", "B. Ki Ageng's harsh refusal to grant shelter to a drenched traveler and a trembling child.", "C. The clockmaker selling his workshop to a foreign investor.", "D. The arrival of monsoon rain in November.", "E. The purchase of new tools from Europe."],
        "key": "B", "explanation": "Penolakan terhadap pengembara miskin yang kedinginan merupakan komplikasi etika yang memicu krisis karma moral berikutnya."
    },
    {
        "no": 4, "topic": "Narrative Text", "kktp": "Menganalisis penggunaan Past Tenses",
        "stem": "Examine the sentence: 'While he was putting the final touches to his masterwork, a traveler knocked upon his door.' What does the tense combination indicate?",
        "options": ["A. Two permanent facts in the present.", "B. An ongoing background past activity interrupted by a sudden past action.", "C. Two imaginary conditions in the future.", "D. A planned schedule that never occurred.", "E. A passive scientific process."],
        "key": "B", "explanation": "Past Continuous ('was putting') menggambarkan peristiwa yang sedang berlangsung di masa lampau yang disela oleh aksi Simple Past ('knocked')."
    },
    {
        "no": 5, "topic": "Narrative Text", "kktp": "Menyimpulkan nilai moral (Coda)",
        "stem": "What is the core moral lesson (Coda) conveyed by this parable?",
        "options": ["A. Wealthy individuals should install concrete waterproofing under their houses.", "B. Intellectual and material achievements are hollow without empathy and communal compassion.", "C. Clocks should not be made of bronze during monsoon seasons.", "D. Beggars should be reported to municipal authorities.", "E. Hard work always guarantees immunity from natural disasters."],
        "key": "B", "explanation": "Pesan moral (coda) utama adalah bahwa kecerdasan dan kekayaan materi tidak ada artinya tanpa welas asih dan empati sosial."
    },
    {
        "no": 6, "topic": "Analytical Exposition", "kktp": "Mengidentifikasi pernyataan tesis (Thesis Statement)",
        "stem": "What is the explicit Thesis Statement presented in the analytical exposition text?",
        "options": ["A. Artificial intelligence tools are completely safe and require no teacher supervision.", "B. Integrating critical digital literacy into the high school curriculum is an absolute educational imperative.", "C. Students should stop using computers for homework.", "D. High schools should replace all human teachers with synthetic chatbots.", "E. Plagiarism is acceptable if done using neural networks."],
        "key": "B", "explanation": "Paragraf 1 merumuskan tesis utama bahwa literasi digital kritis bukan lagi pilihan, melainkan keharusan pendidikan mutlak (absolute imperative)."
    },
    {
        "no": 7, "topic": "Analytical Exposition", "kktp": "Menganalisis argumen pendukung (Supporting Arguments)",
        "stem": "How does digital literacy help students combat machine-generated 'hallucinations'?",
        "options": ["A. By teaching them to accept every AI output without question.", "B. By equipping them with cross-referencing methodologies and source attribution skills to verify veracity.", "C. By forcing them to memorize entire encyclopedia volumes.", "D. By disabling their internet connection during school hours.", "E. By allowing them to submit unverified papers."],
        "key": "B", "explanation": "Literasi digital membekali siswa dengan metodologi cek-silang (cross-referencing) dan atribusi sumber agar mampu memverifikasi kebenaran data."
    },
    {
        "no": 8, "topic": "Analytical Exposition", "kktp": "Menganalisis konjungsi kausal (Causal Conjunctions)",
        "stem": "Which word in the text functions as a causal conjunction linking problem to necessity?",
        "options": ["A. Proliferation", "B. Therefore", "C. Unprecedented", "D. Reshaped", "E. Veracity"],
        "key": "B", "explanation": "'Therefore' adalah konjungsi kausal/kesimpulan yang menghubungkan premis risiko AI dengan kebutuhan mendesak kurikulum literasi digital."
    },
    {
        "no": 9, "topic": "Analytical Exposition", "kktp": "Membedakan fakta dan opini analitis",
        "stem": "Which of the following represents an analytical claim supported by logic, rather than a mere subjective preference?",
        "options": ["A. 'I believe robots are much cooler than human history books.'", "B. 'Critical cognition cannot be outsourced to neural networks because authentic problem-solving requires ethical judgment.'", "C. 'Smartphones are boring when it rains.'", "D. 'Nobody in our class likes studying science.'", "E. 'Video games are the best inventions ever.'"],
        "key": "B", "explanation": "Pernyataan B memuat klaim analitis yang didukung oleh penalaran logis dan relasi sebab-akibat (because authentic problem-solving requires...)."
    },
    {
        "no": 10, "topic": "Analytical Exposition", "kktp": "Menjelaskan fungsi bagian Reiteration",
        "stem": "How does the 'Reiteration' in an analytical exposition differ fundamentally from the 'Recommendation' in a hortatory exposition?",
        "options": ["A. Reiteration provides a funny joke to conclude.", "B. Reiteration restates the author's stance and thesis, whereas Recommendation urges a specific policy action to be taken.", "C. Reiteration introduces a completely new third topic.", "D. Reiteration must be written in the future continuous tense.", "E. There is no difference between them."],
        "key": "B", "explanation": "Reiterasi menegaskan kembali tesis penulis tanpa mendesak perintah tindakan (call to action), sedangkan rekomendasi mendesak apa yang harus dilakukan."
    }
]

# 20 PG Questions for SAS (4 Topics: 5 each)
sas_pg_questions = [
    # Topik 1: Narrative Text (Rawa Pening Reinterpreted)
    {
        "no": 1, "topic": "Narrative Text", "stimulus": "Teks Sastra: Baru Klinthing and the Greed of Pathok Village",
        "stem": "Why did the villagers of Pathok reject and humiliate Baru Klinthing when he appeared in their banquet?",
        "options": ["A. Because he refused to eat their sacrificial goat meat.", "B. Because they judged him superficiality based on his sickly, wound-covered physical appearance.", "C. Because he spoke an unrecognizable foreign dialect.", "D. Because he challenged the village chief to a physical duel.", "E. Because he stole their silver ceremonial chalice."],
        "key": "B", "explanation": "Warga desa menilai secara dangkal berdasarkan rupa fisik Baru Klinthing yang kumal dan penuh luka sehingga mengusirnya."
    },
    {
        "no": 2, "topic": "Narrative Text", "stimulus": "Teks Sastra: Baru Klinthing and the Greed of Pathok Village",
        "stem": "The wooden stick inserted into the ground by Baru Klinthing represents a narrative device known as...",
        "options": ["A. A comic relief prop.", "B. A moral test (crucible) that exposes the physical and spiritual impotence of arrogant people.", "C. A magical weapon to assassinate the village elder.", "D. A compass for finding lost cattle.", "E. An ordinary fence post without significance."],
        "key": "B", "explanation": "Lidi yang ditancapkan berfungsi sebagai 'moral crucible' (alat uji nurani) yang membuktikan ketidakberdayaan orang-orang sombong."
    },
    {
        "no": 3, "topic": "Narrative Text", "stimulus": "Teks Sastra: Baru Klinthing and the Greed of Pathok Village",
        "stem": "Why was the impoverished widow Nyai Latung spared from the cataclysmic flood that formed Lake Rawa Pening?",
        "options": ["A. She paid Baru Klinthing with bags of gold.", "B. She had shown authentic hospitality and kindness by offering him warm rice and accepting his warning.", "C. She was away trading spices in Semarang.", "D. She climbed a tall banyan tree before the feast.", "E. She challenged Baru Klinthing's magical powers."],
        "key": "B", "explanation": "Nyai Latung selamat karena ketulusan hatinya menyambut orang asing yang lapar dan mematuhi petunjuk keselamatan."
    },
    {
        "no": 4, "topic": "Narrative Text", "stimulus": "Teks Sastra: Baru Klinthing and the Greed of Pathok Village",
        "stem": "Which literary conflict is most prominent in the legend of Rawa Pening?",
        "options": ["A. Human versus Alien technology.", "B. Arrogant communal hypocrisy versus Sacred moral justice.", "C. Economic competition between two merchant guilds.", "D. Scientific skepticism versus superstitious rumors.", "E. Romantic dispute between royal courtiers."],
        "key": "B", "explanation": "Konflik dominan adalah kemunafikan dan keserakahan sosial masyarakat melawan keadilan moral sakral alam semesta."
    },
    {
        "no": 5, "topic": "Narrative Text", "stimulus": "Teks Sastra: Baru Klinthing and the Greed of Pathok Village",
        "stem": "Choose the correct indirect speech form of: Baru Klinthing said, 'He who lives with arrogance will drown in his own malice.'",
        "options": ["A. Baru Klinthing said that he who lived with arrogance would drown in his own malice.", "B. Baru Klinthing says that people are drowning yesterday.", "C. Baru Klinthing told that arrogance is good.", "D. Baru Klinthing asking whether arrogance was drowning.", "E. Baru Klinthing said he lived arrogant."],
        "key": "A", "explanation": "Aturan indirect speech: pergeseran tenses lampau (lives -> lived, will drown -> would drown) dengan konjungsi 'that'."
    },

    # Topik 2: Analytical Exposition (Renewable Energy in Rural Agriculture)
    {
        "no": 6, "topic": "Analytical Exposition", "stimulus": "Artikel Ilmiah: Solar-Powered Irrigation for Rural Farmers",
        "stem": "What is the primary socioeconomic thesis presented in the article on rural agricultural energy?",
        "options": ["A. Farmers should abandon rice cultivation and buy imported grain.", "B. Decentralized solar energy systems dramatically reduce agricultural operational costs and stabilize farmer livelihoods.", "C. Fossil fuel diesel pumps are superior in every environmental metric.", "D. Rural communities are incapable of managing micro-grid technology.", "E. Solar panels should only be installed in metropolitan capitals."],
        "key": "B", "explanation": "Tesis teks menegaskan bahwa sistem tenaga surya terdesentralisasi memangkas biaya operasional petani dan menstabilkan ketahanan pangan."
    },
    {
        "no": 7, "topic": "Analytical Exposition", "stimulus": "Artikel Ilmiah: Solar-Powered Irrigation for Rural Farmers",
        "stem": "According to the article's first argument, how does solar irrigation protect farmers against financial shocks?",
        "options": ["A. By offering lottery rewards every harvest.", "B. By liberating them from volatile diesel fuel prices and recurrent engine maintenance expenses.", "C. By guaranteeing free international flight tickets.", "D. By encouraging farmers to borrow high-interest loans.", "E. By making rain fall every morning."],
        "key": "B", "explanation": "Pompa tenaga surya membebaskan petani dari fluktuasi harga bahan bakar solar dan biaya perawatan mesin diesel yang mahal."
    },
    {
        "no": 8, "topic": "Analytical Exposition", "stimulus": "Artikel Ilmiah: Solar-Powered Irrigation for Rural Farmers",
        "stem": "Which modal adverb in the passage strengthens the author's conviction regarding environmental decarbonization?",
        "options": ["A. Questionably", "B. Indisputably", "C. Barely", "D. Reluctantly", "E. Accidentally"],
        "key": "B", "explanation": "'Indisputably' (secara tak terbantahkan) adalah modal adverb yang menegaskan kepastian tinggi (high certainty) dalam teks eksposisi."
    },
    {
        "no": 9, "topic": "Analytical Exposition", "stimulus": "Artikel Ilmiah: Solar-Powered Irrigation for Rural Farmers",
        "stem": "In an analytical exposition, what makes an 'Elaboration' robust and persuasive?",
        "options": ["A. Emotional insults directed at political figures.", "B. Empirical data, case study evidence, logical deductions, and real-world statistics.", "C. Rhyming poetry and humorous wordplay.", "D. Copying rumors from unverified social media blogs.", "E. Repeating the same sentence five times."],
        "key": "B", "explanation": "Elaborasi yang kokoh dibangun atas dasar data empiris, studi kasus nyata, dan penalaran logis deduktif."
    },
    {
        "no": 10, "topic": "Analytical Exposition", "stimulus": "Artikel Ilmiah: Solar-Powered Irrigation for Rural Farmers",
        "stem": "What is the main difference between Analytical Exposition and Hortatory Exposition regarding their endings?",
        "options": ["A. Analytical ends with Reiteration of the thesis; Hortatory ends with a Recommendation urging action.", "B. Analytical ends with a poem; Hortatory ends with a dialogue.", "C. Analytical has no ending at all.", "D. Hortatory never uses arguments.", "E. They are identical in structure and purpose."],
        "key": "A", "explanation": "Eksposisi analitis diakhiri Reiterasi (penegasan tesis), sedangkan eksposisi hortatori diakhiri Rekomendasi (anjuran tindakan nyata)."
    },

    # Topik 3: Asking/Giving Opinions & Stating Agreement/Disagreement
    {
        "no": 11, "topic": "Opinions & Agreement", "stimulus": "Diskusi Akademik Siswa SMAN 1 Bae",
        "stem": "Farhan: 'In my perspective, generative AI tools should be completely banned from high school essay competitions.' \nCitra: '________. While I acknowledge the risks of plagiarism, an outright ban ignores the reality that students must learn to use technology ethically.'",
        "options": ["A. I could not agree more with your statement.", "B. I see your point, but I respectfully disagree.", "C. Absolutely, that is the best idea ever.", "D. You are totally wrong and ignorant.", "E. I don't care about essay competitions."],
        "key": "B", "explanation": "'I see your point, but I respectfully disagree' menunjukkan penolakan diplomatis dan santun dalam wacana akademik formal."
    },
    {
        "no": 12, "topic": "Opinions & Agreement", "stimulus": "Diskusi Akademik Siswa SMAN 1 Bae",
        "stem": "Which expression is the most formal way to ask for a teacher's academic opinion during a seminar?",
        "options": ["A. 'Hey, what do you think about this?'", "B. 'What is your take on this stuff?'", "C. 'Could you kindly share your perspective regarding the pedagogical impact of this policy?'", "D. 'Tell me if this is good or bad right now.'", "E. 'Why are you thinking like that?'"],
        "key": "C", "explanation": "Opsi C menggunakan register bahasa akademik formal dan santun ('Could you kindly share your perspective regarding...')."
    },
    {
        "no": 13, "topic": "Opinions & Agreement", "stimulus": "Diskusi Akademik Siswa SMAN 1 Bae",
        "stem": "Rian: 'Implementing solar panels across our school buildings will significantly cut our carbon footprint.' \nMaya: '________. The data from our science club project clearly supports that conclusion.'",
        "options": ["A. That sounds totally absurd.", "B. I am of the same opinion; that is precisely correct.", "C. I doubt whether the sun ever shines here.", "D. I disagree because solar panels are too blue.", "E. I have no thought on electricity."],
        "key": "B", "explanation": "'I am of the same opinion; that is precisely correct' menyatakan persetujuan mutlak (strong agreement) berbasis data."
    },
    {
        "no": 14, "topic": "Opinions & Agreement", "stimulus": "Diskusi Akademik Siswa SMAN 1 Bae",
        "stem": "Identify the sentence that expresses qualified / partial agreement:",
        "options": ["A. 'I reject your proposal unconditionally.'", "B. 'I agree with your primary premise, provided that adequate funding is allocated transparently.'", "C. 'Whatever you say is always wrong.'", "D. 'I completely concur with everything without hesitation.'", "E. 'Let us cancel the whole debate.'"],
        "key": "B", "explanation": "'I agree... provided that...' menunjukkan persetujuan bersyarat/sebagian (qualified agreement)."
    },
    {
        "no": 15, "topic": "Opinions & Agreement", "stimulus": "Diskusi Akademik Siswa SMAN 1 Bae",
        "stem": "Which phrase is an inappropriate way to express disagreement in an academic school debate?",
        "options": ["A. 'I am afraid I have a divergent viewpoint on this matter.'", "B. 'That is a ridiculous argument and only a foolish person would say that.'", "C. 'With all due respect, the evidence points in a different direction.'", "D. 'I understand your reasoning; however, we must also consider the costs.'", "E. 'Allow me to offer an alternative interpretation.'"],
        "key": "B", "explanation": "Opsi B memuat serangan pribadi (ad hominem) yang tidak pantas dan melanggar etika debat akademik santun."
    },

    # Topik 4: Conditional Sentences Type 1 & Type 2
    {
        "no": 16, "topic": "Conditional 1 & 2", "stimulus": "Analisis Pengandaian Kontekstual",
        "stem": "Complete with Conditional Type 1 (Future Real Possibility): \n'If the school committee ________ (approve) our proposal, we ________ (install) a rainwater harvesting system next semester.'",
        "options": ["A. approves - will install", "B. will approve - install", "C. approved - would install", "D. had approved - would have installed", "E. approve - installed"],
        "key": "A", "explanation": "Rumus Conditional Type 1: If + Simple Present (approves), Result: Subject + will + V1 (will install)."
    },
    {
        "no": 17, "topic": "Conditional 1 & 2", "stimulus": "Analisis Pengandaian Kontekstual",
        "stem": "Complete with Conditional Type 2 (Unreal Present Condition): \n'If high school students ________ (possess) stronger critical digital literacy, they ________ (not fall) for fake news so easily.'",
        "options": ["A. possess - will not fall", "B. possessed - would not fall", "C. had possessed - would not have fallen", "D. possess - did not fall", "E. will possess - do not fall"],
        "key": "B", "explanation": "Rumus Conditional Type 2: If + Simple Past (possessed), Result: Subject + would/could + V1 (would not fall)."
    },
    {
        "no": 18, "topic": "Conditional 1 & 2", "stimulus": "Analisis Pengandaian Kontekstual",
        "stem": "Fact: 'Ki Ageng Semito is too arrogant to welcome the poor traveler, so he loses all his wealth.' \nWhat is the correct Conditional Type 2 expressing the hypothetical present reality?",
        "options": ["A. If Ki Ageng were not so arrogant, he would welcome the traveler warmly.", "B. If Ki Ageng is not arrogant, he will lose his workshop.", "C. If Ki Ageng had not been arrogant, he was rich.", "D. Ki Ageng will be humble if he becomes poor.", "E. Unless Ki Ageng welcomes the traveler, he is wealthy."],
        "key": "A", "explanation": "Kondisi hipotesis masa kini (Conditional 2): 'were' digunakan untuk semua subjek dalam ragam formal + 'would welcome'."
    },
    {
        "no": 19, "topic": "Conditional 1 & 2", "stimulus": "Analisis Pengandaian Kontekstual",
        "stem": "Choose the correct sentence following the standard rules of formal English conditionals:",
        "options": ["A. If I was you, I will buy that solar panel.", "B. If I were you, I would verify the online news source before sharing it.", "C. If she will come, we would start.", "D. If they studied, they will pass.", "E. If you heats water, it boiled."],
        "key": "B", "explanation": "Dalam bahasa Inggris formal, 'If I were you, I would...' adalah bentuk baku untuk memberikan saran hipotesis (Type 2)."
    },
    {
        "no": 20, "topic": "Conditional 1 & 2", "stimulus": "Analisis Pengandaian Kontekstual",
        "stem": "Teacher: 'What will happen if we fail to curb plastic waste on our campus?' \nStudent: 'If we do not take collective action now, our drainage system ________ (clog) during every heavy downpour.'",
        "options": ["A. would clog", "B. will clog", "C. clogged", "D. had clogged", "E. would have clogged"],
        "key": "B", "explanation": "Kalimat ini mengacu pada kemungkinan nyata masa depan (Conditional 1): 'If we do not take..., our drainage will clog'."
    }
]

# 5 Essay Questions for SAS
sas_essay_questions = [
    {
        "no": 1, "topic": "Narrative Text",
        "context": "Studi Kasus Sastra: Menulis Ulang Akhir Cerita (Alternative Narrative Resolution)",
        "prompt": "Tuliskan 1 paragraf narasi kreatif (70–90 kata) yang menyajikan 'Alternative Resolution' untuk kisah Ki Ageng Semito. Bayangkan jika Ki Ageng memilih menyambut pengembara tersebut sejak awal. Gunakan minimal 3 kata kerja lampau tindakan (action verbs) dan 1 kalimat pengandaian Type 2!",
        "rubric": "Skor 6: Alur resolusi kreatif koheren, memuat 3+ action verbs lampau, 1 conditional type 2 tepat secara gramatikal.\nSkor 4: Narasi baik, ada action verbs, namun conditional sentence kurang tepat.\nSkor 2: Terlalu singkat dan tidak memuat target tata bahasa.\nSkor 0: Kosong.",
        "key_model": "Model Jawaban Ideal:\nHearing the frail traveler’s cry amidst the howling storm, Ki Ageng Semito paused his golden timepiece and unlatched the oaken door. He invited the shivering grandmother and child inside, stoked the fireplace with dry cedar, and served steaming ginger tea. Gazing at their relieved faces, he realized: 'If I prioritized cold gears over human warmth, my soul would truly be impoverished.' By embracing empathy, Ki Ageng discovered that the most precious craftsmanship in life is the art of nurturing human compassion."
    },
    {
        "no": 2, "topic": "Analytical Exposition",
        "context": "Studi Kasus Opini Publik: Urgensi Etika Kecerdasan Buatan bagi Pelajar SMA",
        "prompt": "Susunlah 1 paragraf 'Argument' (Point + Elaboration) berstandar akademik (80–100 kata) yang membuktikan bahwa pemanfaatan AI generatif tanpa kontrol etis dapat mengikis kemampuan bernalar kritis (critical thinking) siswa. Gunakan minimal 2 konjungsi kausal/penegas wacana (e.g., consequently, furthermore, indisputably)!",
        "rubric": "Skor 6: Point dan Elaboration sangat logis, 2 konjungsi wacana terintegrasi alami, register akademik formal.\nSkor 4: Argumen baik, ada konjungsi, elaborasi agak umum.\nSkor 2: Opini dangkal tanpa rantai penalaran.\nSkor 0: Kosong.",
        "key_model": "Model Jawaban Ideal:\nFirstly, unchecked dependency on generative AI tools substantially erodes high school students' intrinsic cognitive endurance. When learners routinely outsource analytical essay composition and mathematical problem-solving to automated algorithms, they bypass the essential intellectual friction necessary for synaptic development. Furthermore, uncritical reliance fosters cognitive complacency, leading students to mistake algorithmic synthesis for authentic intellectual mastery. Consequently, unless adolescents actively engage in rigorous manual synthesis, their capacity for autonomous problem-solving will indisputably deteriorate in complex real-world environments."
    },
    {
        "no": 3, "topic": "Opinions & Agreement",
        "context": "Transkrip Debat Akademik SMA Negeri 1 Bae Kudus",
        "prompt": "Buatlah transkrip debat singkat (4 giliran bicara) antara dua pembicara (Bima dan Nadia) yang membahas topik: 'Should High School Curriculums Replace Printed Textbooks Completely with E-Books?' Pembicara harus menunjukkan: (1) Ungkapan memberi opini formal, (2) Ungkapan sanggahan santun (polite disagreement), dan (3) Solusi jalan tengah berbasis argumen logis.",
        "rubric": "Skor 6: Struktur debat formal, ungkapan opini dan sanggahan santun tepat, ada solusi jalan tengah.\nSkor 4: Dialog memadai, namun bahasa sanggahan kurang formal.\nSkor 2: Percakapan informal tanpa kaidah debat akademik.\nSkor 0: Kosong.",
        "key_model": "Model Transkrip Debat:\nBima: 'From my perspective, transitioning entirely to digital textbooks is imperative because it reduces printing costs and broadens information access.'\nNadia: 'I appreciate your reasoning, Bima; however, I respectfully disagree with an absolute elimination. Extensive research demonstrates that deep reading comprehension and memory retention remain significantly superior on physical paper.'\nBima: 'That is a compelling observation. Could we perhaps adopt a hybrid pedagogical model instead?'\nNadia: 'I completely concur with that compromise; utilizing e-books for research while preserving physical texts for foundational analysis balances both advantages.'"
    },
    {
        "no": 4, "topic": "Conditional Sentences",
        "context": "Analisis Pengandaian Kebijakan Lingkungan Sekolah",
        "prompt": "Diberikan fakta riil sekolah: 'SMAN 1 Bae currently generates significant single-use plastic waste because the canteen vendors do not provide reusable containers.' \nBuatlah:\n1. Satu kalimat Conditional Type 1 (Solusi masa depan yang realistis).\n2. Satu kalimat Conditional Type 2 (Refleksi hipotesis situasi saat ini).",
        "rubric": "Skor 6: Kedua kalimat pengandaian tersusun dengan kaidah rumus Conditional 1 & 2 yang sempurna dan kontekstual.\nSkor 4: Satu kalimat benar, satu kalimat memiliki galat tenses minor.\nSkor 2: Rumus conditional salah.\nSkor 0: Kosong.",
        "key_model": "Model Jawaban Ideal:\n1. [Conditional Type 1]: If the school administration mandates reusable food containers next month, our campus will substantially reduce single-use plastic waste.\n2. [Conditional Type 2]: If the canteen vendors provided biodegradable packaging right now, our school would not produce such alarming volumes of garbage."
    },
    {
        "no": 5, "topic": "Comparative Synthesis",
        "context": "Sintesis Lintas Teks: Narrative vs Analytical Exposition",
        "prompt": "Jelaskan dalam 3–4 kalimat analitis bagaimana 'Teks Naratif' dan 'Teks Eksposisi Analitis' memiliki cara yang berbeda dalam memengaruhi pikiran dan moral pembaca mengenai isu ketamakan manusia (greed)!",
        "rubric": "Skor 6: Mampu membandingkan mekanisme retorika kedua teks secara mendalam dan analitis.\nSkor 4: Penjelasan cukup jelas namun komparasi kurang tajam.\nSkor 2: Hanya mendefinisikan masing-masing teks tanpa analisis perbandingan.\nSkor 0: Kosong.",
        "key_model": "Model Jawaban Ideal:\nWhile a narrative text persuades readers through emotional immersion—using character flaws, dramatic tragedy, and relatable moral consequences to evoke empathy—an analytical exposition influences the audience through intellectual logic, deploying structured arguments, causal connectors, and empirical facts to prove the objective destructiveness of human greed. Narrative engages the human heart, whereas analytical exposition convinces the human intellect."
    }
]

# Build sheets
build_sts_sheet(ws_sts, "Kelas XI (Fase F)", "Narrative Text", t1_passage, "Analytical Exposition", t2_passage, sts_questions)
build_sas_sheet(ws_sas, "Kelas XI (Fase F)", "Narrative Text, Analytical Exposition, Asking/Giving Opinions & Agreement, Conditional Sentences Type 1 & 2", sas_pg_questions, sas_essay_questions)
build_gradebook_sheet(ws_gradebook, "Kelas XI (Fase F)")

out_path = "C:/Users/LILI/OneDrive - PDK Jateng/Dokumen/Perangkat_SMAN_1_Bae/Kelas_11/Latihan_Asesmen_Bahasa_Inggris_Kelas_11_SMAN_1_Bae.xlsx"
wb.save(out_path)
print("Workbook Kelas 11 generated successfully at:", out_path)
