import openpyxl
from build_core import build_sts_sheet, build_sas_sheet, build_gradebook_sheet

wb = openpyxl.Workbook()

# Setup sheets
ws_sts = wb.active
ws_sts.title = "STS_2_Topik"
ws_sas = wb.create_sheet(title="SAS_4_Topik")
ws_gradebook = wb.create_sheet(title="Rekap_Nilai_Siswa")

# Passages for Kelas 12
t1_passage = {
    "source": "Adapted from International Literary Archive & Contemporary Asian Review (2025)",
    "text": [
        "In the labyrinthine alleys of Semarang’s Old Town, Master Sastro had spent six decades hand-binding antique manuscripts with waxed hemp thread, breathing cedar, beeswax, and crushed clove oil.",
        "When Nadia, an ambitious young software developer armed with a holographic scanner, offered to digitize his life's work in twelve minutes, Sastro gently held a sun-bleached Javanese chronicle to her ear.",
        "'A machine captures the geometry of letters with ruthless precision,' he murmured, 'but it cannot register the trembling hand of the poet composing by candlelight in 1845; cultural memory is human vulnerability across generations, not pixel vectors.'"
    ]
}

t2_passage = {
    "source": "Adapted from World Economic Forum & Global Technological Deliberation Report (2025)",
    "text": [
        "The integration of autonomous cognitive algorithms into white-collar professions has sparked contentious global debate between technophilic optimism and socioeconomic anxiety.",
        "Proponents contend that automated workflow engines liberate professionals from clerical monotony, redirecting human cognition toward strategic reasoning, empathetic diplomacy, and new algorithmic auditing sectors.",
        "Conversely, labor economists emphasize that automated systems dismantle entry-level mentorship pathways and perpetuate training biases, necessitating strict multi-stakeholder governance to ensure human-centered augmentation rather than unbridled worker replacement."
    ]
}

# 10 Questions for STS
sts_questions = [
    {
        "no": 1, "topic": "Advanced Narrative", "kktp": "Menganalisis subteks filosofis dalam dialog tokoh",
        "stem": "What is the core philosophical distinction Master Sastro draws between digital scanning and physical manuscripts?",
        "options": ["A. Digital scanning is too expensive for local museums.", "B. Machines record optical data accurately but fail to transmit the visceral human soul, suffering, and historical vulnerability embedded in tactile artifacts.", "C. Waxed hemp thread is mechanically superior to cloud storage.", "D. Software engineers should never enter antique workshops.", "E. Dutch colonial parchment cannot be illuminated by LED lights."],
        "key": "B", "explanation": "Dialog Sastro menekankan bahwa mesin merekam geometri piksel, namun tidak dapat mentransmisikan jiwa, kerapuhan nurani, dan sejarah manusia pembuatnya."
    },
    {
        "no": 2, "topic": "Advanced Narrative", "kktp": "Menafsirkan perubahan batin tokoh (Epiphany)",
        "stem": "How did Nadia's perception transform after listening to Master Sastro's demonstration?",
        "options": ["A. She laughed sarcastically and broke his bone folder.", "B. She realized that cultural preservation transcends mere algorithmic data conversion and requires reverence for human experience.", "C. She resigned from her computer engineering company immediately.", "D. She purchased all of Sastro's manuscripts for her private collection.", "E. She replaced her scanner with a medieval printing press."],
        "key": "B", "explanation": "Momen pencerahan batin (epiphany) menyadarkan Nadia bahwa pelestarian warisan budaya melampaui hitungan piksel dan menuntut rasa hormat pada nilai kemanusiaan."
    },
    {
        "no": 3, "topic": "Advanced Narrative", "kktp": "Menganalisis gaya bahasa Participle Clause",
        "stem": "Examine the sentence: 'Tasked by a municipal archive consortium, Nadia viewed the ancient volumes merely as unindexed data sets.' What does the participle phrase indicate?",
        "options": ["A. A future intention she had planned.", "B. Her professional mandate and the utilitarian mindset through which she initially approached the task.", "C. That she was trespassing on private property.", "D. A biological reaction to ancient dust.", "E. An apology for arriving late."],
        "key": "B", "explanation": "Past participle clause 'Tasked by...' menerangkan latar belakang tugas profesional dan pola pikir utilitarian Nadia saat memulai pekerjaannya."
    },
    {
        "no": 4, "topic": "Advanced Narrative", "kktp": "Mengidentifikasi metafora dan citraan sensorik",
        "stem": "The description 'Dutch colonial archways whose plaster flaked like desiccated parchment' functions as...",
        "options": ["A. An hyperbole criticizing Dutch building materials.", "B. A sensory simile connecting the decaying architecture with the fragility of ancient paper manuscripts.", "C. An irrelevant technical digression.", "D. A scientific measurement of humidity.", "E. A dialogue marker."],
        "key": "B", "explanation": "Simile 'like desiccated parchment' menghubungkan kerapuhan arsitektur kota lama dengan kerapuhan naskah kertas kuno yang digeluti Sastro."
    },
    {
        "no": 5, "topic": "Advanced Narrative", "kktp": "Menyimpulkan tema sastra (Literary Theme)",
        "stem": "Which overarching literary theme is explored in 'The Memory Weaver of Kota Lama'?",
        "options": ["A. The inevitable victory of artificial intelligence over traditional arts.", "B. The profound tension between rapid technological efficiency and the enduring sanctity of human craftsmanship.", "C. The economic benefits of tourism in Central Java.", "D. The failure of public libraries in Southeast Asia.", "E. The legal patents of software algorithms."],
        "key": "B", "explanation": "Tema inti cerpen sastra ini adalah ketegangan dialektis antara efisiensi mesin digital dan kesucian nilai karya tangan manusia."
    },
    {
        "no": 6, "topic": "Discussion Text", "kktp": "Menganalisis perumusan isu kontroversial (Statement of Issue)",
        "stem": "What is the controversial issue introduced in the opening paragraph of the discussion text?",
        "options": ["A. The construction of new highways between metropolitan cities.", "B. Whether cognitive workplace automation represents human talent emancipation or mass socioeconomic disenfranchisement.", "C. The decline of physical desktop computers in secondary schools.", "D. Banning all computer science degrees in universities.", "E. Decreasing minimum wage rates for manufacturing workers."],
        "key": "B", "explanation": "Isu kontroversial yang dirumuskan secara berimbang adalah perdebatan apakah otomatisasi AI membebaskan talenta manusia atau menimbulkan peminggiran sosial ekonomi."
    },
    {
        "no": 7, "topic": "Discussion Text", "kktp": "Menganalisis argumen pendukung (Arguments For)",
        "stem": "According to proponents, what is the primary cognitive benefit of delegating repetitive analytical tasks to automated engines?",
        "options": ["A. Professionals can work zero hours per week while earning double salary.", "B. Human cognitive capacity can be redirected toward high-level strategic reasoning, ethical diplomacy, and interpersonal empathy.", "C. Employees can play video games during business meetings.", "D. Corporations can eliminate customer feedback channels.", "E. Computers will write all laws without human judges."],
        "key": "B", "explanation": "Pendukung otomatisasi menekankan pembebasan energi kognitif manusia agar terfokus pada penalaran strategis tingkat tinggi dan empati antarpribadi."
    },
    {
        "no": 8, "topic": "Discussion Text", "kktp": "Menganalisis argumen penentang (Arguments Against)",
        "stem": "What critical risk do labor economists highlight regarding entry-level professionals?",
        "options": ["A. Young workers will refuse to buy smartphones.", "B. The collapse of traditional apprenticeship pathways through which young graduates develop professional discernment.", "C. A complete shortage of office chairs in large corporations.", "D. Excessive manual paperwork slowing down internet routers.", "E. Senior managers being forced to retire at age thirty."],
        "key": "B", "explanation": "Kritik ekonomi buruh menyoroti terputusnya jalur magang karir pemula (loss of entry-level apprenticeship) akibat tugas dasar dikerjakan oleh perangkat lunak."
    },
    {
        "no": 9, "topic": "Discussion Text", "kktp": "Menganalisis penanda wacana kontras (Contrastive Markers)",
        "stem": "Which discourse marker is strategically utilized to shift the discussion from the optimistic pro-perspective to the critical counter-perspective?",
        "options": ["A. In addition", "B. Consequently", "C. Conversely", "D. For instance", "E. Similarly"],
        "key": "C", "explanation": "'Conversely' (sebaliknya) adalah konjungsi wacana kontras formal untuk beralih secara tajam dari sudut pandang pro ke sudut pandang kontra."
    },
    {
        "no": 10, "topic": "Discussion Text", "kktp": "Menyimpulkan sintesis jalan tengah (Balanced Conclusion)",
        "stem": "What balanced governance model does the author advocate in the concluding recommendation?",
        "options": ["A. Banning all automated software permanently across global markets.", "B. Prioritizing human-centered augmentation, employee reskilling, and mandatory human oversight in consequential decisions.", "C. Allowing tech corporations to replace workers without any state regulation.", "D. Taxing manual workers to subsidize robotics factories.", "E. Abandoning all higher education programs."],
        "key": "B", "explanation": "Kesimpulan berimbang menawarkan jalan tengah: otomatisasi yang berpusat pada manusia (*human-centered augmentation*), pelatihan ulang, dan pengawasan manusia."
    }
]

# 20 PG Questions for SAS (4 Topics: 5 each)
sas_pg_questions = [
    # Topik 1: Advanced Narrative Text (The Wayang Master)
    {
        "no": 1, "topic": "Advanced Narrative", "stimulus": "Teks Sastra: The Shadow and the Leather",
        "stem": "Master Ki Gondho chisels leather puppets not merely as theatrical props, but as 'living vessels of spiritual philosophy.' What does the shadow cast on the screen (kelir) symbolize in Javanese metaphysics?",
        "options": ["A. A commercial cinema projection.", "B. The transient, illusory nature of worldly existence reflecting an invisible divine light.", "C. The defeat of agricultural harvests.", "D. A scientific demonstration of solar eclipses.", "E. A secret code for royal rebellions."],
        "key": "B", "explanation": "Dalam filosofi wayang kulit, bayangan pada kelir melambangkan kefanaan dunia materi (*maya/bayangan*) yang bersumber dari cahaya sejati (*blencong/tuhan*)."
    },
    {
        "no": 2, "topic": "Advanced Narrative", "stimulus": "Teks Sastra: The Shadow and the Leather",
        "stem": "When his grandson suggests mass-producing puppets using laser cutters and acrylic sheets, Ki Gondho replies: 'A puppet carved without prayer carries no weight in the dark.' What does he mean?",
        "options": ["A. Acrylic is physically too heavy for the puppeteer's arm.", "B. Cultural artifacts severed from spiritual intentionality and ritual patience lack authentic aesthetic and moral resonance.", "C. Laser cutters use too much electricity in rural hamlets.", "D. He wants to patent his chisel designs.", "E. The dark screen requires special glass."],
        "key": "B", "explanation": "Ki Gondho menegaskan bahwa karya yang dibuat secara mekanik tanpa kesabaran spiritual kehilangan resonansi jiwa dan nilai sakralnya."
    },
    {
        "no": 3, "topic": "Advanced Narrative", "stimulus": "Teks Sastra: The Shadow and the Leather",
        "stem": "Identify the literary figure of speech: 'The ancient chisel whispered secrets to the dried buffalo hide.'",
        "options": ["A. Hyperbole", "B. Personification", "C. Euphemism", "D. Oxymoron", "E. Irony"],
        "key": "B", "explanation": "Personifikasi memberikan sifat manusia ('whispered secrets') kepada benda mati ('chisel')."
    },
    {
        "no": 4, "topic": "Advanced Narrative", "stimulus": "Teks Sastra: The Shadow and the Leather",
        "stem": "Which structural phase in advanced narration is marked by a deep, introspective evaluation before the conflict concludes?",
        "options": ["A. Dramatic Exposition", "B. Evaluative Epiphany", "C. Sudden Abrupt Interruption", "D. Introductory Heading", "E. Bibliography Citation"],
        "key": "B", "explanation": "'Evaluative Epiphany' adalah fase perenungan mendalam tokoh terhadap hakikat krisis sebelum narasi menemukan resolusinya."
    },
    {
        "no": 5, "topic": "Advanced Narrative", "stimulus": "Teks Sastra: The Shadow and the Leather",
        "stem": "In advanced literary prose, authors often utilize Cleft Sentences for rhetorical emphasis. Which of the following is an example?",
        "options": ["A. He carved the leather slowly.", "B. It was not the glittering gold leaf, but the sincere devotion of the craftsman that breathed life into the shadow.", "C. The puppet fell on the floor yesterday.", "D. Because it was night, the performance started.", "E. They left the village after the monsoon."],
        "key": "B", "explanation": "'It was not X, but Y that...' merupakan struktur Cleft Sentence untuk memberikan penekanan retoris tajam pada esensi pesan."
    },

    # Topik 2: Discussion Text (Nuclear Energy Transition)
    {
        "no": 6, "topic": "Discussion Text", "stimulus": "Wacana Debat Kebijakan: Nuclear Energy Transition in Southeast Asia",
        "stem": "What is the primary scientific premise advanced by proponents of nuclear energy in Indonesia?",
        "options": ["A. Nuclear reactors produce zero toxic waste in their entire lifespan.", "B. Nuclear energy provides dense, reliable, high-capacity base-load electricity with negligible operational carbon emissions.", "C. Nuclear plants can be constructed in less than forty-eight hours.", "D. Uranium fuel is completely free of charge.", "E. It eliminates the need for any electrical power grid."],
        "key": "B", "explanation": "Pendukung energi nuklir menonjolkan kemampuan menghasilkan listrik baseload berdaya raksasa tanpa emisi karbon operasional."
    },
    {
        "no": 7, "topic": "Discussion Text", "stimulus": "Wacana Debat Kebijakan: Nuclear Energy Transition in Southeast Asia",
        "stem": "Conversely, what ecological and geographical argument do environmental critics pose against nuclear reactors in the Indonesian archipelago?",
        "options": ["A. Indonesia has too much cold weather for nuclear cooling.", "B. The islands are located directly on the tectonic Ring of Fire, posing catastrophic risks of seismic liquefaction, tsunamis, and long-term spent fuel storage hazards.", "C. Indonesians do not use electricity during nighttime.", "D. Nuclear energy causes excessive rainfall.", "E. Solar panels are already installed on every tree."],
        "key": "B", "explanation": "Penentang menyoroti posisi geologis Indonesia di Ring of Fire yang rawan gempa, tsunami, likuefaksi, serta risiko limbah radioaktif jangka panjang."
    },
    {
        "no": 8, "topic": "Discussion Text", "stimulus": "Wacana Debat Kebijakan: Nuclear Energy Transition in Southeast Asia",
        "stem": "Which hedging expression in the text illustrates objective scholarly neutrality?",
        "options": ["A. 'It is undisputedly proven that everyone will die.'", "B. 'Evidence arguably suggests that small modular reactors (SMRs) may offer enhanced passive safety mechanisms.'", "C. 'Anyone who supports coal is completely insane.'", "D. 'I guarantee nothing bad will ever happen.'", "E. 'Nuclear physics is simply impossible.'"],
        "key": "B", "explanation": "'Evidence arguably suggests that... may offer...' merupakan bentuk hedging yang hati-hati dan objektif dalam wacana sains akademik."
    },
    {
        "no": 9, "topic": "Discussion Text", "stimulus": "Wacana Debat Kebijakan: Nuclear Energy Transition in Southeast Asia",
        "stem": "How should a balanced conclusion in a discussion text synthesize opposing viewpoints?",
        "options": ["A. By declaring one side completely triumphant and insulting the loser.", "B. By constructing a nuanced regulatory framework, such as deploying next-generation Small Modular Reactors (SMRs) only under stringent international IAEA oversight while accelerating renewables.", "C. By ignoring all factual evidence presented in the body paragraphs.", "D. By advising citizens to stop using electricity altogether.", "E. By postponing the conclusion for another fifty years."],
        "key": "B", "explanation": "Kesimpulan diskusi yang berimbang memadukan kehati-hatian risiko dengan kebutuhan transisi energi melalui kerangka regulasi ketat (IAEA oversight)."
    },
    {
        "no": 10, "topic": "Discussion Text", "stimulus": "Wacana Debat Kebijakan: Nuclear Energy Transition in Southeast Asia",
        "stem": "What is the structural role of 'Arguments For' and 'Arguments Against' in a Discussion Text?",
        "options": ["A. To trick the reader into buying foreign products.", "B. To demonstrate balanced critical inquiry by fairly presenting competing perspectives before reaching a reasoned judgment.", "C. To lengthen the document to meet bureaucratic page quotas.", "D. To replace scientific experiments with opinions.", "E. To confuse the grading committee."],
        "key": "B", "explanation": "Kedua bagian tersebut menyajikan dialektika pemikiran yang adil agar pembaca dapat menilai kompleksitas masalah secara objektif."
    },

    # Topik 3: Complex Argumentation & Diplomatic Agreement/Disagreement
    {
        "no": 11, "topic": "Argumentation & Agreement", "stimulus": "Simulasi Sidang Model United Nations (MUN)",
        "stem": "Delegate of Indonesia: 'While we acknowledge the imperative of global decarbonization, developing economies require substantial climate transition financing from industrialized nations.' \nDelegate of Germany: '________. The principle of Common But Differentiated Responsibilities necessitates equitable financial solidarity.'",
        "options": ["A. We categorically reject your unreasonable demands.", "B. Our delegation aligns itself entirely with that assessment; your premise is both just and indispensable.", "C. That is an irrelevant technical detail.", "D. We refuse to discuss climate change today.", "E. Why do you always complain about finances?"],
        "key": "B", "explanation": "'Our delegation aligns itself entirely with that assessment...' adalah ungkapan diplomasi multilateral untuk menyatakan persetujuan formal dan kokoh."
    },
    {
        "no": 12, "topic": "Argumentation & Agreement", "stimulus": "Simulasi Sidang Model United Nations (MUN)",
        "stem": "Which phrasing embodies the highest degree of diplomatic tact when expressing profound disagreement in international negotiations?",
        "options": ["A. 'Your resolution is a pack of lies.'", "B. 'With due deference to the distinguished delegate, we find ourselves unable to subscribe to the conclusions articulated in clause four.'", "C. 'You do not know what you are talking about.'", "D. 'We will walk out if you do not agree with us.'", "E. 'Nobody in the world supports your foolish idea.'"],
        "key": "B", "explanation": "Opsi B menggunakan register diplomasi tinggi ('With due deference... unable to subscribe to...') untuk menolak tanpa memicu konflik personal."
    },
    {
        "no": 13, "topic": "Argumentation & Agreement", "stimulus": "Simulasi Sidang Model United Nations (MUN)",
        "stem": "Identify the statement that constructs a Concessive Argument (mengakui poin lawan sebelum menyajikan bantahan utama):",
        "options": ["A. 'Nuclear energy is wonderful in every single aspect.'", "B. 'Notwithstanding the undeniable thermodynamic efficiency of nuclear power, the unresolved question of millenia-long geological waste repositories renders large-scale deployment contentious.'", "C. 'Coal is dirty, so we must stop today.'", "D. 'I love solar energy because the sun is bright.'", "E. 'Both sides are completely right and wrong.'"],
        "key": "B", "explanation": "'Notwithstanding the undeniable efficiency [konsesi], the unresolved question... renders it contentious [bantahan]' adalah argumen konsesif tingkat tinggi."
    },
    {
        "no": 14, "topic": "Argumentation & Agreement", "stimulus": "Simulasi Sidang Model United Nations (MUN)",
        "stem": "Which connective phrase signals that the speaker is summarizing their argument to build final consensus?",
        "options": ["A. On the contrary", "B. Taking all demonstrated evidence into collective consideration, we propose...", "C. Out of nowhere", "D. Suddenly and violently", "E. Before we were born"],
        "key": "B", "explanation": "Frasa B berfungsi sebagai sintesis konsensus dalam musyawarah formal sebelum pengambilan keputusan."
    },
    {
        "no": 15, "topic": "Argumentation & Agreement", "stimulus": "Simulasi Sidang Model United Nations (MUN)",
        "stem": "In an academic debate, 'Hedging' serves primarily to...",
        "options": ["A. Disguise ignorance with meaningless jargon.", "B. Express nuanced intellectual modesty, protect claims from categorical falsification, and invite constructive dialogue.", "C. Insult the opposing team subtly.", "D. Change the debate topic when cornered.", "E. Ensure the speaker talks for longer than the time limit."],
        "key": "B", "explanation": "Hedging menunjukkan kehati-hatian intelektual (*academic modesty*) dan mencegah klaim mutlak yang mudah dipatahkan data ekstrem."
    },

    # Topik 4: Conditional Sentences Type 1, 2, & 3 (including Inversion)
    {
        "no": 16, "topic": "Conditional 1, 2, & 3", "stimulus": "Analisis Pengandaian Tingkat Lanjut",
        "stem": "Complete with Conditional Type 3 (Unreal Past Condition & Regret): \n'If the engineers ________ (reinforce) the seawall adequately, the tsunami in Fukushima ________ (not trigger) the reactor meltdown.'",
        "options": ["A. reinforced - would not trigger", "B. had reinforced - would not have triggered", "C. reinforce - will not trigger", "D. have reinforced - would not trigger", "E. were reinforced - did not trigger"],
        "key": "B", "explanation": "Rumus Conditional Type 3: If + Past Perfect (had reinforced), Result: Subject + would/could + have + V3 (would not have triggered)."
    },
    {
        "no": 17, "topic": "Conditional 1, 2, & 3", "stimulus": "Analisis Pengandaian Tingkat Lanjut",
        "stem": "Examine the formal inverted conditional: 'Had Ki Ageng Semito listened to his inner conscience, his workshop would have remained intact.' This sentence is structurally identical to...",
        "options": ["A. If Ki Ageng listens to his conscience, his workshop remains intact.", "B. If Ki Ageng had listened to his inner conscience, his workshop would have remained intact.", "C. Did Ki Ageng listen to his conscience when the workshop fell?", "D. Although Ki Ageng had listened, his workshop fell.", "E. If Ki Ageng was listening, he built clocks."],
        "key": "B", "explanation": "Inversi Conditional Type 3: 'Had + Subject + V3' adalah bentuk formal yang setara persis dengan 'If + Subject + had + V3'."
    },
    {
        "no": 18, "topic": "Conditional 1, 2, & 3", "stimulus": "Analisis Pengandaian Tingkat Lanjut",
        "stem": "Which sentence demonstrates a Mixed Conditional (kondisi lampau yang berdampak pada kondisi masa kini)?",
        "options": ["A. If I study hard, I will pass the test.", "B. If our ancestors had not preserved our indigenous folklore, we would not understand our cultural identity today.", "C. If water reaches 100 degrees, it boils.", "D. If she had come yesterday, she would have seen him.", "E. If I were you, I would leave."],
        "key": "B", "explanation": "Mixed Conditional menggabungkan kondisi lampau (If our ancestors had not preserved...) dengan akibat di masa kini (...we would not understand today)."
    },
    {
        "no": 19, "topic": "Conditional 1, 2, & 3", "stimulus": "Analisis Pengandaian Tingkat Lanjut",
        "stem": "Complete the inverted Conditional Type 1 for diplomatic treaties: '________ you require further clarification regarding our decarbonization roadmap, our delegation remains entirely at your disposal.'",
        "options": ["A. Were", "B. Had", "C. Should", "D. Would", "E. Could"],
        "key": "C", "explanation": "Inversi formal untuk Conditional Type 1 menggunakan 'Should': 'Should you require...' (setara dengan 'If you require...')."
    },
    {
        "no": 20, "topic": "Conditional 1, 2, & 3", "stimulus": "Analisis Pengandaian Tingkat Lanjut",
        "stem": "Identify the inverted Conditional Type 2 structure expressing a hypothetical present scenario:",
        "options": ["A. Had I known the truth, I would have spoken.", "B. Were the government to mandate carbon taxes tomorrow, corporations would rapidly adjust their emission quotas.", "C. Should the rain stop, we will go.", "D. If he was arriving, call me.", "E. Will you help if I ask?"],
        "key": "B", "explanation": "Inversi formal Conditional Type 2 menggunakan 'Were + Subject + to Verb-1': 'Were the government to mandate...'."
    }
]

# 5 Essay Questions for SAS
sas_essay_questions = [
    {
        "no": 1, "topic": "Advanced Narrative",
        "context": "Studi Kasus Sastra: Analisis Karakter & Simbolisme (The Memory Weaver)",
        "prompt": "Jelaskan dalam 80–100 kata bagaimana penulis menggunakan kontras sensorik (sensory contrast) antara instrumen digital Nadia ('ultrathin holographic scanner, hum of cooling fan, vectors') dan bengkel Master Sastro ('cedar, beeswax, crushed clove oil, waxed hemp thread') untuk menyampaikan kritik terhadap dehumanisasi modern!",
        "rubric": "Skor 6: Analisis simbolik sangat tajam, menguraikan makna kontras sensorik kedua tokoh, bahasa akademik rapi.\nSkor 4: Analisis baik, menyebutkan kontras bau dan mesin, telaah kritik modernitas cukup jelas.\nSkor 2: Hanya menyalin teks tanpa analisis makna simbolisme.\nSkor 0: Kosong.",
        "key_model": "Model Analisis Sastra Ideal:\nThe author strategically establishes a profound sensory dichotomy to critique the sterile dehumanization of hyper-modernity. Nadia’s digital equipment is characterized by cold, frictionless, and synthetic attributes—the intrusive mechanical hum of cooling fans and disembodied pixel vectors—which reduce human heritage to transactional data points. In stark contrast, Master Sastro’s sanctuary is steeped in warm, organic, and tactile imagery—fragrant beeswax, aged cedar, and waxed hemp thread. Through these opposing sensory textures, the narrative underscores that authentic cultural memory is rooted in the visceral, embodied warmth of mortal human struggle, which cold algorithmic computation inevitably fails to capture."
    },
    {
        "no": 2, "topic": "Discussion Text",
        "context": "Studi Kasus Kebijakan: Menyusun Paragraf Sintesis Berimbang (Balanced Recommendation)",
        "prompt": "Berdasarkan perdebatan mengenai otomatisasi tempat kerja dan kecerdasan buatan, susunlah 1 paragraf 'Balanced Conclusion' (80–100 kata) yang menawarkan formula win-win solution antara efisiensi korporasi dan perlindungan martabat pekerja. Wajib memuat minimal 2 hedging expressions (e.g., arguably, tend to suggest) dan 1 kalimat pengandaian Type 3 atau Inversi!",
        "rubric": "Skor 6: Sintesis kebijakan berimbang, memuat 2+ hedging expressions, memuat 1 conditional 3/inversi sempurna.\nSkor 4: Rekomendasi baik, ada hedging, conditional sentence memadai.\nSkor 2: Memihak satu sisi ekstrem tanpa sintesis jalan tengah.\nSkor 0: Kosong.",
        "key_model": "Model Sintesis Kebijakan Ideal:\nIn conclusion, rather than succumbing to polarized dogma between unbridled automation and technological refusal, a mature society must champion human-centered augmentation. Evidence arguably suggests that productivity gains can sustainably co-exist with labor security if corporations allocate equitable margins toward lifelong employee reskilling. Were policymakers to establish enforceable ethical frameworks mandating human agency in critical decisions, generative technologies would serve as empowering cognitive scaffolds rather than instruments of displacement. Had previous industrial revolutions prioritized social safety nets from their inception, historical worker exploitation would have been substantially averted."
    },
    {
        "no": 3, "topic": "Diplomatic Argumentation",
        "context": "Simulasi Diplomasi Transisi Energi Berkeadilan",
        "prompt": "Tulislah sebuah pernyataan pembuka diplomatis (opening statement) sepanjang 70–90 kata untuk delegasi pemuda Indonesia dalam forum lingkungan hidup internasional. Pernyataan harus: (1) Menggunakan register diplomasi resmi, (2) Mengakui tanggung jawab global bersama, dan (3) Menekankan hak negara berkembang atas bantuan alih teknologi hijau.",
        "rubric": "Skor 6: Register diplomasi sangat formal dan santun, retorika berbobot, memadukan tanggung jawab dan keadilan iklim.\nSkor 4: Bahasa formal memadai, pesan tersampaikan dengan baik.\nSkor 2: Bahasa informal dan tidak mencerminkan protokol diplomasi.\nSkor 0: Kosong.",
        "key_model": "Model Pernyataan Diplomatis Ideal:\nHonorable delegates and esteemed chair, the youth delegation of the Republic of Indonesia reaffirms its resolute commitment to global climate decarbonization. While our nation vigorously safeguards our biodiverse carbon sinks and accelerates renewable infrastructures, justice dictates that the global climate burden be shouldered equitably. We respectfully urge industrialized partners to move beyond rhetorical declarations and deliver substantive green technology transfers alongside concessional financing. Only through authentic, transparent solidarity can developing economies achieve industrial decarbonization without compromising the socioeconomic well-being of our agrarian communities."
    },
    {
        "no": 4, "topic": "Conditional Sentences & Inversion",
        "context": "Transformasi Tata Bahasa Formal: Inversi Pengandaian",
        "prompt": "Ubah ketiga kalimat pengandaian standar berikut menjadi kalimat Inversi Formal (tanpa kata 'If'):\n1. 'If international inspectors visit the nuclear facility, they will observe strict safety protocols.'\n2. 'If the government were to mandate automated auditing, data fraud would plummet.'\n3. 'If our school had established an AI ethics guideline earlier, students would not have engaged in plagiarism.'",
        "rubric": "Skor 6: Ketiga kalimat diubah menjadi inversi 'Should...', 'Were...', dan 'Had...' secara sempurna.\nSkor 4: Dua kalimat inversi tepat, satu memiliki kesalahan struktur minor.\nSkor 2: Gagal menghilangkan kata 'if' atau susunan kata salah.\nSkor 0: Kosong.",
        "key_model": "Model Transformasi Inversi Baku:\n1. [Should]: Should international inspectors visit the nuclear facility, they will observe strict safety protocols.\n2. [Were]: Were the government to mandate automated auditing, data fraud would plummet.\n3. [Had]: Had our school established an AI ethics guideline earlier, students would not have engaged in plagiarism."
    },
    {
        "no": 5, "topic": "Comprehensive Analytical Reflection",
        "context": "Refleksi Akhir Semester Bahasa Inggris Fase F Tingkat Lanjut",
        "prompt": "Sebagai siswa kelas XII yang bersiap memasuki perguruan tinggi dan dunia profesional global, uraikan mengapa penguasaan teks wacana 'Discussion' dan kemampuan berargumentasi diplomatis merupakan keterampilan esensial abad ke-21! (Tuliskan dalam 60–80 kata bahasa Inggris akademis).",
        "rubric": "Skor 6: Refleksi komprehensif, menghubungkan kompetensi bahasa Inggris dengan tuntutan dunia akademik/profesional global.\nSkor 4: Refleksi cukup baik namun argumentasi agak sederhana.\nSkor 2: Terlalu singkat dan dangkal.\nSkor 0: Kosong.",
        "key_model": "Model Refleksi Komprehensif:\nMastering discussion discourse and diplomatic argumentation is paramount in the twenty-first century because global professional environments are characterized by complex, multidimensional challenges that cannot be resolved through rigid dogmatism. In higher education and international careers, professionals must critically deconstruct polar perspectives, accommodate cultural nuances, and articulate qualified concessions. Cultivating this rhetorical agility transforms high school graduates from passive information consumers into persuasive leaders capable of engineering consensus amid global uncertainty."
    }
]

# Build sheets
build_sts_sheet(ws_sts, "Kelas XII (Fase F)", "Advanced Narrative", t1_passage, "Discussion Text", t2_passage, sts_questions)
build_sas_sheet(ws_sas, "Kelas XII (Fase F)", "Advanced Narrative Text, Discussion Text, Complex Argumentation & Agreement, Conditional Sentences Type 1, 2, & 3", sas_pg_questions, sas_essay_questions)
build_gradebook_sheet(ws_gradebook, "Kelas XII (Fase F)")

out_path = "C:/Users/LILI/OneDrive - PDK Jateng/Dokumen/Perangkat_SMAN_1_Bae/Kelas_12/Latihan_Asesmen_Bahasa_Inggris_Kelas_12_SMAN_1_Bae.xlsx"
wb.save(out_path)
print("Workbook Kelas 12 generated successfully at:", out_path)
