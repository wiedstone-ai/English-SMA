import openpyxl
from build_core import build_sts_sheet, build_sas_sheet, build_gradebook_sheet

wb = openpyxl.Workbook()

# Setup sheets
ws_sts = wb.active
ws_sts.title = "STS_2_Topik"
ws_sas = wb.create_sheet(title="SAS_4_Topik")
ws_gradebook = wb.create_sheet(title="Rekap_Nilai_Siswa")

# Passages for Kelas 10
t1_passage = {
    "source": "Adapted from Indonesia Travel & Marine Conservation Bulletin (2025)",
    "text": [
        "Karimunjawa Marine National Park is a pristine archipelago of twenty-seven coral islands situated 80 km northwest of Jepara, Central Java.",
        "The park is renowned for its turquoise waters, fringing coral reefs, and extensive mangrove ecosystems that shelter clownfish, sea turtles, and white-bellied sea eagles.",
        "Local fishermen actively manage community shark conservation pools on Menjangan Besar Island, demonstrating how biodiversity and sustainable tourism flourish together."
    ]
}

t2_passage = {
    "source": "Adapted from World Wildlife & Tropical Ecology Research (2025)",
    "text": [
        "The Javan Rhinoceros (Rhinoceros sondaicus) is a critically endangered mammal of the family Rhinocerotidae, strictly confined to Ujung Kulon National Park in Java.",
        "It possesses armor-plated dusky skin and typically a single horn rarely exceeding twenty-five centimeters in length.",
        "As a solitary herbivore, it relies on natural mud wallows to maintain body temperature, with fewer than eighty individuals currently surviving in the wild."
    ]
}

# 10 Questions for STS
sts_questions = [
    {
        "no": 1, "topic": "Descriptive Text", "kktp": "Menganalisis ide pokok deskripsi geografis",
        "stem": "What is the primary ecological purpose of Karimunjawa Marine National Park as depicted in the passage?",
        "options": ["A. To function as an industrial fishing harbor for Central Java.", "B. To preserve pristine coral archipelago ecosystems alongside sustainable community eco-tourism.", "C. To replace natural mangroves with modern beachside resorts.", "D. To breed captive sea turtles for commercial export.", "E. To restrict all maritime access from local fishing communities."],
        "key": "B", "explanation": "Teks secara eksplisit menguraikan Karimunjawa sebagai suaka pulau karang terlindungi yang memadukan keanekaragaman hayati dengan pariwisata berkelanjutan."
    },
    {
        "no": 2, "topic": "Descriptive Text", "kktp": "Menemukan informasi rinci tentang konservasi hiu",
        "stem": "According to the passage, what specific community-led conservation effort takes place on Menjangan Besar Island?",
        "options": ["A. Construction of deep-sea drilling platforms.", "B. Management of community shark conservation and rehabilitation pools.", "C. Commercial harvesting of rare scleractinian corals.", "D. Deforestation of tropical rain forests for farmland.", "E. Banning all eco-tourist visits to Tanjung Gelam."],
        "key": "B", "explanation": "Paragraf terakhir menyebutkan bahwa nelayan lokal di Pulau Menjangan Besar mengelola kolam konservasi hiu secara partisipatif."
    },
    {
        "no": 3, "topic": "Descriptive Text", "kktp": "Menganalisis fungsi adjective phrase dalam kalimat deskriptif",
        "stem": "In the phrase 'a pristine archipelago of twenty-seven coral islands', the word 'pristine' is closest in meaning to...",
        "options": ["A. Densely populated", "B. Heavily polluted", "C. Pure, unspoiled and untouched", "D. Artificially engineered", "E. Temporarily abandoned"],
        "key": "C", "explanation": "'Pristine' bermakna murni, alami, dan belum terjamah kerusakan atau polusi lingkungan."
    },
    {
        "no": 4, "topic": "Descriptive Text", "kktp": "Mengidentifikasi fitur kebahasaan Simple Present",
        "stem": "Why does the writer predominantly employ the Simple Present Tense throughout the descriptive text?",
        "options": ["A. To narrate a chronological series of past events.", "B. To establish factual descriptions, permanent geographical features, and timeless truths.", "C. To predict future technological advancements in marine navigation.", "D. To present balanced arguments between conflicting stakeholders.", "E. To issue procedural commands for setting up scuba equipment."],
        "key": "B", "explanation": "Simple Present Tense dalam teks deskriptif berfungsi menyatakan fakta permanen dan karakteristik fisik yang berlangsung hingga kini."
    },
    {
        "no": 5, "topic": "Descriptive Text", "kktp": "Menyimpulkan sikap penulis (Tone/Attitude)",
        "stem": "What is the author's tone toward the community of Karimunjawa?",
        "options": ["A. Critical and suspicious of their intentions.", "B. Indifferent and emotionally detached.", "C. Appreciative of their harmonious integration of preservation and tourism.", "D. Sarcastic about their traditional fishing methods.", "E. Pessimistic about the survival of the national park."],
        "key": "C", "explanation": "Penulis secara positif mengapresiasi keharmonisan antara pelestarian alam dan keramahan tradisi masyarakat lokal."
    },
    {
        "no": 6, "topic": "Report Text", "kktp": "Membedakan General Classification pada Report Text",
        "stem": "Which statement best serves as the scientific 'General Classification' of the Javan Rhinoceros?",
        "options": ["A. It loves playing in the mud when the weather is warm.", "B. It is a critically endangered mammal of the family Rhinocerotidae and one of five extant rhinoceros species.", "C. The rhino has thick skin that looks like European knight armor.", "D. Fewer than eighty individuals were counted by forest rangers last month.", "E. Tourists should avoid approaching solitary rhinos during mating seasons."],
        "key": "B", "explanation": "General Classification memuat pernyataan klasifikasi taksonomi biologi ilmiah (ordo/famili) secara umum, bukan deskripsi perilaku insidental."
    },
    {
        "no": 7, "topic": "Report Text", "kktp": "Menemukan fakta biologis tentang ciri fisik",
        "stem": "How does the physical morphology of the Javan rhinoceros differ significantly from African rhino species?",
        "options": ["A. Javan rhinos have two massive horns while African rhinos have none.", "B. Javan rhinos typically possess only a single small horn, unlike two-horned African rhinos.", "C. Javan rhinos have feather-like fur covering their entire limbs.", "D. Javan rhinos are strictly aquatic mammals that swim in open oceans.", "E. Javan rhinos weigh less than one hundred kilograms."],
        "key": "B", "explanation": "Teks mencatat bahwa berbeda dengan badak Afrika yang bertanduk ganda, badak Jawa biasanya hanya memiliki satu cula kecil."
    },
    {
        "no": 8, "topic": "Report Text", "kktp": "Menganalisis fungsi ekologis perilaku satwa",
        "stem": "Why is the habit of wallowing in natural mud pools vital for the survival of the Javan rhino?",
        "options": ["A. It helps them camouflage from satellite cameras.", "B. It regulates body temperature and protects dermal layers against insect parasites.", "C. It is a ritual to mark territory against sea eagles.", "D. It replaces their nutritional requirement for drinking freshwater.", "E. It facilitates rapid vocal communication across jungle valleys."],
        "key": "B", "explanation": "Teks menyatakan berkubang di lumpur penting untuk mengatur suhu tubuh dan melindungi kulit dari gigitan serangga dan parasit."
    },
    {
        "no": 9, "topic": "Report Text", "kktp": "Menganalisis penggunaan Passive Voice ilmiah",
        "stem": "Examine the sentence: 'The entire global population is now strictly confined to Ujung Kulon National Park.' Why is the passive voice used?",
        "options": ["A. Because the writer forgot who confined the rhinos.", "B. To emphasize the population and their critical geographical state rather than the external agent.", "C. To make the sentence sound informal and conversational.", "D. Because active voice cannot be used with geographical names.", "E. To suggest that the rhinos chose to build their own confinement."],
        "key": "B", "explanation": "Passive voice pada teks laporan ilmiah menekankan objek kajian dan status sebarannya secara objektif daripada pelaku penjaga hutan."
    },
    {
        "no": 10, "topic": "Report Text", "kktp": "Menyimpulkan tindakan konservasi mendesak",
        "stem": "Based on the text, what urgent inference can be drawn regarding the future of Rhinoceros sondaicus?",
        "options": ["A. Their population is expanding rapidly across Southeast Asia.", "B. Without intensive continuous habitat protection and anti-poaching vigilance, irreversible extinction is imminent.", "C. They should be domesticated and sold as exotic pets.", "D. The species no longer requires conservation funding due to natural resilience.", "E. Relocating them to urban zoos will permanently solve their survival challenges."],
        "key": "B", "explanation": "Dengan populasi kurang dari 80 ekor, ancaman kepunahan permanen sangat nyata tanpa pengawasan anti-perburuan yang ketat."
    }
]

# 20 PG Questions for SAS (4 Topics: 5 each)
sas_pg_questions = [
    # Topik 1: Descriptive Text (Borobudur)
    {
        "no": 1, "topic": "Descriptive Text", "stimulus": "Teks: The Architectural Splendor of Borobudur",
        "stem": "Borobudur Temple, constructed in the 9th century during the Sailendra Dynasty, is designed in the shape of a magnificent three-dimensional mandala. What does this design symbolize?",
        "options": ["A. A commercial marketplace for medieval spice trade.", "B. The Buddhist cosmological journey from the realm of desires (Kamadhatu) to formlessness (Arupadhatu).", "C. A military fortress defending Central Java against maritime invaders.", "D. A royal palace dedicated to agricultural harvest festivals.", "E. A temporary monument erected for foreign diplomatic summits."],
        "key": "B", "explanation": "Secara kosmologis dan arsitektural, mandala Borobudur melambangkan tingkatan alam spiritual Kamadhatu, Rupadhatu, hingga Arupadhatu."
    },
    {
        "no": 2, "topic": "Descriptive Text", "stimulus": "Teks: The Architectural Splendor of Borobudur",
        "stem": "The stone bas-reliefs decorating the galleries of Borobudur depict...",
        "options": ["A. Modern industrial machinery and steam engines.", "B. Sacred narratives of Buddhist scriptures, folklore, and daily life in ancient Java.", "C. Fictional accounts of space voyages.", "D. European colonial treaties of the nineteenth century.", "E. Mathematical formulas for irrigation canals."],
        "key": "B", "explanation": "Relief batu pada dinding Borobudur memuat kisah kitab suci Buddha (seperti Lalitavistara) serta cerminan kehidupan masyarakat Jawa kuno."
    },
    {
        "no": 3, "topic": "Descriptive Text", "stimulus": "Teks: The Architectural Splendor of Borobudur",
        "stem": "Which adjective phrase best captures the visual grandeur of Borobudur at sunrise?",
        "options": ["A. A completely deserted, dilapidated concrete bunker.", "B. An awe-inspiring stone sanctuary silhouetted against the morning mist.", "C. A brightly painted contemporary leisure park.", "D. A chaotic, noisy urban intersection.", "E. An inaccessible underwater coral mound."],
        "key": "B", "explanation": "Frasa 'an awe-inspiring stone sanctuary silhouetted against the morning mist' memuat sensory imagery yang tepat untuk fajar Borobudur."
    },
    {
        "no": 4, "topic": "Descriptive Text", "stimulus": "Teks: The Architectural Splendor of Borobudur",
        "stem": "What is the primary function of the 'Identification' section in a descriptive text about a world heritage site?",
        "options": ["A. To list the chronological historical battles fought on the premises.", "B. To introduce the specific monument, its geographical location, and historic significance.", "C. To argue whether ticket prices are too high for domestic visitors.", "D. To provide step-by-step instructions on cleaning stone moss.", "E. To conclude that the monument should be demolished."],
        "key": "B", "explanation": "Identification bertujuan memperkenalkan objek spesifik, letak geografis, dan latar belakang signifikansi sebelum rincian detail dipaparkan."
    },
    {
        "no": 5, "topic": "Descriptive Text", "stimulus": "Teks: The Architectural Splendor of Borobudur",
        "stem": "In descriptive writing, sensory adjectives such as 'intricate', 'porous', and 'majestic' serve to...",
        "options": ["A. Confuse the reader with irrelevant jargon.", "B. Enable the reader to mentally visualize and sense the physical texture and aura of the subject.", "C. State personal arguments in a persuasive debate.", "D. Change the grammatical tense of the paragraph to past perfect.", "E. Shorten the word count of the essay."],
        "key": "B", "explanation": "Sensory adjectives memicu daya imajinasi indrawi pembaca agar dapat membayangkan tekstur dan keagungan objek secara nyata."
    },
    
    # Topik 2: Report Text (Geothermal Energy)
    {
        "no": 6, "topic": "Report Text", "stimulus": "Teks: Geothermal Energy in Indonesia",
        "stem": "Indonesia holds approximately forty percent of the world's geothermal energy reserves because it is situated along...",
        "options": ["A. The Saharan desert belt.", "B. The Pacific Ring of Fire characterized by extensive tectonic and volcanic activity.", "C. The Arctic polar circle.", "D. The Mariana deep ocean trench.", "E. The Eurasian limestone plateaus."],
        "key": "B", "explanation": "Potensi panas bumi Indonesia melimpah karena posisinya di cincin api Pasifik (Ring of Fire) yang kaya aktivitas vulkanik."
    },
    {
        "no": 7, "topic": "Report Text", "stimulus": "Teks: Geothermal Energy in Indonesia",
        "stem": "How do geothermal power plants generate base-load electricity?",
        "options": ["A. By burning coal and petroleum in high-pressure furnaces.", "B. By harnessing high-temperature subterranean steam to drive turbine generators.", "C. By capturing wind gusts on high mountain ridges.", "D. By collecting solar radiation through photovoltaic arrays.", "E. By splitting uranium atoms in nuclear fission reactors."],
        "key": "B", "explanation": "Pembangkit geotermal memanfaatkan uap panas bumi bertekanan tinggi dari bawah permukaan bumi untuk memutar turbin listrik."
    },
    {
        "no": 8, "topic": "Report Text", "stimulus": "Teks: Geothermal Energy in Indonesia",
        "stem": "Which characteristic classifies geothermal energy as a 'renewable and sustainable' resource?",
        "options": ["A. The subterranean water can be reinjected into the reservoir and naturally reheated by the Earth's magma.", "B. It completely destroys the surrounding forest within five months.", "C. It can only generate power during dry sunny afternoons.", "D. Its fuel must be imported from overseas oil refineries.", "E. It leaves behind radioactive nuclear waste."],
        "key": "A", "explanation": "Fluida panas bumi dapat diinjeksikan kembali ke dalam reservoir tanah untuk dipanaskan ulang secara alami oleh magma bumi."
    },
    {
        "no": 9, "topic": "Report Text", "stimulus": "Teks: Geothermal Energy in Indonesia",
        "stem": "What is the structural distinction between the 'General Classification' and 'Description' in this report text?",
        "options": ["A. Classification tells an imaginative story, while Description gives moral advice.", "B. Classification categorizes geothermal as a thermal renewable energy, while Description elaborates on mechanisms, reservoirs, and environmental benefits.", "C. Classification lists complaints, while Description gives apologies.", "D. Classification gives numerical data only, while Description gives author opinions.", "E. There is no distinction in report texts."],
        "key": "B", "explanation": "General classification mengelompokkan jenis energi, sedangkan description membedah cara kerja sistem dan potensinya."
    },
    {
        "no": 10, "topic": "Report Text", "stimulus": "Teks: Geothermal Energy in Indonesia",
        "stem": "Which of the following sentences exhibits the formal scientific passive voice typical of report texts?",
        "options": ["A. Yesterday we went to see a geothermal well in Dieng.", "B. Steam is extracted from deep production wells and channeled into moisture separators.", "C. I really love renewable energy because it is cool.", "D. If you visit Kamojang, you will see steam everywhere.", "E. The engineer shouted loudly to start the engine."],
        "key": "B", "explanation": "'Steam is extracted... and channeled...' merupakan bentuk passive voice objektif yang umum digunakan dalam teks sains teknis."
    },

    # Topik 3: Simple Present Tense & Active / Passive Voice
    {
        "no": 11, "topic": "Grammar: Present & Passive", "stimulus": "Konteks Ekologi & Tata Bahasa",
        "stem": "Complete the scientific sentence: 'Mangrove roots ________ (trap) sediment and ________ (protect) shorelines from sea erosion.'",
        "options": ["A. traps - protects", "B. trapped - protected", "C. trap - protect", "D. are trapping - are protecting", "E. have trapped - will protect"],
        "key": "C", "explanation": "Subjek 'Mangrove roots' adalah jamak (plural), sehingga kata kerja Simple Present berbentuk dasar (Verb-1): 'trap' dan 'protect'."
    },
    {
        "no": 12, "topic": "Grammar: Present & Passive", "stimulus": "Konteks Ekologi & Tata Bahasa",
        "stem": "Transform into Passive Voice: 'Environmental researchers closely monitor the endangered wildlife population.'",
        "options": ["A. The endangered wildlife population monitored by environmental researchers.", "B. The endangered wildlife population is closely monitored by environmental researchers.", "C. The endangered wildlife population was closely monitored by environmental researchers.", "D. The endangered wildlife population has been closely monitored.", "E. Environmental researchers are being monitored by the wildlife."],
        "key": "B", "explanation": "Pola passive Simple Present: Subject (singular: population) + is + V3 (closely monitored) + by Agent."
    },
    {
        "no": 13, "topic": "Grammar: Present & Passive", "stimulus": "Konteks Ekologi & Tata Bahasa",
        "stem": "Identify the grammatically correct passive sentence:",
        "options": ["A. Solar panels is installed on top of the school roof.", "B. Organic compost are produced by agricultural students.", "C. Geothermal energy is utilized for clean electricity generation across Central Java.", "D. Endangered species is protect by national legislation.", "E. Many historic artifacts was discovered in the valley."],
        "key": "C", "explanation": "'Geothermal energy' (singular/uncountable) berpasangan dengan 'is utilized' (V3) secara benar."
    },
    {
        "no": 14, "topic": "Grammar: Present & Passive", "stimulus": "Konteks Ekologi & Tata Bahasa",
        "stem": "Choose the correct verb form: 'Each member of the eco-tourism committee ________ (strive) to minimize plastic waste.'",
        "options": ["A. strive", "B. strives", "C. striving", "D. are striving", "E. were striven"],
        "key": "B", "explanation": "Frasa 'Each member...' memiliki subjek singular ('Each member'), sehingga kata kerja Simple Present harus berakhiran -s ('strives')."
    },
    {
        "no": 15, "topic": "Grammar: Present & Passive", "stimulus": "Konteks Ekologi & Tata Bahasa",
        "stem": "Which sentence correctly changes 'Local craftsmen carve intricate wooden reliefs in Kudus' into passive?",
        "options": ["A. Intricate wooden reliefs are carved by local craftsmen in Kudus.", "B. Intricate wooden reliefs is carved by local craftsmen in Kudus.", "C. Intricate wooden reliefs were carved by local craftsmen in Kudus.", "D. Local craftsmen are carved by intricate wooden reliefs in Kudus.", "E. In Kudus, intricate wooden reliefs has been carved."],
        "key": "A", "explanation": "'Reliefs' bersifat jamak, maka to be passive-nya adalah 'are' + V3 'carved'."
    },

    # Topik 4: Adjective Phrases & Stating Intentions/Feelings
    {
        "no": 16, "topic": "Grammar: Adjective Phrases & Intentions", "stimulus": "Percakapan Sekolah & Frasa Sifat",
        "stem": "Rafi: 'What are your plans after graduating from SMA Negeri 1 Bae?' \nSalma: 'I ________ environmental engineering at a state university because I want to conserve water resources.'",
        "options": ["A. will have planned", "B. am intending to study", "C. studied definitely", "D. would rather to study", "E. have been studied"],
        "key": "B", "explanation": "'am intending to study' / 'intend to study' menyatakan niat atau rencana masa depan (stating intention) secara tepat."
    },
    {
        "no": 17, "topic": "Grammar: Adjective Phrases & Intentions", "stimulus": "Percakapan Sekolah & Frasa Sifat",
        "stem": "Teacher: 'How do you feel about presenting our eco-tourism campaign in front of the regional council tomorrow?' \nNadia: 'To be honest, I feel ________ about the feedback, but I am confident in our team's research.'",
        "options": ["A. a little anxious", "B. completely destructively", "C. without feeling", "D. boring very much", "E. delightfully cold"],
        "key": "A", "explanation": "Kata 'feel' berpasangan dengan adjective feeling: 'a little anxious' (merasa sedikit cemas/tegang namun percaya diri)."
    },
    {
        "no": 18, "topic": "Grammar: Adjective Phrases & Intentions", "stimulus": "Percakapan Sekolah & Frasa Sifat",
        "stem": "Which of the following contains a well-structured Adjective Phrase describing a tourism destination?",
        "options": ["A. The park is located near the highway.", "B. SMAN 1 Bae visits the museum by bus.", "C. We stayed at a remarkably tranquil, eco-friendly wooden lodge.", "D. Because of heavy rain, tourists cancelled their booking.", "E. Swimming without life jackets is dangerous."],
        "key": "C", "explanation": "'a remarkably [Adv] tranquil [Adj], eco-friendly [Adj] wooden lodge [Noun]' adalah frasa sifat bertingkat yang terstruktur rapi."
    },
    {
        "no": 19, "topic": "Grammar: Adjective Phrases & Intentions", "stimulus": "Percakapan Sekolah & Frasa Sifat",
        "stem": "Dewi: 'We are organizing an excursion to Menjangan Besar next month.' \nBima: 'That sounds fantastic! I ________ join you if my parents permit me.'",
        "options": ["A. definitely intend to", "B. has intended", "C. am not will to", "D. intended yesterday", "E. will have been"],
        "key": "A", "explanation": "'definitely intend to' menunjukkan ungkapan kesungguhan niat (strong intention)."
    },
    {
        "no": 20, "topic": "Grammar: Adjective Phrases & Intentions", "stimulus": "Percakapan Sekolah & Frasa Sifat",
        "stem": "Which expression best conveys deep emotional satisfaction after completing a conservation project?",
        "options": ["A. 'We felt utterly disappointed by the result.'", "B. 'It feels immensely rewarding to witness the coral reef revitalized.'", "C. 'I do not care about the marine sanctuary.'", "D. 'The project made us extremely frustrated and exhausted without any point.'", "E. 'We are intending to quit the club immediately.'"],
        "key": "B", "explanation": "'It feels immensely rewarding...' mengekspresikan perasaan puas dan berharga (stating profound positive feeling)."
    }
]

# 5 Essay Questions for SAS
sas_essay_questions = [
    {
        "no": 1, "topic": "Descriptive Text",
        "context": "Studi Kasus: Promosi Wisata Alam Berkelanjutan di Kabupaten Kudus (e.g., Air Terjun Montel / Wisata Colo Muria)",
        "prompt": "Susunlah 1 (satu) paragraf teks deskriptif otentik (60–80 kata) yang mempromosikan destinasi alam di lereng Gunung Muria. Paragraf wajib memuat: (1) Identification yang jelas, (2) Minimal 3 adjective phrases bernuansa sensorik (visual/auditory), dan (3) Penggunaan Simple Present Tense secara konsisten.",
        "rubric": "Skor 6: Memuat identifikasi jelas, 3+ adjective phrases tepat, tata bahasa Simple Present sempurna.\nSkor 4: Memuat deskripsi baik, ada 2 frasa sifat, sedikit galat tenses minor.\nSkor 2: Deskripsi umum tanpa frasa sensorik memadai.\nSkor 0: Tidak menjawab.",
        "key_model": "Contoh Jawaban Ideal: Nestled serenely along the verdant slopes of Mount Muria, Montel Waterfall in Kudus represents a breathtakingly tranquil mountain sanctuary. The cascade features crystal-clear subterranean water plummeting sixty meters into a cool, rock-strewn pool blanketed by dense tropical fern foliage. As visitors inhale the invigorating pine-scented air, the soothing roar of rushing water creates an enchanting atmosphere that revitalizes the spirit of every weary traveler."
    },
    {
        "no": 2, "topic": "Report Text",
        "context": "Studi Kasus: Laporan Observasi Keanekaragaman Hayati (Burung Kacamata Jawa / Flora Endemik)",
        "prompt": "Berdasarkan prinsip Report Text ilmiah, tuliskan bagian 'General Classification' dan 1 aspek 'Description' (Habitat atau Diet) mengenai satwa endemik lokal Jawa. Jelaskan perbedaan mendasar paragraf yang Anda buat ini dibandingkan teks deskriptif!",
        "rubric": "Skor 6: Klasifikasi ilmiah jelas, deskripsi faktual objektif, penjelasan perbedaan dengan teks deskriptif tepat.\nSkor 4: Klasifikasi memadai, deskripsi baik, penjelasan perbedaan kurang komprehensif.\nSkor 2: Cenderung menyerupai cerita fiksi atau opini pribadi.\nSkor 0: Kosong.",
        "key_model": "Contoh Jawaban Ideal:\n[General Classification] The Javan White-eye (Zosterops flavus) is an endemic passerine bird belonging to the family Zosteropidae, native to coastal mangroves and tropical lowland shrublands of Java and Borneo.\n[Description: Diet] As an omnivorous forager, this species predominantly feeds on nectar, ripe berries, and small soft-bodied insects among mangrove canopies.\n[Perbedaan] Teks report ini mengklasifikasikan spesies secara biologis umum (seluruh populasi), bukan burung peliharaan tertentu dengan nama khusus."
    },
    {
        "no": 3, "topic": "Grammar: Passive Voice",
        "context": "Transformasi Wacana Ilmiah dari Aktif ke Pasif dalam Konteks Pengolahan Sampah Sekolah",
        "prompt": "Ubah wacana aktif berikut menjadi wacana pasif formal berstandar akademik:\n'The students collect organic waste daily. The green club processes the compost in special bio-reactors. Finally, gardeners apply the natural fertilizer to improve school soil fertility.'",
        "rubric": "Skor 6: Ketiga kalimat diubah menjadi passive voice dengan subject-verb agreement tepat (is/are + V3).\nSkor 4: Dua kalimat pasif benar, satu keliru minor.\nSkor 2: Masih dominan kalimat aktif.\nSkor 0: Kosong.",
        "key_model": "Model Jawaban Ideal:\n'Organic waste is collected daily by the students. The compost is processed in special bio-reactors by the green club. Finally, the natural fertilizer is applied by gardeners to improve school soil fertility.'"
    },
    {
        "no": 4, "topic": "Adjective Phrases",
        "context": "Pengembangan Kosakata & Gaya Penulisan Bahasa Inggris",
        "prompt": "Kembangkan kata benda (noun) berikut menjadi Adjective Phrase bertingkat (Determiner + Adverb + Adjective + Noun):\n1. 'heritage building'\n2. 'marine sanctuary'\n3. 'scientific discovery'\nBuatlah masing-masing 1 kalimat utuh yang kontekstual!",
        "rubric": "Skor 6: Ketiga frasa tersusun sempurna sesuai kaidah adv+adj+noun dan digunakan dalam kalimat bermakna.\nSkor 4: Dua frasa dan kalimat tepat.\nSkor 2: Susunan kata terbalik atau tidak membentuk kalimat utuh.\nSkor 0: Kosong.",
        "key_model": "Model Jawaban:\n1. 'an exceptionally majestic heritage building' -> The old town square features an exceptionally majestic heritage building from the colonial era.\n2. 'a remarkably pristine marine sanctuary' -> Karimunjawa remains a remarkably pristine marine sanctuary guarded by local residents.\n3. 'a profoundly revolutionary scientific discovery' -> Harnessing geothermal heat represents a profoundly revolutionary scientific discovery for clean energy."
    },
    {
        "no": 5, "topic": "Intentions & Feelings",
        "context": "Dialog Transaksional & Interpersonal Siswa SMA Negeri 1 Bae",
        "prompt": "Buatlah teks percakapan singkat (4 pertukaran giliran bicara) antara dua orang siswa (Adit dan Ratna) yang sedang merencanakan proyek kampanye literasi lingkungan hidup. Percakapan harus mengekspresikan: (1) Ungkapan niat yang kuat (stating intention) dan (2) Perasaan antusias serta komitmen moral (stating feeling & commitment).",
        "rubric": "Skor 6: Dialog mengalir natural, ungkapan niat dan perasaan tepat secara gramatikal dan kontekstual.\nSkor 4: Memuat ungkapan niat dan rasa, namun tata bahasa ada kekeliruan minor.\nSkor 2: Percakapan terlalu kaku dan tidak memuat target kebahasaan.\nSkor 0: Kosong.",
        "key_model": "Model Dialog Ideal:\nAdit: 'Ratna, I am definitely intending to launch a zero-plastic challenge across all tenth-grade classes next Monday.'\nRatna: 'That is wonderful, Adit! I feel profoundly enthusiastic about collaborating on this environmental initiative.'\nAdit: 'We plan to produce bilingual infographic posters to educate our peers on waste sorting.'\nRatna: 'I feel deeply grateful for this opportunity because protecting our school ecosystem is our shared responsibility.'"
    }
]

# Build sheets
build_sts_sheet(ws_sts, "Kelas X (Fase E)", "Descriptive Text", t1_passage, "Report Text", t2_passage, sts_questions)
build_sas_sheet(ws_sas, "Kelas X (Fase E)", "Descriptive Text, Report Text, Simple Present & Passive Voice, Adjective Phrases & Intentions", sas_pg_questions, sas_essay_questions)
build_gradebook_sheet(ws_gradebook, "Kelas X (Fase E)")

out_path = "C:/Users/LILI/OneDrive - PDK Jateng/Dokumen/Perangkat_SMAN_1_Bae/Kelas_10/Latihan_Asesmen_Bahasa_Inggris_Kelas_10_SMAN_1_Bae.xlsx"
wb.save(out_path)
print("Workbook Kelas 10 generated successfully at:", out_path)
