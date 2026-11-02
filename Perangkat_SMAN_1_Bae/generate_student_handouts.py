# -*- coding: utf-8 -*-
import os
from build_student_docx_lib import (
    create_student_doc, add_student_header, add_heading_1, add_heading_2,
    add_concept_box, add_elements_table, add_annotated_model_text,
    add_grammar_section, add_pronunciation_section
)

# ==============================================================================
# 1. KELAS X GENERATION
# ==============================================================================
def generate_kelas_10():
    doc = create_student_doc()
    add_student_header(doc, "MODUL MATERI BELAJAR KELAS X (FASE E)", "DESCRIPTIVE TEXT & REPORT TEXT")
    
    # PERTEMUAN 1: DESCRIPTIVE TEXT
    add_heading_1(doc, "PERTEMUAN 1: DESCRIPTIVE TEXT (TEKS DESKRIPTIF)")
    add_concept_box(doc, 
        "Descriptive Text adalah jenis teks yang bertujuan untuk menggambarkan ciri-ciri, sifat, bentuk, suasana, dan keunikan suatu objek khusus (seperti tempat wisata tertentu, orang/tokoh inspiratif, atau bangunan bersejarah spesifik) secara rinci dan hidup sehingga pembaca seolah-olah dapat melihat, mendengar, atau merasakannya sendiri.",
        "To describe and reveal a particular person, place, or thing. Fokusnya adalah pada objek partikular/spesifik (bukan kelompok umum), misalnya Candi Prambanan atau Taman Nasional Karimunjawa."
    )
    
    add_elements_table(doc, [
        ("1. Identification (Pengenalan Objek)", 
         "Paragraf pembuka yang memperkenalkan objek spesifik yang akan dideskripsikan, menyebutkan nama objek, lokasi geografisnya, dan alasan mengapa objek tersebut istimewa/terkenal.",
         "Terletak di paragraf pertama. Gunakan pertanyaan pemandu: 'What/Who is being described?' dan 'Where is it located?'"),
        ("2. Description (Uraian Ciri & Suasana)", 
         "Rangkaian paragraf yang menguraikan rincian fisik objek, bagian-bagian penting, keindahan lanskap, kualitas karakter, atau suasana sekitar dengan memanfaatkan panca indra (penglihatan, pendengaran, penciuman).",
         "Terletak di paragraf kedua dan seterusnya. Ditandai oleh banyaknya kata sifat sensorik (sensory adjectives) dan frasa sifat (adjective phrases).")
    ])
    
    p1_annotated = [
        ("[PARAGRAF 1: IDENTIFICATION]\nMemperkenalkan objek spesifik, letak geografis, dan status suaka alam.",
         "Karimunjawa Marine National Park, an archipelago of twenty-seven pristine islands situated approximately 80 kilometers northwest of Jepara in Central Java, represents one of Indonesia’s most ecologically vital maritime sanctuaries. Declared a protected marine reserve, this tranquil haven attracts both local and international researchers who seek to observe sustainable marine ecosystems harmoniously preserved by the local coastal communities."),
        ("[PARAGRAF 2: DESCRIPTION (Underwater & Terrestrial Landscape)]\nMendeskripsikan pemandangan bawah laut, terumbu karang, dan satwa daratan.",
         "The archipelago features breathtaking marine topography characterized by crystalline turquoise waters, vibrant fringing coral reefs, and extensive mangrove ecosystems. Beneath the calm surface, over two hundred species of scleractinian hard corals thrive alongside vibrant schools of clownfish, hawksbill sea turtles, and majestic spotted stingrays. The terrestrial landscape is equally enchanting, blanketed by verdant tropical rain forests that provide a natural habitat for the endemic red-breasted parakeet and white-bellied sea eagles."),
        ("[PARAGRAF 3: DESCRIPTION (Community Conservation & Atmosphere)]\nMendeskripsikan kolam konservasi hiu oleh nelayan dan suasana senja di pantai.",
         "Beyond its magnificent physical scenery, Karimunjawa offers visitors an authentic experience of community-based eco-tourism. Local fishermen on Pulau Menjangan Besar actively manage community-led shark rehabilitation pools, where wounded marine life is nurtured before being safely released into the wild. As dusk approaches, the horizon radiates an amber glow, casting long silhouettes across the white powdered sands of Tanjung Gelam Beach. The seamless integration of biological preservation and hospitable cultural traditions makes Karimunjawa a remarkable national treasure of Central Java.")
    ]
    add_annotated_model_text(doc, "Karimunjawa: The Living Coral Haven of the Java Sea", 
                            "Indonesia Travel & Marine Conservation Bulletin (CEFR Level A2+/B1)", p1_annotated)
                            
    add_grammar_section(doc, [
        {
            "rule_name": "Simple Present Tense (Menyatakan Fakta Permanen & Ciri Objek)",
            "formula": "Kalimat Positif (+): Subject + Verb-1 (s/es) + Object/Complement\nKalimat Nominal (+): Subject + is / am / are + Adjective/Noun/Adverb",
            "explanation": "Teks deskriptif wajib menggunakan Simple Present Tense karena mendeskripsikan fakta fisik dan karakteristik yang masih berlangsung dan nyata hingga saat ini.",
            "examples": [
                "The archipelago features breathtaking marine topography. (Subject singular 'archipelago' + verb with -s 'features')",
                "Over two hundred species of corals thrive beneath the calm surface. (Subject plural 'species' + verb-1 'thrive')",
                "Karimunjawa is a protected marine reserve in Central Java. (Nominal sentence with 'is')"
            ]
        },
        {
            "rule_name": "Adjective Phrases (Frasa Sifat Bertingkat & Bahasa Sensorik)",
            "formula": "Determiner + Adverb (Tingkat) + Adjective (Sifat) + Noun (Kata Benda)",
            "explanation": "Frasa sifat digunakan untuk memberikan gambaran sensorik yang jauh lebih tajam dan kaya daripada hanya menggunakan satu kata sifat sederhana.",
            "examples": [
                "a [Det] + remarkably [Adv] + pristine [Adj] + sanctuary [Noun] = suaka alam yang sangat murni.",
                "an [Det] + ecologically [Adv] + vital [Adj] + reserve [Noun] = suaka yang vital secara ekologis.",
                "crystalline turquoise waters = air berwarna toska jernih berkilau.",
                "white powdered sands = pasir putih halus bagai bubuk."
            ]
        }
    ])
    
    add_pronunciation_section(doc, [
        ("Archipelago", "/ˌɑːkɪˈpeləɡəʊ/", "ar-chi-PEL-a-go", "Kepulauan / gugusan pulau"),
        ("Pristine", "/ˈprɪstiːn/", "PRIS-teen", "Murni, alami, belum rusak"),
        ("Sanctuary", "/ˈsæŋktʃuəri/", "SANGK-chu-er-ee", "Suaka margasatwa / tempat perlindungan"),
        ("Endemic", "/enˈdemɪk/", "en-DEM-ik", "Spesies asli khas daerah tertentu"),
        ("Breathtaking", "/ˈbreθteɪkɪŋ/", "BRETH-tay-king", "Sangat menakjubkan / memukau"),
        ("Turquoise", "/ˈtɜːkwɔɪz/", "TUR-kwoyz", "Warna biru kehijauan cerah (toska)"),
        ("Silhouette", "/ˌsɪluˈet/", "sil-oo-ET", "Bayangan hitam berlatar terang"),
        ("Harmoniously", "/hɑːˈməʊniəsli/", "har-MOH-nee-us-lee", "Secara selaras dan serasi")
    ], [
        "Huruf 'ch' pada kata 'Archipelago' dilafalkan /k/ (bukan /tʃ/), seperti kata 'chemistry' atau 'character'.",
        "Huruf 's' pada kata 'Island' bersifat SILENT (tidak dibaca!). Ucapkan /ˈaɪlənd/ (ai-lend), JANGAN ucapkan 'is-land'.",
        "Bunyi 'th' pada kata 'Breathtaking' dilafalkan tak bersuara /θ/ (lidah diletakkan di antara gigi atas dan bawah).",
        "Perhatikan suku kata dengan huruf kapital (STRESS) untuk memberi penekanan suara yang lebih kuat dan jelas."
    ])
    
    # PERTEMUAN 2: REPORT TEXT
    add_heading_1(doc, "PERTEMUAN 2: REPORT TEXT (TEKS LAPORAN ILMIAH)")
    add_concept_box(doc,
        "Report Text adalah teks faktual ilmiah yang menyajikan informasi apa adanya tentang suatu fenomena alam, biologis, sosial, atau teknologi secara UMUM (General) berdasarkan hasil pengamatan ilmiah, riset, atau klasifikasi sistematis.",
        "To present information about something generally, as it is. Berbeda dengan Descriptive Text yang membahas satu objek spesifik milik pribadi/lokasi tunggal, Report Text membahas sekelompok/golongan fenomena (misal: Badak Jawa secara umum, Hutan Mangrove secara umum, atau Gempa Bumi secara umum)."
    )
    
    add_elements_table(doc, [
        ("1. General Classification (Klasifikasi Umum)",
         "Pernyataan pembuka yang menempatkan subjek bahasan ke dalam taksonomi, kelompok ilmiah, famili biologis, atau kategori umum yang diakui sains.",
         "Terletak di paragraf pertama. Ciri khas: Menyebutkan nama ilmiah (scientific name), definisi taksonomi, dan wilayah sebaran umum."),
        ("2. Description (Uraian Faktual Sistematis)",
         "Serangkaian paragraf yang menguraikan aspek-aspek umum subjek secara ilmiah, meliputi: morfologi fisik, habitat alami, pola perilaku, pola makan (diet), serta ancaman kepunahan dan upaya konservasi.",
         "Terletak di paragraf kedua dan seterusnya. Ciri khas: Bersifat objektif, bebas dari opini pribadi/perasaan emosional, serta menggunakan kata-kata teknis ilmiah (technical terms).")
    ])
    
    p2_annotated = [
        ("[PARAGRAF 1: GENERAL CLASSIFICATION]\nKlasifikasi taksonomi famili Rhinocerotidae dan habitat sebaran global.",
         "The Javan Rhinoceros (Rhinoceros sondaicus), also known as the lesser one-horned rhinoceros, is a critically endangered member of the family Rhinocerotidae and one of five extant rhinoceros species worldwide. Historically distributed across the lush tropical forests of Southeast Asia, Java, and Sumatra, the entire global population is now strictly confined to Ujung Kulon National Park at the westernmost tip of Java Island, Indonesia."),
        ("[PARAGRAF 2: DESCRIPTION (Morphology & Anatomy)]\nCiri fisik umum: kulit berlipat mirip zirah, cula tunggal, bobot tubuh, kelincahan.",
         "Morphologically, the Javan rhino possesses a dusky grey, heavily folded skin that resembles armor plating, which provides natural protection against dense tropical vegetation and thorny thickets. Unlike African rhinoceros species which develop two prominent horns, the Javan rhino typically bears a single small horn rarely exceeding twenty-five centimeters in length, often found only in males. Adults can weigh between 900 and 2,300 kilograms, yet despite their massive physical stature, they are exceptionally agile movers capable of navigating dense jungle mudflats and steep slopes."),
        ("[PARAGRAF 3: DESCRIPTION (Ecology, Behavior, & Conservation Threat)]\nPola makan herbivora, perilaku berkubang lumpur, dan upaya konservasi anti-perburuan liar.",
         "Ecologically, Rhinoceros sondaicus is a solitary, herbivorous browser that feeds on a diverse variety of shoots, fallen fruits, twigs, and tender young leaves. They rely heavily on wallowing in natural mud pools to regulate body temperature and protect their sensitive dermal layers from parasites and insect bites. Due to severe habitat destruction and historical poaching for illegal wildlife trade, fewer than eighty individuals remain in the wild. Consequently, intensive satellite monitoring, biological habitat restoration, and anti-poaching patrols are continuously implemented to prevent the irreversible extinction of this prehistoric mammal.")
    ]
    add_annotated_model_text(doc, "The Javan Rhinoceros: Ecology and Survival Frontiers",
                            "World Wildlife & Tropical Ecology Research Journal (CEFR Level B1)", p2_annotated)
                            
    add_grammar_section(doc, [
        {
            "rule_name": "Passive Voice pada Penulisan Ilmiah (Menjaga Objektivitas Sains)",
            "formula": "Subject + is / am / are + Verb-3 (Past Participle) + (by Agent)",
            "explanation": "Dalam teks sains, objek yang diamati atau proses ilmiah jauh lebih penting daripada orang yang mengamatinya. Oleh karena itu, Passive Voice sangat dominan digunakan.",
            "examples": [
                "The entire population is strictly confined to Ujung Kulon National Park. (Objek 'population' singular + is + V3 'confined')",
                "Anti-poaching patrols are continuously implemented by rangers. (Subjek plural 'patrols' + are + V3 'implemented')",
                "Mud wallowing is required to regulate body temperature."
            ]
        },
        {
            "rule_name": "Timeless Generic Present & Technical Vocabulary",
            "formula": "Generic Noun Plural/Uncountable + Base Verb (Verb-1)",
            "explanation": "Report Text tidak menceritakan peristiwa kemarin atau besok, melainkan kebiasaan alamiah satwa/fenomena yang berlaku sepanjang masa (timeless truth).",
            "examples": [
                "Rhinos feed on shoots, fallen fruits, and tender leaves. (Fakta makanan permanen)",
                "They rely heavily on wallowing in natural mud pools.",
                "Istilah teknis biologi: 'herbivorous browser', 'dermal layers', 'extant species', 'Rhinocerotidae'."
            ]
        }
    ])
    
    add_pronunciation_section(doc, [
        ("Rhinoceros", "/raɪˈnɒsərəs/", "rye-NOSS-er-us", "Badak"),
        ("Herbivorous", "/hɜːˈbɪvərəs/", "her-BIV-or-us", "Pemakan tumbuhan"),
        ("Extant", "/ˈekstənt/", "EK-stuhnt", "Masih hidup / ada di bumi saat ini (lawan kata extinct)"),
        ("Solitary", "/ˈsɒlətri/", "SOL-i-tree", "Menyendiri / hidup tidak berkelompok"),
        ("Dermal", "/ˈdɜːml/", "DUR-muhl", "Berhubungan dengan lapisan kulit"),
        ("Wallow", "/ˈwɒləʊ/", "WOL-oh", "Berkubang di lumpur/air"),
        ("Poaching", "/ˈpəʊtʃɪŋ/", "POH-ching", "Perburuan satwa liar secara ilegal"),
        ("Critically", "/ˈkrɪtɪkli/", "KRIT-ik-lee", "Secara kritis / genting")
    ], [
        "Huruf 'rh' pada 'Rhinoceros' cukup dilafalkan sebagai bunyi /r/ (huruf 'h' tidak bersuara): /raɪˈnɒsərəs/.",
        "Pelafalan akhiran lampau/passive '-ed':",
        "  1. Dilafalkan /t/ setelah bunyi tak bersuara (voiceless): e.g., 'developed' -> /dɪˈveləpt/.",
        "  2. Dilafalkan /d/ setelah bunyi bersuara (voiced): e.g., 'confined' -> /kənˈfaɪnd/.",
        "  3. Dilafalkan /ɪd/ hanya jika kata kerja berakhiran bunyi /t/ atau /d/: e.g., 'protected' -> /prəˈtektɪd/.",
        "Kata 'Herbivorous' memiliki penekanan suku kata kedua: her-BIV-or-us (bukan her-bi-VOR-us)."
    ])
    
    # Save Kelas 10
    out_docx = "C:/Users/LILI/OneDrive - PDK Jateng/Dokumen/Perangkat_SMAN_1_Bae/Kelas_10/Ringkasan_Materi_2_Pertemuan_Kelas_10.docx"
    doc.save(out_docx)
    print("New Student Handout DOCX Kelas 10 generated successfully!")

# ==============================================================================
# 2. KELAS XI GENERATION
# ==============================================================================
def generate_kelas_11():
    doc = create_student_doc()
    add_student_header(doc, "MODUL MATERI BELAJAR KELAS XI (FASE F)", "NARRATIVE TEXT & ANALYTICAL EXPOSITION")
    
    # PERTEMUAN 1: NARRATIVE TEXT
    add_heading_1(doc, "PERTEMUAN 1: NARRATIVE TEXT (TEKS NARATIF & KISAH MORAL)")
    add_concept_box(doc,
        "Narrative Text adalah teks cerita fiksi atau kisah nyata masa lampau yang menyajikan rangkaian peristiwa kronologis dengan masalah/krisis (Complication) yang memuncak dan akhirnya menemukan jalan keluar (Resolution) untuk menghibur dan menyampaikan hikmah moral.",
        "To entertain, amuse, and deal with an actual or vicarious experience in different ways. Cerita naratif mengajak pembaca menyelami emosi dan pergulatan etika tokoh utama."
    )
    
    add_elements_table(doc, [
        ("1. Orientation (Pengenalan)", 
         "Memperkenalkan tokoh utama (Who), latar tempat (Where), latar waktu (When), dan watak dasar tokoh.", 
         "Terletak di awal cerita (Paragraf 1). Menjawab: Siapa tokohnya dan di mana kisahnya bermula?"),
        ("2. Complication (Krisis / Masalah)", 
         "Rangkaian peristiwa ketika tokoh menghadapi masalah moral, godaan, atau benturan dengan alam/masyarakat hingga mencapai puncak klimaks.", 
         "Terletak di bagian tengah cerita. Ditandai munculnya konflik, ketegangan, dan pilihan sulit tokoh."),
        ("3. Resolution (Penyelesaian)", 
         "Bagian akhir ketika masalah terselesaikan, baik berakhir bahagia (happy ending), tragis (sad ending), atau perenungan mendalam.", 
         "Terletak di bagian akhir. Menjelaskan nasib akhir tokoh dan kondisi setelah krisis."),
        ("4. Coda (Pesan Moral)", 
         "Pelajaran hidup atau pesan moral yang dapat dipetik oleh pembaca dari perjalanan hidup sang tokoh.", 
         "Bisa tersirat dalam perubahan sikap tokoh atau tersurat di penutup cerita.")
    ])
    
    p1_annotated = [
        ("[PARAGRAF 1: ORIENTATION]\nPengenalan tokoh Ki Ageng Semito di lereng Gunung Muria, keahlian jam mekanik, dan watak angkuh.",
         "In the mist-shrouded foothills of Mount Muria during the late eighteenth century, there lived a master clockmaker named Ki Ageng Semito. Renowned across Central Java for constructing intricate bronze timepieces that ticked with astonishing precision, he grew increasingly arrogant. He firmly believed that human intellect could subjugate time itself, treating every second as a commodity and refusing hospitality to anyone who could not afford his exorbitant services."),
        ("[PARAGRAF 2: COMPLICATION]\nKonflik moral: menolak pengembara lansia dan anak kecil yang kedinginan di tengah badai.",
         "One stormy evening during the height of the monsoon, while Ki Ageng was putting the final touches to his masterwork—a mechanical clock adorned with golden gears and automated puppets—a frail, soaked traveler knocked upon his oaken door, begging for shelter and dry embers to warm her trembling grandchild. Ki Ageng furrowed his brow and retorted harshly, 'My time is measured in gold, not wasted on penniless wanderers!' He slammed the door, oblivious to the ominous lightning that fractured the twilight sky."),
        ("[PARAGRAF 3: RESOLUTION & CODA]\nBanjir mata air bawah tanah, jam berputar balik, dan pertobatan Ki Ageng merengkuh kerendahan hati.",
         "Moments later, a deafening tremor shook the workshop. A subterranean spring burst through the stone floor, directly beneath his golden mechanism. Terrified, Ki Ageng watched as the pristine waters rapidly engulfed his coveted tools and blueprints. Strangely, as the water submerged the great clock, its hands began spinning wildly in reverse. In that terrifying moment of realization, Ki Ageng understood that all his accumulated wealth and mechanical mastery were powerless against the unyielding forces of nature and moral compassion. Escaping empty-handed into the pouring rain, he found the weeping child and shared his dry cloak, weeping tears of genuine humility.")
    ]
    add_annotated_model_text(doc, "The Clockmaker’s Spring: A Parable of Hubris and Humility",
                            "Contemporary Indonesian Literary Anthology (CEFR Level B1/B1+)", p1_annotated)
                            
    add_grammar_section(doc, [
        {
            "rule_name": "Past Tenses Sequence (Alur Waktu Peristiwa Lampau)",
            "formula": "1. Simple Past (Aksi berurutan): Subject + Verb-2\n2. Past Continuous (Aksi latar yang sedang berlangsung): Subject + was/were + Verb-ing\n3. Past Perfect (Aksi yang telah selesai sebelum peristiwa lampau lain): Subject + had + Verb-3",
            "explanation": "Narasi menggunakan perpaduan Past Tenses untuk menciptakan alur waktu dramatik yang bertingkat.",
            "examples": [
                "Interrupted Past Action: While he was putting the final touches (sedang bekerja), a traveler knocked (aksi tiba-tiba).",
                "Past Perfect: He had accumulated immense wealth before the tragedy struck his workshop.",
                "Action Verbs: slammed, retorted, engulfed, fractured, submerged."
            ]
        },
        {
            "rule_name": "Direct and Indirect Speech (Kalimat Langsung & Tak Langsung)",
            "formula": "Direct: Speaker + said, \"[Present/Past Tense quote]\"\nIndirect: Speaker + said that + [Shifted Past Tense]",
            "explanation": "Dialog langsung mempertegas emosi tokoh, sedangkan dialog tak langsung mempercepat jalannya cerita narasi.",
            "examples": [
                "Direct Speech: Ki Ageng retorted harshly, \"My time is measured in gold!\"",
                "Indirect Speech: Ki Ageng retorted harshly that his time was measured in gold.",
                "Direct: The traveler pleaded, \"Please give us shelter.\"",
                "Indirect: The traveler pleaded with him to give them shelter."
            ]
        }
    ])
    
    add_pronunciation_section(doc, [
        ("Craftsman", "/ˈkrɑːftsmən/", "KRAHFT-sman", "Pengrajin / ahli kriya"),
        ("Subterranean", "/ˌsʌbtəˈreɪniən/", "sub-tuh-RAY-nee-uhn", "Bawah tanah / di dalam perut bumi"),
        ("Exorbitant", "/ɪɡˈzɔːbɪtənt/", "ig-ZOR-bi-tuhnt", "Sangat mahal / melampaui batas wajar"),
        ("Subjugate", "/ˈsʌbdʒʊɡeɪt/", "SUB-joo-gayt", "Menundukkan / menguasai secara paksa"),
        ("Compassion", "/kəmˈpæʃn/", "kuhm-PASH-uhn", "Welas asih / rasa belas kasih"),
        ("Oblivious", "/əˈblɪviəs/", "uh-BLIV-ee-uhs", "Lengah / tidak menyadari bahaya"),
        ("Parable", "/ˈpærəbl/", "PAR-uh-bl", "Kisah perumpamaan moral"),
        ("Humility", "/hjuːˈmɪləti/", "hyoo-MIL-i-tee", "Kerendahan hati")
    ], [
        "Kata kerja lampau tak beraturan (Irregular Past Verbs) yang harus dihafal pelafalannya:",
        "  - Shook /ʃʊk/ (vokal pendek seperti 'book', bukan 'shoot').",
        "  - Burst /bɜːst/ (vokal panjang tengah, bentuk V1, V2, dan V3-nya sama persis!).",
        "  - Wept /wept/ (bentuk lampau dari 'weep').",
        "Intonasi Bercerita (Narrative Delivery): Naikkan sedikit nada bicara saat mencapai kalimat ketegangan (Suspense), dan turunkan nada suara saat menyampaikan resolusi dan pesan moral penutup."
    ])
    
    # PERTEMUAN 2: ANALYTICAL EXPOSITION
    add_heading_1(doc, "PERTEMUAN 2: ANALYTICAL EXPOSITION (TEKS EKSPOSISI ANALITIS)")
    add_concept_box(doc,
        "Analytical Exposition adalah teks argumentatif formal yang bertujuan meyakinkan pembaca (to persuade the audience) bahwa suatu ide, fenomena, atau topik yang diangkat adalah benar, penting, dan layak disikapi secara kritis melalui serangkaian argumen logis berbasis data.",
        "To persuade the reader that the idea is an important matter. Ciri khas mutlaknya: Diakhiri dengan Reiteration (penegasan kembali tesis), BUKAN rekomendasi saran aksi (karena saran aksi adalah ranah Hortatory)."
    )
    
    add_elements_table(doc, [
        ("1. Thesis (Pernyataan Pendirian)",
         "Paragraf pembuka yang memperkenalkan topik krusial dan secara tegas menyatakan sudut pandang atau pendirian penulis (author's stance/claim).",
         "Terletak di akhir Paragraf 1. Ciri khas: Memuat kalimat tesis utama: 'Therefore, it is indisputable that... is an absolute imperative.'"),
        ("2. Series of Arguments (Argumen & Elaborasi)",
         "Rangkaian paragraf pendukung. Setiap paragraf memuat: Point (klaim ide pokok argumen) dan Elaboration (bukti data, alasan logis, fakta ilmiah).",
         "Biasanya diawali oleh konjungsi urutan (Firstly, Secondly, Furthermore). Menjawab: 'Mengapa klaim tesis tersebut benar?'"),
        ("3. Reiteration (Penegasan Kembali Tesis)",
         "Paragraf penutup yang menegaskan dan merangkum kembali posisi tesis utama penulis menggunakan pilihan kata (diksi) yang berbeda tanpa mengubah esensi makna.",
         "Terletak di paragraf terakhir. Diawali frasa simpulan: 'In conclusion,...', 'To sum up,...', 'From the arguments above,...'.")
    ])
    
    p2_annotated = [
        ("[PARAGRAF 1: THESIS]\nMengangkat isu AI generatif dan merumuskan tesis pentingnya literasi digital di SMA.",
         "The rapid proliferation of generative artificial intelligence and synthetic media has fundamentally transformed how young individuals consume, synthesize, and disseminate information. While these technological breakthroughs offer unprecedented learning conveniences, they simultaneously expose students to sophisticated misinformation, automated plagiarism, and deepfake content. Therefore, it is indisputable that integrating critical digital literacy into the high school curriculum is no longer merely optional, but an absolute educational imperative."),
        ("[PARAGRAF 2: ARGUMENT 1 (Point & Elaboration)]\nPoint: Evaluasi kebenaran sumber digital. Elaboration: Mengatasi machine hallucinations dengan metode cek-silang.",
         "Firstly, critical digital literacy equips adolescents with essential analytical scrutiny to evaluate the credibility of digital sources. In an era where algorithmic feeds prioritize sensationalism over factual veracity, uncritical consumers frequently mistake machine-generated hallucinations for verified truths. By understanding cross-referencing methodologies and source attribution, students can distinguish authentic academic consensus from commercially driven digital falsehoods."),
        ("[PARAGRAF 3: ARGUMENT 2 (Point & Elaboration)]\nPoint: Integritas akademik & nalar manusia. Elaboration: Membedakan AI sebagai mitra riset vs plagiasi otomatis.",
         "Secondly, institutionalizing digital literacy fosters intellectual integrity and ethical awareness in the contemporary classroom. When learners understand the profound differences between using AI as an exploratory brainstorming partner versus engaging in automated academic dishonesty, they develop authentic academic resilience. They realize that authentic problem-solving and critical cognition cannot be outsourced to a predictive neural network."),
        ("[PARAGRAF 4: REITERATION]\nMenegaskan kembali posisi tesis: Pendidikan harus mencetak navigator digital yang cerdas dan bertanggung jawab.",
         "In conclusion, secondary educational institutions must not perceive technological advancements in isolation from ethical reasoning. Enhancing students' capacity to critically deconstruct digital media ensures that young citizens emerge as discerning, responsible, and empowered digital navigators in an increasingly automated global society.")
    ]
    add_annotated_model_text(doc, "Why Critical Digital Literacy Must Be Prioritized in Secondary Education",
                            "Global Educational Technology Review (CEFR Level B1+/B2)", p2_annotated)
                            
    add_grammar_section(doc, [
        {
            "rule_name": "Causal & Contrastive Conjunctions (Penanda Hubungan Kausalitas)",
            "formula": "Cause / Premise + [therefore / consequently / as a result] + Effect / Conclusion",
            "explanation": "Eksposisi analitis mengandalkan konjungsi logis untuk menghubungkan masalah dengan urgensi tindakan.",
            "examples": [
                "Algorithmic feeds prioritize sensationalism; therefore, critical digital literacy is an absolute imperative.",
                "Students outsourced their critical cognition; consequently, their independent problem-solving skills deteriorated.",
                "Contrast Conjunction: While AI offers great conveniences, it simultaneously creates new ethical risks."
            ]
        },
        {
            "rule_name": "Modal Adverbs of Certainty & Mental Verbs",
            "formula": "Subject + [indisputably / evidently / certainly] + Verb  |  Subject + Mental Verb (believe/conclude/assess)",
            "explanation": "Modal adverbs memperkuat derajat kepastian argumen, sedangkan mental verbs mengekspresikan proses pemikiran analitis.",
            "examples": [
                "It is indisputably proven that digital literacy strengthens student integrity.",
                "Scholars conclude that authentic problem-solving requires human critical judgment.",
                "Learners recognize the profound difference between AI assistance and plagiarism."
            ]
        }
    ])
    
    add_pronunciation_section(doc, [
        ("Proliferation", "/prəˌlɪfəˈreɪʃn/", "pro-lif-er-AY-shun", "Perkembangan pesat / lonjakan tajam"),
        ("Veracity", "/vəˈræsəti/", "vuh-RASS-i-tee", "Kebenaran fakta / keakuratan data"),
        ("Hallucination", "/həˌluːsɪˈneɪʃn/", "huh-loo-si-NAY-shun", "Output keliru AI yang dikira fakta"),
        ("Imperative", "/ɪmˈperətɪv/", "im-PEHR-uh-tiv", "Keharusan mutlak / hal yang wajib"),
        ("Indisputable", "/ˌɪndɪˈspjuːtəbl/", "in-di-SPYOO-tuh-bl", "Pasti benar / tak terbantahkan"),
        ("Scrutiny", "/ˈskruːtəni/", "SKROO-ti-nee", "Pemeriksaan cermat dan teliti"),
        ("Resilience", "/rɪˈzɪliəns/", "ri-ZIL-yuhns", "Ketahanan mental dan daya juang"),
        ("Discerning", "/dɪˈsɜːnɪŋ/", "dih-SURN-ing", "Kritis dan bijak membedakan hal benar")
    ], [
        "Aturan Penekanan Kata Berakhiran '-tion' dan '-ity':",
        "  - Kata yang berakhiran '-tion' (e.g., pro-lif-er-AY-tion, hal-lu-ci-NA-tion) selalu mendapat penekanan suara (stress) tepat pada SUKU KATA SEBELUMNYA.",
        "  - Kata yang berakhiran '-ity' (e.g., ve-RASS-i-ty, hu-MIL-i-ty) selalu mendapat stress pada suku kata KETIGA DARI BELAKANG.",
        "Pelafalan bunyi vokal /ɜː/ pada 'Discerning': Ucapkan seperti vokal pada kata 'bird' atau 'learn'."
    ])
    
    # Save Kelas 11
    out_docx = "C:/Users/LILI/OneDrive - PDK Jateng/Dokumen/Perangkat_SMAN_1_Bae/Kelas_11/Ringkasan_Materi_2_Pertemuan_Kelas_11.docx"
    doc.save(out_docx)
    print("New Student Handout DOCX Kelas 11 generated successfully!")

# ==============================================================================
# 3. KELAS XII GENERATION
# ==============================================================================
def generate_kelas_12():
    doc = create_student_doc()
    add_student_header(doc, "MODUL MATERI BELAJAR KELAS XII (FASE F)", "ADVANCED NARRATIVE & DISCUSSION TEXT")
    
    # PERTEMUAN 1: ADVANCED NARRATIVE
    add_heading_1(doc, "PERTEMUAN 1: ADVANCED NARRATIVE TEXT (ALEGORI & SASTRA TINGKAT LANJUT)")
    add_concept_box(doc,
        "Advanced Narrative Text adalah karya fiksi sastra tingkat lanjut yang mengeksplorasi kedalaman psikologis karakter, dilema moral yang ambigu, serta menggunakan bahasa simbolik (allegory & symbolism) untuk mengkritik realitas modern atau menggugah perenungan batin.",
        "To explore existential, philosophical, and moral dilemmas through multi-layered narration. Tidak hanya mengisahkan peristiwa lahiriah, tetapi mengungkap gejolak batin manusia melawan perubahan zaman."
    )
    
    add_elements_table(doc, [
        ("1. Atmospheric Orientation (Latar Batin & Simbolik)",
         "Membangun suasana emosional yang kuat sejak kalimat pertama melalui citraan sensorik arsitektur, aroma, dan nuansa psikologis tokoh.",
         "Terletak di Paragraf 1. Ciri khas: Penggunaan metafora, simile, dan latar sejarah yang kaya simbol."),
        ("2. Multi-layered Complication (Benturan Paradigma)",
         "Konflik tidak hanya berupa pertengkaran fisik, melainkan benturan ideologi/paradigma (misalnya efisiensi teknologi dingin versus ketulusan karya tangan manusia).",
         "Terletak di Paragraf 2. Ciri khas: Karakter antagonis/penantang tidak selalu jahat, melainkan mewakili pola pikir dunia modern."),
        ("3. Epiphany / Turning Point (Momen Pencerahan Batin)",
         "Momen hening ketika tokoh menyadari hakikat kebenaran yang lebih dalam melalui dialog atau benda pusaka.",
         "Terletak di Paragraf 3. Ciri khas: Perubahan sudut pandang drastis dari dangkal menjadi mendalam."),
        ("4. Reflective Open Resolution (Resolusi Kontemplatif)",
         "Akhir cerita yang memberi ruang perenungan bagi pembaca tanpa menyederhanakan masalah menjadi hitam-putih.",
         "Terletak di Paragraf 4. Ciri khas: Menekankan nilai kemanusiaan abadi yang melampaui batas generasi.")
    ])
    
    p1_annotated = [
        ("[PARAGRAF 1: ATMOSPHERIC ORIENTATION]\nMembangun suasana Kota Lama Semarang; dedikasi Master Sastro terhadap naskah fisik di tengah dunia digital.",
         "Deep within the labyrinthine alleyways of Semarang’s Old Town, flanked by Dutch colonial archways whose plaster flaked like desiccated parchment, sat Master Sastro. For over six decades, he had practiced the vanishing art of hand-binding antique manuscripts, stitching fragile leather covers with waxed hemp thread. While the surrounding world surrendered itself to the dazzling ephemeral glow of digital screens and cloud servers, Sastro's sanctuary smelled persistently of aged cedar, beeswax, and crushed clove oil. To him, every tactile page possessed an indelible soul that binary code could never replicate."),
        ("[PARAGRAF 2: COMPLICATION (Philosophical Collision)]\nBenturan teknologi vs seni tangan: Nadia menawarkan digitalisasi 12 menit dengan pemindai holografik.",
         "One sweltering afternoon, a young software developer named Nadia entered his workshop carrying an ultrathin holographic scanner. Tasked by a municipal archive consortium to digitize historical artifacts, she viewed Sastro’s ancient volumes merely as unindexed data sets awaiting conversion into high-density neural vectors. 'Bapak,' Nadia remarked politely, displaying her sleek handheld scanner, 'my machine can preserve your entire life's work in less than twelve minutes. If you allow me, physical decay will never threaten these texts again.'"),
        ("[PARAGRAF 3: EPIPHANY (Momen Pencerahan Batin)]\nMaster Sastro memperdengarkan naskah kuno ke telinga Nadia: jiwa karya melebihi kalkulasi piksel.",
         "Sastro paused, setting down his polished bone folder. He picked up a sun-bleached Javanese chronicle transcribed in golden botanical ink and gently held it near her ear. 'Listen closely, young lady,' he whispered gently. 'A machine can capture the shape of these letters with ruthless optical precision, but can it register the trembling hand of the poet who composed these verses by candlelight during the famine of 1845? Can your server transmit the scent of rain trapped between these dried leaves?'"),
        ("[PARAGRAF 4: REFLECTIVE RESOLUTION]\nNadia menyadari bahwa warisan budaya adalah transmisi kerapuhan dan cinta manusia antar-generasi.",
         "Nadia froze, the hum of her cooling fan suddenly sounding intrusive amidst the solemn silence of the workshop. For the first time, she looked beyond the optical surface of the document, recognizing that cultural preservation is not merely an algorithmic calculation of pixels, but a sacred transmission of human vulnerability across generations.")
    ]
    add_annotated_model_text(doc, "The Memory Weaver of Kota Lama",
                            "Asian Writers Review & Literary Archive (CEFR Level B2)", p1_annotated)
                            
    add_grammar_section(doc, [
        {
            "rule_name": "Participle Clauses (Frasa Partisipatif untuk Memperkaya Gaya Bahasa)",
            "formula": "1. Present Participle (V-ing): Menyatakan dua aksi simultan oleh subjek yang sama.\n2. Past Participle (V-3): Menyatakan latar belakang pasif/kondisi subjek.",
            "explanation": "Participle clauses menggantikan klausa anak yang panjang sehingga tulisan narasi terasa lebih dinamis, anggun, dan berkelas sastra.",
            "examples": [
                "Tasked by a municipal archive consortium, Nadia viewed the books merely as data sets. (Past Participle clause)",
                "Stitching fragile leather covers with waxed hemp thread, Sastro preserved cultural history. (Present Participle)",
                "Gazing at the sun-bleached manuscript, she recognized the sacred human heritage."
            ]
        },
        {
            "rule_name": "Cleft Sentences & Literary Inversion (Penekanan Khusus Sastra)",
            "formula": "Cleft: It is / was + [Elemen yang Ditekankan] + that / who + [Sisa Kalimat]\nInversion: Negative Adverb (Rarely/Never/Seldom) + Auxiliary + Subject + Verb",
            "explanation": "Struktur ini digunakan penulis untuk memusatkan perhatian pembaca pada nilai filosofis paling mendasar.",
            "examples": [
                "Cleft Sentence: It was not the optical scanner, but the poet's sincere devotion that breathed life into the text.",
                "Literary Inversion: Rarely had Nadia witnessed such quiet, unwavering conviction in manual craftsmanship."
            ]
        }
    ])
    
    add_pronunciation_section(doc, [
        ("Labyrinthine", "/ˌlæbəˈrɪnθaɪn/", "lab-uh-RIN-thyne", "Berliku-liku rumit seperti labirin"),
        ("Ephemeral", "/ɪˈfemərəl/", "ih-FEM-er-uhl", "Hanya berlangsung sekejap / fana"),
        ("Desiccated", "/ˈdesɪkeɪtɪd/", "DES-i-kay-tid", "Kering kerontang / layu kehabisan cairan"),
        ("Indelible", "/ɪnˈdeləbl/", "in-DEL-uh-bl", "Abadi / tidak dapat terhapus oleh waktu"),
        ("Vulnerability", "/ˌvʌlnərəˈbɪləti/", "vul-ner-uh-BIL-i-tee", "Kerapuhan nurani & kepekaan kemanusiaan"),
        ("Consortium", "/kənˈsɔːtiəm/", "kuhn-SOR-tee-uhm", "Konsorsium / asosiasi gabungan lembaga"),
        ("Transcribe", "/trænˈskraɪb/", "tran-SKRYBE", "Menyalin tulisan tangan secara cermat"),
        ("Resonance", "/ˈrezənəns/", "REZ-uh-nuhns", "Gema rasa / getaran makna yang mendalam")
    ], [
        "Artikulasi Gugus Konsonan Kompleks (Consonant Clusters):",
        "  - 'Transcribe' /trænˈskraɪb/: Latih transisi cepat dari bunyi /n/ ke /skr/ tanpa menyisipkan vokal sisipan.",
        "Vokal Lemah Schwa /ə/:",
        "  - Kata 'Ephemeral' dilafalkan /ɪˈfemərəl/, bukan /e-fe-me-ral/. Suku kata terakhir menggunakan bunyi vokal netral /əl/.",
        "Penekanan Suku Kata pada Kata Majemuk:",
        "  - 'Vul-ner-a-BIL-i-ty': Tekan suku kata 'BIL' dengan nada lebih tinggi dan volume sedikit lebih kuat."
    ])
    
    # PERTEMUAN 2: DISCUSSION TEXT
    add_heading_1(doc, "PERTEMUAN 2: DISCUSSION TEXT (TEKS DISKUSI WACANA BERIMBANG)")
    add_concept_box(doc,
        "Discussion Text adalah teks argumentatif berimbang yang mengkaji suatu isu kontroversial atau dilematis dari setidaknya dua sudut pandang yang berlawanan (Arguments For / Pro versus Arguments Against / Contra) sebelum menyajikan kesimpulan sintesis netral.",
        "To present differing viewpoints on an issue before formulating a balanced conclusion. Penulis bertindak sebagai penengah yang objektif dan adil (neutral arbiter), bukan pembuat opini sepihak."
    )
    
    add_elements_table(doc, [
        ("1. Statement of Issue (Rumusan Isu)",
         "Paragraf pembuka yang memaparkan isu kontroversial yang sedang hangat diperdebatkan masyarakat dunia beserta latar belakangnya.",
         "Terletak di Paragraf 1. Ciri khas: Memperkenalkan dua kubu yang berdebat (misal: 'optimisme teknologi vs kecemasan sosial ekonomi')."),
        ("2. Arguments For / Pro (Sudut Pandang Pendukung)",
         "Paragraf yang menjabarkan poin-poin keuntungan, potensi positif, dan argumen pihak yang mendukung isu.",
         "Diawali penanda pihak pendukung: 'Proponents contend that...', 'Advocates point out that...', disertai bukti peningkatan produktivitas."),
        ("3. Arguments Against / Contra (Sudut Pandang Pengkritik)",
         "Paragraf yang membedah kelemahan, potensi risiko bahaya, atau kerugian bagi masyarakat/pekerja.",
         "Diawali konjungsi kontras tajam: 'Conversely,...', 'On the other hand,...', 'Detractors argue that...'."),
        ("4. Balanced Conclusion (Kesimpulan Sintesis Berimbang)",
         "Paragraf penutup yang merumuskan solusi jalan tengah (win-win solution), rekomendasi regulasi, atau kebijakan kompromi yang adil.",
         "Terletak di paragraf terakhir. Tidak memihak secara ekstrem, melainkan menawarkan tata kelola etis terpadu (governance framework).")
    ])
    
    p2_annotated = [
        ("[PARAGRAF 1: STATEMENT OF ISSUE]\nPerdebatan global: Otomatisasi tempat kerja sebagai emansipasi talenta vs peminggiran sosial ekonomi.",
         "The relentless integration of autonomous algorithmic systems and generative artificial intelligence into white-collar professions has ignited fierce international debate. Across legal, financial, and creative sectors, cognitive tasks once deemed exclusively human are now executed with breathtaking speed by neural networks. Consequently, economists, policymakers, and labor leaders find themselves sharply divided over whether automation represents an unprecedented emancipation of human talent or a perilous catalyst for widespread socioeconomic disenfranchisement."),
        ("[PARAGRAF 2: ARGUMENTS FOR (Perspektif Pendukung)]\nLonjakan produktivitas, pembebasan dari tugas monoton, dan terciptanya profesi baru (prompt engineering, auditing).",
         "Proponents of rapid workplace automation contend that cognitive AI elevates human productivity to historic heights. By delegating repetitive computational analysis, data verification, and drafting procedures to automated engines, professionals can redirect their cognitive energy toward high-level strategic reasoning, ethical diplomacy, and interpersonal empathy. Furthermore, business analysts argue that augmented enterprises generate entirely new technological ecosystems, creating millions of specialized roles in algorithmic auditing, prompt engineering, and cyber-security oversight that previously did not exist."),
        ("[PARAGRAF 3: ARGUMENTS AGAINST (Perspektif Pengkritik)]\nPengikisan jalur karir magang pemula, risiko bias algoritma, dan hilangnya akuntabilitas manusia.",
         "Conversely, detractors and labor unions present grave concerns regarding structural displacement and the erosion of entry-level career pathways. When junior analyst and associate tasks are swallowed by software subscriptions, the traditional apprenticeship model through which young graduates hone professional judgment collapses. Furthermore, critics emphasize that automated systems frequently perpetuate historical biases embedded within training datasets, resulting in systemic discrimination in recruitment and credit scoring without transparent human accountability."),
        ("[PARAGRAF 4: BALANCED CONCLUSION]\nSintesis jalan tengah: Tata kelola multi-pihak, pelatihan ulang pekerja, dan pengawasan manusia atas mesin.",
         "In conclusion, neither blind technophilia nor reflexive resistance provides an adequate response to the fourth industrial revolution. A balanced strategy requires robust multi-stakeholder governance: corporations must reinvest productivity gains into lifelong employee reskilling, while governments formulate strict labor protections that mandate human oversight in consequential decisions. The ultimate objective must be human-centered augmentation rather than unbridled worker replacement.")
    ]
    add_annotated_model_text(doc, "The Autonomous Workplace: Collaborative Evolution or Mass Disenfranchisement?",
                            "World Economic Forum Deliberation Report (CEFR Level B2)", p2_annotated)
                            
    add_grammar_section(doc, [
        {
            "rule_name": "Contrastive Discourse Markers (Penanda Transisi Kontras Wacana)",
            "formula": "Kalimat Argumen Pro. [Conversely, / On the other hand, / Nevertheless,] + Kalimat Argumen Kontra.",
            "explanation": "Penanda wacana ini sangat krusial untuk memandu pembaca saat berpindah sudut pandang secara adil dan halus.",
            "examples": [
                "Proponents celebrate economic efficiency; conversely, labor unions raise alarms over structural unemployment.",
                "While AI boosts computational speed, it cannot replicate nuanced human emotional empathy.",
                "On the one hand, new industries emerge; on the other hand, traditional apprenticeships collapse."
            ]
        },
        {
            "rule_name": "Academic Hedging & Reporting Verbs (Menjaga Sikap Netral Akademis)",
            "formula": "Subject + [contend / argue / assert / observe] that...  |  Evidence + [tends to suggest / arguably implies] that...",
            "explanation": "Hedging mencegah penulis membuat klaim mutlak yang sombong atau tidak ilmiah. Reporting verbs mengatribusikan klaim kepada kubu terkait.",
            "examples": [
                "Proponents contend that automated engines elevate cognitive productivity.",
                "Evidence arguably suggests that a transitional period of frictional unemployment is inevitable.",
                "Scholars observe that algorithm-driven hiring may perpetuate training biases."
            ]
        }
    ])
    
    add_pronunciation_section(doc, [
        ("Emancipation", "/ɪˌmænsɪˈpeɪʃn/", "ih-man-si-PAY-shun", "Pembebasan dari belenggu tugas monoton"),
        ("Disenfranchisement", "/ˌdɪsɪnˈfræntʃaɪzmənt/", "dis-en-FRAN-chyz-muhnt", "Hilangnya hak ekonomi atau rasa berdaya"),
        ("Contend", "/kənˈtend/", "kuhn-TEND", "Berpendapat / menyatakan argumen"),
        ("Adversary", "/ˈædvəsəri/", "AD-ver-ser-ee", "Pihak lawan / oposisi dalam perdebatan"),
        ("Conversely", "/ˈkɒnvɜːsli/", "KON-vers-lee", "Sebaliknya / dari sudut pandang berlawanan"),
        ("Apprenticeship", "/əˈprentɪsʃɪp/", "uh-PREN-tis-ship", "Masa magang dan penempaan keterampilan"),
        ("Augmentation", "/ˌɔːɡmenˈteɪʃn/", "awg-men-TAY-shun", "Pemberdayaan dan peningkatan kapasitas"),
        ("Governance", "/ˈɡʌvənəns/", "GUV-er-nuhns", "Tata kelola regulasi yang sah dan adil")
    ], [
        "Thought Grouping & Pausing dalam Debat Formal:",
        "  - Selalu beri jeda singkat (pause) 1 detik setelah kata penghubung kontras: 'Conversely, [jeda sejenak] detractors emphasize that...'.",
        "Penurunan Nada di Akhir Kalimat Argumen (Falling Intonation):",
        "  - Dalam menyampaikan argumen ilmiah, turunkan nada bicara pada akhir kalimat untuk menegaskan otoritas akademis dan keyakinan faktual.",
        "Pelafalan Suku Kata 'Conversely':",
        "  - Berikan penekanan utama pada suku kata pertama: KON-vers-lee (/ˈkɒnvɜːsli/)."
    ])
    
    # Save Kelas 12
    out_docx = "C:/Users/LILI/OneDrive - PDK Jateng/Dokumen/Perangkat_SMAN_1_Bae/Kelas_12/Ringkasan_Materi_2_Pertemuan_Kelas_12.docx"
    doc.save(out_docx)
    print("New Student Handout DOCX Kelas 12 generated successfully!")

if __name__ == "__main__":
    generate_kelas_10()
    generate_kelas_11()
    generate_kelas_12()
    print("All Student Handout DOCX files created perfectly!")
