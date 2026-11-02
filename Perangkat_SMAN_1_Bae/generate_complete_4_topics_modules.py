# -*- coding: utf-8 -*-
"""
Generator Lengkap 4 Materi Esensial Semester Gasal (Berdasarkan Pemetaan MGMP Kudus 2025-2026)
Format: Modul Teoretis Belajar Siswa (Student Handout)
Setiap Materi Esensial Memuat:
1. Konsep Teori & Unsur-Unsur Teks yang Relevan
2. Contoh Teks / Dialog Otentik Model Teranotasi
3. Grammar (Tata Bahasa: Rumus Baku, Penjelasan, & Contoh Kontekstual)
4. Pronunciation Guide (Fonetik IPA, Penekanan Suku Kata, Arti, & Tips Pengucapan)
"""

import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from build_student_docx_lib import (
    create_student_doc, add_student_header, add_heading_1, add_heading_2,
    add_concept_box, add_elements_table, add_annotated_model_text,
    add_grammar_section, add_pronunciation_section
)

# ==============================================================================
# 1. KELAS X (FASE E): 4 MATERI ESENSIAL GASAL
# 1. Descriptive Text
# 2. Report Text
# 3. Simple Present Tense & Active / Passive Voice
# 4. Adjective Phrases & Expressing Intentions / Feelings
# ==============================================================================
def generate_docx_kelas_10():
    doc = create_student_doc()
    add_student_header(doc, "MODUL MATERI BELAJAR KELAS X (FASE E)", 
                      "4 MATERI ESENSIAL LENGKAP SEMESTER GASAL")
    
    # --------------------------------------------------------------------------
    # MATERI 1: DESCRIPTIVE TEXT
    # --------------------------------------------------------------------------
    add_heading_1(doc, "MATERI 1: DESCRIPTIVE TEXT (TEKS DESKRIPTIF)")
    add_concept_box(doc,
        "Descriptive Text adalah teks yang bertujuan untuk menggambarkan ciri fisik, keindahan, sifat, dan suasana suatu objek khusus secara spesifik (orang tertentu, tempat wisata khusus, atau bangunan bersejarah) sehingga pembaca seolah-olah melihat dan merasakannya sendiri.",
        "To describe and reveal a particular person, place, or thing. Fokus pada satu objek partikular (misal: Karimunjawa Marine National Park), bukan kelompok umum."
    )
    add_elements_table(doc, [
        ("1. Identification (Pengenalan Objek)", 
         "Memperkenalkan objek spesifik yang dideskripsikan, letak geografisnya, dan alasan mengapa objek tersebut terkenal atau istimewa.", 
         "Terletak di paragraf pertama. Menjawab: What/Who is being described? Where is it located?"),
        ("2. Description (Uraian Rinci)", 
         "Menguraikan rincian bagian-bagian fisik, pemandangan, atmosfer, serta kualitas khas menggunakan panca indra (visual, suara, aroma).", 
         "Terletak di paragraf kedua dan seterusnya. Sarat dengan sensory adjectives dan frasa sifat (adjective phrases).")
    ])
    p1_annotated = [
        ("[PARAGRAF 1: IDENTIFICATION]\nPengenalan objek Karimunjawa, lokasi 80 km barat laut Jepara, dan status suaka alam.",
         "Karimunjawa Marine National Park, an archipelago of twenty-seven pristine islands situated approximately 80 kilometers northwest of Jepara in Central Java, represents one of Indonesia’s most ecologically vital maritime sanctuaries. Declared a protected marine reserve, this tranquil haven attracts both local and international researchers who seek to observe sustainable marine ecosystems harmoniously preserved by the local coastal communities."),
        ("[PARAGRAF 2: DESCRIPTION]\nPemandangan bawah laut, terumbu karang, ikan badut, dan satwa daratan.",
         "The archipelago features breathtaking marine topography characterized by crystalline turquoise waters, vibrant fringing coral reefs, and extensive mangrove ecosystems. Beneath the calm surface, over two hundred species of scleractinian hard corals thrive alongside vibrant schools of clownfish, hawksbill sea turtles, and majestic spotted stingrays. The terrestrial landscape is equally enchanting, blanketed by verdant tropical rain forests that provide a natural habitat for the endemic red-breasted parakeet and white-bellied sea eagles."),
        ("[PARAGRAF 3: DESCRIPTION]\nKonservasi hiu oleh nelayan Menjangan Besar dan panorama senja Pantai Tanjung Gelam.",
         "Beyond its magnificent physical scenery, Karimunjawa offers visitors an authentic experience of community-based eco-tourism. Local fishermen on Pulau Menjangan Besar actively manage community-led shark rehabilitation pools, where wounded marine life is nurtured before being safely released into the wild. As dusk approaches, the horizon radiates an amber glow, casting long silhouettes across the white powdered sands of Tanjung Gelam Beach.")
    ]
    add_annotated_model_text(doc, "Karimunjawa: The Living Coral Haven of the Java Sea", 
                            "Indonesia Travel Bulletin (CEFR A2+/B1)", p1_annotated)
    add_grammar_section(doc, [
        {
            "rule_name": "Simple Present Tense pada Teks Deskriptif",
            "formula": "Verbal (+): Subject + Verb-1(s/es) + Object  |  Nominal (+): Subject + is/am/are + Adjective/Noun",
            "explanation": "Digunakan untuk menyatakan fakta fisik permanen dan kebenaran umum tentang objek yang dideskripsikan.",
            "examples": [
                "The archipelago features breathtaking marine topography. (Subject singular 'archipelago' + verb with -s)",
                "Over two hundred species of corals thrive beneath the surface. (Subject plural 'species' + verb-1)",
                "Karimunjawa is a protected marine reserve in Central Java."
            ]
        }
    ])
    add_pronunciation_section(doc, [
        ("Archipelago", "/ˌɑːkɪˈpeləɡəʊ/", "ar-chi-PEL-a-go", "Kepulauan / gugusan pulau"),
        ("Pristine", "/ˈprɪstiːn/", "PRIS-teen", "Murni, alami, belum terjamah"),
        ("Sanctuary", "/ˈsæŋktʃuəri/", "SANGK-chu-er-ee", "Suaka / suaka margasatwa"),
        ("Silhouette", "/ˌsɪluˈet/", "sil-oo-ET", "Bayangan gelap berlatar terang")
    ], [
        "Huruf 'ch' pada 'Archipelago' dilafalkan /k/ (bukan /tʃ/).",
        "Huruf 's' pada kata 'Island' bersifat SILENT (tidak dibaca!). Ucapkan /ˈaɪlənd/."
    ])

    # --------------------------------------------------------------------------
    # MATERI 2: REPORT TEXT
    # --------------------------------------------------------------------------
    add_heading_1(doc, "MATERI 2: REPORT TEXT (TEKS LAPORAN ILMIAH)")
    add_concept_box(doc,
        "Report Text adalah teks faktual ilmiah yang menyajikan informasi apa adanya tentang suatu fenomena alam, biologis, sosial, atau buatan secara UMUM (General) berdasarkan hasil pengamatan ilmiah, riset, atau klasifikasi taksonomi.",
        "To present information about something generally, as it is. Membahas seluruh kelompok spesies/fenomena (misal: Badak Jawa secara umum), bukan hewan peliharaan tertentu."
    )
    add_elements_table(doc, [
        ("1. General Classification (Klasifikasi Umum)", 
         "Pernyataan ilmiah pembuka yang mengelompokkan subjek ke dalam taksonomi, famili biologis, atau kategori umum yang sah.", 
         "Terletak di paragraf pertama. Memuat nama ilmiah (scientific name) dan wilayah sebaran alami."),
        ("2. Description (Uraian Faktual Sistematis)", 
         "Uraian objektif tentang bagian tubuh (morfologi), habitat, pola makan (diet), tingkah laku alami, dan ancaman konservasi.", 
         "Terletak di paragraf kedua dan seterusnya. Bebas dari perasaan emosional pribadi dan menggunakan istilah teknis.")
    ])
    p2_annotated = [
        ("[PARAGRAF 1: GENERAL CLASSIFICATION]\nKlasifikasi famili Rhinocerotidae dan habitat populasi global di Ujung Kulon.",
         "The Javan Rhinoceros (Rhinoceros sondaicus), also known as the lesser one-horned rhinoceros, is a critically endangered member of the family Rhinocerotidae and one of five extant rhinoceros species worldwide. Historically distributed across the lush tropical forests of Southeast Asia, Java, and Sumatra, the entire global population is now strictly confined to Ujung Kulon National Park at the westernmost tip of Java Island, Indonesia."),
        ("[PARAGRAF 2: DESCRIPTION (Morphology & Habits)]\nCiri fisik kulit zirah, cula tunggal (<25 cm), bobot tubuh, dan kebiasaan berkubang lumpur.",
         "Morphologically, the Javan rhino possesses a dusky grey, heavily folded skin that resembles armor plating, which provides natural protection against dense tropical vegetation and thorny thickets. Unlike African rhinoceros species which develop two prominent horns, the Javan rhino typically bears a single small horn rarely exceeding twenty-five centimeters in length, often found only in males. Adults can weigh between 900 and 2,300 kilograms, yet they are exceptionally agile movers capable of navigating dense jungle mudflats."),
        ("[PARAGRAF 3: DESCRIPTION (Diet & Conservation Threats)]\nMakanan herbivora, perburuan liar, dan pemantauan satelit intensif.",
         "Ecologically, Rhinoceros sondaicus is a solitary, herbivorous browser that feeds on a diverse variety of shoots, fallen fruits, twigs, and tender young leaves. They rely heavily on wallowing in natural mud pools to regulate body temperature. Due to severe habitat destruction and historical poaching for illegal wildlife trade, fewer than eighty individuals remain in the wild. Consequently, intensive satellite monitoring and anti-poaching patrols are continuously implemented.")
    ]
    add_annotated_model_text(doc, "The Javan Rhinoceros: Ecology and Survival Frontiers", 
                            "World Wildlife Journal (CEFR B1)", p2_annotated)
    add_grammar_section(doc, [
        {
            "rule_name": "Timeless Generic Present & Technical Terms",
            "formula": "Generic Noun Plural + Base Verb (Verb-1)  |  Technical Biological Terminology",
            "explanation": "Menyajikan kebiasaan alami satwa yang berlaku sepanjang masa (timeless truth).",
            "examples": [
                "Rhinos feed on shoots, fallen fruits, and tender young leaves. (Generic plural)",
                "They rely heavily on wallowing in natural mud pools.",
                "Technical Terms: 'herbivorous browser', 'dermal layers', 'extant species', 'Rhinocerotidae'."
            ]
        }
    ])
    add_pronunciation_section(doc, [
        ("Rhinoceros", "/raɪˈnɒsərəs/", "rye-NOSS-er-us", "Badak"),
        ("Herbivorous", "/hɜːˈbɪvərəs/", "her-BIV-or-us", "Pemakan tumbuhan / daun"),
        ("Extant", "/ˈekstənt/", "EK-stuhnt", "Masih hidup / ada di bumi saat ini"),
        ("Solitary", "/ˈsɒlətri/", "SOL-i-tree", "Menyendiri / tidak berkelompok")
    ], [
        "Huruf 'rh' pada 'Rhinoceros' dilafalkan /r/ (huruf 'h' tidak dibaca): /raɪˈnɒsərəs/.",
        "Kata 'Herbivorous' mendapat stress pada suku kata kedua: her-BIV-or-us."
    ])

    # --------------------------------------------------------------------------
    # MATERI 3: SIMPLE PRESENT TENSE & ACTIVE / PASSIVE VOICE
    # --------------------------------------------------------------------------
    add_heading_1(doc, "MATERI 3: SIMPLE PRESENT TENSE & ACTIVE / PASSIVE VOICE")
    add_concept_box(doc,
        "Dalam penulisan ilmiah dan akademik, siswa harus mampu mengubah kalimat aktif (di mana subjek melakukan aksi) menjadi kalimat pasif (di mana objek yang dikenai tindakan menjadi fokus utama) guna menciptakan gaya penulisan yang objektif dan formal.",
        "To shift focus from the actor to the action or receiver. Digunakan dalam laporan penelitian, pengolahan limbah, dan prosedur konservasi."
    )
    add_elements_table(doc, [
        ("Kalimat Aktif (Active Voice)", 
         "Subjek bertindak sebagai pelaku (doer of the action). Fokusnya adalah siapa yang melakukan pekerjaan.", 
         "Contoh: Rangers protect the national park. (Rangers = Subjek Pelaku)."),
        ("Kalimat Pasif (Passive Voice)", 
         "Objek penerima tindakan diangkat menjadi subjek utama di awal kalimat. Pelaku asli bisa diletakkan di akhir dengan 'by' atau dihilangkan bila sudah jelas/tidak penting.", 
         "Contoh: The national park is protected by rangers. (Fokus pada taman nasionalnya).")
    ])
    add_grammar_section(doc, [
        {
            "rule_name": "Rumus Baku Transformasi Aktif ke Pasif (Simple Present)",
            "formula": "Aktif: Subject + Verb-1(s/es) + Object\nPasif: Object (jadi Subjek Baru) + is / am / are + Verb-3 (Past Participle) + (by Agent)",
            "explanation": "Pilihlah to be (is/am/are) sesuai dengan jumlah Subjek Baru (is untuk tunggal/uncountable, are untuk jamak), lalu ubah kata kerja utama ke bentuk ketiga (Verb-3).",
            "examples": [
                "Aktif: Forest rangers closely monitor the rhino population. -> Pasif: The rhino population is closely monitored by forest rangers.",
                "Aktif: Students collect organic waste every Friday. -> Pasif: Organic waste is collected by students every Friday.",
                "Aktif: Gardeners cultivate rare orchid species in the greenhouse. -> Pasif: Rare orchid species are cultivated in the greenhouse."
            ]
        }
    ])
    add_pronunciation_section(doc, [
        ("Protected", "/prəˈtektɪd/", "pruh-TEK-tid", "Dilindungi (akhiran -ed dibaca /ɪd/)"),
        ("Confined", "/kənˈfaɪnd/", "kuhn-FYND", "Dibatasi / dikurung (akhiran -ed dibaca /d/)"),
        ("Cultivated", "/ˈkʌltɪveɪtɪd/", "KUL-ti-vay-tid", "Dibudidayakan (akhiran -ed dibaca /ɪd/)"),
        ("Implemented", "/ˈɪmplɪmentɪd/", "IM-pli-men-tid", "Diterapkan / dilaksanakan")
    ], [
        "Aturan Akhiran '-ed' pada Passive Voice:",
        "  1. Dibaca /ɪd/ jika kata kerja dasar berakhiran bunyi /t/ atau /d/ (e.g. protect -> protected /prəˈtektɪd/, plant -> planted /plɑːntɪd/).",
        "  2. Dibaca /d/ jika kata dasar berakhiran bunyi bersuara (e.g. confine -> confined /kənˈfaɪnd/, monitor -> monitored /ˈmɒnɪtəd/).",
        "  3. Dibaca /t/ jika kata dasar berakhiran bunyi tak bersuara (e.g. trap -> trapped /træpt/)."
    ])

    # --------------------------------------------------------------------------
    # MATERI 4: ADJECTIVE PHRASES & STATING INTENTIONS / FEELINGS
    # --------------------------------------------------------------------------
    add_heading_1(doc, "MATERI 4: ADJECTIVE PHRASES & STATING INTENTIONS / FEELINGS")
    add_concept_box(doc,
        "Materi ini memadukan pembentukan frasa sifat bertingkat untuk memperkaya deskripsi sensorik dengan ungkapan interpersonal untuk menyatakan niat, rencana masa depan (Intentions), serta perasaan antusias dan komitmen moral (Feelings) dalam konteks interaksi sosial siswa.",
        "To enrich descriptive details with compound adjectives and express future personal plans and emotional commitments clearly and politely."
    )
    add_elements_table(doc, [
        ("1. Adjective Phrases (Frasa Sifat)", 
         "Gabungan kata penjelas (adverb/determiner) dengan kata sifat utama dan kata benda untuk memberikan gambaran spesifik dan hidup.", 
         "Rumus: Determiner + Adverb + Adjective + Noun. Contoh: 'a remarkably pristine marine sanctuary'."),
        ("2. Stating Intentions (Ungkapan Niat)", 
         "Ungkapan untuk menyampaikan rencana kuat atau kehendak masa depan.", 
         "Frasa kunci: 'I intend to...', 'I am planning to...', 'I would like to...', 'I am definitely going to...'."),
        ("3. Stating Feelings (Ungkapan Perasaan)", 
         "Ungkapan untuk menyampaikan kondisi emosional positif/antusias terhadap suatu proyek atau pengalaman.", 
         "Frasa kunci: 'I feel immensely proud/excited...', 'It feels profoundly rewarding to...', 'We are enthusiastic about...'.")
    ])
    p4_dialog = [
        ("[GILIRAN BICARA 1: STATING INTENTION]\nAdit menyampaikan niat kuatnya meluncurkan program ramah lingkungan di sekolah.",
         "Adit: 'Ratna, after visiting the recycling center, I definitely intend to launch a zero-plastic challenge across all tenth-grade classes next month. We plan to produce bilingual infographic posters to educate our peers on waste sorting.'"),
        ("[GILIRAN BICARA 2: STATING FEELING & SUPPORT]\nRatna menyampaikan perasaan antusias dan komitmen berpartisipasi.",
         "Ratna: 'That is a wonderful initiative, Adit! I feel immensely enthusiastic about collaborating on this green project. It feels profoundly rewarding to know that our small actions can protect our school environment.'"),
        ("[GILIRAN BICARA 3: ADJECTIVE PHRASES & PLANNING]\nAdit menguraikan tempat pameran dengan frasa sifat bertingkat.",
         "Adit: 'We are going to organize the exhibition in our remarkably spacious school hall, showcasing creative handicrafts made from recycled materials.'"),
        ("[GILIRAN BICARA 4: MUTUAL COMMITMENT]\nRatna menegaskan kesiapan memulai kegiatan bersama.",
         "Ratna: 'I would like to invite our English teacher to review our English presentation slides. I am genuinely confident that this event will inspire everyone.'")
    ]
    add_annotated_model_text(doc, "Model Percakapan: Planning an Environmental Campaign", 
                            "SMAN 1 Bae Interpersonal Dialogue Bank", p4_dialog)
    add_grammar_section(doc, [
        {
            "rule_name": "Pola Kalimat Ungkapan Niat & Perasaan",
            "formula": "Intention: Subject + [intend to / plan to / am going to / would like to] + Verb-1\nFeeling: Subject + feel(s) + [immensely / profoundly / extremely] + Adjective",
            "explanation": "Gunakan kata keterangan derajat (immensely, profoundly) untuk memperkuat kesungguhan perasaan Anda.",
            "examples": [
                "I definitely intend to study environmental science at university.",
                "We are planning to visit Karimunjawa next semester.",
                "I feel profoundly grateful for this learning opportunity.",
                "The volunteers feel immensely rewarded after cleaning Tanjung Gelam Beach."
            ]
        }
    ])
    add_pronunciation_section(doc, [
        ("Intention", "/ɪnˈtenʃn/", "in-TEN-shun", "Niat / maksud / tujuan"),
        ("Enthusiastic", "/ɪnˌθjuːziˈæstɪk/", "en-thoo-zee-AS-tik", "Sangat antusias / bersemangat"),
        ("Profoundly", "/prəˈfaʊndli/", "pruh-FOWND-lee", "Secara mendalam / sangat tulus"),
        ("Rewarding", "/rɪˈwɔːdɪŋ/", "ri-WOR-ding", "Memberi kepuasan batin / berharga")
    ], [
        "Kata 'Enthusiastic' dilafalkan dengan bunyi /θ/ (lidah di gigi depan) dan penekanan suku kata pada 'AS': en-thoo-zee-AS-tik.",
        "Intonasi menyatakan perasaan antusias: Berikan nada naik yang bersemangat pada kata sifat emosi (*immensely excited!*)."
    ])

    out = "C:/Users/LILI/OneDrive - PDK Jateng/Dokumen/Perangkat_SMAN_1_Bae/Kelas_10/Ringkasan_Materi_2_Pertemuan_Kelas_10.docx"
    doc.save(out)
    print("DOCX Kelas 10 (4 Materi Esensial) generated successfully!")

# ==============================================================================
# 2. KELAS XI (FASE F): 4 MATERI ESENSIAL GASAL
# 1. Narrative Text
# 2. Analytical Exposition Text
# 3. Asking & Giving Opinions and Stating Agreement / Disagreement
# 4. Conditional Sentences Type 1 & Type 2
# ==============================================================================
def generate_docx_kelas_11():
    doc = create_student_doc()
    add_student_header(doc, "MODUL MATERI BELAJAR KELAS XI (FASE F)", 
                      "4 MATERI ESENSIAL LENGKAP SEMESTER GASAL")
    
    # --------------------------------------------------------------------------
    # MATERI 1: NARRATIVE TEXT
    # --------------------------------------------------------------------------
    add_heading_1(doc, "MATERI 1: NARRATIVE TEXT (TEKS NARATIF & KISAH MORAL)")
    add_concept_box(doc,
        "Narrative Text adalah teks cerita masa lampau (fiksi/legenda/fabel) yang menyajikan rangkaian peristiwa kronologis dengan konflik/krisis (Complication) yang memuncak dan akhirnya diselesaikan (Resolution) untuk menghibur pembaca serta menanamkan nilai moral.",
        "To entertain, amuse, and deal with actual or vicarious experiences. Mengajak pembaca merenungi dilema etika dan konsekuensi dari pilihan hidup sang tokoh."
    )
    add_elements_table(doc, [
        ("1. Orientation (Pengenalan)", 
         "Memperkenalkan tokoh utama (Who), latar tempat (Where), latar waktu (When), dan watak dasar tokoh.", 
         "Terletak di awal cerita (Paragraf 1)."),
        ("2. Complication (Krisis / Masalah)", 
         "Munculnya krisis, godaan, atau benturan moral yang dihadapi tokoh hingga mencapai puncak ketegangan (climax).", 
         "Terletak di paragraf tengah. Ditandai masalah etika dan pilihan sulit tokoh."),
        ("3. Resolution (Penyelesaian)", 
         "Bagian akhir ketika masalah terselesaikan, baik berakhir bahagia (happy ending), tragis (sad ending), atau pertobatan moral.", 
         "Terletak di bagian akhir cerita."),
        ("4. Coda (Nilai Moral)", 
         "Pesan moral atau hikmah hidup yang dapat dipetik pembaca dari kisah sang tokoh.", 
         "Tersurat di akhir atau tersirat dari perubahan batin tokoh.")
    ])
    p1_annotated = [
        ("[PARAGRAF 1: ORIENTATION]\nPengenalan Master Ki Ageng Semito di lereng Muria, keahlian jam mekanik, watak angkuh dan materialistis.",
         "In the mist-shrouded foothills of Mount Muria during the late eighteenth century, there lived a master clockmaker named Ki Ageng Semito. Renowned across Central Java for constructing intricate bronze timepieces that ticked with astonishing precision, he grew increasingly arrogant. He firmly believed that human intellect could subjugate time itself, treating every second as a commodity and refusing hospitality to anyone who could not afford his exorbitant services."),
        ("[PARAGRAF 2: COMPLICATION]\nKonflik moral: mengusir pengembara lansia dan anak kecil yang basah kuyup di tengah badai.",
         "One stormy evening during the height of the monsoon, while Ki Ageng was putting the final touches to his masterwork—a mechanical clock adorned with golden gears and automated puppets—a frail, soaked traveler knocked upon his oaken door, begging for shelter and dry embers to warm her trembling grandchild. Ki Ageng furrowed his brow and retorted harshly, 'My time is measured in gold, not wasted on penniless wanderers!' He slammed the door, oblivious to the ominous lightning that fractured the twilight sky."),
        ("[PARAGRAF 3: RESOLUTION & CODA]\nBanjir mata air bawah tanah, jam berputar balik, dan pertobatan Ki Ageng merengkuh kerendahan hati.",
         "Moments later, a deafening tremor shook the workshop. A subterranean spring burst through the stone floor, directly beneath his golden mechanism. Terrified, Ki Ageng watched as the pristine waters rapidly engulfed his coveted tools and blueprints. Strangely, as the water submerged the great clock, its hands began spinning wildly in reverse. In that terrifying moment of realization, Ki Ageng understood that all his accumulated wealth and mechanical mastery were powerless against the unyielding forces of nature and moral compassion. Escaping empty-handed into the pouring rain, he found the weeping child and shared his dry cloak, weeping tears of genuine humility.")
    ]
    add_annotated_model_text(doc, "The Clockmaker’s Spring: A Parable of Hubris and Humility",
                            "Contemporary Indonesian Literary Anthology (CEFR B1/B1+)", p1_annotated)
    add_grammar_section(doc, [
        {
            "rule_name": "Rangkaian Past Tenses & Direct/Indirect Speech",
            "formula": "Simple Past (V2)  |  Past Continuous (was/were + V-ing)  |  Past Perfect (had + V3)",
            "explanation": "Past continuous melukiskan aksi latar yang sedang berlangsung saat aksi tiba-tiba (Simple Past) memotongnya.",
            "examples": [
                "While he was putting the final touches, a traveler knocked upon the oaken door.",
                "Direct: Ki Ageng retorted, \"My time is measured in gold!\"",
                "Indirect: Ki Ageng retorted that his time was measured in gold."
            ]
        }
    ])
    add_pronunciation_section(doc, [
        ("Subterranean", "/ˌsʌbtəˈreɪniən/", "sub-tuh-RAY-nee-uhn", "Bawah tanah / di dalam perut bumi"),
        ("Exorbitant", "/ɪɡˈzɔːbɪtənt/", "ig-ZOR-bi-tuhnt", "Sangat mahal melampaui batas wajar"),
        ("Subjugate", "/ˈsʌbdʒʊɡeɪt/", "SUB-joo-gayt", "Menundukkan / menguasai secara paksa"),
        ("Humility", "/hjuːˈmɪləti/", "hyoo-MIL-i-tee", "Kerendahan hati sejati")
    ], [
        "Shook /ʃʊk/ adalah vokal pendek (seperti 'book', bukan 'shoot').",
        "Burst /bɜːst/ memiliki bentuk V1, V2, dan V3 yang persis sama."
    ])

    # --------------------------------------------------------------------------
    # MATERI 2: ANALYTICAL EXPOSITION TEXT
    # --------------------------------------------------------------------------
    add_heading_1(doc, "MATERI 2: ANALYTICAL EXPOSITION (TEKS EKSPOSISI ANALITIS)")
    add_concept_box(doc,
        "Analytical Exposition adalah teks argumentatif formal yang bertujuan meyakinkan pembaca (to persuade the audience) bahwa suatu isu atau fenomena penting adalah benar dan layak disikapi melalui serangkaian argumen rasional berbasis data.",
        "To persuade the reader that the idea is an important matter. Ciri khas mutlaknya: Diakhiri dengan Reiteration (penegasan kembali tesis), BUKAN rekomendasi saran aksi."
    )
    add_elements_table(doc, [
        ("1. Thesis (Pernyataan Pendirian)", 
         "Memperkenalkan topik krusial dan secara tegas menyatakan pendirian atau klaim penulis (writer's stance).", 
         "Terletak di akhir Paragraf 1. Memuat klaim utama."),
        ("2. Series of Arguments (Argumen & Bukti)", 
         "Rangkaian paragraf pendukung: Point (klaim argumen) dan Elaboration (bukti data, penalaran logis, fakta).", 
         "Diawali konjungsi urutan (Firstly, Secondly, Furthermore)."),
        ("3. Reiteration (Penegasan Kembali Tesis)", 
         "Menegaskan kembali posisi tesis penulis di akhir tulisan dengan pilihan kata (diksi) yang berbeda.", 
         "Terletak di paragraf terakhir. Diawali: 'In conclusion,...', 'To sum up,...'.")
    ])
    p2_annotated = [
        ("[PARAGRAF 1: THESIS]\nProliferasi AI dan klaim tesis mutlaknya literasi digital di SMA.",
         "The rapid proliferation of generative artificial intelligence and synthetic media has fundamentally transformed how young individuals consume, synthesize, and disseminate information. While these technological breakthroughs offer unprecedented learning conveniences, they simultaneously expose students to sophisticated misinformation, automated plagiarism, and deepfake content. Therefore, it is indisputable that integrating critical digital literacy into the high school curriculum is no longer merely optional, but an absolute educational imperative."),
        ("[PARAGRAF 2: ARGUMENT 1]\nPoint: Validasi sumber data; Elaboration: Mengatasi AI hallucinations dengan metode cek silang.",
         "Firstly, critical digital literacy equips adolescents with essential analytical scrutiny to evaluate the credibility of digital sources. In an era where algorithmic feeds prioritize sensationalism over factual veracity, uncritical consumers frequently mistake machine-generated hallucinations for verified truths. By understanding cross-referencing methodologies and source attribution, students can distinguish authentic academic consensus from commercially driven digital falsehoods."),
        ("[PARAGRAF 3: ARGUMENT 2]\nPoint: Integritas akademik; Elaboration: Nalar manusia tidak dapat dialihdayakan ke jaringan neural AI.",
         "Secondly, institutionalizing digital literacy fosters intellectual integrity and ethical awareness in the contemporary classroom. When learners understand the profound differences between using AI as an exploratory brainstorming partner versus engaging in automated academic dishonesty, they develop authentic academic resilience. They realize that authentic problem-solving and critical cognition cannot be outsourced to a predictive neural network."),
        ("[PARAGRAF 4: REITERATION]\nPenegasan kembali tesis: Sekolah harus membekali generasi muda menjadi navigator digital yang bertanggung jawab.",
         "In conclusion, secondary educational institutions must not perceive technological advancements in isolation from ethical reasoning. Enhancing students' capacity to critically deconstruct digital media ensures that young citizens emerge as discerning, responsible, and empowered digital navigators in an increasingly automated global society.")
    ]
    add_annotated_model_text(doc, "Why Critical Digital Literacy Must Be Prioritized in Secondary Education",
                            "Global Educational Technology Review (CEFR B1+/B2)", p2_annotated)
    add_grammar_section(doc, [
        {
            "rule_name": "Causal Conjunctions & Modal Adverbs of Certainty",
            "formula": "Premis + [therefore / consequently / leading to] + Kesimpulan  |  Subject + [indisputably / evidently] + Verb",
            "explanation": "Causal conjunctions menghubungkan premis dengan kesimpulan logis, sedangkan modal adverbs menegaskan keyakinan tesis.",
            "examples": [
                "Algorithmic feeds prioritize sensationalism; therefore, digital literacy is an absolute imperative.",
                "Students outsourced cognition; consequently, independent problem-solving skills deteriorated.",
                "It is indisputably proven that digital ethics fosters academic resilience."
            ]
        }
    ])
    add_pronunciation_section(doc, [
        ("Proliferation", "/prəˌlɪfəˈreɪʃn/", "pro-lif-er-AY-shun", "Perkembangan pesat / lonjakan tajam"),
        ("Veracity", "/vəˈræsəti/", "vuh-RASS-i-tee", "Kebenaran fakta / keakuratan data"),
        ("Indisputable", "/ˌɪndɪˈspjuːtəbl/", "in-di-SPYOO-tuh-bl", "Pasti benar / tak terbantahkan"),
        ("Discerning", "/dɪˈsɜːnɪŋ/", "dih-SURN-ing", "Kritis dan bijak membedakan hal benar")
    ], [
        "Kata berakhiran '-tion' selalu mendapat penekanan (stress) pada suku kata tepat sebelum akhiran (pro-lif-er-AY-tion).",
        "Kata berakhiran '-ity' mendapat stress pada suku kata ketiga dari belakang (ve-RASS-i-ty)."
    ])

    # --------------------------------------------------------------------------
    # MATERI 3: ASKING & GIVING OPINIONS, STATING AGREEMENT / DISAGREEMENT
    # --------------------------------------------------------------------------
    add_heading_1(doc, "MATERI 3: ASKING & GIVING OPINIONS, STATING AGREEMENT / DISAGREEMENT")
    add_concept_box(doc,
        "Dalam wacana akademik dan debat sekolah, siswa harus mampu menanyakan pendapat (asking opinion), menyatakan pandangan berbasis argumen logis (giving opinion), serta menyatakan persetujuan (agreement) dan sanggahan santun (polite disagreement) secara terhormat.",
        "To express personal viewpoints respectfully and navigate intellectual debates without personal animosity."
    )
    add_elements_table(doc, [
        ("1. Asking Opinions (Menanyakan Pendapat)", 
         "Ungkapan formal/santun untuk meminta sudut pandang rekan atau guru.", 
         "Contoh: 'What is your perspective on...?', 'Could you share your views regarding...?'"),
        ("2. Giving Opinions (Menyatakan Pendapat)", 
         "Menyatakan sikap didukung oleh alasan rasional (bukan sekadar selera pribadi).", 
         "Contoh: 'From my perspective,...', 'In my reasoned assessment,...', 'I firmly maintain that...'"),
        ("3. Stating Agreement (Menyatakan Persetujuan)", 
         "Persetujuan penuh atau persetujuan bersyarat (qualified agreement).", 
         "Contoh: 'I am of the same opinion.', 'I completely concur with your premise, provided that...'"),
        ("4. Polite Disagreement (Sanggahan Santun)", 
         "Menyanggah tanpa menyerang pribadi (mengakui poin lawan terlebih dahulu sebelum menyajikan bantahan).", 
         "Contoh: 'I see your point; however, the evidence points in a different direction.', 'With all due respect, I disagree.'")
    ])
    p3_dialog = [
        ("[GILIRAN 1: ASKING OPINION FORMAL]\nFarhan membuka sesi diskusi mengenai pemanfaatan kecerdasan buatan.",
         "Farhan: 'Citra, could you share your perspective regarding the proposal to completely ban generative AI from high school essay competitions?'"),
        ("[GILIRAN 2: GIVING OPINION & QUALIFIED DISAGREEMENT]\nCitra memberikan pandangan analitis dan menyanggah usulan larangan total secara santun.",
         "Citra: 'In my reasoned assessment, an outright ban is counterproductive. While I acknowledge the risks of automated plagiarism, prohibiting the technology ignores the pedagogical reality that students must be taught to use digital tools ethically.'"),
        ("[GILIRAN 3: QUALIFIED AGREEMENT]\nFarhan menyepakati perlunya regulasi daripada larangan mutlak.",
         "Farhan: 'I completely concur with that compromise. Provided that students submit their prompt logs and drafting histories, AI can serve as a legitimate brainstorming partner.'"),
        ("[GILIRAN 4: SYNTHESIS & CONSENSUS]\nCitra menyimpulkan konsensus musyawarah bersama.",
         "Citra: 'Precisely. Transparent accountability transforms a potential ethical hazard into an empowering learning opportunity.'")
    ]
    add_annotated_model_text(doc, "Model Wacana: Academic Seminar on Technology in Education", 
                            "SMAN 1 Bae Academic Debate Forum", p3_dialog)
    add_grammar_section(doc, [
        {
            "rule_name": "Frasa Ungkapan Opini & Kesantunan Diplomatis",
            "formula": "Giving Opinion: In my reasoned assessment / From my perspective, [Clause]\nPolite Disagreement: I see your point; however, / With all due respect, [Counter-argument]",
            "explanation": "Hindari menyerang orangnya (ad hominem); selalu serang atau bedah argumen dan datanya.",
            "examples": [
                "From my perspective, digital textbooks enhance information accessibility.",
                "I am of the same opinion; that conclusion is strongly supported by research data.",
                "With all due respect, I find myself unable to agree with an outright ban."
            ]
        }
    ])
    add_pronunciation_section(doc, [
        ("Perspective", "/pəˈspektɪv/", "per-SPEK-tiv", "Sudut pandang / perspektif"),
        ("Concur", "/kənˈkɜː/", "kuhn-KUR", "Setuju / sepakat"),
        ("Counterproductive", "/ˌkaʊntəprəˈdʌktɪv/", "kown-ter-pruh-DUK-tiv", "Bumerang / justru merugikan"),
        ("Accountability", "/əˌkaʊntəˈbɪləti/", "uh-kown-tuh-BIL-i-tee", "Akuntabilitas / tanggung jawab terbuka")
    ], [
        "Kata 'Concur' dilafalkan dengan penekanan pada suku kata kedua: kuhn-KUR /kənˈkɜː/.",
        "Intonasi Sanggahan Santun: Turunkan sedikit volume dan gunakan nada bicara datar-lembut saat mengucapkan 'I see your point, however...'."
    ])

    # --------------------------------------------------------------------------
    # MATERI 4: CONDITIONAL SENTENCES TYPE 1 & TYPE 2
    # --------------------------------------------------------------------------
    add_heading_1(doc, "MATERI 4: CONDITIONAL SENTENCES TYPE 1 & TYPE 2")
    add_concept_box(doc,
        "Kalimat pengandaian (Conditionals) digunakan untuk membicarakan kemungkinan hasil dari suatu kondisi tertentu. Type 1 membahas kemungkinan masa depan yang realistis (Future Real Possibility), sedangkan Type 2 membahas situasi pengandaian yang bertentangan dengan kenyataan saat ini (Unreal Present Condition).",
        "To express real future conditions versus hypothetical, contrary-to-fact present situations."
    )
    add_elements_table(doc, [
        ("Conditional Type 1 (Kemungkinan Nyata Masa Depan)", 
         "Kondisi yang sangat mungkin terjadi di masa depan jika syaratnya terpenuhi saat ini.", 
         "Rumus: If + Simple Present (V1), Result: Subject + will + Verb-1. Contoh: 'If we study hard, we will pass.'"),
        ("Conditional Type 2 (Pengandaian Bertolak Belakang dengan Realitas)", 
         "Kondisi khayalan/hipotesis saat ini yang kenyataannya TIDAK terjadi (unreal present).", 
         "Rumus: If + Simple Past (V2 / were), Result: Subject + would/could + Verb-1. Contoh: 'If I were you, I would verify the news.'")
    ])
    add_grammar_section(doc, [
        {
            "rule_name": "Perbandingan Rumus & Fakta Riil (Real Fact)",
            "formula": "Type 1: If + Subject + Verb-1(s/es), Subject + will + Verb-1\nType 2: If + Subject + Verb-2 (were untuk semua subjek), Subject + would/could + Verb-1",
            "explanation": "Pada Conditional Type 2 ragam formal, to be yang digunakan selalu 'were' (bukan 'was'), bahkan untuk subjek tunggal (I, he, she, it).",
            "examples": [
                "[Type 1]: If the school committee approves our proposal, we will install solar panels next semester. (Realistis)",
                "[Type 2]: If high school students possessed stronger digital literacy, they would not fall for fake news. (Fakta saat ini: Siswa belum memiliki literasi yang memadai, sehingga masih tertipu).",
                "[Type 2]: If I were the school principal, I would mandate a zero-plastic canteen policy."
            ]
        }
    ])
    add_pronunciation_section(doc, [
        ("Possibility", "/ˌpɒsəˈbɪləti/", "pos-uh-BIL-i-tee", "Kemungkinan nyata"),
        ("Hypothetical", "/ˌhaɪpəˈθetɪkl/", "hy-puh-THET-i-kuhl", "Berdasarkan dugaan / khayalan"),
        ("Mandate", "/ˈmændeɪt/", "MAN-dayt", "Mewajibkan secara resmi / mandat"),
        ("Contrary", "/ˈkɒntrəri/", "KON-truh-ree", "Bertentangan dengan kenyataan")
    ], [
        "Dalam pengandaian 'If I were you', hubungkan suara (linking sound): /ɪf aɪ wɜː juː/ tanpa jeda terputus.",
        "Kata 'Hypothetical' diucapkan dengan bunyi /θ/ (lidah di gigi) pada suku kata 'THET'."
    ])

    out = "C:/Users/LILI/OneDrive - PDK Jateng/Dokumen/Perangkat_SMAN_1_Bae/Kelas_11/Ringkasan_Materi_2_Pertemuan_Kelas_11.docx"
    doc.save(out)
    print("DOCX Kelas 11 (4 Materi Esensial) generated successfully!")

# ==============================================================================
# 3. KELAS XII (FASE F): 4 MATERI ESENSIAL GASAL
# 1. Narrative Text (Advanced / Allegory)
# 2. Argument & Stating Agreement
# 3. Discussion Text
# 4. Conditional Sentences Type 1, 2, 3 & Agreement (including Inversion)
# ==============================================================================
def generate_docx_kelas_12():
    doc = create_student_doc()
    add_student_header(doc, "MODUL MATERI BELAJAR KELAS XII (FASE F)", 
                      "4 MATERI ESENSIAL LENGKAP SEMESTER GASAL")
    
    # --------------------------------------------------------------------------
    # MATERI 1: ADVANCED NARRATIVE TEXT
    # --------------------------------------------------------------------------
    add_heading_1(doc, "MATERI 1: ADVANCED NARRATIVE TEXT (ALEGORI & SASTRA TINGKAT LANJUT)")
    add_concept_box(doc,
        "Advanced Narrative Text adalah karya fiksi sastra tingkat lanjut yang mengeksplorasi kedalaman psikologis karakter, dilema moral ambigu, serta bahasa simbolik (allegory & symbolism) untuk mengkritik realitas modern atau menggugah perenungan eksistensial pembaca.",
        "To explore existential, philosophical, and moral dilemmas through multi-layered narration. Tidak hanya mengisahkan peristiwa lahiriah, melainkan benturan nilai kemanusiaan abadi melawan arus zaman."
    )
    add_elements_table(doc, [
        ("1. Atmospheric Orientation (Latar Batin & Simbolik)", 
         "Membangun suasana emosional yang pekat melalui citraan sensorik arsitektur kuno, aroma khas, dan nuansa psikologis tokoh.", 
         "Terletak di Paragraf 1. Ciri khas: Penuh metafora dan simbolisme sejarah."),
        ("2. Multi-layered Complication (Benturan Paradigma)", 
         "Konflik batin atau benturan ideologi (misalnya efisiensi teknologi digital dingin versus ketulusan karya tangan manusia).", 
         "Terletak di Paragraf 2. Ciri khas: Karakter penantang mewakili pola pikir dunia modern yang utilitarian."),
        ("3. Epiphany / Turning Point (Momen Pencerahan Batin)", 
         "Momen hening ketika tokoh menyadari hakikat kebenaran yang lebih dalam melalui dialog atau benda pusaka.", 
         "Terletak di Paragraf 3. Ciri khas: Perubahan sudut pandang drastis dari dangkal menjadi mendalam."),
        ("4. Reflective Open Resolution (Resolusi Kontemplatif)", 
         "Akhir cerita yang memberi ruang refleksi bagi pembaca tanpa menyederhanakan masalah menjadi hitam-putih.", 
         "Terletak di Paragraf 4. Ciri khas: Menekankan nilai kemanusiaan abadi yang melampaui batas generasi.")
    ])
    p1_annotated = [
        ("[PARAGRAF 1: ATMOSPHERIC ORIENTATION]\nLatar Kota Lama Semarang; dedikasi Master Sastro terhadap naskah fisik di tengah dunia digital.",
         "Deep within the labyrinthine alleyways of Semarang’s Old Town, flanked by Dutch colonial archways whose plaster flaked like desiccated parchment, sat Master Sastro. For over six decades, he had practiced the vanishing art of hand-binding antique manuscripts, stitching fragile leather covers with waxed hemp thread. While the surrounding world surrendered itself to the dazzling ephemeral glow of digital screens and cloud servers, Sastro's sanctuary smelled persistently of aged cedar, beeswax, and crushed clove oil. To him, every tactile page possessed an indelible soul that binary code could never replicate."),
        ("[PARAGRAF 2: COMPLICATION (Philosophical Collision)]\nNadia menawarkan pemindaian digital 12 menit dengan scanner holografik.",
         "One sweltering afternoon, a young software developer named Nadia entered his workshop carrying an ultrathin holographic scanner. Tasked by a municipal archive consortium to digitize historical artifacts, she viewed Sastro’s ancient volumes merely as unindexed data sets awaiting conversion into high-density neural vectors. 'Bapak,' Nadia remarked politely, displaying her sleek handheld scanner, 'my machine can preserve your entire life's work in less than twelve minutes. If you allow me, physical decay will never threaten these texts again.'"),
        ("[PARAGRAF 3: EPIPHANY (Pencerahan Batin Melalui Naskah)]\nSastro memperdengarkan naskah kuno ke telinga Nadia: jiwa naskah melampaui kalkulasi piksel.",
         "Sastro paused, setting down his polished bone folder. He picked up a sun-bleached Javanese chronicle transcribed in golden botanical ink and gently held it near her ear. 'Listen closely, young lady,' he whispered gently. 'A machine can capture the shape of these letters with ruthless optical precision, but can it register the trembling hand of the poet who composed these verses by candlelight during the famine of 1845? Can your server transmit the scent of rain trapped between these dried leaves?'"),
        ("[PARAGRAF 4: REFLECTIVE RESOLUTION]\nNadia menyadari pelestarian warisan budaya adalah transmisi kerapuhan dan cinta manusia antar-generasi.",
         "Nadia froze, the hum of her cooling fan suddenly sounding intrusive amidst the solemn silence of the workshop. For the first time, she looked beyond the optical surface of the document, recognizing that cultural preservation is not merely an algorithmic calculation of pixels, but a sacred transmission of human vulnerability across generations.")
    ]
    add_annotated_model_text(doc, "The Memory Weaver of Kota Lama", 
                            "Asian Writers Review (CEFR B2)", p1_annotated)
    add_grammar_section(doc, [
        {
            "rule_name": "Participle Clauses & Cleft Sentences (Gaya Sastra Formal)",
            "formula": "Participle: [V-ing / V-3 Phrase], Subject + Main Verb\nCleft Sentence: It was / is + [Elemen Penekanan] + that / who + [Klausa]",
            "explanation": "Participle clauses memadatkan kalimat dan memperkaya ritme narasi sastra tingkat lanjut.",
            "examples": [
                "Tasked by a municipal archive consortium, Nadia viewed the ancient volumes merely as data sets.",
                "Stitching fragile leather covers with waxed hemp thread, Sastro preserved cultural history.",
                "Cleft: It was not the optical scanner, but the poet's sincere devotion that breathed life into the text."
            ]
        }
    ])
    add_pronunciation_section(doc, [
        ("Labyrinthine", "/ˌlæbəˈrɪnθaɪn/", "lab-uh-RIN-thyne", "Berliku-liku rumit bagai labirin"),
        ("Ephemeral", "/ɪˈfemərəl/", "ih-FEM-er-uhl", "Fana / berlangsung sekejap"),
        ("Indelible", "/ɪnˈdeləbl/", "in-DEL-uh-bl", "Abadi / tidak dapat dihapus waktu"),
        ("Vulnerability", "/ˌvʌlnərəˈbɪləti/", "vul-ner-uh-BIL-i-tee", "Kerapuhan nurani & kepekaan kemanusiaan")
    ], [
        "Kata 'Ephemeral' menggunakan vokal lemah schwa /ə/ pada suku kata terakhir: ih-FEM-er-uhl.",
        "Kata 'Vulnerability' memiliki penekanan utama pada suku kata 'BIL': vul-ner-a-BIL-i-ty."
    ])

    # --------------------------------------------------------------------------
    # MATERI 2: ARGUMENT & STATING AGREEMENT
    # --------------------------------------------------------------------------
    add_heading_1(doc, "MATERI 2: ARGUMENT & STATING AGREEMENT (WACANA ARGUMENTASI & DIPLOMASI)")
    add_concept_box(doc,
        "Teks argumentasi tingkat lanjut menuntut penguasaan struktur penalaran ilmiah (Claim, Evidence, Warrant, Concession) serta retorika diplomatis santun dalam menyatakan persetujuan formal dan penyelarasan sikap (Stating Agreement) di forum akademik internasional.",
        "To formulate rigorous academic arguments with evidentiary backing and articulate diplomatic alignment in high-stakes deliberative contexts."
    )
    add_elements_table(doc, [
        ("1. Claim (Pernyataan Klaim)", 
         "Pernyataan pendirian utama yang didukung data dan penalaran.", 
         "Contoh: 'Decarbonization mandates equitable financial solidarity between nations.'"),
        ("2. Evidence & Warrant (Bukti & Prinsip Pembenar)", 
         "Fakta empiris atau prinsip hukum internasional yang menghubungkan klaim dengan data.", 
         "Contoh: 'Under the Paris Agreement principle of Common But Differentiated Responsibilities...'"),
        ("3. Concession (Pengakuan Konsesif)", 
         "Mengakui validitas sebagian poin lawan sebelum memperkuat klaim mandiri.", 
         "Contoh: 'Notwithstanding the high upfront capital cost of renewables,...'"),
        ("4. Diplomatic Agreement (Persetujuan Diplomatik)", 
         "Penyelarasan sikap delegasi secara resmi dan berwibawa.", 
         "Contoh: 'Our delegation aligns itself entirely with that assessment; your premise is both just and indispensable.'")
    ])
    p2_dialog = [
        ("[GILIRAN 1: OPENING ARGUMENT WITH CONCESSION]\nDelegasi Indonesia menyampaikan argumen transisi energi berkeadilan.",
         "Delegate of Indonesia: 'Honorable Chair, notwithstanding the urgent imperative of global decarbonization, developing economies cannot sacrifice basic rural livelihoods. The principle of Common But Differentiated Responsibilities dictates that industrialized nations must provide substantive green technology transfers and concessional financing.'"),
        ("[GILIRAN 2: DIPLOMATIC AGREEMENT & ALIGNMENT]\nDelegasi Jerman menyatakan persetujuan penuh terhadap klaim keadilan iklim.",
         "Delegate of Germany: 'Our delegation aligns itself entirely with that assessment. The premise articulated by the distinguished delegate of Indonesia is both equitable and indispensable. Sustainable decarbonization can only succeed when financial solidarity accompanies ecological mandates.'"),
        ("[GILIRAN 3: CO-SPONSORING RESOLUTION]\nKedua delegasi merumuskan kesepakatan bersama.",
         "Delegate of Indonesia: 'We warmly welcome this consensus and invite all member states to co-sponsor the clean energy equity framework.'")
    ]
    add_annotated_model_text(doc, "Model Diplomasi: Model United Nations Climate Deliberation", 
                            "International Youth Diplomacy Proceedings", p2_dialog)
    add_grammar_section(doc, [
        {
            "rule_name": "Concessive Connectors & Diplomatic Register",
            "formula": "Notwithstanding + Noun Phrase, [Main Clause]  |  Albeit + Adjective, [Main Clause]",
            "explanation": "Frasa konsesif menunjukkan kematangan berpikir akademis tingkat tinggi.",
            "examples": [
                "Notwithstanding the high capital costs, renewable microgrids stabilize farming economies.",
                "The proposal, albeit ambitious, offers the only viable roadmap toward net-zero emissions.",
                "Our delegation aligns itself entirely with your core assessment."
            ]
        }
    ])
    add_pronunciation_section(doc, [
        ("Notwithstanding", "/ˌnɒtwɪθˈstændɪŋ/", "not-with-STAN-ding", "Meskipun demikian / terlepas dari"),
        ("Concessional", "/kənˈseʃənl/", "kuhn-SESH-uh-nuhl", "Bersifat lunak / berbunga sangat rendah"),
        ("Indispensable", "/ˌɪndɪˈspensəbl/", "in-di-SPEN-suh-bl", "Sangat mutlak / tak dapat ditiadakan"),
        ("Equitable", "/ˈekwɪtəbl/", "EK-wi-tuh-bl", "Berkeadilan / merata sesuai hak")
    ], [
        "Kata 'Notwithstanding' mendapat penekanan suara pada 'STAN': not-with-STAN-ding.",
        "Kata 'Equitable' diawali bunyi /e/ pendek: EK-wi-tuh-bl (bukan ee-kwi-tuh-bl)."
    ])

    # --------------------------------------------------------------------------
    # MATERI 3: DISCUSSION TEXT
    # --------------------------------------------------------------------------
    add_heading_1(doc, "MATERI 3: DISCUSSION TEXT (TEKS DISKUSI WACANA BERIMBANG)")
    add_concept_box(doc,
        "Discussion Text adalah teks argumentatif yang membahas suatu isu kontroversial atau dilematis dari setidaknya dua sudut pandang yang berlawanan (Arguments For / Pro vs Arguments Against / Contra) sebelum menyajikan kesimpulan sintesis netral.",
        "To present differing viewpoints on an issue before formulating a balanced conclusion. Penulis bertindak sebagai penengah yang adil dan objektif (neutral arbiter)."
    )
    add_elements_table(doc, [
        ("1. Statement of Issue (Rumusan Isu)", 
         "Memperkenalkan perdebatan kontroversial dunia beserta dua kubu yang berselisih paham.", 
         "Terletak di Paragraf 1. Ciri khas: Memuat perdebatan 'antara optimisme versus kecemasan sosial'."),
        ("2. Arguments For / Pro (Sudut Pandang Pendukung)", 
         "Uraian keunggulan, lonjakan produktivitas, dan peluang lapangan kerja baru dari penerapan isu.", 
         "Diawali: 'Proponents contend that...', 'Advocates point out that...'."),
        ("3. Arguments Against / Contra (Sudut Pandang Pengkritik)", 
         "Uraian risiko, potensi bahaya bias algoritma, dan ancaman pemutusan hubungan kerja.", 
         "Diawali: 'Conversely,...', 'On the other hand,...', 'Detractors argue that...'."),
        ("4. Balanced Conclusion (Kesimpulan Sintesis Berimbang)", 
         "Rekomendasi jalan tengah berupa kerangka regulasi dan pengawasan etis multi-pemangku kepentingan.", 
         "Terletak di paragraf terakhir. Tidak memihak ekstrem, melainkan menawarkan solusi kompromi.")
    ])
    p3_annotated = [
        ("[PARAGRAF 1: STATEMENT OF ISSUE]\nPerdebatan global: Otomatisasi tempat kerja sebagai emansipasi talenta vs peminggiran sosial ekonomi.",
         "The relentless integration of autonomous algorithmic systems and generative artificial intelligence into white-collar professions has ignited fierce international debate. Across legal, financial, and creative sectors, cognitive tasks once deemed exclusively human are now executed with breathtaking speed by neural networks. Consequently, economists, policymakers, and labor leaders find themselves sharply divided over whether automation represents an unprecedented emancipation of human talent or a perilous catalyst for widespread socioeconomic disenfranchisement."),
        ("[PARAGRAF 2: ARGUMENTS FOR (PRO)]\nLonjakan produktivitas, pembebasan dari tugas monoton, dan profesi baru (prompt engineering, auditing).",
         "Proponents of rapid workplace automation contend that cognitive AI elevates human productivity to historic heights. By delegating repetitive computational analysis, data verification, and drafting procedures to automated engines, professionals can redirect their cognitive energy toward high-level strategic reasoning, ethical diplomacy, and interpersonal empathy. Furthermore, business analysts argue that augmented enterprises generate entirely new technological ecosystems, creating millions of specialized roles in algorithmic auditing, prompt engineering, and cyber-security oversight that previously did not exist."),
        ("[PARAGRAF 3: ARGUMENTS AGAINST (CONTRA)]\nPengikisan jenjang karir magang pemula, bias algoritma, dan hilangnya akuntabilitas manusia.",
         "Conversely, detractors and labor unions present grave concerns regarding structural displacement and the erosion of entry-level career pathways. When junior analyst and associate tasks are swallowed by software subscriptions, the traditional apprenticeship model through which young graduates hone professional judgment collapses. Furthermore, critics emphasize that automated systems frequently perpetuate historical biases embedded within training datasets, resulting in systemic discrimination in recruitment and credit scoring without transparent human accountability."),
        ("[PARAGRAF 4: BALANCED CONCLUSION]\nSintesis jalan tengah: Tata kelola multi-pihak, pelatihan ulang pekerja, dan pengawasan manusia atas mesin.",
         "In conclusion, neither blind technophilia nor reflexive resistance provides an adequate response to the fourth industrial revolution. A balanced strategy requires robust multi-stakeholder governance: corporations must reinvest productivity gains into lifelong employee reskilling, while governments formulate strict labor protections that mandate human oversight in consequential decisions. The ultimate objective must be human-centered augmentation rather than unbridled worker replacement.")
    ]
    add_annotated_model_text(doc, "The Autonomous Workplace: Collaborative Evolution or Mass Disenfranchisement?",
                            "World Economic Forum Deliberation Report (CEFR B2)", p3_annotated)
    add_grammar_section(doc, [
        {
            "rule_name": "Contrastive Linkers & Academic Hedging",
            "formula": "Kalimat Pro. [Conversely, / On the other hand, / Nevertheless,] Kalimat Kontra.\nSubject + [contend / argue / assert] that...  |  Evidence + [arguably suggests / tends to indicate] that...",
            "explanation": "Penanda kontras menghubungkan argumen pro dan kontra secara transparan; hedging menjaga kehati-hatian ilmiah.",
            "examples": [
                "Proponents celebrate efficiency; conversely, labor unions raise alarms on unemployment.",
                "Evidence arguably suggests that cognitive tasks will undergo transitional disruptions.",
                "Detractors contend that systemic biases remain embedded within training data."
            ]
        }
    ])
    add_pronunciation_section(doc, [
        ("Disenfranchisement", "/ˌdɪsɪnˈfræntʃaɪzmənt/", "dis-en-FRAN-chyz-muhnt", "Hilangnya hak ekonomi atau kesempatan berdaya"),
        ("Apprenticeship", "/əˈprentɪsʃɪp/", "uh-PREN-tis-ship", "Masa magang dan penempaan keahlian"),
        ("Augmentation", "/ˌɔːɡmenˈteɪʃn/", "awg-men-TAY-shun", "Pemberdayaan dan peningkatan kapasitas"),
        ("Governance", "/ˈɡʌvənəns/", "GUV-er-nuhns", "Tata kelola regulasi yang sah dan adil")
    ], [
        "Thought Grouping & Pausing: Beri jeda 1 detik setelah kata penghubung kontras: 'Conversely, [jeda sejenak] detractors emphasize that...'.",
        "Gunakan Falling Intonation (nada turun) di akhir kalimat kesimpulan untuk menunjukkan objektivitas netral."
    ])

    # --------------------------------------------------------------------------
    # MATERI 4: CONDITIONAL SENTENCES TYPE 1, 2, 3 & AGREEMENT (INVERSION)
    # --------------------------------------------------------------------------
    add_heading_1(doc, "MATERI 4: CONDITIONAL SENTENCES TYPE 1, 2, 3 & INVERSION")
    add_concept_box(doc,
        "Siswa kelas XII harus menguasai seluruh spektrum kalimat pengandaian: Type 1 (Kemungkinan masa depan), Type 2 (Khayalan bertolak belakang dengan realitas saat ini), Type 3 (Penyesalan masa lampau yang tidak mungkin terulang), serta pola Inversi Formal (tanpa kata 'If') untuk penulisan esai akademik.",
        "To express future possibilities, hypothetical present conditions, past counterfactual regrets, and formal stylistic inversions."
    )
    add_elements_table(doc, [
        ("Conditional Type 1 (Future Real)", 
         "Peluang nyata di masa depan bila syarat terpenuhi.", 
         "Rumus: If + Simple Present (V1), will + Verb-1. Contoh: 'If we verify data, we will avoid fake news.'"),
        ("Conditional Type 2 (Unreal Present)", 
         "Khayalan saat ini yang berlawanan dengan kenyataan.", 
         "Rumus: If + Simple Past (were/V2), would + Verb-1. Contoh: 'If he were humble, he would welcome the traveler.'"),
        ("Conditional Type 3 (Unreal Past / Regret)", 
         "Penyesalan masa lampau yang sudah terjadi dan tidak dapat diubah lagi.", 
         "Rumus: If + Past Perfect (had + V3), would have + Verb-3. Contoh: 'If the engineers had reinforced the seawall, the disaster would not have occurred.'"),
        ("Inversion Conditionals (Inversi Formal Tanpa 'If')", 
         "Pembalikan letak kata kerja bantu ke depan subjek untuk gaya bahasa akademik resmi tingkat tinggi.", 
         "Pola: 'Should you require...' (Type 1), 'Were the government to mandate...' (Type 2), 'Had Ki Ageng listened...' (Type 3).")
    ])
    add_grammar_section(doc, [
        {
            "rule_name": "Aturan Transformasi Inversi Formal (Stylistic Inversion)",
            "formula": "1. Type 1: Should + Subject + Verb-1, [Result Clause]\n2. Type 2: Were + Subject + to Verb-1, [Result Clause]\n3. Type 3: Had + Subject + Verb-3, [Result Clause with would have V3]",
            "explanation": "Pada ragam formal akademik internasional, kata 'If' dihilangkan dan diganti dengan memindahkan Should, Were, atau Had ke posisi paling depan kalimat.",
            "examples": [
                "[Type 1 Inversion]: Should international inspectors visit the plant, they will observe strict safety protocols.",
                "[Type 2 Inversion]: Were policymakers to mandate carbon taxes tomorrow, corporations would rapidly adjust.",
                "[Type 3 Inversion]: Had Ki Ageng Semito embraced moral compassion, his workshop would have remained intact.",
                "[Mixed Conditional]: If our ancestors had not preserved ancient chronicles, we would not understand our identity today."
            ]
        }
    ])
    add_pronunciation_section(doc, [
        ("Counterfactual", "/ˌkaʊntəˈfæktʃuəl/", "kown-ter-FAK-choo-uhl", "Bertolak belakang dengan kenyataan sejarah"),
        ("Inversion", "/ɪnˈvɜːʃn/", "in-VUR-shun", "Pembalikan struktur tata bahasa"),
        ("Requisite", "/ˈrekwɪzɪt/", "REK-wi-zit", "Syarat yang mutlak diperlukan"),
        ("Solidarity", "/ˌsɒlɪˈdærəti/", "sol-i-DARE-i-tee", "Solidaritas / kebersamaan rasa")
    ], [
        "Inversi 'Had Ki Ageng listened': Tekan kata 'Had' dan kata kerja utama 'listened' dengan ritme mantap.",
        "Kata 'Requisite' dilafalkan /REK-wi-zit/ (suku kata pertama ditekankan)."
    ])

    out = "C:/Users/LILI/OneDrive - PDK Jateng/Dokumen/Perangkat_SMAN_1_Bae/Kelas_12/Ringkasan_Materi_2_Pertemuan_Kelas_12.docx"
    doc.save(out)
    print("DOCX Kelas 12 (4 Materi Esensial) generated successfully!")

if __name__ == "__main__":
    generate_docx_kelas_10()
    generate_docx_kelas_11()
    generate_docx_kelas_12()
    print("ALL 3 GRADES GENERATED WITH COMPLETE 4 ESSENTIAL TOPICS!")
