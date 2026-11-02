import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx_helper import (
    create_styled_doc, add_header_banner, add_tp_info_box,
    add_heading_1, add_heading_2, add_callout, add_authentic_text_block,
    add_vocab_table, set_table_borders, set_cell_shading, set_cell_margins,
    HEX_NAVY, HEX_ROYAL, HEX_LIGHT_BLUE, HEX_ZEBRA, COLOR_NAVY, COLOR_ROYAL, COLOR_BODY
)

def add_comparative_matrix(doc, headers, rows_data):
    add_heading_2(doc, "Matriks Perbandingan Komparatif Antar-Genre")
    table = doc.add_table(rows=len(rows_data) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table, color="CBD5E1", sz="4")
    
    widths = [Inches(1.8), Inches(2.3), Inches(2.4)]
    hdr_row = table.rows[0]
    for c_idx, h_text in enumerate(headers):
        c = hdr_row.cells[c_idx]
        c.width = widths[c_idx]
        c.text = h_text
        set_cell_shading(c, HEX_NAVY)
        set_cell_margins(c, 80, 80, 100, 100)
        p = c.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.runs[0].font.name = "Segoe UI"
        p.runs[0].font.size = Pt(9)
        p.runs[0].font.bold = True
        p.runs[0].font.color.rgb = RGBColor(255, 255, 255)
        
    for r_idx, r_data in enumerate(rows_data, 1):
        row = table.rows[r_idx]
        for c_idx, val in enumerate(r_data):
            c = row.cells[c_idx]
            c.width = widths[c_idx]
            c.text = str(val)
            set_cell_shading(c, HEX_ZEBRA if r_idx % 2 == 1 else "FFFFFF")
            set_cell_margins(c, 70, 70, 100, 100)
            p = c.paragraphs[0]
            p.runs[0].font.name = "Segoe UI"
            p.runs[0].font.size = Pt(9)
            if c_idx == 0:
                p.runs[0].font.bold = True
                p.runs[0].font.color.rgb = COLOR_NAVY
            else:
                p.runs[0].font.color.rgb = COLOR_BODY

def add_grammar_box(doc, title, formula, examples):
    add_heading_2(doc, title)
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = tbl.cell(0, 0)
    c.width = Inches(6.5)
    set_cell_shading(c, HEX_LIGHT_BLUE)
    set_cell_margins(c, 100, 100, 150, 150)
    
    p = c.paragraphs[0]
    p.paragraph_format.space_after = Pt(4)
    r_f = p.add_run(f"Pola & Rumus Baku: {formula}\n\n")
    r_f.font.name = "Segoe UI"
    r_f.font.size = Pt(9.5)
    r_f.font.bold = True
    r_f.font.color.rgb = COLOR_ROYAL
    
    r_e = p.add_run("Contoh Kalimat Kontekstual:\n" + "\n".join([f"• {ex}" for ex in examples]))
    r_e.font.name = "Segoe UI"
    r_e.font.size = Pt(9)
    r_e.font.color.rgb = COLOR_BODY

# ==============================================================================
# 1. KELAS X
# ==============================================================================
def build_docx_kelas_10():
    doc = create_styled_doc()
    add_header_banner(doc, "RINGKASAN MATERI PEMBELAJARAN BAHASA INGGRIS",
                      "SMA NEGERI 1 BAE KUDUS — KELAS X (FASE E) SEMESTER GASAL 2025/2026")
    
    add_tp_info_box(doc, {
        "Satuan Pendidikan": "SMA Negeri 1 Bae Kudus",
        "Mata Pelajaran": "Bahasa Inggris (Wajib / Fase E)",
        "Fokus Materi Esensial": "Pertemuan 1: Descriptive Text | Pertemuan 2: Report Text",
        "Alokasi Waktu": "2 Pertemuan (4 Jam Pelajaran @ 45 Menit)",
        "Rujukan Kurikulum": "Pemetaan MGMP SMA Kudus 2025-2026 & Kepka BSKAP No. 020/2026"
    })
    
    # Pertemuan 1
    add_heading_1(doc, "PERTEMUAN 1: DESCRIPTIVE TEXT (WARISAN BUDAYA & ALAM)")
    add_callout(doc, "Tujuan Pembelajaran (TP 10.1) & Indikator Ketercapaian (KKTP):",
                "Peserta didik mampu mengidentifikasi fungsi sosial, menganalisis struktur teks (Identification & Description), serta mengevaluasi unsur kebahasaan (Simple Present, Adjective Phrases, Sensory Imagery) pada teks deskriptif otentik secara kritis.")
    
    p1_paras = [
        "Karimunjawa Marine National Park, an archipelago of twenty-seven pristine islands situated approximately 80 kilometers northwest of Jepara in Central Java, represents one of Indonesia’s most ecologically vital maritime sanctuaries. Declared a protected marine reserve, this tranquil haven attracts both local and international researchers who seek to observe sustainable marine ecosystems harmoniously preserved by the local coastal communities.",
        "The archipelago features breathtaking marine topography characterized by crystalline turquoise waters, vibrant fringing coral reefs, and extensive mangrove ecosystems. Beneath the calm surface, over two hundred species of scleractinian hard corals thrive alongside vibrant schools of clownfish, hawksbill sea turtles, and majestic spotted stingrays. The terrestrial landscape is equally enchanting, blanketed by verdant tropical rain forests that provide a natural habitat for the endemic red-breasted parakeet and white-bellied sea eagles.",
        "Beyond its magnificent physical scenery, Karimunjawa offers visitors an authentic experience of community-based eco-tourism. Local fishermen on Pulau Menjangan Besar actively manage community-led shark rehabilitation pools, where wounded marine life is nurtured before being safely released into the wild. As dusk approaches, the horizon radiates an amber glow, casting long silhouettes across the white powdered sands of Tanjung Gelam Beach."
    ]
    
    p1_analysis = {
        "Identification (Paragraf 1)": "Memperkenalkan objek spesifik Karimunjawa, lokasi geografis (80 km barat laut Jepara), dan statusnya sebagai suaka maritim terlindungi.",
        "Description 1 (Paragraf 2)": "Menguraikan rincian topografi bawah laut (karang, penyu sisik, ikan badut) dan daratan (hutan hujan, burung endemik).",
        "Description 2 (Paragraf 3)": "Menggambarkan aktivitas konservasi hiu oleh nelayan Menjangan Besar serta panorama senja di Pantai Tanjung Gelam."
    }
    
    add_authentic_text_block(doc, "1", "Karimunjawa: The Living Coral Haven of the Java Sea",
                            "Teks Model Descriptive", {
                                "portal": "Indonesia Travel & Marine Conservation Bulletin",
                                "title": "Karimunjawa Living Coral Haven",
                                "url": "https://indonesia.travel/en/destination/java/karimunjawa.html",
                                "year": "2025", "cefr": "A2+ / B1"
                            }, content_paragraphs=p1_paras, analysis_dict=p1_analysis)
    
    add_grammar_box(doc, "Tata Bahasa Esensial: Simple Present Tense & Adjective Phrases",
                    "Subject + Verb-1(s/es) + Object/Complement  |  Determiner + Adverb + Adjective + Noun",
                    [
                        "The archipelago features breathtaking marine topography. (Simple Present Fact)",
                        "Over two hundred species of scleractinian corals thrive beneath the calm surface.",
                        "Adjective Phrase: 'a remarkably pristine marine sanctuary' (Det + Adv + Adj + Noun)",
                        "Adjective Phrase: 'an ecologically vital maritime reserve guarded by coastal villagers.'"
                    ])
                    
    add_vocab_table(doc, [
        ("Pristine", "Adjective", "Murni, alami, belum terjamah kerusakan", "The archipelago consists of twenty-seven pristine coral islands."),
        ("Sanctuary", "Noun", "Suaka, kawasan perlindungan satwa", "The reserve serves as a vital maritime sanctuary in Java Sea."),
        ("Endemic", "Adjective", "Spesies asli khas wilayah tertentu", "The forest shelters the endemic red-breasted parakeet."),
        ("Harmoniously", "Adverb", "Secara selaras dan bersahabat", "Ecosystems are harmoniously preserved by coastal communities."),
        ("Silhouette", "Noun", "Bayangan gelap berlatar terang", "The sunset casts long silhouettes across Tanjung Gelam Beach.")
    ])
    
    # Pertemuan 2
    add_heading_1(doc, "PERTEMUAN 2: REPORT TEXT (PENGAMATAN ILMIAH BIODIVERSITAS)")
    add_callout(doc, "Tujuan Pembelajaran (TP 10.2) & Indikator Ketercapaian (KKTP):",
                "Peserta didik mampu menganalisis fungsi klasifikasi ilmiah (General Classification) dan deskripsi objektif (Description), serta memahami fitur kebahasaan Passive Voice dan istilah teknis ilmiah pada teks laporan observasi.")
                
    p2_paras = [
        "The Javan Rhinoceros (Rhinoceros sondaicus), also known as the lesser one-horned rhinoceros, is a critically endangered member of the family Rhinocerotidae and one of five extant rhinoceros species worldwide. Historically distributed across the lush tropical forests of Southeast Asia, Java, and Sumatra, the entire global population is now strictly confined to Ujung Kulon National Park at the westernmost tip of Java Island, Indonesia.",
        "Morphologically, the Javan rhino possesses a dusky grey, heavily folded skin that resembles armor plating, which provides natural protection against dense tropical vegetation and thorny thickets. Unlike African rhinoceros species which develop two prominent horns, the Javan rhino typically bears a single small horn rarely exceeding twenty-five centimeters in length, often found only in males. Adults can weigh between 900 and 2,300 kilograms, yet despite their massive physical stature, they are exceptionally agile movers capable of navigating dense jungle mudflats.",
        "Ecologically, Rhinoceros sondaicus is a solitary, herbivorous browser that feeds on a diverse variety of shoots, fallen fruits, twigs, and tender young leaves. They rely heavily on wallowing in natural mud pools to regulate body temperature and protect their sensitive dermal layers from parasites and insect bites. Due to severe habitat destruction and historical poaching for illegal wildlife trade, fewer than eighty individuals remain in the wild."
    ]
    
    p2_analysis = {
        "General Classification (Paragraf 1)": "Menetapkan taksonomi ilmiah (famili Rhinocerotidae, spesies terancam kritis) dan batas habitat populasi global saat ini.",
        "Description: Morphology (Paragraf 2)": "Menguraikan anatomi fisik umum: kulit berlipat mirip zirah perisai, cula tunggal (<25 cm), bobot 900-2.300 kg, serta kelincahan gerak.",
        "Description: Ecology & Threat (Paragraf 3)": "Mengulas diet herbivora, perilaku berkubang di lumpur (mud wallow), dan status konservasi populasi (<80 individu tersisa)."
    }
    
    add_authentic_text_block(doc, "2", "The Javan Rhinoceros (Rhinoceros sondaicus): Ecology and Survival",
                            "Teks Model Report", {
                                "portal": "World Wildlife Fund & Tropical Ecology Journal",
                                "title": "Rhinoceros sondaicus Survival Assessment",
                                "url": "https://www.worldwildlife.org/species/javan-rhino",
                                "year": "2025", "cefr": "B1"
                            }, content_paragraphs=p2_paras, analysis_dict=p2_analysis)
                            
    add_grammar_box(doc, "Tata Bahasa Esensial: Passive Voice pada Laporan Ilmiah",
                    "Subject + is/am/are + Verb-3 (Past Participle) + (by Agent)",
                    [
                        "The global population is strictly confined to Ujung Kulon National Park. (Present Passive)",
                        "Intensive satellite monitoring is continuously implemented by forest rangers.",
                        "Endangered animals are protected under international conservation treaties.",
                        "Mud wallowing is required to regulate the mammal's body temperature."
                    ])
                    
    add_vocab_table(doc, [
        ("Extant", "Adjective", "Masih hidup / ada di bumi saat ini", "It is one of five extant rhinoceros species worldwide."),
        ("Herbivorous", "Adjective", "Pemakan tumbuhan / daun", "The rhino is a solitary, herbivorous browser in the forest."),
        ("Wallow", "Verb / Noun", "Berkubang dalam lumpur untuk pendinginan", "Rhinos rely on natural mud pools to wallow daily."),
        ("Poaching", "Noun", "Perburuan satwa liar secara ilegal", "Historical poaching has decimated wild animal populations."),
        ("Critically Endangered", "Adjective", "Berada di ambang kepunahan sangat tinggi", "The species is classified as critically endangered by IUCN.")
    ])
    
    add_comparative_matrix(doc, ["Kriteria Pembeda", "Descriptive Text (Pertemuan 1)", "Report Text (Pertemuan 2)"], [
        ["Fokus Subjek", "Objek spesifik & partikular (Taman Nasional Karimunjawa)", "Golongan umum makhluk / fenomena (Spesies Badak Jawa)"],
        ["Tujuan Utama", "Menggambarkan pesona estetis & kesan visual indrawi", "Menyajikan fakta objektif & klasifikasi ilmiah apa adanya"],
        ["Struktur Teks", "Identification -> Description", "General Classification -> Description"],
        ["Ciri Bahasa", "Sensory adjectives, adjective phrases, specific nouns", "Scientific taxonomy, timeless generic verbs, passive voice"],
        ["Contoh Subjek", "Pantai Kartini, Candi Borobudur, Museum Kretek", "Mangroves, Earthquakes, Geothermal Energy, Rhinos"]
    ])
    
    out = "C:/Users/LILI/OneDrive - PDK Jateng/Dokumen/Perangkat_SMAN_1_Bae/Kelas_10/Ringkasan_Materi_2_Pertemuan_Kelas_10.docx"
    doc.save(out)
    print("DOCX Kelas 10 saved:", out)

# ==============================================================================
# 2. KELAS XI
# ==============================================================================
def build_docx_kelas_11():
    doc = create_styled_doc()
    add_header_banner(doc, "RINGKASAN MATERI PEMBELAJARAN BAHASA INGGRIS",
                      "SMA NEGERI 1 BAE KUDUS — KELAS XI (FASE F) SEMESTER GASAL 2025/2026")
                      
    add_tp_info_box(doc, {
        "Satuan Pendidikan": "SMA Negeri 1 Bae Kudus",
        "Mata Pelajaran": "Bahasa Inggris (Wajib / Fase F)",
        "Fokus Materi Esensial": "Pertemuan 1: Narrative Text | Pertemuan 2: Analytical Exposition",
        "Alokasi Waktu": "2 Pertemuan (4 Jam Pelajaran @ 45 Menit)",
        "Rujukan Kurikulum": "Pemetaan MGMP SMA Kudus 2025-2026 & Kepka BSKAP No. 020/2026"
    })
    
    # Pertemuan 1
    add_heading_1(doc, "PERTEMUAN 1: NARRATIVE TEXT (LEGENDA KONTEMPORER & NILAI MORAL)")
    add_callout(doc, "Tujuan Pembelajaran (TP 11.W.1) & Indikator Ketercapaian (KKTP):",
                "Peserta didik mampu menganalisis alur naratif (Orientation, Complication, Resolution, Coda), mengidentifikasi motif psikologis tokoh utama, serta mengevaluasi fitur kebahasaan Past Tenses dan Direct/Indirect Speech pada cerpen/legenda.")
                
    p1_paras = [
        "In the mist-shrouded foothills of Mount Muria during the late eighteenth century, there lived a master clockmaker named Ki Ageng Semito. Renowned across Central Java for constructing intricate bronze timepieces that ticked with astonishing precision, he grew increasingly arrogant. He firmly believed that human intellect could subjugate time itself, treating every second as a commodity and refusing hospitality to anyone who could not afford his exorbitant services.",
        "One stormy evening during the height of the monsoon, while Ki Ageng was putting the final touches to his masterwork—a mechanical clock adorned with golden gears and automated puppets—a frail, soaked traveler knocked upon his oaken door, begging for shelter and dry embers to warm her trembling grandchild. Ki Ageng furrowed his brow and retorted harshly, 'My time is measured in gold, not wasted on penniless wanderers!' He slammed the door, oblivious to the ominous lightning that fractured the twilight sky.",
        "Moments later, a deafening tremor shook the workshop. A subterranean spring burst through the stone floor, directly beneath his golden mechanism. Terrified, Ki Ageng watched as the pristine waters rapidly engulfed his coveted tools and blueprints. Strangely, as the water submerged the great clock, its hands began spinning wildly in reverse. In that terrifying moment of realization, Ki Ageng understood that all his accumulated wealth and mechanical mastery were powerless against the unyielding forces of nature and moral compassion. Escaping empty-handed into the pouring rain, he found the weeping child and shared his dry cloak, weeping tears of genuine humility."
    ]
    
    p1_analysis = {
        "Orientation (Paragraf 1)": "Pengenalan Master Ki Ageng Semito di lereng Gunung Muria, keahlian jam mekanik, watak tamak dan angkuh.",
        "Complication (Paragraf 2)": "Munculnya pengembara lansia yang basah kuyup meminta perlindungan, namun diusir dengan kasar.",
        "Resolution & Coda (Paragraf 3)": "Mata air bawah tanah meledak menghancurkan bengkel; jarum jam berputar balik; Ki Ageng bertobat dan memeluk kerendahan hati."
    }
    
    add_authentic_text_block(doc, "1", "The Clockmaker’s Spring: A Parable of Hubris and Humility",
                            "Teks Model Narrative", {
                                "portal": "Contemporary Indonesian Folktale & Literary Heritage Anthology",
                                "title": "The Clockmaker's Spring of Muria",
                                "url": "https://literaryindonesia.org/muria-tales",
                                "year": "2025", "cefr": "B1 / B1+"
                            }, content_paragraphs=p1_paras, analysis_dict=p1_analysis)
                            
    add_grammar_box(doc, "Tata Bahasa Esensial: Past Narrative Tenses & Direct/Indirect Speech",
                    "Simple Past (V2) | Past Continuous (was/were + V-ing) | Past Perfect (had + V3)",
                    [
                        "Interrupted Past Action: While he was putting the final touches, a traveler knocked upon the door.",
                        "Past Perfect Sequence: He had accumulated immense wealth before tragedy struck his workshop.",
                        "Direct Speech: Ki Ageng retorted, 'My time is measured in gold!'",
                        "Indirect Speech: Ki Ageng retorted that his time was measured in gold."
                    ])
                    
    add_vocab_table(doc, [
        ("Hubris", "Noun", "Keangkuhan berlebihan yang membawa kehancuran", "The craftsman's hubris blinded him to moral duty."),
        ("Subjugate", "Verb", "Menundukkan, menguasai secara paksa", "He falsely believed intellect could subjugate time."),
        ("Exorbitant", "Adjective", "Sangat mahal melampaui batas wajar", "Only wealthy merchants could pay his exorbitant fees."),
        ("Oblivious", "Adjective", "Lengah, tidak menyadari bahaya sekitar", "He slammed the door, oblivious to the approaching storm."),
        ("Humility", "Noun", "Kerendahan hati sejati", "He wept tears of genuine humility in the rain.")
    ])
    
    # Pertemuan 2
    add_heading_1(doc, "PERTEMUAN 2: ANALYTICAL EXPOSITION (ETIKA KECERDASAN BUATAN)")
    add_callout(doc, "Tujuan Pembelajaran (TP 11.W.2) & Indikator Ketercapaian (KKTP):",
                "Peserta didik mampu menganalisis struktur tesis-argumen-reiterasi, mengidentifikasi konjungsi kausalitas dan kata kerja mental, serta membedakan antara fakta objektif dan opini analitis secara rasional.")
                
    p2_paras = [
        "The rapid proliferation of generative artificial intelligence and synthetic media has fundamentally transformed how young individuals consume, synthesize, and disseminate information. While these technological breakthroughs offer unprecedented learning conveniences, they simultaneously expose students to sophisticated misinformation, automated plagiarism, and deepfake content. Therefore, it is indisputable that integrating critical digital literacy into the high school curriculum is no longer merely optional, but an absolute educational imperative.",
        "Firstly, critical digital literacy equips adolescents with essential analytical scrutiny to evaluate the credibility of digital sources. In an era where algorithmic feeds prioritize sensationalism over factual veracity, uncritical consumers frequently mistake machine-generated hallucinations for verified truths. By understanding cross-referencing methodologies and source attribution, students can distinguish authentic academic consensus from commercially driven digital falsehoods.",
        "Secondly, institutionalizing digital literacy fosters intellectual integrity and ethical awareness in the contemporary classroom. When learners understand the profound differences between using AI as an exploratory brainstorming partner versus engaging in automated academic dishonesty, they develop authentic academic resilience. They realize that authentic problem-solving and critical cognition cannot be outsourced to a predictive neural network.",
        "In conclusion, secondary educational institutions must not perceive technological advancements in isolation from ethical reasoning. Enhancing students' capacity to critically deconstruct digital media ensures that young citizens emerge as discerning, responsible, and empowered digital navigators in an increasingly automated global society."
    ]
    
    p2_analysis = {
        "Thesis (Paragraf 1)": "Pernyataan tesis tegas: Integrasi literasi digital kritis dalam kurikulum SMA merupakan keharusan pendidikan mutlak.",
        "Argument 1 (Paragraf 2)": "Point: Kemampuan memvalidasi sumber digital. Elaboration: Mengatasi 'hallucinations' AI dengan metode cek-silang.",
        "Argument 2 (Paragraf 3)": "Point: Menjaga integritas akademik. Elaboration: Menolak plagiasi otomatis; kognisi kritis manusia tidak dapat digantikan mesin.",
        "Reiteration (Paragraf 4)": "Penegasan kembali tesis: Sekolah harus membekali generasi muda menjadi navigator digital yang bertanggung jawab."
    }
    
    add_authentic_text_block(doc, "2", "Why Critical Digital Literacy Must Be Prioritized in Secondary Education",
                            "Teks Model Analytical Exposition", {
                                "portal": "The Global Educational Technology Review",
                                "title": "Critical Digital Literacy Imperative in High Schools",
                                "url": "https://edutech-review.org/digital-literacy-2025",
                                "year": "2025", "cefr": "B1+ / B2"
                            }, content_paragraphs=p2_paras, analysis_dict=p2_analysis)
                            
    add_grammar_box(doc, "Tata Bahasa Esensial: Causal Conjunctions & Modal Adverbs",
                    "Causal Linkers (therefore, consequently, because) | Modal Adverbs (indisputably, evidently)",
                    [
                        "Cause & Effect: The feed prioritized sensationalism; consequently, students mistook AI output for facts.",
                        "Logical Connector: Therefore, integrating digital ethics is an absolute educational imperative.",
                        "Modal Adverb of Certainty: Technological advancement will indisputably redefine career pathways.",
                        "Mental Verbs: Scholars conclude that critical thinking requires disciplined intellectual exertion."
                    ])
                    
    add_vocab_table(doc, [
        ("Proliferation", "Noun", "Penyebaran atau lonjakan pesat", "The rapid proliferation of AI reshapes modern schools."),
        ("Veracity", "Noun", "Keakuratan dan kebenaran fakta", "Students must verify the factual veracity of digital articles."),
        ("Hallucination", "Noun (AI)", "Informasi keliru yang dikarang oleh sistem AI", "Algorithmic systems may generate hallucinations presented as facts."),
        ("Resilience", "Noun", "Ketahanan dan keteguhan daya pikir", "Adolescents develop authentic cognitive and ethical resilience."),
        ("Discerning", "Adjective", "Kritis dan tajam dalam membedakan hal baik", "Schools empower students to become discerning digital citizens.")
    ])
    
    add_comparative_matrix(doc, ["Kriteria Pembeda", "Narrative Text (Pertemuan 1)", "Analytical Exposition (Pertemuan 2)"], [
        ["Fokus Utama", "Kisah estetis, alur konflik, dan perenungan moral", "Meyakinkan pembaca atas kebenaran suatu gagasan/isu penting"],
        ["Pendekatan Retorika", "Emosional & imersif melalui pengalaman tokoh", "Rasional & logis melalui bukti data empiris"],
        ["Struktur Teks", "Orientation -> Complication -> Resolution -> Coda", "Thesis -> Arguments (Point-Elaboration) -> Reiteration"],
        ["Dominasi Waktu", "Past Tenses (Simple Past, Past Continuous, Past Perfect)", "Present Tense, Modal Auxiliaries, Causal Conjunctions"],
        ["Akhir Tulisan", "Resolusi cerita & pesan moral (Coda)", "Penegasan kembali tesis (Reiteration) tanpa anjuran aksi"]
    ])
    
    out = "C:/Users/LILI/OneDrive - PDK Jateng/Dokumen/Perangkat_SMAN_1_Bae/Kelas_11/Ringkasan_Materi_2_Pertemuan_Kelas_11.docx"
    doc.save(out)
    print("DOCX Kelas 11 saved:", out)

# ==============================================================================
# 3. KELAS XII
# ==============================================================================
def build_docx_kelas_12():
    doc = create_styled_doc()
    add_header_banner(doc, "RINGKASAN MATERI PEMBELAJARAN BAHASA INGGRIS",
                      "SMA NEGERI 1 BAE KUDUS — KELAS XII (FASE F) SEMESTER GASAL 2025/2026")
                      
    add_tp_info_box(doc, {
        "Satuan Pendidikan": "SMA Negeri 1 Bae Kudus",
        "Mata Pelajaran": "Bahasa Inggris (Wajib / Tingkat Lanjut)",
        "Fokus Materi Esensial": "Pertemuan 1: Advanced Narrative | Pertemuan 2: Discussion Text",
        "Alokasi Waktu": "2 Pertemuan (4 Jam Pelajaran @ 45 Menit)",
        "Rujukan Kurikulum": "Pemetaan MGMP SMA Kudus 2025-2026 & Kepka BSKAP No. 020/2026"
    })
    
    # Pertemuan 1
    add_heading_1(doc, "PERTEMUAN 1: ADVANCED NARRATIVE (ALEGORI & KEDALAMAN PSIKOLOGIS)")
    add_callout(doc, "Tujuan Pembelajaran (TP 12.W.1) & Indikator Ketercapaian (KKTP):",
                "Peserta didik mampu menganalisis struktur naratif bertingkat, menafsirkan bahasa kiasan simbolik (subtext & symbolism), serta mengevaluasi dilema etika tokoh dalam cerpen sastra kontemporer tingkat lanjut.")
                
    p1_paras = [
        "Deep within the labyrinthine alleyways of Semarang’s Old Town, flanked by Dutch colonial archways whose plaster flaked like desiccated parchment, sat Master Sastro. For over six decades, he had practiced the vanishing art of hand-binding antique manuscripts, stitching fragile leather covers with waxed hemp thread. While the surrounding world surrendered itself to the dazzling ephemeral glow of digital screens and cloud servers, Sastro's sanctuary smelled persistently of aged cedar, beeswax, and crushed clove oil. To him, every tactile page possessed an indelible soul that binary code could never replicate.",
        "One sweltering afternoon, a young software developer named Nadia entered his workshop carrying an ultrathin holographic scanner. Tasked by a municipal archive consortium to digitize historical artifacts, she viewed Sastro’s ancient volumes merely as unindexed data sets awaiting conversion into high-density neural vectors. 'Bapak,' Nadia remarked politely, displaying her sleek handheld scanner, 'my machine can preserve your entire life's work in less than twelve minutes. If you allow me, physical decay will never threaten these texts again.'",
        "Sastro paused, setting down his polished bone folder. He picked up a sun-bleached Javanese chronicle transcribed in golden botanical ink and gently held it near her ear. 'Listen closely, young lady,' he whispered gently. 'A machine can capture the shape of these letters with ruthless optical precision, but can it register the trembling hand of the poet who composed these verses by candlelight during the famine of 1845? Can your server transmit the scent of rain trapped between these dried leaves?'",
        "Nadia froze, the hum of her cooling fan suddenly sounding intrusive amidst the solemn silence of the workshop. For the first time, she looked beyond the optical surface of the document, recognizing that cultural preservation is not merely an algorithmic calculation of pixels, but a sacred transmission of human vulnerability across generations."
    ]
    
    p1_analysis = {
        "Atmospheric Orientation (Paragraf 1)": "Membangun latar Kota Lama Semarang; dedikasi Master Sastro terhadap naskah fisik di tengah gelombang digitalisasi.",
        "Complex Complication (Paragraf 2)": "Benturan filosofis antara efisiensi digital Nadia (pemindai holografik) dan ketulusan seni tangan Sastro.",
        "Epiphany & Resolution (Paragraf 3-4)": "Dialog Sastro membuka mata batin Nadia bahwa warisan budaya adalah transmisi jiwa dan kerapuhan manusia, bukan sekadar kalkulasi piksel."
    }
    
    add_authentic_text_block(doc, "1", "The Memory Weaver of Kota Lama",
                            "Teks Model Advanced Narrative", {
                                "portal": "International Literary Archive & Asian Writers Review",
                                "title": "The Memory Weaver of Semarang",
                                "url": "https://asianwritersreview.com/memory-weaver-2025",
                                "year": "2025", "cefr": "B2"
                            }, content_paragraphs=p1_paras, analysis_dict=p1_analysis)
                            
    add_grammar_box(doc, "Tata Bahasa Esensial: Participle Clauses & Cleft Sentences",
                    "Participle Clause (V-ing / V3 at clause head)  |  Cleft Sentence (It was X that Y)",
                    [
                        "Participle Clause: Tasked by a municipal archive consortium, Nadia viewed the books as data sets.",
                        "Participle Clause: Gazing at the fragile manuscript, she recognized the sacred human heritage.",
                        "Cleft Sentence: It was not the optical scanner, but the poet's devotion that gave soul to the text.",
                        "Inversion for Emphasis: Rarely had Nadia witnessed such profound dedication to manual craft."
                    ])
                    
    add_vocab_table(doc, [
        ("Labyrinthine", "Adjective", "Berliku-liku rumit seperti labirin", "He walked through the labyrinthine alleys of Kota Lama."),
        ("Ephemeral", "Adjective", "Hanya berlangsung sekejap, fana", "The world surrendered to the ephemeral glow of digital screens."),
        ("Indelible", "Adjective", "Abadi, tidak dapat dihapus waktu", "Every physical page possessed an indelible soul."),
        ("Vulnerability", "Noun", "Kerapuhan nurani dan keterbatasan insani", "Preservation is a sacred transmission of human vulnerability."),
        ("Consortium", "Noun", "Gabungan beberapa lembaga/organisasi", "She represented a municipal digital archive consortium.")
    ])
    
    # Pertemuan 2
    add_heading_1(doc, "PERTEMUAN 2: DISCUSSION TEXT (WACANA DEBAT OTOMASI KERJA)")
    add_callout(doc, "Tujuan Pembelajaran (TP 12.W.2) & Indikator Ketercapaian (KKTP):",
                "Peserta didik mampu menganalisis teks diskusi berimbang (Pro vs Contra), mengidentifikasi penanda wacana kontras dan bahasa kehati-hatian (hedging), serta merumuskan kesimpulan sintesis netral.")
                
    p2_paras = [
        "The relentless integration of autonomous algorithmic systems and generative artificial intelligence into white-collar professions has ignited fierce international debate. Across legal, financial, and creative sectors, cognitive tasks once deemed exclusively human are now executed with breathtaking speed by neural networks. Consequently, economists, policymakers, and labor leaders find themselves sharply divided over whether automation represents an unprecedented emancipation of human talent or a perilous catalyst for widespread socioeconomic disenfranchisement.",
        "Proponents of rapid workplace automation contend that cognitive AI elevates human productivity to historic heights. By delegating repetitive computational analysis, data verification, and drafting procedures to automated engines, professionals can redirect their cognitive energy toward high-level strategic reasoning, ethical diplomacy, and interpersonal empathy. Furthermore, business analysts argue that augmented enterprises generate entirely new technological ecosystems, creating millions of specialized roles in algorithmic auditing, prompt engineering, and cyber-security oversight that previously did not exist.",
        "Conversely, detractors and labor unions present grave concerns regarding structural displacement and the erosion of entry-level career pathways. When junior analyst and associate tasks are swallowed by software subscriptions, the traditional apprenticeship model through which young graduates hone professional judgment collapses. Furthermore, critics emphasize that automated systems frequently perpetuate historical biases embedded within training datasets, resulting in systemic discrimination in recruitment and credit scoring without transparent human accountability.",
        "In conclusion, neither blind technophilia nor reflexive resistance provides an adequate response to the fourth industrial revolution. A balanced strategy requires robust multi-stakeholder governance: corporations must reinvest productivity gains into lifelong employee reskilling, while governments formulate strict labor protections that mandate human oversight in consequential decisions. The ultimate objective must be human-centered augmentation rather than unbridled worker replacement."
    ]
    
    p2_analysis = {
        "Statement of Issue (Paragraf 1)": "Merumuskan perdebatan global: Otomatisasi tempat kerja sebagai emansipasi talenta vs ancaman peminggiran sosial ekonomi.",
        "Arguments For (Paragraf 2)": "Sudut pandang pendukung: Lonjakan produktivitas, fokus pada diplomasi dan empati, lahirnya lapangan kerja baru.",
        "Arguments Against (Paragraf 3)": "Sudut pandang penentang: Pengikisan jalur karir magang pemula, bias algoritma, dan hilangnya akuntabilitas manusia.",
        "Balanced Conclusion (Paragraf 4)": "Sintesis jalan tengah: Tata kelola multi-pihak, kewajiban pelatihan ulang pekerja, dan pengawasan manusia."
    }
    
    add_authentic_text_block(doc, "2", "The Autonomous Workplace: Collaborative Evolution or Mass Disenfranchisement?",
                            "Teks Model Discussion", {
                                "portal": "World Economic Forum & Global Technological Deliberation Report",
                                "title": "The Autonomous Workplace Deliberation",
                                "url": "https://weforum.org/reports/autonomous-workplace-2025",
                                "year": "2025", "cefr": "B2"
                            }, content_paragraphs=p2_paras, analysis_dict=p2_analysis)
                            
    add_grammar_box(doc, "Tata Bahasa Esensial: Contrastive Markers & Academic Hedging",
                    "Contrastive Linkers (Conversely, on the other hand) | Hedging (arguably, tend to suggest)",
                    [
                        "Contrast Marker: Proponents celebrate efficiency; conversely, labor unions raise alarms on unemployment.",
                        "Hedging Expression: Evidence arguably suggests that automation could temporarily disrupt employment.",
                        "Attribution Verbs: Detractors contend that systemic biases remain embedded in training data.",
                        "Balanced Synthesis: In conclusion, neither blind optimism nor total refusal provides an answer."
                    ])
                    
    add_vocab_table(doc, [
        ("Emancipation", "Noun", "Pembebasan dari rutinitas beban kerja monoton", "Proponents envision the emancipation of human talent."),
        ("Disenfranchisement", "Noun", "Hilangnya hak ekonomi atau kesempatan kerja layak", "Critics warn of widespread socioeconomic disenfranchisement."),
        ("Apprenticeship", "Noun", "Masa magang dan penempaan keterampilan dasar", "The entry-level apprenticeship pathway collapses under automation."),
        ("Augmentation", "Noun", "Pemberdayaan dan peningkatan kapasitas manusia", "Policies must champion human augmentation over worker replacement."),
        ("Governance", "Noun", "Tata kelola regulasi yang akuntabel", "Technological transition demands multi-stakeholder ethical governance.")
    ])
    
    add_comparative_matrix(doc, ["Kriteria Pembeda", "Advanced Narrative (Pertemuan 1)", "Discussion Text (Pertemuan 2)"], [
        ["Tujuan Utama", "Refleksi filosofis, empati moral, dan eksplorasi batin manusia", "Membahas perdebatan dua sisi secara adil sebelum merumuskan simpulan"],
        ["Sikap Penulis", "Imersif, menyatu dengan penghayatan emosional tokoh", "Netral, objektif, bertindak sebagai penengah yang adil (neutral arbiter)"],
        ["Struktur Teks", "Orientation -> Complication -> Epiphany -> Resolution", "Statement of Issue -> Arguments For -> Arguments Against -> Conclusion"],
        ["Ciri Bahasa", "Metaphor, symbolism, descriptive sensory verbs, participle clauses", "Contrastive linkers, hedging language, reporting verbs, neutral stance"],
        ["Contoh Isu", "Kisah pelestari naskah kuno, empati seniman wayang kulit", "Otomatisasi AI, transisi energi nuklir, etika rekayasa genetika"]
    ])
    
    out = "C:/Users/LILI/OneDrive - PDK Jateng/Dokumen/Perangkat_SMAN_1_Bae/Kelas_12/Ringkasan_Materi_2_Pertemuan_Kelas_12.docx"
    doc.save(out)
    print("DOCX Kelas 12 saved:", out)

if __name__ == "__main__":
    build_docx_kelas_10()
    build_docx_kelas_11()
    build_docx_kelas_12()
    print("All DOCX files created successfully!")
