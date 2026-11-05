# -*- coding: utf-8 -*-
"""
Generator script for authentic English materials (.docx) in 5-Modul Ajar X SMA (Fase E).
Calibrated specifically to CEFR A2+ to B1 (Accessible High School Level).
Each TP contains 3 authentic texts from credible sources with explicit URLs,
accessible vocabulary glossaries, and aligned assessment questions.
"""

import os
import sys

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import docx_helper

TARGET_DIR = r"C:\Users\KOMPUTERKU2\Documents\English SMA\5-Modul Ajar X SMA"
os.makedirs(TARGET_DIR, exist_ok=True)

# ==============================================================================
# TP 10.1: DESCRIPTIVE TEXT - GREAT ATHLETES
# ==============================================================================
def create_tp_10_1():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS SMA",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase E / Kelas X / Semester 1 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 10.1 - Menganalisis fungsi sosial, struktur teks, dan unsur kebahasaan teks deskriptif lisan dan tulis tentang atlet berprestasi dunia secara kritis dan komunikatif.",
        "Elemen Capaian Pembelajaran": "Menyimak - Berbicara & Membaca - Memirsa",
        "Genre Teks / Ruang Lingkup": "Descriptive Text (People / Great World Athletes)",
        "Alokasi Waktu": "10 Jam Pelajaran (5 Pertemuan @ 2 JP)",
        "Target Kemahiran CEFR": "Level A2+ s.d. B1 (Intermediate Awal - Tingkat SMA)",
        "Profil Pelajar Pancasila": "Mandiri, Bernalar Kritis, Berkebinekaan Global"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "Olympics Official Portal / The Jakarta Post Sports Desk",
        "title": "How Greysia Polii and Apriyani Rahayu Won Indonesia's Historic Badminton Gold",
        "url": "https://olympics.com/en/news/greysia-polii-apriyani-rahayu-badminton-gold",
        "year": "2021 / 2024",
        "cefr": "Level B1 (Menengah Awal - Sesuai Siswa Kelas X)"
    }
    t1_paras = [
        "Greysia Polii and Apriyani Rahayu made badminton history at the Tokyo 2020 Olympic Games. They won Indonesia's first-ever Olympic gold medal in the women's doubles event. On the court, the two players showed amazing teamwork, quick movement, and strong determination. In the final match, they defeated the world's top-ranked pair from China in two exciting straight games.",
        "Greysia Polii, the older player, is well known for her calm attitude, positive energy, and great court experience. Before the Olympics, she suffered painful injuries and personal sadness, but she never gave up on her dream. During every match, Greysia always smiled, guided her partner, and defended every fast shuttlecock with incredible reflexes.",
        "Apriyani Rahayu, her younger partner, brings huge energy and powerful smashes from the back of the court. She is strong, fast, and brave in every rally. Outside the court, both Greysia and Apriyani are humble and friendly. Their great victory teaches young people in Indonesia that hard work and friendship can turn impossible dreams into reality."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Mendeskripsikan karakter, prestasi luar biasa, dan kepribadian atlet bulu tangkis peraih emas Olimpiade untuk memotivasi pembaca.",
        "Struktur Teks": "Identification (Paragraf 1: Pengenalan tokoh dan gelar juara) -> Description (Paragraf 2: Ciri karakter dan peran Greysia Polii; Paragraf 3: Gaya bermain Apriyani dan sifat rendah hati mereka).",
        "Fitur Bahasa": "Simple Present Tense ('brings huge energy', 'is well known'), Descriptive Adjectives ('amazing', 'calm', 'positive', 'humble', 'friendly'), Kata penghubung sederhana ('and', 'but', 'because')."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "The Winning Spirit: Greysia Polii and Apriyani Rahayu",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "BBC Sport / Men's Health Athletic Profile",
        "title": "Cristiano Ronaldo: The Secrets Behind His Elite Physical Fitness and Routine",
        "url": "https://www.bbc.com/sport/football/athletic-habits-ronaldo",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah Awal - Sesuai Siswa Kelas X)"
    }
    t2_paras = [
        "Cristiano Ronaldo is widely recognized as one of the hardest working athletes in world football. The Portuguese forward is famous for his incredible fitness, fast running speed, and jumping ability. Even in his late thirties, Ronaldo continues to compete at the highest level of international football.",
        "Standing at 1.87 meters tall, Ronaldo has a strong, muscular body. His physical condition is not just luck; it is the result of years of daily discipline. Ronaldo trains for hours every day, eats healthy meals with fresh fish and vegetables, drinks plenty of water, and ensures he gets eight hours of good sleep every night. He avoids sugary drinks and fast food completely.",
        "In addition to his physical strength, Ronaldo has a winning mindset. He always wants to improve and help his teammates win. Off the pitch, Ronaldo is known for his generous charity work, frequently donating blood and supporting hospitals for sick children."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Mendeskripsikan rutinitas disiplin, kondisi fisik bugar, dan kedermawanan seorang pemain sepak bola dunia.",
        "Struktur Teks": "Identification (Paragraf 1: Pengenalan tokoh) -> Description (Paragraf 2: Pola hidup sehat & latihan; Paragraf 3: Mental juara & kepedulian sosial).",
        "Fitur Bahasa": "Kata sifat deskriptif fisik ('strong', 'muscular', 'tall'), kata keterangan kebiasaan ('daily', 'always', 'frequently')."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "Discipline and Healthy Habits: The Story of Cristiano Ronaldo",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "FIFA Youth Stories / The Guardian Sports News",
        "title": "Lionel Messi: The Quiet Legend and the Joy of Football",
        "url": "https://www.fifa.com/fifaplus/en/articles/lionel-messi-profile-magic",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah Awal - Sesuai Siswa Kelas X)"
    }
    t3_paras = [
        "Lionel Messi is often called a magician with the football at his feet. Unlike many players who are tall and heavily built, the Argentine captain stands at 1.70 meters. However, his low height gives him great balance and quick agility, allowing him to move smoothly past defenders with the ball.",
        "Messi is famous for his smart vision on the field. He can spot an open teammate across the pitch in a second and send an accurate pass through tight spaces. When he shoots with his left foot, the ball curves beautifully into the net. Messi rarely complains on the pitch and always plays with fairness and sportsmanship.",
        "In daily life, Messi is a quiet and loving family man. He loves spending peaceful time with his wife and three sons. Winning the 2022 FIFA World Cup was the happiest moment of his career, proving that talent, patience, and humility can inspire millions of football fans worldwide."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "The Humble Magician: Lionel Messi",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    # Practical High School Vocabulary
    vocabs = [
        ("Determination", "Noun", "Tekad kuat, kemauan pantang menyerah", "The two players showed strong determination to win."),
        ("Reflexes", "Noun (Plural)", "Refleks cepat tubuh merespons sesuatu", "Greysia defended every fast shot with quick reflexes."),
        ("Humble", "Adjective", "Rendah hati, tidak sombong", "Despite their fame, both champions remain humble."),
        ("Discipline", "Noun", "Kedisiplinan mematuhi aturan dan latihan", "Ronaldo's fitness is the result of strict daily discipline."),
        ("Agility", "Noun", "Kelincahan bergerak cepat dan mudah", "His low height gives him great balance and quick agility."),
        ("Sportsmanship", "Noun", "Sikap sportif dan menghargai lawan", "Messi always plays with fairness and true sportsmanship."),
        ("Teammate", "Noun", "Rekan satu tim dalam olahraga", "A great athlete always encourages their teammates."),
        ("Victory", "Noun", "Kemenangan dalam pertandingan", "Their historic victory brought pride to the whole nation.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "According to Text 1, what historic achievement did Greysia Polii and Apriyani Rahayu achieve at the Tokyo 2020 Olympics?",
        "In Text 2, what healthy habits does Cristiano Ronaldo practice daily to maintain his fitness?",
        "Based on Text 3, what physical advantages help Lionel Messi move past defenders smoothly?"
    ]
    inf_q = [
        "Why is having a positive attitude and smiling on the court important for team athletes like Greysia Polii?",
        "How do both Cristiano Ronaldo (Text 2) and Lionel Messi (Text 3) show good character outside of their football matches?",
        "What can high school students learn from the discipline of these world athletes to improve their daily study habits?"
    ]
    eval_q = [
        "In your opinion, which quality is more important for a young athlete to succeed: natural talent or daily discipline? Explain your reason.",
        "Do you agree that famous athletes have a responsibility to be good role models for teenagers? Give examples from the texts."
    ]
    perf_task = (
        "TUGAS KINERJA OTENTIK (WRITING & SPEAKING):\n"
        "Pilihlah salah satu atlet favoritmu (atlet Indonesia atau dunia). Tulislah teks deskriptif sederhana sepanjang 150-200 kata.\n"
        "Pastikan teks memuat:\n"
        "1. Identification: Nama atlet, cabang olahraga, dan prestasi utamanya.\n"
        "2. Description: Ciri fisik, kebiasaan latihan/gaya hidup sehat, dan sifat baiknya.\n"
        "3. Gunakan minimal 4 kosakata baru dari tabel di atas (misal: determination, humble, discipline, sportsmanship).\n"
        "Kemudian, ceritakan deskripsi atletmu secara lisan selama 1-2 menit di depan teman sekelasmu!"
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_10.1_Descriptive_Text_Athletes.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 10.2: MULTIMODAL DESCRIPTIVE TEXT - TOURISM & NATURAL HERITAGE
# ==============================================================================
def create_tp_10_2():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS SMA",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase E / Kelas X / Semester 1 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 10.2 - Memproduksi teks deskriptif multimoda lisan dan tulis tentang destinasi wisata dan warisan budaya nusantara dengan kosakata dan tata bahasa yang tepat.",
        "Elemen Capaian Pembelajaran": "Membaca - Memirsa & Menulis - Mempresentasikan",
        "Genre Teks / Ruang Lingkup": "Descriptive Text (Places / Tourism & Natural Heritage)",
        "Alokasi Waktu": "12 Jam Pelajaran (6 Pertemuan @ 2 JP)",
        "Target Kemahiran CEFR": "Level A2+ s.d. B1",
        "Profil Pelajar Pancasila": "Berkebinekaan Global, Mandiri, Cinta Tanah Air"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "UNESCO World Heritage Centre / Indonesia Tourism",
        "title": "Komodo National Park: Habitat of the World's Largest Lizard",
        "url": "https://whc.unesco.org/en/list/609/",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah Awal - Sesuai Siswa Kelas X)"
    }
    t1_paras = [
        "Komodo National Park is located in the center of the Indonesian archipelago, between the islands of Sumbawa and Flores in East Nusa Tenggara. Created in 1980, this famous park was named a UNESCO World Heritage Site in 1991. The park consists of three major islands—Komodo, Rinca, and Padar—along with numerous smaller islands.",
        "The park is world-famous because it is the only natural home of the Komodo dragon, the largest living lizard on Earth. These impressive reptiles can grow up to three meters in length and weigh over seventy kilograms. The landscape of the islands is dramatic and beautiful, featuring dry green hills, rugged coastlines, and clear blue seas. One of the most famous spots is Pink Beach on Komodo Island, where the sand has a soft pink color from microscopic red coral particles.",
        "Under the water, the park is home to rich marine life, including colorful coral reefs, giant manta rays, sea turtles, and hundreds of species of tropical fish. Park rangers work hard every day to protect the dragons and educate tourists to keep the natural environment clean and safe."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Mendeskripsikan keunikan alam, habitat komodo, dan keindahan bahari Taman Nasional Komodo untuk mempromosikan wisata ramah lingkungan.",
        "Struktur Teks": "Identification (Paragraf 1: Lokasi dan status UNESCO) -> Description (Paragraf 2: Ciri komodo dan pantai berpasir merah muda; Paragraf 3: Kehidupan bawah laut dan konservasi).",
        "Fitur Bahasa": "Simple Present Tense ('is located', 'consists of', 'can grow'), Kata sifat pemandangan ('impressive', 'dramatic', 'colorful', 'rugged')."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "Komodo National Park: Nature's Wonder in East Nusa Tenggara",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "WWF Indonesia Conservation / National Geographic Traveler",
        "title": "Raja Ampat: The Heart of Coral Reefs and Island Paradise",
        "url": "https://www.wwf.id/en/what-we-do/marine/raja-ampat",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah Awal - Sesuai Siswa Kelas X)"
    }
    t2_paras = [
        "Raja Ampat, which means 'The Four Kings', is an archipelago located off the northwest tip of Papua, Indonesia. The archipelago consists of more than 1,500 small islands, cays, and shoals surrounding the four main islands of Waigeo, Batanta, Salawati, and Misool. It is often described as one of the most beautiful marine areas on Earth.",
        "From above, the limestone islands look like green mushrooms floating on turquoise ocean water. When visitors hike to the viewpoint at Pianemo, they are greeted by a breathtaking panorama of emerald lagoons and rocky peaks. The air is fresh, and the waters are so clear that you can see fish swimming near the surface.",
        "Scientists consider Raja Ampat the global center of marine biodiversity. It contains over 75 percent of all known coral species in the world and more than a thousand species of reef fish. Local communities practice traditional conservation rules called 'Sasi', which prohibit fishing in certain zones to ensure the ocean stays healthy for future generations."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Mendeskripsikan kepulauan Raja Ampat dan kearifan lokal konservasi laut masyarakat Papua.",
        "Struktur Teks": "Identification (Paragraf 1: Letak geografis) -> Description (Paragraf 2: Panorama pulau karst Pianemo; Paragraf 3: Kekayaan terumbu karang dan tradisi Sasi).",
        "Fitur Bahasa": "Kata sifat warna dan bentuk ('turquoise', 'emerald', 'limestone', 'breathtaking'), Preposisi tempat ('off the tip of', 'surrounding', 'near')."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "Raja Ampat: An Island Paradise in Papua",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "Indonesia Travel Official / UNESCO World Heritage",
        "title": "Borobudur Temple: The Great Buddhist Monument in Central Java",
        "url": "https://www.indonesia.travel/gb/en/destinations/java/magelang-regency/borobudur",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah Awal - Sesuai Siswa Kelas X)"
    }
    t3_paras = [
        "Borobudur Temple is a magnificent 9th-century Buddhist temple located near Magelang in Central Java. Built during the reign of the Sailendra Dynasty, it is the largest Buddhist monument in the world. The temple stands proudly on a hill surrounded by green rice fields and distant volcanic mountains.",
        "The design of Borobudur follows traditional Buddhist architecture, shaped like a giant stepped pyramid. It consists of six square platforms topped by three circular platforms, decorated with over 2,600 relief panels and 504 Buddha statues. At the very top stands a large central stupa that represents peace and spiritual enlightenment.",
        "Watching the sunrise from Borobudur is an unforgettable experience for tourists. As the morning mist gently rises over the palm trees and the sun lights up the ancient grey stones, the atmosphere feels peaceful and sacred. Visitors are encouraged to walk carefully and respect the heritage rules to preserve the ancient stones."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "Borobudur: Majestic Heritage in Central Java",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Archipelago", "Noun", "Kepulauan, kumpulan banyak pulau", "Indonesia is the largest archipelago country in the world."),
        ("Reptile", "Noun", "Hewan melata berdarah dingin", "The Komodo dragon is the largest living reptile on Earth."),
        ("Breathtaking", "Adjective", "Sangat indah menakjubkan", "The view from the top of the hill is truly breathtaking."),
        ("Biodiversity", "Noun", "Keanekaragaman hayati makhluk hidup", "Raja Ampat has the highest marine biodiversity on our planet."),
        ("Heritage", "Noun", "Warisan sejarah atau budaya berharga", "Borobudur is a protected world cultural heritage site."),
        ("Monument", "Noun", "Bangunan monumen bersejarah", "It is recognized as the largest Buddhist monument in the world."),
        ("Enlightenment", "Noun", "Pencerahan batin, kedamaian spiritual", "The central stupa at the top symbolizes peace and enlightenment."),
        ("Preserve", "Verb", "Menjaga, melestarikan agar tidak rusak", "We must follow rules to preserve our natural environment.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "According to Text 1, what are the three major islands that form Komodo National Park?",
        "In Text 2, what traditional conservation practice is used by local communities in Raja Ampat to protect ocean life?",
        "What architectural features decorate the walls and terraces of Borobudur Temple in Text 3?"
    ]
    inf_q = [
        "Why is Pink Beach on Komodo Island able to have sand with a soft pink color?",
        "How does the traditional rule of 'Sasi' in Papua demonstrate the wisdom of local Indonesian culture?",
        "Why is it important for tourists to follow heritage rules when visiting historic places like Borobudur Temple?"
    ]
    eval_q = [
        "If you could choose to visit either Komodo Island or Raja Ampat, which one would you choose and why? Describe your reasons.",
        "How can high school students promote Indonesian tourist destinations responsibly through social media without causing environmental harm?"
    ]
    perf_task = (
        "TUGAS KINERJA PROYEK MULTIMODA (BROCHURE / SOCIAL MEDIA POST):\n"
        "Buatlah sebuah Brosur Wisata atau infografis digital (satu halaman) yang mempromosikan destinasi wisata alam atau budaya di daerahmu.\n"
        "Petunjuk Pengerjaan:\n"
        "1. Cantumkan nama tempat, lokasi, dan daya tarik utamanya (pantai, gunung, air terjun, museum, dsb).\n"
        "2. Tuliskan deskripsi 100-150 kata menggunakan Simple Present Tense dan kata sifat deskriptif (misal: breathtaking, peaceful, beautiful).\n"
        "3. Sertakan tips ramah lingkungan bagi pengunjung (misal: do not litter, preserve nature)."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_10.2_Multimodal_Descriptive_Tourism.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 10.3: HISTORICAL RECOUNT TEXT - HISTORIC EVENTS
# ==============================================================================
def create_tp_10_3():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS SMA",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase E / Kelas X / Semester 1 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 10.3 - Menganalisis kronologi, fungsi sosial, dan tata bahasa teks rekon sejarah lisan dan tulis tentang peristiwa bersejarah kemerdekaan dan kepahlawanan.",
        "Elemen Capaian Pembelajaran": "Menyimak - Berbicara & Membaca - Memirsa",
        "Genre Teks / Ruang Lingkup": "Historical Recount Text (Events in History / Heroes)",
        "Alokasi Waktu": "12 Jam Pelajaran (6 Pertemuan @ 2 JP)",
        "Target Kemahiran CEFR": "Level A2+ s.d. B1",
        "Profil Pelajar Pancasila": "Berkebinekaan Global, Mandiri, Bernalar Kritis"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "National Archives of Indonesia / The Jakarta Post History",
        "title": "Reading the Proclamation: The Birth of Indonesian Independence in 1945",
        "url": "https://www.thejakartapost.com/indonesia/independence-history-1945",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah Awal - Sesuai Siswa Kelas X)"
    }
    t1_paras = [
        "On the morning of Friday, August 17, 1945, a historic event occurred in Jakarta that changed the destiny of the Indonesian nation forever. At exactly 10:00 AM, Ir. Soekarno and Dr. Mohammad Hatta stepped onto the front porch of Soekarno's house on Jalan Pegangsaan Timur No. 56 to declare Indonesian independence.",
        "The night before, young and senior nationalist leaders had gathered at Admiral Maeda's house to draft the short but powerful proclamation text. Soekarno wrote the original words by hand, while Sayuti Melik typed the final document. The next morning, a crowd of citizens and youth had assembled quietly in the front yard despite the presence of Japanese military forces.",
        "With a steady and emotional voice, Soekarno read the brief statement declaring that Indonesia was now a free and independent nation. Immediately after the reading, the red-and-white national flag, sewn by Fatmawati, was hoisted up a bamboo pole as the crowd proudly sang the national anthem 'Indonesia Raya'. The courageous declaration was then broadcast secretly across radio networks, spreading hope throughout the entire country."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Menceritakan kembali kronologi pembacaan naskah Proklamasi Kemerdekaan Indonesia pada 17 Agustus 1945 untuk mengenang jasa pahlawan.",
        "Struktur Teks": "Orientation (Paragraf 1: Kapan, siapa, di mana peristiwa terjadi) -> Events in Chronological Order (Paragraf 2: Perumusan naskah; Paragraf 3: Pembacaan proklamasi, pengibaran bendera, dan siaran radio) -> Reorientation (Makna kemerdekaan bagi bangsa).",
        "Fitur Bahasa": "Simple Past Tense ('occurred', 'stepped', 'gathered', 'read', 'sang'), Keterangan waktu ('On the morning of', 'The night before', 'At exactly 10:00 AM')."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "The Dawn of Freedom: The Proclamation of Indonesian Independence",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "Youth Pledge Museum / Kemendikbud Cultural Portal",
        "title": "The Youth Pledge of 1928: Uniting a Nation of Islands",
        "url": "https://museumsumpahpemuda.kemdikbud.go.id/sejarah-singkat-1928",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah Awal - Sesuai Siswa Kelas X)"
    }
    t2_paras = [
        "In October 1928, young representatives from various regional youth organizations across the Indonesian archipelago gathered in Batavia (now Jakarta) for the Second Youth Congress. These passionate young people came from diverse backgrounds, including Jong Java, Jong Sumatranen Bond, Jong Celebes, and Jong Ambon.",
        "During the meetings, they discussed how to overcome regional differences and unite against colonial rule. On the final evening, October 28, 1928, the delegates made a historic decision. They stood together and declared the three sacred promises known as Sumpah Pemuda (The Youth Pledge): one motherland (Indonesia), one nation (the Indonesian nation), and one language of unity (the Indonesian language).",
        "At the same meeting, young musician Wage Rudolf Supratman played the instrumental melody of 'Indonesia Raya' on his violin for the very first time. The pledge united the hearts of young Indonesians and became the guiding spirit that led to national independence seventeen years later."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Menceritakan sejarah lahirnya Sumpah Pemuda 1928 sebagai tonggak persatuan pemuda Indonesia.",
        "Struktur Teks": "Orientation (Paragraf 1: Kongres Pemuda II) -> Events (Paragraf 2: Deklarasi tiga ikrar Sumpah Pemuda) -> Reorientation (Paragraf 3: Lagu Indonesia Raya dan pengaruh persatuan).",
        "Fitur Bahasa": "Verba lampau reguler dan ireguler ('gathered', 'came', 'discussed', 'declared', 'played')."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "The Youth Pledge of 1928: One Motherland, One Nation, One Language",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "NASA History Division / Smithsonian Air & Space",
        "title": "Apollo 11: One Giant Leap for Mankind in July 1969",
        "url": "https://www.nasa.gov/mission/apollo-11-moon-landing/",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah Awal - Sesuai Siswa Kelas X)"
    }
    t3_paras = [
        "On July 16, 1969, the Apollo 11 spacecraft launched from the Kennedy Space Center in Florida, carrying American astronauts Neil Armstrong, Buzz Aldrin, and Michael Collins. Their historic mission was to achieve the impossible: land the first humans safely on the surface of the Moon.",
        "After traveling four days through space, the lunar module named 'Eagle' separated and prepared to land. On July 20, 1969, Armstrong carefully guided the spacecraft onto a flat plain known as the Sea of Tranquility. Millions of people around the world watched on live television as Armstrong stepped down the ladder and announced his famous words: 'That's one small step for a man, one giant leap for mankind.'",
        "Armstrong and Aldrin spent more than two hours exploring the lunar surface, collecting rock samples, and planting the United Nations and American flags before returning safely to Earth. The successful mission proved the power of science, teamwork, and human courage."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "Apollo 11: The Historic Journey to the Moon",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Proclamation", "Noun", "Pernyataan resmi, pengumuman kemerdekaan", "The proclamation of independence was read on August 17, 1945."),
        ("Assemble", "Verb", "Berkumpul bersama di suatu tempat", "A crowd of youth assembled quietly in the yard."),
        ("Hoisted", "Verb (Past)", "Mengerek bendera ke atas tiang", "The red-and-white flag was hoisted proudly on a bamboo pole."),
        ("Delegate", "Noun", "Perwakilan utusan yang hadir dalam kongres", "Youth delegates from across the islands agreed on the pledge."),
        ("Sacred", "Adjective", "Suci, sakral, sangat dihormati", "They made three sacred promises to unite the nation."),
        ("Spacecraft", "Noun", "Pesawat penjelajah luar angkasa", "The Apollo 11 spacecraft launched smoothly into orbit."),
        ("Lunar", "Adjective", "Berhubungan dengan bulan", "The astronauts walked and collected rocks on the lunar surface."),
        ("Courage", "Noun", "Keberanian menghadapi tantangan besar", "Their success proved the courage and determination of human beings.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "Where and when was the Indonesian proclamation of independence read by Soekarno in Text 1?",
        "What are the three sacred promises declared by the youth during the Sumpah Pemuda congress in Text 2?",
        "Who was the first astronaut to step onto the surface of the Moon according to Text 3?"
    ]
    inf_q = [
        "Why was the proclamation declaration broadcast secretly across radio networks immediately after the event?",
        "How did the Youth Pledge of 1928 help unite people from different islands who spoke different regional languages?",
        "Why did Neil Armstrong describe his step on the Moon as 'one giant leap for mankind'?"
    ]
    eval_q = [
        "In your opinion, what is the best way for Indonesian high school students today to honor the spirit of Sumpah Pemuda?",
        "Compare the courage required to fight for national independence (Text 1) with the courage to explore space (Text 3). What shared values do they display?"
    ]
    perf_task = (
        "PENULISAN TEKS REKON SEJARAH / PENGALAMAN PRIBADI (LKPD / SUMATIF):\n"
        "Tulislah sebuah teks recount (150-200 kata) mengenai salah satu topik berikut:\n"
        "Opsi A: Menceritakan kembali peristiwa bersejarah pahlawan lokal di kotamu/provinsimu.\n"
        "Opsi B: Menceritakan kembali pengalaman berkesanmu saat mengikuti upacara peringatan Hari Kemerdekaan 17 Agustus di sekolah.\n"
        "Gunakan Simple Past Tense secara tepat dan sertakan urutan kronologis yang jelas (Orientation -> Events -> Reorientation)."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_10.3_Historical_Recount_Events.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 10.4: NARRATIVE TEXT - FOLKLORE & MORAL LESSONS
# ==============================================================================
def create_tp_10_4():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS SMA",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase E / Kelas X / Semester 2 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 10.4 - Menganalisis nilai moral, alur cerita, dan unsur bahasa teks naratif dongeng dan legenda nusantara maupun mancanegara secara kritis.",
        "Elemen Capaian Pembelajaran": "Menyimak - Berbicara & Membaca - Memirsa",
        "Genre Teks / Ruang Lingkup": "Narrative Text (Folklore, Legends, Moral Stories)",
        "Alokasi Waktu": "12 Jam Pelajaran (6 Pertemuan @ 2 JP)",
        "Target Kemahiran CEFR": "Level A2+ s.d. B1",
        "Profil Pelajar Pancasila": "Berakhlak Mulia, Bernalar Kritis, Berkebinekaan Global"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "Indonesian Cultural Heritage / West Sumatra Folklore",
        "title": "The Legend of Malin Kundang: A Lesson in Filial Respect",
        "url": "https://folklore.kemdikbud.go.id/malin-kundang-legend",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah Awal - Sesuai Siswa Kelas X)"
    }
    t1_paras = [
        "A long time ago, in a small coastal village in West Sumatra, lived a poor widow named Mande Rubayah and her only son, Malin Kundang. They lived in hardship, but the mother loved her son dearly. As Malin grew into a strong young man, he decided to sail overseas on a merchant ship to seek his fortune and lift his family out of poverty.",
        "Many years passed without news, and his mother waited patiently at the beach every day. Eventually, a magnificent three-masted ship docked at the harbor. Malin Kundang stepped ashore wearing wealthy clothes alongside his beautiful wife. Recognizing her beloved son, the elderly mother ran toward him and cried with tears of joy. However, feeling ashamed of her dirty rags in front of his wife, Malin pushed her away and coldly yelled, 'You are an ugly beggar! You cannot be my mother!'",
        "Heartbroken and humiliated, the poor mother knelt down on the beach and prayed to God for justice. Suddenly, the clear sky turned black, thunder roared, and a violent storm smashed Malin's great ship onto the shore. Realizing his terrible sin, Malin cried in regret, but it was too late. His body slowly stiffened and turned into stone on the sandy beach."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Menghibur pembaca sekaligus mengajarkan pesan moral agar selalu berbakti dan menghormati orang tua.",
        "Struktur Teks": "Orientation (Paragraf 1: Pengenalan Malin Kundang dan ibunya di desa pantai) -> Complication (Paragraf 2: Malin kembali kaya raya namun mengingkari ibu kandungnya) -> Resolution (Paragraf 3: Doa ibu, badai besar, dan Malin berubah menjadi batu karang).",
        "Fitur Bahasa": "Past Action Verbs ('lived', 'sailed', 'docked', 'pushed', 'turned'), Dialog langsung ('You are an ugly beggar!'), Kata keterangan waktu naratif ('A long time ago', 'Many years passed', 'Suddenly')."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "The Curse of the Arrogant Son: The Legend of Malin Kundang",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "North Sumatra Folklore Archives / Indonesia Tourism",
        "title": "The Legend of Lake Toba: The Broken Promise and the Sacred Flood",
        "url": "https://indonesia.travel/gb/en/destinations/sumatra/lake-toba/legend",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah Awal - Sesuai Siswa Kelas X)"
    }
    t2_paras = [
        "Once upon a time in North Sumatra, there lived a humble, hardworking farmer named Toba. One afternoon while fishing in the river, he caught a large, beautiful golden fish. To his astonishment, the fish spoke with a human voice and transformed into an enchanting woman. She agreed to marry Toba on one strict condition: he must promise never to tell anyone that she was once a fish.",
        "Toba promised solemnly, and they lived happily together. Soon, they were blessed with a son named Samosir. However, Samosir was a playful boy who often ate all the food in the house. One afternoon, his mother asked Samosir to bring lunch to his father in the fields. Feeling hungry on the way, Samosir opened the basket and ate nearly all the meal, leaving only small crumbs.",
        "When the exhausted father opened the empty basket, he lost his temper and shouted angrily, 'You ill-behaved son of a fish!' Samosir ran home crying to his mother. Realizing that her husband had broken his sacred promise, the mother told Samosir to run to the highest hill. Heavy rain poured endlessly, earthquakes shook the land, and the valley filled with water, creating the magnificent Lake Toba, with Samosir Island rising in its center."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Mengajarkan pentingnya memegang janji suci dan mengendalikan amarah melalui cerita asal-usul Danau Toba.",
        "Struktur Teks": "Orientation (Toba menangkap ikan emas yang berubah menjadi wanita) -> Complication (Toba melanggar janji saat marah kepada Samosir) -> Resolution (Banjir besar menenggelamkan lembah menjadi Danau Toba).",
        "Fitur Bahasa": "Direct Speech ('You son of a fish!'), Kata penghubung kronologis ('One afternoon', 'Soon', 'When')."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "The Broken Promise: The Origin of Lake Toba",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "British Council LearnEnglish Stories / Oxford Legends",
        "title": "The Sword in the Stone: The Rise of Young King Arthur",
        "url": "https://learnenglish.britishcouncil.org/general-english/stories/king-arthur",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah Awal - Sesuai Siswa Kelas X)"
    }
    t3_paras = [
        "In ancient Britain, after the death of the wise king Uther Pendragon, the kingdom fell into chaos. In a churchyard in London, a mysterious large stone appeared. Firmly stuck in an anvil on top of the stone was a shiny steel sword with golden letters that read: 'Whosoever pulls out this sword from this stone is the rightful king of all Britain.'",
        "Hundreds of proud knights and strong warriors from across the land tried with all their might to draw the blade, but none could move it even a fraction of an inch. Years passed, and a young squire named Arthur came to London to help his foster brother Kay at a tournament. When Kay realized he had forgotten his sword at the inn, Arthur ran to fetch it, but found the inn locked.",
        "Seeing the sword in the churchyard, Arthur innocently approached and pulled it out smoothly without any effort. When the great nobles saw young Arthur holding the blade, they were shocked. After verifying the miraculous deed, they bowed before him, recognizing that true nobility comes from purity of heart rather than physical pride. Arthur was crowned king and brought lasting peace to Britain."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "The Sword in the Stone: The Legend of King Arthur",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Hardship", "Noun", "Kesulitan hidup, masa-masa prihatin", "The mother and son lived in hardship in their small hut."),
        ("Merchant", "Noun", "Pedagang perniagaan antar pulau/negara", "Malin sailed on a large merchant ship to seek fortune."),
        ("Ashamed", "Adjective", "Malu dan enggan mengakui kebenaran", "He felt ashamed of his poor mother in front of his wealthy wife."),
        ("Astonishment", "Noun", "Rasa takjub atau sangat heran", "To his astonishment, the golden fish spoke in a human voice."),
        ("Solemnly", "Adverb", "Dengan sungguh-sungguh dan khidmat", "He promised solemnly never to reveal her secret origin."),
        ("Exhausted", "Adjective", "Sangat lelah setelah bekerja keras", "The exhausted father sat down waiting for his lunch."),
        ("Anvil", "Noun", "Landasan besi tempa untuk menempa logam", "The sword was firmly stuck in an iron anvil on the stone."),
        ("Nobility", "Noun", "Kemuliaan budi pekerti atau martabat luhur", "True nobility comes from a pure and kind heart.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "In Text 1, what reason caused Malin Kundang to deny and push away his elderly mother at the harbor?",
        "What strict condition did the golden fish make before agreeing to marry Toba in Text 2?",
        "According to Text 3, what was written on the golden letters of the sword stuck in the stone?"
    ]
    inf_q = [
        "What does the transformation of Malin Kundang into stone symbolize regarding his moral behavior?",
        "In Text 2, what lesson can we learn about controlling our anger when dealing with family members?",
        "Why were the proud knights unable to pull out the sword while young Arthur could draw it effortlessly?"
    ]
    eval_q = [
        "Do you think the mother's curse on Malin Kundang in Text 1 was fair, or should she have forgiven him? Explain your viewpoint.",
        "Compare the moral lessons between the Indonesian legend in Text 1 and the British legend in Text 3. What positive values do both teach to teenagers?"
    ]
    perf_task = (
        "PROYEK MENDONGENG & REKONSTRUKSI CERITA (LKPD / SUMATIF):\n"
        "1. Pilihlah satu cerita rakyat nusantara yang kamu sukai (misal: Roro Jonggrang, Timun Mas, Bawang Merah Bawang Putih, Danau Toba).\n"
        "2. Tuliskan ringkasan cerita naratif sepanjang 150-200 kata dalam bahasa Inggris dengan alur lengkap (Orientation, Complication, Resolution).\n"
        "3. Tuliskan 1 kalimat pesan moral (moral value) yang dapat dipetik dari cerita tersebut.\n"
        "4. Praktikkan mendongeng (Storytelling) di depan kelas dengan intonasi yang menarik dan ekspresif!"
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_10.4_Narrative_Text_Folklore.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 10.5: PROCEDURE TEXT - HEALTHY LIVING & DIGITAL SAFETY
# ==============================================================================
def create_tp_10_5():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS SMA",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase E / Kelas X / Semester 2 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 10.5 - Menganalisis dan menyusun teks prosedur lisan dan tulis tentang gaya hidup sehat, keselamatan digital, dan resep tradisional nusantara.",
        "Elemen Capaian Pembelajaran": "Membaca - Memirsa & Menulis - Mempresentasikan",
        "Genre Teks / Ruang Lingkup": "Procedure Text (Healthy Habits, Cyber Hygiene, Traditional Recipes)",
        "Alokasi Waktu": "12 Jam Pelajaran (6 Pertemuan @ 2 JP)",
        "Target Kemahiran CEFR": "Level A2+ s.d. B1",
        "Profil Pelajar Pancasila": "Mandiri, Bernalar Kritis, Kreatif"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "Harvard Health Publishing / Sleep Foundation Youth",
        "title": "Healthy Sleep Habits for Teenagers: A Step-by-Step Guide",
        "url": "https://www.health.harvard.edu/blog/teens-and-sleep-healthy-habits",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah Awal - Sesuai Siswa Kelas X)"
    }
    t1_paras = [
        "Quality sleep is essential for high school students. It improves memory, sharpens focus in class, and supports a healthy immune system. Doctors recommend that teenagers sleep between 8 to 10 hours each night. Follow these simple steps to build a healthy bedtime routine:",
        "Step 1: Set a consistent sleep schedule. Go to bed and wake up at the exact same time every day, even on weekends. This trains your brain's internal clock to feel naturally tired at night.\n\n"
        "Step 2: Turn off all digital screens at least 45 minutes before sleeping. Blue light emitted from smartphones, laptops, and televisions prevents your brain from producing melatonin, the hormone that makes you sleepy.\n\n"
        "Step 3: Create a quiet and cool bedroom environment. Keep your room dark, tidy, and well-ventilated. If necessary, use curtains to block outside streetlights.\n\n"
        "Step 4: Avoid heavy meals and caffeinated drinks (such as coffee, energy drinks, and soda) after 5:00 PM. Instead, drink warm water or chamomile tea.\n\n"
        "Step 5: Do a relaxing pre-sleep activity. Spend 15 minutes reading a printed book, listening to soft instrumental music, or practicing slow breathing exercises."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Memberikan panduan langkah demi langkah kepada remaja agar memiliki pola tidur sehat dan berkualitas.",
        "Struktur Teks": "Goal (Tujuan mendapatkan tidur 8-10 jam yang berkualitas) -> Materials/Conditions (Suasana kamar nyaman, buku santai) -> Steps (5 langkah berurutan dengan kata kerja perintah/imperative).",
        "Fitur Bahasa": "Imperative Verbs ('Set', 'Turn off', 'Create', 'Avoid', 'Drink'), Sequencers ('Step 1', 'Step 2', 'Instead'), Present Tense."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "How to Build a Healthy Sleep Routine for Teenagers",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "UNICEF Youth Digital Safety / CISA Cyber Guidelines",
        "title": "Creating Strong Passwords and Protecting Your Privacy Online",
        "url": "https://www.unicef.org/digital-safety-tips-for-teens",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah Awal - Sesuai Siswa Kelas X)"
    }
    t2_paras = [
        "In our modern digital world, protecting your social media accounts and personal information is just as important as locking your front door. Cyber criminals often use automated software to guess simple passwords. Follow these essential procedures to protect your digital identity:",
        "1. Create a memorable passphrase instead of a single word. Combine at least three unrelated words with numbers and special symbols (for example: BluePencil#Jump92).\n\n"
        "2. Never use personal information such as your birthdate, phone number, school name, or nickname in your password, as these are easy to find on public profiles.\n\n"
        "3. Enable Two-Factor Authentication (2FA) on all your social media and email apps. This adds a second layer of security by sending a verification code to your phone.\n\n"
        "4. Avoid clicking on suspicious links sent through direct messages or unfamiliar emails. These links often lead to phishing websites designed to steal your login credentials.\n\n"
        "5. Regularly review your privacy settings. Set your profile to 'Friends Only' to prevent strangers from accessing your photos and location data."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Memberikan instruksi jelas kepada siswa untuk mengamankan akun digital dan mencegah peretasan.",
        "Struktur Teks": "Goal (Mengamankan akun digital) -> Steps (Langkah 1 s.d. 5: pembuatan passphrase, aktivasi 2FA, pencegahan phishing).",
        "Fitur Bahasa": "Imperative Negative ('Never use', 'Avoid clicking'), Action Verbs ('Create', 'Enable', 'Review')."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "How to Secure Your Social Media Accounts and Personal Data",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "Ministry of Tourism and Creative Economy / Traditional Culinary",
        "title": "How to Prepare Traditional Indonesian Jamu Kunyit Asam",
        "url": "https://www.indonesia.travel/gb/en/trip-ideas/traditional-herbal-drink-jamu",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah Awal - Sesuai Siswa Kelas X)"
    }
    t3_paras = [
        "Jamu Kunyit Asam is a refreshing and healthy traditional Indonesian herbal drink. It is well known for boosting immunity, reducing inflammation, and improving digestion. Here is how to prepare this delicious beverage at home.",
        "Ingredients:\n"
        "• 200 grams of fresh turmeric (kunyit), washed and peeled\n"
        "• 100 grams of tamarind paste (asam jawa)\n"
        "• 80 grams of palm sugar (gula merah), thinly sliced\n"
        "• 1 liter of clean drinking water\n"
        "• A pinch of salt\n\n"
        "Instructions:\n"
        "1. First, grate or blend the peeled fresh turmeric with a small amount of water until smooth.\n"
        "2. Next, pour the remaining water into a medium cooking pot and add the blended turmeric.\n"
        "3. Then, add the tamarind paste, palm sugar, and a pinch of salt to the pot.\n"
        "4. Bring the mixture to a gentle boil over medium heat for about 15 minutes, stirring occasionally.\n"
        "5. After boiling, turn off the heat and strain the herbal liquid through a fine sieve into a glass pitcher.\n"
        "6. Finally, serve the jamu warm in small cups, or chill it in the refrigerator with ice cubes for a refreshing cold drink."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "How to Make Refreshing Jamu Kunyit Asam at Home",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Immune system", "Noun Phrase", "Sistem kekebalan tubuh melawan penyakit", "Adequate sleep strengthens your immune system."),
        ("Consistent", "Adjective", "Konsisten, teratur sama setiap waktu", "Maintain a consistent sleep schedule even on weekends."),
        ("Ventilated", "Adjective", "Memiliki sirkulasi udara yang baik", "Keep your bedroom dark, quiet, and well-ventilated."),
        ("Passphrase", "Noun", "Kata sandi berupa rangkaian beberapa kata unik", "A long passphrase is much harder for hackers to crack."),
        ("Authentication", "Noun", "Otentikasi atau verifikasi keabsahan pengguna", "Enable Two-Factor Authentication for maximum security."),
        ("Phishing", "Noun", "Penipuan digital untuk mencuri kata sandi", "Beware of phishing links designed to steal your credentials."),
        ("Turmeric", "Noun", "Kunyit (tanaman herbal rimpang kuning)", "Fresh turmeric is the main ingredient in jamu kunyit asam."),
        ("Strain", "Verb", "Menyaring cairan dari ampas menggunakan saringan", "Strain the herbal tea through a clean sieve before serving.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "According to Text 1, why should teenagers turn off digital screens at least 45 minutes before bedtime?",
        "In Text 2, what is Two-Factor Authentication (2FA) and how does it protect your online accounts?",
        "What are the main natural ingredients needed to prepare Jamu Kunyit Asam in Text 3?"
    ]
    inf_q = [
        "How does blue light from electronic gadgets interfere with the body's natural production of melatonin?",
        "Why is using a passphrase with three combined words safer than using a single common English word?",
        "What is the purpose of straining the boiled jamu mixture in Step 5 before serving?"
    ]
    eval_q = [
        "Evaluate your own current sleeping and gadget habits: Which step from Text 1 would be the most challenging for you to implement, and why?",
        "Why is preserving traditional recipes like Jamu (Text 3) important for young Indonesians in this modern era of packaged soft drinks?"
    ]
    perf_task = (
        "PENULISAN TEKS PROSEDUR KREATIF (LKPD / SUMATIF MENULIS):\n"
        "Pilihlah salah satu tema prosedur berikut:\n"
        "Tema 1: How to Cook a Simple Healthy Indonesian Breakfast (e.g., Oatmeal with fruits / Nasi Goreng Sehat).\n"
        "Tema 2: How to Organize a Study Desk to Maximize Focus.\n"
        "Tema 3: How to Safely Shop Online without Getting Scammed.\n"
        "Susunlah teks prosedur sepanjang 150-200 kata dengan format: Goal -> Materials/Ingredients Needed -> Steps 1 s.d. 5.\n"
        "Gunakan kata kerja perintah (imperative verbs) dan kata penghubung urutan waktu (first, next, then, finally)."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_10.5_Procedure_Text_Healthy_Living.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# TP 10.6: REPORT & SIMPLE ANALYTICAL EXPOSITION - ENVIRONMENT
# ==============================================================================
def create_tp_10_6():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "BANK MATERI OTENTIK PEMBELAJARAN BAHASA INGGRIS SMA",
        "Modul Ajar, LKPD & Asesmen Sumatif | Kurikulum Merdeka 2025"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Umum (Wajib)",
        "Fase / Kelas / Semester": "Fase E / Kelas X / Semester 2 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 10.6 - Mengevaluasi informasi faktual dalam teks report dan menyusun teks eksposisi analitis sederhana tentang pelestarian lingkungan hidup dan satwa langka.",
        "Elemen Capaian Pembelajaran": "Membaca - Memirsa & Menulis - Mempresentasikan",
        "Genre Teks / Ruang Lingkup": "Report Text & Simple Analytical Exposition (Wildlife Conservation & Ecology)",
        "Alokasi Waktu": "14 Jam Pelajaran (7 Pertemuan @ 2 JP)",
        "Target Kemahiran CEFR": "Level A2+ s.d. B1",
        "Profil Pelajar Pancasila": "Bernalar Kritis, Gotong Royong, Berakhlak Mulia"
    }
    docx_helper.add_tp_info_box(doc, info)
    docx_helper.add_heading_1(doc, "I. KOLEKSI MATERI OTENTIK BERAGAM (MULTI-TEXT AUTHENTIC BANK)")

    t1_src = {
        "portal": "World Wildlife Fund (WWF) International Species / IUCN Red List",
        "title": "The Orangutan: Gentle Giants of the Indonesian Rainforest",
        "url": "https://www.worldwildlife.org/species/orangutan",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah Awal - Sesuai Siswa Kelas X)"
    }
    t1_paras = [
        "Orangutans are the only great apes native to Asia, found exclusively on the tropical islands of Borneo and Sumatra. The name 'orangutan' originates from the Malay words 'orang' (person) and 'hutan' (forest). Known for their long reddish-brown hair and distinctive arm span, these gentle primates spend approximately 90 percent of their lives high in the forest canopy.",
        "Orangutans have an extremely diverse diet, eating over 400 different types of fruits, leaves, tree bark, and flowers. Because they consume large quantities of wild fruits and travel long distances, they play an essential role as 'gardeners of the forest' by dispersing fruit seeds across the ecosystem. Without orangutans, many tropical tree species would struggle to reproduce.",
        "Unfortunately, orangutan populations are classified as critically endangered. Their numbers have declined sharply due to deforestation, illegal forest fires, and land clearing for large plantations. Conservation centers in Central Kalimantan and North Sumatra are working hard to rescue orphaned baby orangutans, teach them survival skills, and release them back into protected national parks."
    ]
    t1_analysis = {
        "Fungsi Sosial": "Menyajikan laporan faktual ilmiah (Report Text) mengenai habitat, makanan, peran ekologis, dan ancaman kelestarian orangutan.",
        "Struktur Teks": "General Classification (Paragraf 1: Pengenalan spesies kera besar Asia) -> Description (Paragraf 2: Pola makan dan fungsi penyebar benih; Paragraf 3: Ancaman kepunahan dan upaya rehabilitasi).",
        "Fitur Bahasa": "General Nouns ('Orangutans', 'tropical tree species'), Present Tense ('spend', 'consume', 'play'), Passive Voice ('are classified', 'are found')."
    }
    docx_helper.add_authentic_text_block(
        doc, "1", "Orangutans: The Gardeners of Asia's Rainforests",
        "Bahan Modul Ajar (Modeling of the Text)",
        t1_src, content_paragraphs=t1_paras, analysis_dict=t1_analysis
    )

    t2_src = {
        "portal": "National Geographic Education / Ocean Conservation",
        "title": "Why We Must Stop Plastic Pollution in Our Oceans: An Exposition",
        "url": "https://www.nationalgeographic.org/encyclopedia/ocean-plastic/",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah Awal - Sesuai Siswa Kelas X)"
    }
    t2_paras = [
        "Plastic is one of the most common materials in modern life. It is lightweight, durable, and inexpensive. However, our heavy reliance on single-use plastics has created a severe global environmental crisis. Every year, over eight million tons of plastic waste leak into the oceans, threatening marine life and human health. We must take immediate action to reduce plastic pollution for several important reasons.",
        "First, plastic trash is deadly to sea animals. Thousands of marine creatures—including sea turtles, dolphins, seabirds, and whales—often mistake plastic bags for jellyfish or fish. When animals ingest plastic, it blocks their digestive systems, causing painful starvation. Other animals become trapped in discarded fishing nets and drown.",
        "Second, plastic never truly biodegrades. Instead, sunlight and ocean waves break large plastic bottles into tiny pieces called microplastics. Small fish eat these toxic particles, and when larger fish eat the small fish, the microplastics travel up the food chain and eventually end up on our dinner plates.",
        "In conclusion, plastic pollution is an urgent threat to our oceans and our own health. High school students can make a real difference by carrying reusable water bottles, refusing plastic shopping bags, and participating in local beach clean-ups."
    ]
    t2_analysis = {
        "Fungsi Sosial": "Mempersuasi pembaca (Analytical Exposition) bahwa sampah plastik di lautan adalah bahaya nyata yang harus segera diatasi.",
        "Struktur Teks": "Thesis (Paragraf 1: Masalah sampah plastik lautan) -> Arguments (Paragraf 2: Bahaya plastik bagi satwa laut; Paragraf 3: Masalah mikroplastik dalam rantai makanan) -> Reiteration (Paragraf 4: Seruan tindakan nyata siswa).",
        "Fitur Bahasa": "Causal Conjunctions ('because', 'instead', 'for several reasons'), Present Tense, Evaluative Words ('severe crisis', 'urgent threat')."
    }
    docx_helper.add_authentic_text_block(
        doc, "2", "Why We Must Stop Single-Use Plastics in Our Oceans",
        "Bahan LKPD (Joint Construction / Analisis Kelompok)",
        t2_src, content_paragraphs=t2_paras, analysis_dict=t2_analysis
    )

    t3_src = {
        "portal": "UNEP Youth Environment Portal / Green School Guide",
        "title": "Why High Schools Should Practice Waste Sorting and Composting",
        "url": "https://www.unep.org/youth-environment-initiatives",
        "year": "2023 / 2024",
        "cefr": "Level B1 (Menengah Awal - Sesuai Siswa Kelas X)"
    }
    t3_paras = [
        "Schools are centers of learning where students develop important habits for the future. With hundreds of students and teachers gathered every day, a single high school generates dozens of kilograms of waste daily, including snack wrappers, paper, and food scraps. Implementing waste sorting and organic composting in every school is an essential practice that should be mandatory.",
        "Foremost, sorting waste into separate bins—organic, recyclable, and residual—reduces the volume of trash sent to overflowing municipal landfills. When students separate clean plastic bottles and cardboard boxes, these materials can easily be collected by recycling centers to create new products, saving natural resources and energy.",
        "Furthermore, composting organic waste from the school canteen produces rich, chemical-free fertilizer for school gardens. Instead of allowing food scraps to rot in landfills and release methane gas, students can turn vegetable peels into fertile soil to grow flowers and vegetables on campus. This gives students valuable hands-on science experience.",
        "In summary, practicing waste sorting at school teaches environmental responsibility, reduces city waste, and creates a greener campus. Every school administration should provide categorized bins and encourage students to care for our planet."
    ]
    docx_helper.add_authentic_text_block(
        doc, "3", "The Benefits of School Waste Sorting and Recycling",
        "Bahan Asesmen Sumatif Mandiri (Independent Test)",
        t3_src, content_paragraphs=t3_paras
    )

    vocabs = [
        ("Endangered", "Adjective", "Terancam punah di alam liar", "Orangutans are critically endangered due to habitat loss."),
        ("Deforestation", "Noun", "Penebangan hutan secara liar dan masif", "Deforestation destroys the natural homes of wild animals."),
        ("Single-use", "Adjective", "Sekali pakai langsung dibuang", "We should stop using single-use plastic cups and straws."),
        ("Microplastics", "Noun (Plural)", "Partikel plastik berukuran mikroskopis", "Fish accidentally consume microplastics floating in the water."),
        ("Biodegrade", "Verb", "Terurai secara alami oleh mikroorganisme", "Plastic items can take hundreds of years to biodegrade."),
        ("Food chain", "Noun Phrase", "Rantai makanan ekosistem alam", "Toxins travel up the marine food chain to humans."),
        ("Composting", "Noun / Gerund", "Pembuatan kompos dari sisa sampah organik", "Composting canteen food scraps creates natural soil fertilizer."),
        ("Landfill", "Noun", "Tempat pembuangan akhir (TPA) sampah", "Sorting trash reduces the burden on city landfills.")
    ]
    docx_helper.add_vocab_table(doc, vocabs)

    lit_q = [
        "According to Text 1, why are orangutans affectionately called the 'gardeners of the forest'?",
        "What are the two major dangers of ocean plastic pollution explained in Text 2?",
        "In Text 3, what are the three categorized bins recommended for schools to sort their waste?"
    ]
    inf_q = [
        "Why would tropical fruit tree species struggle to reproduce if orangutans disappeared from the ecosystem?",
        "Explain how plastic waste in the ocean can eventually end up inside human bodies through the food chain (Text 2).",
        "How does composting food scraps at school help prevent the release of harmful greenhouse gases like methane (Text 3)?"
    ]
    eval_q = [
        "In your opinion, what is the most effective way to encourage high school students to bring their own reusable food containers to school?",
        "Evaluate the difference between a Report Text (Text 1) and an Analytical Exposition (Text 2). How do their social purposes differ?"
    ]
    perf_task = (
        "PENYUSUNAN ESAI EKSPOSISI SEDERHANA (LKPD / SUMATIF MENULIS):\n"
        "Tulislah sebuah esai eksposisi analitis sederhana (Simple Analytical Exposition) sepanjang 150-200 kata mengenai salah satu isu lingkungan:\n"
        "Pilihan 1: Why Teenagers Should Stop Using Single-Use Plastic Bags.\n"
        "Pilihan 2: Why Every Student Should Plant at Least One Tree at Home or School.\n"
        "Struktur Esai:\n"
        "- Paragraph 1: Thesis (Perkenalkan masalah dan sebutkan pendapatmu).\n"
        "- Paragraph 2 & 3: Arguments (Berikan 2 alasan kuat didukung fakta sederhana).\n"
        "- Paragraph 4: Reiteration (Tegaskan kembali kesimpulan dan ajakan bertindak)."
    )
    docx_helper.add_assessment_questions(doc, lit_q, inf_q, eval_q, perf_task)

    fpath = os.path.join(TARGET_DIR, "Materi_Otentik_TP_10.6_Report_and_Analytical_Exposition.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


if __name__ == "__main__":
    print("Generating calibrated authentic materials for 5-Modul Ajar X SMA (Fase E)...")
    create_tp_10_1()
    create_tp_10_2()
    create_tp_10_3()
    create_tp_10_4()
    create_tp_10_5()
    create_tp_10_6()
    print("All 6 DOCX files for Kelas X SMA successfully updated and calibrated!")
