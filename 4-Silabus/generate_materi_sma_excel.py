# -*- coding: utf-8 -*-
"""
Generator script for 'Materi Bahasa Inggris SMA.xlsx' in 4-Silabus.
Peta materi terstruktur dan sederhana untuk 4 Keterampilan:
1) Listening (Menyimak)
2) Speaking (Berbicara)
3) Reading (Membaca & Memirsa)
4) Writing (Menulis & Mempresentasikan)
Terpisah secara jelas antara Bahasa Inggris Wajib (Kelas X, XI, XII)
dan Bahasa Inggris Tingkat Lanjut (Kelas XI, XII).
"""

import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Palette & Fonts
FONT_NAME = "Segoe UI"
FONT_TITLE = Font(name=FONT_NAME, size=15, bold=True, color="0F2942")
FONT_SUBTITLE = Font(name=FONT_NAME, size=9.5, italic=True, color="475569")
FONT_SEC_HEADER = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
FONT_TH = Font(name=FONT_NAME, size=9.5, bold=True, color="FFFFFF")
FONT_DATA = Font(name=FONT_NAME, size=9, color="1E293B")
FONT_DATA_BOLD = Font(name=FONT_NAME, size=9, bold=True, color="0F2942")

FILL_NAVY = PatternFill(start_color="0F2942", end_color="0F2942", fill_type="solid")
FILL_ROYAL = PatternFill(start_color="1E40AF", end_color="1E40AF", fill_type="solid")
FILL_TEAL = PatternFill(start_color="0D9488", end_color="0D9488", fill_type="solid")
FILL_SLATE = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
FILL_HEADER_WAJIB = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
FILL_HEADER_TL = PatternFill(start_color="065F46", end_color="065F46", fill_type="solid")
FILL_ZEBRA = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
FILL_ACCENT_BLUE = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
FILL_ACCENT_GREEN = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
FILL_SECTION_WAJIB = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
FILL_SECTION_TL = PatternFill(start_color="D1FAE5", end_color="D1FAE5", fill_type="solid")

BORDER_THIN = Side(border_style="thin", color="CBD5E1")
CELL_BORDER = Border(left=BORDER_THIN, right=BORDER_THIN, top=BORDER_THIN, bottom=BORDER_THIN)

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="top", wrap_text=True)
ALIGN_CENTER_TOP = Alignment(horizontal="center", vertical="top", wrap_text=True)

def apply_title(ws, title, subtitle, max_cols=8):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_cols)
    ws.cell(row=1, column=1, value=title).font = FONT_TITLE
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_cols)
    ws.cell(row=2, column=1, value=subtitle).font = FONT_SUBTITLE
    ws.cell(row=2, column=1).alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[3].height = 10

def auto_fit_columns(ws, min_w=10, max_w=48):
    ws.views.sheetView[0].showGridLines = True
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            if cell.row in [1, 2]: # skip title
                continue
            val = str(cell.value or '')
            lines = val.split('\n')
            for l in lines:
                if len(l) > max_len:
                    max_len = len(l)
        col_width = max(min_w, min(max_len + 3, max_w))
        ws.column_dimensions[col_letter].width = col_width


# ==============================================================================
# DATA: BAHASA INGGRIS WAJIB (KELAS X, XI, XII)
# ==============================================================================
DATA_WAJIB_4_SKILLS = [
    # KELAS X - SEMESTER 1
    ("1", "Kelas X - Sem 1", "TP 10.1", "Descriptive Text (Great Athletes & Tourism)",
     "Menyimak monolog audio biografi atlet dunia; mengidentifikasi ciri fisik, karakter mental, dan prestasi kunci.",
     "Mendeskripsikan secara lisan tokoh atlet idola dalam dialog tanya-jawab berpasangan dengan intonasi wajar.",
     "Membaca teks deskripsi destinasi wisata dunia; menganalisis struktur Identification & Description, adjectives.",
     "Menyusun teks deskriptif 100-150 kata tentang objek wisata/atlet lokal dengan simple present tense yang akurat.",
     "Simple Present Tense, Descriptive Adjectives, Noun Phrases, Sensory Details.",
     "Lembar analisis audio & teks deskriptif profil atlet / destinasi wisata."),

    ("2", "Kelas X - Sem 1", "TP 10.2", "Descriptive Text Multimodal (Promosi Budaya & Wisata)",
     "Menyimak video promosi pariwisata nusantara; mencatat diksi sensory words dan daya tarik destinasi.",
     "Mempresentasikan brosur promosi wisata secara lisan dengan artikulasi jelas, percaya diri, dan kontak mata.",
     "Membaca dan membedah tata letak (layout) brosur wisata profesional, headline, takarir gambar, dan persuasi.",
     "Merancang brosur digital (Canva/Poster) memadukan teks deskriptif otentik dan visual estetis cagar budaya.",
     "Adjectives of Quality, Spatial Ordering, Persuasive Language, Descriptive Captions.",
     "Brosur promosi wisata nusantara digital & rekaman presentasi lisan."),

    ("3", "Kelas X - Sem 1", "TP 10.3", "Recount Text (Historical & Personal Events)",
     "Menyimak monolog kilas balik peristiwa bersejarah (Battle of Surabaya / Sumpah Pemuda); memetakan linimasa waktu.",
     "Menceritakan kembali pengalaman berkesan masa lalu secara lisan dalam kelompok secara runtut dan ekspresif.",
     "Membaca teks recount sejarah; menganalisis struktur Orientation, Events, Re-orientation, dan kata kerja lampau.",
     "Menulis jurnal memoar pribadi atau laporan peristiwa sejarah 150-200 kata dengan konjungsi kronologis.",
     "Simple Past Tense (Regular/Irregular), Past Continuous, Time Connectors (first, then, finally).",
     "Linimasa peristiwa sejarah & esai recount memoar pribadi (Learning Journal)."),

    # KELAS X - SEMESTER 2
    ("4", "Kelas X - Sem 2", "TP 10.4", "Narrative Text (Folklore & Legends)",
     "Menyimak penceritaan dongeng/legenda nusantara (Malin Kundang / Roro Jonggrang); mencermati karakter dan intonasi.",
     "Mendemonstrasikan storytelling lisan ekspresif membawakan legenda rakyat dengan variasi suara tokoh.",
     "Membaca teks cerita rakyat; menganalisis Orientation, Complication, Resolution, Coda, dan direct/indirect speech.",
     "Menulis naskah dongeng pendek 150-200 kata dengan pesan moral bijak atau merancang alternative ending kreatif.",
     "Past Action Verbs, Direct & Indirect Speech, Adverbs of Time, Moral Evaluation (Coda).",
     "Rekaman video storytelling & lembar analisis unsur intrinsik cerita."),

    ("5", "Kelas X - Sem 2", "TP 10.5", "Procedure Text (Healthy Living & Cyber Safety)",
     "Menyimak siniar (podcast) audio panduan pola hidup sehat remaja dan langkah keselamatan berinternet.",
     "Memperagakan instruksi lisan tutorial cara membuat menu sehat atau panduan mengamankan privasi akun medsos.",
     "Membaca manual teks prosedur; mengidentifikasi Goal, Materials/Equipment, dan Sequenced Steps berurutan.",
     "Menyusun teks prosedur infografis atau booklet mini tutorial hidup sehat/digital safety (100-150 kata).",
     "Imperative Sentences (Cut, Wash, Avoid), Sequencers (First, Next, Then), Adverbials of Manner.",
     "Infografis panduan hidup sehat / poster tutorial digital safety."),

    ("6", "Kelas X - Sem 2", "TP 10.6", "Report Text & Intro to Analytical Exposition",
     "Menyimak audio dokumenter konservasi fauna endemik Indonesia (Komodo, Orangutan) dan inovasi sains.",
     "Menyampaikan fakta ilmiah hasil pengamatan satwa langka secara lisan di depan kelompok kelas.",
     "Membaca teks report observasi sains; menganalisis General Classification, Description, dan kalimat pasif.",
     "Menulis laporan ilmiah mini 3 paragraf tentang satwa langka dan merumuskan tesis awal pelestarian lingkungan.",
     "General Nouns, Timeless Present Tense, Passive Voice, Technical Scientific Vocabulary.",
     "Makalah laporan observasi fauna ilmiah mini & poster kampanye lingkungan."),

    # KELAS XI - SEMESTER 3
    ("7", "Kelas XI - Sem 3", "TP 11.W.1", "Analytical Exposition (Digital Literacy & Wellness)",
     "Menyimak diskusi radio tentang bahaya kecanduan gawai dan pentingnya digital wellness bagi remaja.",
     "Mengemukakan pendapat dan mempertahankan argumen pro-keseimbangan digital dalam sesi diskusi kelas.",
     "Membaca artikel eksposisi analitis; mengurai Thesis, Arguments (bukti fakta empiris), dan Reiteration.",
     "Menulis esai eksposisi analitis 200-250 kata meyakinkan pembaca mengenai urgensi membatasi screen time.",
     "Internal Conjunctions (Furthermore, In addition), Causal Connectors (Therefore, Consequently), Evaluative Words.",
     "Esai eksposisi analitis bertema digital literacy & lembar telaah argumen."),

    ("8", "Kelas XI - Sem 3", "TP 11.W.2", "Hortatory Exposition (Environmental Advocacy)",
     "Menyimak orasi audio kampanye pemilahan sampah dan program sekolah bebas plastik (zero-waste school).",
     "Menyampaikan pidato anjuran persuasif (hortatory speech) menggunakan intonasi ajakan dan modalitas tegas.",
     "Membaca teks eksposisi hortatori; menganalisis Thesis, Arguments, dan Recommendation (tindakan nyata).",
     "Menyusun naskah advokasi hortatori 200-250 kata merekomendasikan solusi pengelolaan sampah sekolah.",
     "Modal Auxiliaries (should, must, ought to), Action Verbs, Persuasive Rhetoric.",
     "Naskah pidato advokasi lingkungan & poster kampanye aksi nyata di sekolah."),

    ("9", "Kelas XI - Sem 3", "TP 11.W.3", "Contemporary Narrative & Fractured Tales",
     "Menyimak sandiwara audio dongeng modern (fractured tale) yang diceritakan dari sudut pandang tokoh antagonis.",
     "Memeragakan percakapan peran (roleplay) dialog dongeng modern bersama rekan kelompok dengan ekspresi hidup.",
     "Membaca cerpen narasi kontemporer; mengkaji pembongkaran stereotip karakter, suspense, dan foreshadowing.",
     "Menulis cerita pendek modern 250 kata dengan plot twist orisinal dan refleksi nilai-nilai kemanusiaan.",
     "Past Perfect Tense, Descriptive Adverbs, Figurative Language, Dialogue Tags.",
     "Naskah cerita pendek narasi kontemporer & kartu evaluasi karakter tokoh."),

    # KELAS XI - SEMESTER 4
    ("10", "Kelas XI - Sem 4", "TP 11.W.4", "Explanation Text (Natural & Societal Phenomena)",
     "Menyimak penjelasan audio ilmiah mengenai proses terjadinya pemanasan global dan siklus hidrologi bumi.",
     "Menjelaskan alur sebab-akibat (cause-and-effect) fenomena alam/sosial menggunakan diagram alir di depan kelas.",
     "Membaca teks eksplanasi ilmiah; menganalisis General Statement, Sequenced Explanation, dan konjungsi kausal.",
     "Menulis teks eksplanasi kausal 200-250 kata menjelaskan mekanisme terjadinya banjir perkotaan atau gempa bumi.",
     "Passive Voice, Cause and Effect Connectors (because of, due to, as a result), Technical Terminology.",
     "Makalah eksplanasi ilmiah bergambar diagram alir sebab-akibat."),

    ("11", "Kelas XI - Sem 4", "TP 11.W.5", "Discussion Text (Everyday Teen Issues & School Policies)",
     "Menyimak gelar wicara lisan dua narasumber membahas pro dan kontra penerapan zonasi sekolah dan jam masuk pagi.",
     "Berpartisipasi aktif dalam debat kelompok membahas aturan sekolah dengan etika kesantunan saling menghargai.",
     "Membaca teks diskusi; mengurai Issue, Arguments For, Arguments Against, dan Conclusion/Balanced Recommendation.",
     "Menyusun esai diskusi berimbang 200-250 kata yang merumuskan solusi jalan tengah terhadap aturan sekolah.",
     "Contrast Connectors (On the other hand, Conversely, However), Concessive Clauses (Although), Modality.",
     "Esai teks diskusi berimbang isu remaja & lembar observasi diskusi kelas."),

    ("12", "Kelas XI - Sem 4", "TP 11.W.6", "Review Text (Arts, Books & Cinema)",
     "Menyimak ulasan audio kritikus film mengenai sinema inspiratif persahabatan karya sutradara nasional.",
     "Menyampaikan review lisan singkat (movie talk) merekomendasikan buku sastra atau film favorit kepada teman.",
     "Membaca teks reviu kritis; menganalisis Orientation, Interpretative Recount, Evaluation, dan Evaluative Summation.",
     "Menulis naskah reviu kritis 200-250 kata menilai kelebihan dan kekurangan novel/film lokal secara objektif.",
     "Evaluative Adjectives (remarkable, captivating, flawed), Metaphor, Recommendation Phrasing.",
     "Resensi naskah reviu film/buku & katalog rekomendasi karya literasi kelas."),

    # KELAS XII - SEMESTER 5
    ("13", "Kelas XII - Sem 5", "TP 12.W.1", "Discussion Text (Future Work & Workplace Automation)",
     "Menyimak siniar ketenagakerjaan mengenai dampak kecerdasan buatan terhadap pergeseran profesi masa depan.",
     "Mengutarakan pandangan berimbang dalam forum meja diskusi kelas tentang peluang dan ancaman otomasi kerja.",
     "Membaca artikel diskusi berbobot (World Economic Forum, BBC) dengan bahasa berpagar (hedging).",
     "Menyusun esai diskusi berimbang 250-300 kata disertai rekomendasi program pelatihan keterampilan kerja baru.",
     "Hedging Devices (point out, emphasize, evidence shows), Concessive Clauses (In spite of), Contrast Markers.",
     "Esai diskusi masa depan dunia kerja & transkrip diskusi meja bundar kelas."),

    ("14", "Kelas XII - Sem 5", "TP 12.W.2", "News Item Text & Media Literacy",
     "Menyimak siaran berita radio/TV mengenai peresmian infrastruktur transportasi kereta cepat Whoosh.",
     "Mempraktikkan siaran berita televisi (News Anchor simulation) dengan intonasi berita resmi dan wibawa siaran.",
     "Membaca teks berita aktual; menganalisis Newsworthy Event, Background/Elaboration, Sources, dan indirect quotes.",
     "Menyusun naskah berita siaran TV 120-150 kata berformat piramida terbalik dengan kutipan resmi narasumber.",
     "Reported Speech, Saying Verbs (stated, reported, confirmed), Passive Voice in Headlines, Concise Lead.",
     "Naskah siaran berita TV & rekaman video pembaca berita (News Casting)."),

    ("15", "Kelas XII - Sem 5", "TP 12.W.3", "Review Text (Creative Heritage & Local Craft)",
     "Menyimak wawancara desainer mengenai inovasi batik modern dan kerajinan nusantara menembus pasar ekspor.",
     "Mempresentasikan ulasan komersial produk kreatif daerah secara meyakinkan dalam pameran kewirausahaan kelas.",
     "Membaca artikel ulasan produk seni kriya (Orientation, Evaluation 1-2, Recommendation/Verdict).",
     "Menulis ulasan kritis 200-250 kata mengevaluasi keunikan, kualitas bahan, dan nilai estetika produk lokal.",
     "Descriptive Noun Phrases, Persuasive Diction, Appraising Lexis, Constructive Recommendations.",
     "Lembar ulasan produk ekonomi kreatif & portofolio apresiasi budaya lokal."),

    # KELAS XII - SEMESTER 6
    ("16", "Kelas XII - Sem 6", "TP 12.W.4", "Job Application Letter & Modern CV",
     "Menyimak tips wawancara HRD perusahaan multinasional mengenai kriteria utama surat lamaran yang dilirik rekruter.",
     "Menyampaikan perkenalan diri profesional singkat (elevator pitch 60 detik) merangkum kualifikasi kunci.",
     "Membaca dan membedah format surat lamaran bisnis resmi dan resume modern (action verbs, reverse chronological).",
     "Menyusun surat lamaran kerja eksekutif (200-250 kata) dan merancang Curriculum Vitae profesional siap kerja.",
     "Formal Business Register, Action Verbs, Standard Salutation & Sign-off, Professional Diction.",
     "Berkas surat lamaran kerja formal & Curriculum Vitae (CV) siap kirim."),

    ("17", "Kelas XII - Sem 6", "TP 12.W.5", "Job Interview Simulation & Oral Fluency",
     "Menyimak rekaman audio tanya-jawab wawancara kerja profesional dan pertanyaan jebakan kepribadian.",
     "Mempraktikkan simulasi wawancara kerja (Job Interview) merespons pertanyaan kompetensi dengan taktis dan santun.",
     "Membaca panduan strategi menjawab pertanyaan wawancara (STAR technique dasar, body language, composure).",
     "Menuliskan lembar naskah persiapan jawaban wawancara (skenario kekuatan diri, kelemahan, dan resolusi krisis).",
     "Polite Formulaic Language, Past Tense for Achievements, Tactical Fillers (Certainly, In my experience).",
     "Transkrip persiapan wawancara kerja & lembar penilaian unjuk kerja interview."),

    ("18", "Kelas XII - Sem 6", "TP 12.W.6", "Captions, Data Graphs & Multimodal Infographics",
     "Menyimak presentasi analis data yang mendeskripsikan tren demografi pemuda dan investasi energi hijau.",
     "Menjelaskan makna takarir (caption) dan tren grafik secara lisan di depan kelas dengan bahasa formal.",
     "Membaca infografis, takarir media massa (Lead + Narrative caption), dan grafik batang/garis analitis.",
     "Menulis takarir foto jurnalistik 2-3 kalimat dan menyusun laporan singkat 100-150 kata mendeskripsikan grafik.",
     "Trend Verbs (increased, declined, fluctuated, plateaued), Prepositions of Data (by, at, from...to), Prepositional Phrases.",
     "Takarir foto jurnalistik edukatif & laporan singkat analisis data grafik.")
]


# ==============================================================================
# DATA: BAHASA INGGRIS TINGKAT LANJUT (KELAS XI & XII TL)
# ==============================================================================
DATA_TL_4_SKILLS = [
    # KELAS XI TL - SEMESTER 3
    ("1", "Kelas XI TL - Sem 3", "TP 11.TL.1", "Advanced Analytical Exposition (AI Ethics & Privacy)",
     "Menyimak kuliah akademik MIT/Harvard mengenai bias algoritma AI, data mining, dan privasi digital.",
     "Memimpin seminar akademik kelas mempertahankan tesis etika kecerdasan buatan dengan argumen berbasis data.",
     "Membaca jurnal ilmiah eksposisi analitis tingkat lanjut (Academic Word List, complex nominal clauses, sitasi in-text).",
     "Menulis esai eksposisi analitis riset 400-500 kata dengan pembuktian data empiris dan penalaran kausal ketat.",
     "Academic Word List (AWL), Complex Nominal Clauses, In-text Citations, Evaluative Stance Markers.",
     "Makalah esai eksposisi analitis akademik 500 kata & lembar notulensi seminar."),

    ("2", "Kelas XI TL - Sem 3", "TP 11.TL.2", "Advanced Hortatory Exposition (Climate Diplomacy)",
     "Menyimak pidato delegasi diplomatik PBB mengenai negosiasi target emisi net-zero dan pajak karbon global.",
     "Berperan sebagai delegasi pemuda membawakan pidato resolusi iklim di sidang simulasi Model United Nations (MUN).",
     "Membaca naskah advokasi kebijakan lingkungan global (hortatory discourse, modalitas tingkat tinggi: imperative, mandatory).",
     "Menyusun policy paper advokasi 400-500 kata merekomendasikan skema transisi energi terbarukan berkeadilan.",
     "High Modality (imperative, strictly required), Action-oriented Verbs, International Treaty Diction.",
     "Naskah resolusi kebijakan lingkungan (Policy Paper) & unjuk kerja diplomasi MUN."),

    ("3", "Kelas XI TL - Sem 3", "TP 11.TL.3", "Advanced Literary Narrative & Psychological Stories",
     "Menyimak sandiwara audio sastra psikologis dunia; mencermati monolog batin dan atmosfer ketegangan cerita.",
     "Menganalisis perkembangan watak karakter sastra secara lisan dalam forum diskusi bedah buku (Literary Circle).",
     "Mengkaji cerpen sastra klasik/modern (symbolism, allegory, character arc, unreliable narrator, stream of consciousness).",
     "Menulis cerita pendek orisinal 500-750 kata mengeksplorasi tema moralitas batin, show-don't-tell, dan gaya figuratif.",
     "Literary Devices (Metaphor, Allegory, Symbolism), Sensory Imagery, Interior Monologue Syntax.",
     "Karya cerpen sastra orisinal 750 kata & esai apresiasi bedah karakter tokoh."),

    ("4", "Kelas XI TL - Sem 3", "TP 11.TL.4", "Stylistics, Poetry & Protest Songs",
     "Menyimak rekaman lagu-lagu protes bersejarah dunia (Bob Dylan, John Lennon) dan pembacaan puisi lirik kemanusiaan.",
     "Mendeklamasikan puisi bahasa Inggris (Poetry Recitation) dengan variasi rima, ritme, jeda, dan penghayatan emosional.",
     "Menganalisis stilistika teks puisi (imagery, enjambment, meter, figurative devices, kritik sosial masyarakat).",
     "Menggubah puisi orisinal bertema perdamaian dunia atau menulis esai analisis stilistika puisi 300 kata.",
     "Poetic Meters, Rhyme Schemes, Figurative Tropes (Personification, Irony), Prosodic Rhythm.",
     "Antologi puisi orisinal siswa & rekaman audio deklamasi puisi (Poetry Recital)."),

    # KELAS XI TL - SEMESTER 4
    ("5", "Kelas XI TL - Sem 4", "TP 11.TL.5", "Discussion Text (Official Puskurjar - Gaming Issues)",
     "Menyimak transkrip audio debat para neurolog dan sosiolog tentang dampak kognitif game online dan industri esports.",
     "Memimpin diskusi panel meja bundar mengurai argumen multi-sudut pandang mengenai batas sehat bermain game.",
     "Membedah teks model resmi Puskurjar (Bab D) teks diskusi multi-layered argumentation dan interpretasi data.",
     "Menyusun esai teks diskusi komprehensif 400 kata berimbang dengan rekomendasi regulasi screen-time yang solutif.",
     "Multi-perspective Signposting, Academic Hedging (suggests, indicates), Concessive Connectors.",
     "Esai diskusi komprehensif model Puskurjar & lembar penilaian panel meja bundar."),

    ("6", "Kelas XI TL - Sem 4", "TP 11.TL.6", "Complex Causal Explanation (Science Frontiers)",
     "Menyimak pemaparan saintifik TED Talk tentang fenomena fusi nuklir, komputasi kuantum, dan astrofisika.",
     "Mempresentasikan cara kerja teknologi frontier menggunakan analogi visual sains yang mudah dimengerti audiens.",
     "Membaca teks eksplanasi ilmiah kompleks (causal chains, technical nominalization, passive scientific register).",
     "Menyusun teks eksplanasi ilmiah 400 kata mengurai mekanisme kausal fenomena sains mutakhir secara runtut.",
     "Technical Nominalization, Complex Causal Passive Chains, Scientific Formulaic Transitions.",
     "Makalah eksplanasi teknologi frontier & slide presentasi infografis sains."),

    ("7", "Kelas XI TL - Sem 4", "TP 11.TL.7", "Biography & Historical Memoirs (World Leaders)",
     "Menyimak wawancara sejarah dan memoar lisan tokoh-tokoh pembawa perubahan dunia (Mandela, Marie Curie, Habibie).",
     "Membawakan presentasi biografi inspiratif berformat PechaKucha (20 slide x 20 detik) menyoroti keteguhan tokoh.",
     "Membaca biografi analitis dan memoar sejarah (historical recount tingkat mahir, evaluative reflection, arsip dokumen).",
     "Menulis esai biografi analitis 400-500 kata yang menyoroti filosofi kepemimpinan dan integritas moral tokoh dunia.",
     "Historical Narrative Register, Evaluative Epithets, Quotation Integration, Chronological Synthesizers.",
     "Esai biografi analitis kepemimpinan tokoh dunia & rekaman presentasi PechaKucha."),

    ("8", "Kelas XI TL - Sem 4", "TP 11.TL.8", "Critical Review of International Arts & Cinema",
     "Menyimak siniar kritikus seni internasional mengenai sinema arthouse dan pameran seni rupa kontemporer.",
     "Berdebat mengenai nilai estetika, semiotika visual, dan kritik sosial suatu film dalam simposium perfilman kelas.",
     "Membaca esai kritik seni internasional (The Guardian, Sight & Sound, semiotic analysis, mise-en-scène).",
     "Menulis naskah reviu kritis seni/sinema 400-500 kata dengan dekonstruksi sinematografi, tema, dan dampak sosiologis.",
     "Appraising Vocabulary, Semiotic Terms (mise-en-scène, montage), Evaluative Subordinate Clauses.",
     "Naskah ulasan kritis sinema internasional & esai kuratorial pameran seni."),

    # KELAS XII TL - SEMESTER 5
    ("9", "Kelas XII TL - Sem 5", "TP 12.TL.1", "Global Academic Discussion on Bioethics (CRISPR)",
     "Menyimak debat panel bioetika kedokteran dunia tentang rekayasa genetika germline dan etika triage medis AI.",
     "Memimpin dialog meja bundar kebijakan bioetika (Roundtable deliberation) dengan stance markers dan etika ilmiah.",
     "Membaca jurnal bioetika bereputasi (Nature Biotechnology, The Lancet) mengkaji argumen utilitarian vs hak asasi.",
     "Menulis Position Paper akademik 500 kata mengenai regulasi etika bioteknologi dan pemanfaatan antariksa.",
     "Stance Markers (arguably, fundamentally imperative), Concessive Clauses (Notwithstanding, Whereas), Bioethics Lexis.",
     "Makalah Position Paper bioetika 500 kata & transkrip dialog meja bundar ilmiah."),

    ("10", "Kelas XII TL - Sem 5", "TP 12.TL.2", "Academic Research Essay Writing & Source Synthesis",
     "Menyimak kuliah metodologi penelitian universitas mengenai integrasi multi-sumber literatur dan pencegahan plagiasi.",
     "Mempertahankan working thesis dan bukti sintesis literatur dalam sesi konsultasi akademik kemitraan (peer review).",
     "Menelaah artikel jurnal terakreditasi (mikroplastik, stabilitas jaringan energi terbarukan, uji coba UBI empiris).",
     "Menyusun esai penelitian akademik (Academic Synthesis Essay) 1.000 kata dengan sitasi APA/MLA dan daftar pustaka.",
     "Parenthetical In-text Citations (APA 7th), Reporting Verbs (hypothesized, established), Synthesizing Cohesive Links.",
     "Makalah sintesis penelitian akademik 1.000 kata lengkap dengan daftar pustaka APA."),

    ("11", "Kelas XII TL - Sem 5", "TP 12.TL.3", "Dramatic Literature & Theatrical Script Performance",
     "Menyimak rekaman pementasan teater Broadway (The Crucible, A Raisin in the Sun) mencermati intonasi tokoh.",
     "Mementaskan adegan monolog dramatik atau sandiwara panggung 3-4 menit dengan modulasi vokal dan subteks emosi.",
     "Membedah naskah drama klasik/modern (Arthur Miller, Lorraine Hansberry, Shakespeare's Hamlet, stage directions).",
     "Mengadaptasi naskah drama atau menulis naskah monolog kontemporer 300 kata lengkap dengan petunjuk panggung.",
     "Dramatic Dialogue Conventions, Stage Direction Formats, Subtextual Ellipses, Prosodic Vocal Cues.",
     "Pementasan panggung monolog teatrikal 3-4 menit & naskah drama dengan anotasi vokal."),

    ("12", "Kelas XII TL - Sem 5", "TP 12.TL.4", "Persuasive Oratory & Parliamentary Debating",
     "Menyimak pidato-pidato oratori bersejarah pemimpin dunia dan debat parlemen kejuaraan dunia (WUDC).",
     "Membawakan pidato oratori persuasif di podium kelas (3-4 menit) dan bertanding dalam simulasi debat parlemen.",
     "Mendekonstruksi retorika klasik (Ethos, Pathos, Logos) dan perangkat gaya bahasa (tricolon, parallelism, antithesis).",
     "Menyusun naskah pidato oratori formal 300-400 kata dan merancang berkas kasus afirmasi/oposisi debat parlemen.",
     "Classical Rhetoric Triangle, Tricolon, Antithesis, Parallelism, Parliamentary Signposting & Rebuttal Syntax.",
     "Pidato oratori formal di podium kelas & unjuk kerja simulasi debat parlemen."),

    # KELAS XII TL - SEMESTER 6
    ("13", "Kelas XII TL - Sem 6", "TP 12.TL.5", "Professional Admissions & Career Dossier",
     "Menyimak panduan seleksi beasiswa dunia (Oxford Rhodes, Fulbright, LPDP) dan wawancara admisi universitas terkemuka.",
     "Menyampaikan personal elevator pitch 90 detik di hadapan komite seleksi beasiswa internasional simulasi.",
     "Membedah contoh personal statement sukses mahasiswa dunia, resume ATS-friendly, dan executive cover letter PBB.",
     "Menyusun berkas dossier lengkap: Executive Cover Letter, ATS Resume (Google X-Y-Z formula), dan Personal Statement 600 kata.",
     "Executive Register, Google X-Y-Z Action Formula, Dynamic Action Verbs, ATS Keyword Architecture.",
     "Berkas Dossier lengkap (Cover Letter, ATS CV satu kolom, Personal Statement beasiswa)."),

    ("14", "Kelas XII TL - Sem 6", "TP 12.TL.6", "Job Interview & Academic Oral Defense Simulation",
     "Menyimak rekaman ujian sidang skripsi universitas (viva voce) dan wawancara perilaku (behavioral interview).",
     "Menampilkan simulasi wawancara metode STAR dan mempertahankan karya ilmiah lisan di hadapan dewan penguji.",
     "Mempelajari taktik validasi santun (tactical validation), jeda taktis (tactical fillers), dan penanganan kritik metodologi.",
     "Menuliskan naskah skenario jawaban STAR dan ringkasan pembelaan metodologi penelitian untuk dewan penguji.",
     "STAR Framework Architecture, Spoken Tactical Fillers, Diplomatic Delay Formulae, Academic Composure Lexis.",
     "Simulasi wawancara kerja eksekutif metode STAR & sidang pertahanan karya ilmiah (viva voce)."),

    ("15", "Kelas XII TL - Sem 6", "TP 12.TL.7", "Research Abstract & Scientific Infographic Poster",
     "Menyimak presentasi flash talk 2 menit para peneliti muda di ajang simposium sains internasional (ISEF).",
     "Mempresentasikan poster infografis saintifik (flash talk 2 menit) dan melayani tanya-jawab di stan simposium kelas.",
     "Membedah naskah abstrak jurnal Scopus (struktur IMRaD: Background, Methods, Results, Conclusion) dan Ten-Second Rule.",
     "Menulis naskah abstrak terstruktur maksimal 250 kata dengan nominalisasi padat dan mendesain poster infografis 3 kolom.",
     "IMRaD Structural Sections, Scientific Nominalization, Active Headline Synthesis, Visual Data Hierarchy.",
     "Naskah abstrak penelitian terstruktur 250 kata & stan gelar simposium poster ilmiah."),

    ("16", "Kelas XII TL - Sem 6", "TP 12.TL.8", "Standardized English Proficiency Readiness (B2+)",
     "Menyimak latihan dialog dan monolog akademik terstandar internasional (format IELTS / TOEFL iBT / DET).",
     "Menjawab pertanyaan wawancara kemahiran berbicara terstandar (IELTS Speaking Part 1, 2, 3) secara spontan dan lancar.",
     "Menerapkan teknik membaca cepat (skimming, scanning, inferencing) pada teks akademik paleoklimatologi es kutub.",
     "Menulis laporan analisis data grafik matematis (IELTS Academic Task 1) 150-180 kata dalam batasan waktu 20 menit tepat.",
     "Standardized Inversion Syntax, Trend Movement Verbs, Proportional Data Prepositions, Timed Exam Cohesion.",
     "Simulasi Benchmark Kemahiran Terstandar Internasional (IELTS Academic Reading & Writing Task 1).")
]


# ==============================================================================
# BUILD SHEET 1: RINGKASAN PETA MATERI
# ==============================================================================
def build_sheet_overview(ws):
    ws.title = "Ringkasan_Peta_Materi"
    apply_title(ws, "PETA MATERI & 4 KETERAMPILAN BAHASA INGGRIS SMA",
                "Ikhtisar Komprehensif Kurikulum Merdeka 2025/2026: Bahasa Inggris Wajib vs Bahasa Inggris Tingkat Lanjut", 8)

    # Section 1: Perbedaan Esensial
    r = 4
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws.cell(row=r, column=1, value="I. MATRIKS PERBEDAAN ESENSIAL: BAHASA INGGRIS WAJIB VS TINGKAT LANJUT").font = FONT_SEC_HEADER
    for c in range(1, 9):
        ws.cell(row=r, column=c).fill = FILL_NAVY
    ws.row_dimensions[r].height = 24

    comp_headers = ["Dimensi Kurikuler", "Bahasa Inggris Wajib (General English)", "Bahasa Inggris Tingkat Lanjut (Advanced English)", "Dampak pada 4 Keterampilan Bahasa"]
    r += 1
    ws.row_dimensions[r].height = 24
    col_spans = [(1, 2), (3, 4), (5, 6), (7, 8)]
    for idx, (cs, ce) in enumerate(col_spans):
        ws.merge_cells(start_row=r, start_column=cs, end_row=r, end_column=ce)
        cell = ws.cell(row=r, column=cs, value=comp_headers[idx])
        cell.font = FONT_TH
        cell.alignment = ALIGN_CENTER
        for c in range(cs, ce + 1):
            ws.cell(row=r, column=c).fill = FILL_ROYAL
            ws.cell(row=r, column=c).border = CELL_BORDER

    comp_rows = [
        ("Fokus & Sasaran Kompetensi", 
         "Kemahiran komunikasi fungsional sehari-hari, literasi teks umum masyarakat, dan kesiapan vokasional dasar.",
         "Kemahiran diskursus akademik (English for Academic Purposes), apresiasi karya sastra mendalam, dan literasi riset ilmiah.",
         "Wajib berfokus pada kelancaran komunikasi praktis; Tingkat Lanjut memperdalam akurasi bernalar kritis dan register formal tinggi."),
        ("Sasaran Fase & Alokasi Waktu",
         "Fase E (Kelas X: 72 JP/th) & Fase F Umum (Kelas XI: 72 JP/th, Kelas XII: 64 JP/th) = Total 208 JP.",
         "Fase F Pilihan (Kelas XI: 180 JP/th @ 5 JP/mgg, Kelas XII: 160 JP/th @ 5 JP/mgg) = Total 340 JP.",
         "Alokasi Tingkat Lanjut 2.5 kali lebih intensif per minggu, memungkinkan proyek penulisan 1.000 kata dan pementasan teater."),
        ("Standar Level Kemahiran CEFR",
         "Kelas X: CEFR A2+ s.d. B1\nKelas XI - XII Wajib: CEFR B1 s.d. B1+",
         "Kelas XI TL: CEFR B1+ s.d. B2\nKelas XII TL: CEFR B2 s.d. B2+ (International Readiness)",
         "Tingkat Lanjut melatih kesiapan tes kemahiran global (IELTS Band 6.5 - 7.5+ / TOEFL iBT 90 - 100+)."),
        ("Karakteristik Teks & Bacaan",
         "Teks otentik praktis: brosur wisata, resep, dongeng rakyat, berita transportasi, surat lamaran umum.",
         "Teks jurnal ilmiah terakreditasi (Nature, Lancet), naskah drama panggung (Shakespeare, Miller), dan dokumen kebijakan PBB.",
         "Wajib menyajikan wacana 150-300 kata; Tingkat Lanjut menganalisis teks kompleks 400-800 kata dengan sitasi."),
        ("Kompleksitas Tata Bahasa & Diksi",
         "Simple tenses, modalitas saran umum, konjungsi kronologis, passive voice standar.",
         "Academic Word List (AWL), nominalisasi saintifik padat, klausa konsesif, inversi gramatikal formal, stance markers.",
         "Keterampilan menulis Wajib menghasilkan esai 4 paragraf; Tingkat Lanjut menyusun naskah akademik IMRaD dan sintesis."),
        ("Model Asesmen & Unjuk Kerja",
         "Storytelling dongeng, simulasi pembaca berita TV, presentasi brosur wisata, tes pilihan ganda fungsional.",
         "Debat parlemen resmi, pementasan drama monolog teatrikal, sidang pertahanan skripsi lisan (viva voce), esai 1.000 kata.",
         "Asesmen Tingkat Lanjut menguji ketahanan berpikir spontan lisan di bawah tekanan dan orisinalitas riset mandiri.")
    ]

    r += 1
    for r_data in comp_rows:
        ws.row_dimensions[r].height = 42
        for idx, (cs, ce) in enumerate(col_spans):
            ws.merge_cells(start_row=r, start_column=cs, end_row=r, end_column=ce)
            cell = ws.cell(row=r, column=cs, value=r_data[idx])
            cell.font = FONT_DATA_BOLD if idx == 0 else FONT_DATA
            cell.alignment = ALIGN_CENTER if idx == 0 else ALIGN_LEFT
            for c in range(cs, ce + 1):
                cell_item = ws.cell(row=r, column=c)
                cell_item.border = CELL_BORDER
                if r % 2 == 1:
                    cell_item.fill = FILL_ZEBRA
        r += 1

    r += 1
    # Section 2: Ringkasan 4 Keterampilan
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws.cell(row=r, column=1, value="II. PEMETAAN INTI 4 KETERAMPILAN BAHASA (LISTENING, SPEAKING, READING, WRITING)").font = FONT_SEC_HEADER
    for c in range(1, 9):
        ws.cell(row=r, column=c).fill = FILL_NAVY
    ws.row_dimensions[r].height = 24

    sk_headers = ["Keterampilan Bahasa", "Hakikat & Elemen Kurikulum Merdeka", "Fokus Bahasa Inggris Wajib (X, XI, XII)", "Fokus Bahasa Inggris Tingkat Lanjut (XI, XII)"]
    r += 1
    ws.row_dimensions[r].height = 24
    for idx, (cs, ce) in enumerate(col_spans):
        ws.merge_cells(start_row=r, start_column=cs, end_row=r, end_column=ce)
        cell = ws.cell(row=r, column=cs, value=sk_headers[idx])
        cell.font = FONT_TH
        cell.alignment = ALIGN_CENTER
        for c in range(cs, ce + 1):
            ws.cell(row=r, column=c).fill = FILL_TEAL
            ws.cell(row=r, column=c).border = CELL_BORDER

    sk_rows = [
        ("1) LISTENING\n(Menyimak)",
         "Kemampuan reseptif lisan: memahami ide pokok, detail informasi tersurat/tersirat, nada pembicara, dan tujuan komunikatif dari teks audio/multimodal.",
         "• Menyimak monolog atlet, dongeng cerita rakyat, siniar hidup sehat, dan siaran berita radio/TV.\n• Mengidentifikasi 5W+1H dan instruksi prosedural teratur.\n• Kalibrasi: Audio bertempo sedang (CEFR B1).",
         "• Menyimak kuliah akademik universitas dunia, rekaman debat PBB, dialog drama teater Broadway, dan sidang tesis.\n• Mengidentifikasi bias data, subteks psikologis, dan falasi logika.\n• Kalibrasi: Audio otentik bertempo alami (CEFR B2+)."),
        ("2) SPEAKING\n(Berbicara)",
         "Kemampuan produktif lisan: menyampaikan ide, berdialog interaktif, berorasi di podium, dan mempertahankan argumen dengan kelancaran serta artikulasi akurat.",
         "• Mendeskripsikan destinasi wisata, mempraktikkan storytelling dongeng, dan simulasi presenter berita TV.\n• Mengemukakan pendapat dalam diskusi kelas dan elevator pitch karier.\n• Target: Kelancaran lisan komunikatif tanpa hambatan.",
         "• Berpidato oratori formal (Ethos, Pathos, Logos), debat parlemen WUDC, pementasan drama monolog teatrikal.\n• Menghadapi simulasi wawancara STAR dan sidang viva voce lisan.\n• Target: Wibawa eksekutif, jeda taktis, dan artikulasi presisi."),
        ("3) READING\n(Membaca & Memirsa)",
         "Kemampuan reseptif tulis & visual: mendekonstruksi struktur genre teks, menelaah fitur bahasa, menemukan inferensi tersembunyi, dan mengevaluasi kredibilitas pesan.",
         "• Membaca teks deskripsi, recount sejarah, dongeng, manual prosedur, berita aktual, dan reviu karya seni.\n• Menganalisis urutan peristiwa, kalimat pasif, dan kalimat perintah.\n• Target: Pemahaman membaca kritis teks umum 150-300 kata.",
         "• Membaca artikel jurnal terindeks Scopus (Nature, Elsevier), naskah sastra klasik (Shakespeare, Miller), dan IMRaD abstrak.\n• Menerapkan teknik membaca cepat skimming, scanning, dan inferensi.\n• Target: Dekonstruksi wacana ilmiah padat 400-800 kata."),
        ("4) WRITING\n(Menulis & Presentasi)",
         "Kemampuan produktif tulis & multimodal: menuangkan gagasan ke dalam teks terstruktur, memilih diksi formal yang kaya, menggunakan grammar akurat, serta mendesain media komunikasi.",
         "• Menyusun brosur wisata digital, jurnal memoar pribadi, panduan infografis, teks diskusi isu sekolah, dan surat lamaran kerja.\n• Menguasai tenses dasar, konjungsi kontras, dan format surat formal.\n• Target: Naskah esai 150-250 kata terstruktur dan rapi.",
         "• Menyusun makalah sintesis penelitian akademik 1.000 kata dengan sitasi APA 7th Edition, Position Paper, cerpen 750 kata.\n• Menulis abstrak 250 kata, berkas dossier (ATS resume), dan IELTS Task 1.\n• Target: Karya tulis ilmiah/kreatif bebas plagiasi.")
    ]

    r += 1
    for r_data in sk_rows:
        ws.row_dimensions[r].height = 68
        for idx, (cs, ce) in enumerate(col_spans):
            ws.merge_cells(start_row=r, start_column=cs, end_row=r, end_column=ce)
            cell = ws.cell(row=r, column=cs, value=r_data[idx])
            cell.font = FONT_DATA_BOLD if idx in [0, 1] else FONT_DATA
            cell.alignment = ALIGN_CENTER if idx == 0 else ALIGN_LEFT
            for c in range(cs, ce + 1):
                cell_item = ws.cell(row=r, column=c)
                cell_item.border = CELL_BORDER
                if r % 2 == 1:
                    cell_item.fill = FILL_ZEBRA
        r += 1

    auto_fit_columns(ws, min_w=12, max_w=46)


# ==============================================================================
# BUILD SHEET 2: BAHASA INGGRIS WAJIB (TABEL SEDERHANA 4 SKILLS)
# ==============================================================================
def build_sheet_wajib(ws):
    ws.title = "Bahasa_Inggris_Wajib"
    apply_title(ws, "PEMETAAN MATERI 4 KETERAMPILAN: BAHASA INGGRIS UMUM (WAJIB)",
                "Kelas X, XI, dan XII SMA | Kurikulum Merdeka (Kepka BSKAP No. 020 Tahun 2026)", 10)

    headers = [
        "No", "Kelas & Sem", "Kode TP", "Jenis Teks (Genre) & Tema Kontekstual",
        "1) LISTENING (Menyimak)", "2) SPEAKING (Berbicara)",
        "3) READING (Membaca & Memirsa)", "4) WRITING (Menulis & Mempresentasikan)",
        "Tata Bahasa & Kosakata Kunci", "Target Produk / Bukti Belajar"
    ]

    r = 4
    ws.row_dimensions[r].height = 28
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=r, column=col_idx, value=h)
        cell.font = FONT_TH
        cell.fill = FILL_HEADER_WAJIB
        cell.alignment = ALIGN_CENTER
        cell.border = CELL_BORDER

    # Track sections to add class header rows
    current_class = ""
    r += 1
    for row_data in DATA_WAJIB_4_SKILLS:
        cls_sem = row_data[1]
        cls_name = cls_sem.split(' - ')[0]
        if cls_name != current_class:
            current_class = cls_name
            # Section header for Class
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
            sec_cell = ws.cell(row=r, column=1, value=f"✦ JENJANG {current_class.upper()} (FASE {'E' if 'X' in current_class and not 'XI' in current_class else 'F UMUM'})")
            sec_cell.font = Font(name=FONT_NAME, size=10, bold=True, color="1E3A8A")
            sec_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            for c in range(1, 11):
                ws.cell(row=r, column=c).fill = FILL_SECTION_WAJIB
                ws.cell(row=r, column=c).border = CELL_BORDER
            ws.row_dimensions[r].height = 22
            r += 1

        ws.row_dimensions[r].height = 54
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=r, column=col_idx, value=val)
            cell.font = FONT_DATA_BOLD if col_idx in [1, 2, 3] else FONT_DATA
            cell.alignment = ALIGN_CENTER_TOP if col_idx in [1, 2, 3] else ALIGN_LEFT
            cell.border = CELL_BORDER
            if r % 2 == 1:
                cell.fill = FILL_ZEBRA

        r += 1

    auto_fit_columns(ws, min_w=8, max_w=42)
    ws.freeze_panes = "E5"


# ==============================================================================
# BUILD SHEET 3: BAHASA INGGRIS TINGKAT LANJUT (TABEL SEDERHANA 4 SKILLS)
# ==============================================================================
def build_sheet_tl(ws):
    ws.title = "Bahasa_Inggris_Tingkat_Lanjut"
    apply_title(ws, "PEMETAAN MATERI 4 KETERAMPILAN: BAHASA INGGRIS TINGKAT LANJUT (PILIHAN)",
                "Kelas XI dan XII SMA Fase F Tingkat Lanjut | Kurikulum Merdeka (Kepka BSKAP No. 020 Tahun 2026)", 10)

    headers = [
        "No", "Kelas & Sem", "Kode TP", "Jenis Teks (Genre) & Tema Akademik",
        "1) LISTENING (Menyimak Akademik)", "2) SPEAKING (Berbicara / Oratori / Performansi)",
        "3) READING (Membaca & Analisis Wacana)", "4) WRITING (Menulis Ilmiah & Portofolio)",
        "Tata Bahasa & Kosakata Kompleks", "Target Produk / Bukti Belajar"
    ]

    r = 4
    ws.row_dimensions[r].height = 28
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=r, column=col_idx, value=h)
        cell.font = FONT_TH
        cell.fill = FILL_HEADER_TL
        cell.alignment = ALIGN_CENTER
        cell.border = CELL_BORDER

    current_class = ""
    r += 1
    for row_data in DATA_TL_4_SKILLS:
        cls_sem = row_data[1]
        cls_name = cls_sem.split(' - ')[0]
        if cls_name != current_class:
            current_class = cls_name
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
            sec_cell = ws.cell(row=r, column=1, value=f"✦ JENJANG {current_class.upper()} (FASE F TINGKAT LANJUT / PEMINATAN)")
            sec_cell.font = Font(name=FONT_NAME, size=10, bold=True, color="065F46")
            sec_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            for c in range(1, 11):
                ws.cell(row=r, column=c).fill = FILL_SECTION_TL
                ws.cell(row=r, column=c).border = CELL_BORDER
            ws.row_dimensions[r].height = 22
            r += 1

        ws.row_dimensions[r].height = 56
        for col_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row=r, column=col_idx, value=val)
            cell.font = FONT_DATA_BOLD if col_idx in [1, 2, 3] else FONT_DATA
            cell.alignment = ALIGN_CENTER_TOP if col_idx in [1, 2, 3] else ALIGN_LEFT
            cell.border = CELL_BORDER
            if r % 2 == 1:
                cell.fill = FILL_ZEBRA

        r += 1

    auto_fit_columns(ws, min_w=8, max_w=42)
    ws.freeze_panes = "E5"


# ==============================================================================
# BUILD SHEET 4 s.d. 7: SHEET KHUSUS PER SKILL (LISTENING, SPEAKING, READING, WRITING)
# ==============================================================================
def build_skill_specific_sheet(ws, title_skill, skill_idx, skill_desc):
    apply_title(ws, f"MATERI KHUSUS KETERAMPILAN: {title_skill.upper()}",
                f"{skill_desc} | Pemisahan Komparatif Bahasa Inggris Wajib & Bahasa Inggris Tingkat Lanjut", 8)

    skill_headers = [
        "No", "Jenjang / Kelas", "Semester", "Kode TP",
        "Genre Teks & Topik Bahasan", f"Fokus Materi {title_skill}",
        "Tata Bahasa / Diksi Pendukung", "Aktivitas Belajar & Asesmen Kunci"
    ]

    r = 4
    # Section Wajib
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    sec_wajib = ws.cell(row=r, column=1, value=f"BAGIAN I: BAHASA INGGRIS UMUM (WAJIB) - {title_skill.upper()} KELAS X, XI, XII")
    sec_wajib.font = FONT_SEC_HEADER
    for c in range(1, 9):
        ws.cell(row=r, column=c).fill = FILL_HEADER_WAJIB
    ws.row_dimensions[r].height = 24

    r += 1
    ws.row_dimensions[r].height = 24
    for col_idx, h in enumerate(skill_headers, 1):
        cell = ws.cell(row=r, column=col_idx, value=h)
        cell.font = FONT_TH
        cell.fill = FILL_NAVY
        cell.alignment = ALIGN_CENTER
        cell.border = CELL_BORDER

    r += 1
    for row_data in DATA_WAJIB_4_SKILLS:
        ws.row_dimensions[r].height = 46
        # Map: No, Kelas, Sem, Kode TP, Genre, Skill Content, Grammar, Produk
        cls_parts = row_data[1].split(' - ')
        cls_name = cls_parts[0]
        sem_name = cls_parts[1] if len(cls_parts) > 1 else ""
        vals = [
            row_data[0], cls_name, sem_name, row_data[2],
            row_data[3], row_data[skill_idx], row_data[8], row_data[9]
        ]
        for c_idx, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c_idx, value=v)
            cell.font = FONT_DATA_BOLD if c_idx in [1, 2, 3, 4] else FONT_DATA
            cell.alignment = ALIGN_CENTER_TOP if c_idx in [1, 2, 3, 4] else ALIGN_LEFT
            cell.border = CELL_BORDER
            if r % 2 == 1:
                cell.fill = FILL_ZEBRA
        r += 1

    r += 1
    # Section Tingkat Lanjut
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    sec_tl = ws.cell(row=r, column=1, value=f"BAGIAN II: BAHASA INGGRIS TINGKAT LANJUT (PILIHAN) - {title_skill.upper()} KELAS XI & XII")
    sec_tl.font = FONT_SEC_HEADER
    for c in range(1, 9):
        ws.cell(row=r, column=c).fill = FILL_HEADER_TL
    ws.row_dimensions[r].height = 24

    r += 1
    ws.row_dimensions[r].height = 24
    for col_idx, h in enumerate(skill_headers, 1):
        cell = ws.cell(row=r, column=col_idx, value=h)
        cell.font = FONT_TH
        cell.fill = FILL_SLATE
        cell.alignment = ALIGN_CENTER
        cell.border = CELL_BORDER

    r += 1
    for row_data in DATA_TL_4_SKILLS:
        ws.row_dimensions[r].height = 48
        cls_parts = row_data[1].split(' - ')
        cls_name = cls_parts[0]
        sem_name = cls_parts[1] if len(cls_parts) > 1 else ""
        vals = [
            row_data[0], cls_name, sem_name, row_data[2],
            row_data[3], row_data[skill_idx], row_data[8], row_data[9]
        ]
        for c_idx, v in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c_idx, value=v)
            cell.font = FONT_DATA_BOLD if c_idx in [1, 2, 3, 4] else FONT_DATA
            cell.alignment = ALIGN_CENTER_TOP if c_idx in [1, 2, 3, 4] else ALIGN_LEFT
            cell.border = CELL_BORDER
            if r % 2 == 1:
                cell.fill = FILL_ZEBRA
        r += 1

    auto_fit_columns(ws, min_w=8, max_w=46)
    ws.freeze_panes = "F6"


# ==============================================================================
# MAIN FUNCTION
# ==============================================================================
def main():
    target_dir = r"c:\Users\LILI\OneDrive - PDK Jateng\Dokumen\English-SMA\4-Silabus"
    os.makedirs(target_dir, exist_ok=True)
    out_file = os.path.join(target_dir, "Materi Bahasa Inggris SMA.xlsx")

    print(f"Creating workbook: {out_file}")
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove default sheet

    # 1. Sheet Ringkasan Peta Materi
    ws1 = wb.create_sheet(title="Ringkasan_Peta_Materi")
    build_sheet_overview(ws1)

    # 2. Sheet Bahasa Inggris Wajib
    ws2 = wb.create_sheet(title="Bahasa_Inggris_Wajib")
    build_sheet_wajib(ws2)

    # 3. Sheet Bahasa Inggris Tingkat Lanjut
    ws3 = wb.create_sheet(title="Bahasa_Inggris_Tingkat_Lanjut")
    build_sheet_tl(ws3)

    # 4. Sheet 1_Listening
    ws4 = wb.create_sheet(title="1_Listening")
    build_skill_specific_sheet(
        ws4, "Listening (Menyimak)", 4,
        "Fokus Reseptif Lisan: Pemahaman Ide Utama, Informasi Rinci, dan Sikap Pembicara"
    )

    # 5. Sheet 2_Speaking
    ws5 = wb.create_sheet(title="2_Speaking")
    build_skill_specific_sheet(
        ws5, "Speaking (Berbicara)", 5,
        "Fokus Produktif Lisan: Dialog Interaktif, Monolog, Oratori Pidato, Debat & Wawancara"
    )

    # 6. Sheet 3_Reading
    ws6 = wb.create_sheet(title="3_Reading")
    build_skill_specific_sheet(
        ws6, "Reading (Membaca & Memirsa)", 6,
        "Fokus Reseptif Tulis & Multimodal: Dekonstruksi Struktur Teks, Inferensi & Evaluasi Kritis"
    )

    # 7. Sheet 4_Writing
    ws7 = wb.create_sheet(title="4_Writing")
    build_skill_specific_sheet(
        ws7, "Writing (Menulis & Presentasi)", 7,
        "Fokus Produktif Tulis & Multimodal: Penulisan Esai Genre, Karya Sastra, Sintesis & Portofolio"
    )

    wb.save(out_file)
    print("Workbook successfully created and saved!")

if __name__ == "__main__":
    main()
