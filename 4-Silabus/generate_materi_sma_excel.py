# -*- coding: utf-8 -*-
"""
Generator script for 'Materi Bahasa Inggris SMA.xlsx' in 4-Silabus.
Peta materi terstruktur dan komprehensif untuk:
1) Listening (Menyimak)
2) Speaking (Berbicara)
3) Reading (Membaca & Memirsa)
4) Writing (Menulis & Mempresentasikan)
5) Grammar (Tata Bahasa & Struktur Kalimat yang Seharusnya untuk Siswa SMA)

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
FILL_PURPLE = PatternFill(start_color="581C87", end_color="581C87", fill_type="solid")
FILL_SLATE = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
FILL_HEADER_WAJIB = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
FILL_HEADER_TL = PatternFill(start_color="065F46", end_color="065F46", fill_type="solid")
FILL_ZEBRA = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
FILL_ACCENT_BLUE = PatternFill(start_color="EFF6FF", end_color="EFF6FF", fill_type="solid")
FILL_ACCENT_GREEN = PatternFill(start_color="ECFDF5", end_color="ECFDF5", fill_type="solid")
FILL_ACCENT_PURPLE = PatternFill(start_color="FAF5FF", end_color="FAF5FF", fill_type="solid")
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
# DATA: BAHASA INGGRIS WAJIB (KELAS X, XI, XII) - 4 SKILLS
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
     "Menyusun teks eksplanasi kausal 200-250 kata menjelaskan mekanisme terjadinya banjir perkotaan atau gempa bumi.",
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
# DATA: BAHASA INGGRIS TINGKAT LANJUT (KELAS XI & XII TL) - 4 SKILLS
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
# DATA: MATERI GRAMMAR & STRUKTUR KALIMAT (SCOPE & SEQUENCE)
# ==============================================================================
DATA_GRAMMAR_WAJIB = [
    # KELAS X - SEMESTER 1
    ("1", "Kelas X", "Semester 1", "Simple Present Tense & Subject-Verb Agreement",
     "Menyatakan fakta ilmiah umum, rutinitas, dan deskripsi permanen atlet serta objek wisata.",
     "Verbal: S + V1(-s/-es) + O\nNominal: S + is/am/are + Complement",
     "Cristiano Ronaldo trains relentlessly every morning to maintain peak physical endurance.",
     "TP 10.1 & TP 10.2: Descriptive Text (Writing & Speaking)"),

    ("2", "Kelas X", "Semester 1", "Adjectives & Noun Phrases (OSASCOMP Order)",
     "Mendeskripsikan ciri fisik, karakter, dan daya tarik visual destinasi wisata secara spesifik dan hidup.",
     "Opinion + Size + Age + Shape + Color + Origin + Material + Purpose + Noun",
     "A magnificent ancient stone temple stands peacefully amidst the lush green tropical mountains.",
     "TP 10.1 & TP 10.2: Descriptive Text & Brosur Wisata"),

    ("3", "Kelas X", "Semester 1", "Simple Past Tense (Regular & Irregular Verbs)",
     "Menceritakan peristiwa masa lampau secara kronologis dalam teks recount biografi sejarah & memoar.",
     "Verbal: S + V2 + O/Adv\nNominal: S + was/were + Complement",
     "On November 10, 1945, Indonesian freedom fighters defended Surabaya with unwavering courage.",
     "TP 10.3: Recount Text Sejarah (Reading & Writing)"),

    ("4", "Kelas X", "Semester 1", "Past Continuous Tense & Time Clauses (When/While)",
     "Menyatakan aksi yang sedang berlangsung di masa lampau ketika disela oleh peristiwa lain.",
     "S + was/were + V-ing + while/when + S + V2",
     "While Bung Tomo was delivering his radio speech, citizens were gathering across the city square.",
     "TP 10.3: Recount Memoar & Linimasa (Writing & Listening)"),

    # KELAS X - SEMESTER 2
    ("5", "Kelas X", "Semester 2", "Direct & Indirect Speech (Reported Speech Basics)",
     "Melaporkan percakapan langsung para tokoh cerita rakyat/fabel dengan pergeseran tenses dan kata ganti.",
     "Direct: S + said, \"[Quote]\"\nIndirect: S + said that + S + V(shifted)",
     "The wise old hermit told the greedy prince that arrogance would inevitably destroy his kingdom.",
     "TP 10.4: Narrative Text & Storytelling (Speaking & Writing)"),

    ("6", "Kelas X", "Semester 2", "Imperative Sentences & Sequencing Adverbials",
     "Memberikan instruksi langkah demi langkah, panduan, dan larangan keselamatan berinternet & pola hidup sehat.",
     "Infinitive: Verb 1 + Object\nNegative: Do not (Don't) + Verb 1 + Object\nConnectors: First, Next, After that, Finally",
     "First, create a multi-layered alphanumeric password. Never disclose your personal credentials to strangers.",
     "TP 10.5: Procedure Text & Manual Digital (Writing & Reading)"),

    ("7", "Kelas X", "Semester 2", "Simple Present for Scientific Truths & General Nouns",
     "Mendeskripsikan siklus hidup fauna endemik dan klasifikasi ilmiah dalam teks laporan observasi faktual.",
     "Plural Noun + V1 / Singular Noun + V1(-s/-es) (Timeless present)",
     "Komodo dragons primarily inhabit Komodo and Rinca islands, relying on potent venom to subdue prey.",
     "TP 10.6: Report Text Observasi Satwa (Reading & Writing)"),

    ("8", "Kelas X", "Semester 2", "Passive Voice in Present Tense",
     "Menitikberatkan proses atau objek tindakan daripada pelaku aksi dalam laporan ilmiah dan proses teknis.",
     "S + is/am/are + V3 (Past Participle) + (by Agent)",
     "Organic municipal waste is separated into specialized processing units and converted into compost.",
     "TP 10.6: Report Text Sains & Lingkungan (Reading & Writing)"),

    # KELAS XI - SEMESTER 3
    ("9", "Kelas XI", "Semester 3", "Internal Conjunctions & Addition Connectors",
     "Menghubungkan argumen penjelas secara kohesif untuk memperkuat tesis dalam esai eksposisi analitis.",
     "Furthermore / Moreover / In addition / Aside from that, + Sentence",
     "Furthermore, chronic overexposure to smartphone screens impairs adolescent melatonin production.",
     "TP 11.W.1: Analytical Exposition (Writing & Reading)"),

    ("10", "Kelas XI", "Semester 3", "Causal Connectors & Cause-and-Effect Markers",
     "Menjelaskan relasi kausal dan penalaran logis yang mendasari fenomena literasi digital dan kesehatan.",
     "Clause + therefore / consequently / as a result, + Clause\nDue to / Because of + Noun Phrase",
     "Due to excessive digital screen time, students experience fatigue; consequently, their academic focus declines.",
     "TP 11.W.1: Analytical Exposition (Writing & Speaking)"),

    ("11", "Kelas XI", "Semester 3", "Modal Auxiliaries of Obligation & Recommendation",
     "Menyatakan rekomendasi persuasif, desakan moral, dan usulan aksi nyata dalam eksposisi hortatori.",
     "S + should / ought to / must / had better + Verb 1 (Infinitive)",
     "School administrators must enforce a strict ban on single-use plastics across all cafeteria stalls.",
     "TP 11.W.2: Hortatory Exposition (Writing & Speaking)"),

    ("12", "Kelas XI", "Semester 3", "Past Perfect Tense & Relative Clauses",
     "Menunjukkan kejadian yang telah tuntas sebelum peristiwa lampau lain dan merincikan karakter cerita kontemporer.",
     "S + had + V3 + before + S + V2\nRelative: who / which / that / whose + Clause",
     "The villagers had already vacated the valley before the river broke its banks, which saved hundreds of lives.",
     "TP 11.W.3: Contemporary Narrative & Fractured Tales (Writing & Reading)"),

    # KELAS XI - SEMESTER 4
    ("13", "Kelas XI", "Semester 4", "Passive Voice in Complex & Modal Structures",
     "Menjelaskan mekanisme ilmiah proses alam (pemanasan global, gempa) secara objektif dan impersonal.",
     "S + has/have been + V3\nS + modal (can/must/may) + be + V3",
     "Excessive carbon dioxide has been trapped within the atmosphere and cannot be naturally absorbed quickly.",
     "TP 11.W.4: Explanation Text Fenomena Alam (Reading & Writing)"),

    ("14", "Kelas XI", "Semester 4", "Contrast Connectors & Concessive Clauses",
     "Menyajikan dua sudut pandang pro dan kontra secara berimbang dan adil dalam teks diskusi isu sekolah.",
     "Although / Even though / Despite the fact that + Clause\nOn the one hand..., on the other hand... / However",
     "Although early morning start times align with adult work schedules, they severely curtail adolescent rest.",
     "TP 11.W.5: Discussion Text Aturan Sekolah (Writing & Speaking)"),

    ("15", "Kelas XI", "Semester 4", "Evaluative Adjectives & Appraising Diction",
     "Memberikan evaluasi kritis, interpretasi estetika, dan rekomendasi objektif terhadap karya seni/novel/film.",
     "Adjective of Evaluation + Noun\nLinking Verb + extraordinarily / profoundly + Adjective",
     "The director delivers a visually arresting and emotionally nuanced masterpiece that questions human prejudice.",
     "TP 11.W.6: Review Text Seni & Sinema (Reading & Writing)"),

    ("16", "Kelas XI", "Semester 4", "Conditional Sentences (Type 1 & Type 2)",
     "Mengemukakan skenario kemungkinan masa depan realistis (Type 1) dan situasi pengandaian hipotetis (Type 2).",
     "Type 1: If + Simple Present, S + will + V1\nType 2: If + Simple Past, S + would + V1",
     "If the school council allocated more budget to library facilities, students would read more enthusiastically.",
     "TP 11.W.5: Discussion & Rekomendasi Solusi (Speaking & Writing)"),

    # KELAS XII - SEMESTER 5
    ("17", "Kelas XII", "Semester 5", "Hedging Devices & Cautious Academic Stance",
     "Menyatakan opini secara santun, berhati-hati, dan tidak mutlak dalam teks diskusi ketenagakerjaan dan AI.",
     "It appears that / The evidence indicates that / S + tends to + Verb / arguably",
     "Artificial intelligence arguably tends to augment human productivity rather than outright eliminate entry-level roles.",
     "TP 12.W.1: Discussion Text Otomasi Dunia Kerja (Writing & Reading)"),

    ("18", "Kelas XII", "Semester 5", "Reported Speech with Advanced Reporting Verbs",
     "Mengutip pernyataan narasumber dalam teks berita jurnalistik secara variatif dan akurat.",
     "S + claimed / affirmed / confirmed / warned / maintained + that + Clause",
     "The transport authority confirmed that the express rail system had reached full operational reliability.",
     "TP 12.W.2: News Item Text & Media Literacy (Writing & Speaking)"),

    ("19", "Kelas XII", "Semester 5", "Compound Adjectives & Descriptive Prepositional Phrases",
     "Mengulas keunikan karya seni kriya lokal, detail bahan, dan kualitas pengerjaan produk komersial.",
     "Hyphenated compound adj (hand-woven, eco-friendly) + Noun + with / in + Noun Phrase",
     "The artisan presented exquisitely hand-woven silk fabrics infused with authentic plant-based organic dyes.",
     "TP 12.W.3: Review Text Kerajinan Nusantara (Writing & Reading)"),

    # KELAS XII - SEMESTER 6
    ("20", "Kelas XII", "Semester 6", "Formal Business Register & Salutations",
     "Menulis surat lamaran kerja resmi sesuai konvensi korespondensi profesional internasional.",
     "Formal opening: Dear Hiring Committee,\nFormula: I am writing to express my interest in...\nSign-off: Sincerely yours,",
     "I am writing to respectfully submit my application for the junior administrative officer position at your firm.",
     "TP 12.W.4: Job Application Letter (Writing)"),

    ("21", "Kelas XII", "Semester 6", "Dynamic Action Verbs & Parallel Syntax in CVs",
     "Menyusun butir capaian prestasi resume/CV yang ringkas, terukur, dan berdampak tanpa subjek pribadi 'I'.",
     "Past Action Verb + Noun/Deliverable + Metric/Context (Parallel bullet points)",
     "Coordinated ten community outreach workshops, increasing active youth participation by 35 percent.",
     "TP 12.W.4: Curriculum Vitae Profesional (Writing)"),

    ("22", "Kelas XII", "Semester 6", "Polite Formulaic Language & Spoken Tactical Fillers",
     "Merespons pertanyaan wawancara kerja secara santun, terstruktur, dan memberi jeda berpikir profesional.",
     "Polite openers: Certainly; That is a pertinent question; In my previous experience; I would be privileged to...",
     "Certainly. In my previous role as committee treasurer, I managed financial reconciliations with zero discrepancy.",
     "TP 12.W.5: Job Interview Simulation (Speaking)"),

    ("23", "Kelas XII", "Semester 6", "Prepositions of Data & Trend Movement Verbs",
     "Mendeskripsikan tren grafik, infografis statistik, dan takarir visual media massa secara presisi.",
     "Verbs: increase, decline, fluctuate, stabilize + by [amount] / from [A] to [B] / at [percentage]",
     "Solar energy adoption expanded rapidly by 22 percent between 2021 and 2024, stabilizing at 65 gigawatts.",
     "TP 12.W.6: Captions & Data Graphs (Writing & Reading)")
]

DATA_GRAMMAR_TL = [
    # KELAS XI TL - SEMESTER 3
    ("1", "Kelas XI TL", "Semester 3", "Academic Nominalization (Transforming Verbs/Adj into Nouns)",
     "Mengemas ide ke dalam bentuk kata benda abstrak yang padat informasi, esensial untuk ragam penulisan ilmiah.",
     "Verb -> Noun (analyze -> analysis; implement -> implementation; prioritize -> prioritization)",
     "The systematic implementation of algorithmic oversight prevents the perpetuation of demographic bias.",
     "TP 11.TL.1: Advanced Analytical Exposition (Writing & Reading)"),

    ("2", "Kelas XI TL", "Semester 3", "Complex Nominal Clauses (That-clauses & Wh-clauses)",
     "Menempatkan seluruh klausa sebagai subjek atau objek tata bahasa untuk argumentasi tingkat sarjana.",
     "That + Subject + Verb + acts as Subject / Verb + what + Subject + Verb",
     "That autonomous algorithms frequently reproduce historical societal prejudices constitutes an undeniable bioethical dilemma.",
     "TP 11.TL.1: Jurnal Riset AI Ethics (Reading & Writing)"),

    ("3", "Kelas XI TL", "Semester 3", "High Deontic Modality & Diplomatic Policy Registers",
     "Menyatakan keharusan imperatif, komitmen perjanjian internasional, dan resolusi diplomasi iklim.",
     "It is strictly imperative / mandatory that + S + (should) be + V3\nSignatory states are obligated to + V1",
     "It is mandatory that developing nations receive equitable technology transfers to expedite renewable grid transitions.",
     "TP 11.TL.2: Policy Paper & Model UN (Writing & Speaking)"),

    ("4", "Kelas XI TL", "Semester 3", "Literary Figurative Tropes & Allegorical Syntax",
     "Membangun kedalaman psikologis, simbolisme sastra, dan monolog batin dalam penulisan prosa fiksi.",
     "Extended Metaphor, Allegory, Objective Correlative, Stream of Consciousness",
     "A suffocating silence engulfed the ancient courtyard, whispering unresolved grievances through the trembling autumn leaves.",
     "TP 11.TL.3: Advanced Literary Narrative (Reading & Writing)"),

    ("5", "Kelas XI TL", "Semester 3", "Poetic Inversion & Enjambment in Stylistics",
     "Menciptakan rima berbobot, modulasi prosodik, dan penekanan ritmis pada deklamasi puisi kritik sosial.",
     "Inverted predicate/adverb: Seldom heard was... / High upon the hill stood...\nEnjambment across stanzas",
     "Dark grew the skies above industrial towers, where sorrow wept unheeded beneath the mechanical hum of greed.",
     "TP 11.TL.4: Stylistics, Poetry & Songs (Reading & Speaking)"),

    # KELAS XI TL - SEMESTER 4
    ("6", "Kelas XI TL", "Semester 4", "Multi-Perspective Academic Signposting & Epistemic Hedging",
     "Menavigasi argumen multidimensi dalam teks diskusi akademik model Puskurjar tanpa bias subjektif.",
     "While proponent scholars maintain that..., critical sociologists conversely argue that...\nIt is plausible to posit...",
     "While neuroscientists establish that tactical gaming sharpens spatial reflexes, pediatricians emphasize potential behavioral risks.",
     "TP 11.TL.5: Official Puskurjar Discussion (Reading & Writing)"),

    ("7", "Kelas XI TL", "Semester 4", "Complex Causal Passive Chains & Technical Nominalization",
     "Menguraikan mekanisme kausal sains mutakhir (komputasi kuantum, fusi nuklir) secara impersonal dan presisi.",
     "Plasma is magnetically confined within superconducting coils, thereby initiating fusion under thermonuclear temperatures.",
     "S + is + V3 + by [Agent], thereby inducing [Noun Phrase] which consecutively precipitates [Noun Phrase].",
     "TP 11.TL.6: Science Frontiers Explanation (Writing & Reading)"),

    ("8", "Kelas XI TL", "Semester 4", "Historical Recount Synthesizers & In-text Quotation Integration",
     "Mengintegrasikan kutipan arsip sejarah dan refleksi filosofis kepemimpinan dunia dalam biografi mendalam.",
     "As famously observed by [Leader], \"[Quote]\"—a conviction that irrevocably reshaped modern democratic governance.",
     "As articulated by Nelson Mandela, \"Education is the most powerful weapon,\" an axiom that anchored his post-apartheid vision.",
     "TP 11.TL.7: Biography & Historical Memoirs (Writing & Speaking)"),

    ("9", "Kelas XI TL", "Semester 4", "Semiotic Lexis & Evaluative Subordinate Clauses in Cultural Critique",
     "Mendekonstruksi mise-en-scène sinematografi dan motif artistik dalam esai kuratorial serta telaah film.",
     "By juxtaposing stark chiaroscuro lighting against opulent settings, the auteur underscores the protagonist's moral hypocrisy.",
     "Prepositional phrase of means + Subject + Verb + that-clause of thematic critique",
     "TP 11.TL.8: Critical Review of International Arts (Reading & Writing)"),

    # KELAS XII TL - SEMESTER 5
    ("10", "Kelas XII TL", "Semester 5", "Bioethical Registers & Concessive Clauses of Complex Contrast",
     "Merumuskan argumen berbobot dalam debat bioetika kedokteran CRISPR dan alokasi sumber daya antariksa.",
     "Notwithstanding [Noun Phrase], Subject + Verb\nWhereas [Clause A], [Clause B]",
     "Notwithstanding the therapeutic efficacy of somatic gene therapies, germline alterations present irreversible intergenerational hazards.",
     "TP 12.TL.1: Global Bioethics Discussion (Writing & Speaking)"),

    ("11", "Kelas XII TL", "Semester 5", "In-Text Citation Syntax (APA 7th Edition) & Literature Synthesis",
     "Mensintesis temuan multi-jurnal terindeks Scopus secara etis dan bebas plagiarisme dalam riset akademik 1.000 kata.",
     "According to Author (Year), ...; however, Subsequent Author (Year) disputed this premise, observing that...",
     "Recent clinical trials (Harrison & Chen, 2024; Wijaya, 2025) corroborate the premise that nanoplastics induce cellular cytotoxicity.",
     "TP 12.TL.2: Academic Research Essay Synthesis (Writing & Reading)"),

    ("12", "Kelas XII TL", "Semester 5", "Dramatic Dialogue Conventions & Subtextual Ellipses",
     "Menulis naskah sandiwara panggung dengan instruksi panggung (stage directions) dan jeda bicara tersirat.",
     "CHARACTER NAME (in uppercase) + (stage direction in parentheses/italics) + Dialogue with ellipses (...) and pauses",
     "ARTHUR (hesitates, clutching the unsealed telegram): You knew... from the very beginning, you knew the cost.",
     "TP 12.TL.3: Theatrical Monologue & Drama (Speaking & Writing)"),

    ("13", "Kelas XII TL", "Semester 5", "Classical Rhetoric Devices (Tricolon, Antithesis, Anaphora)",
     "Membangun pidato oratori persuasif berdaya pikat tinggi di podium kelas dan debat parlemen resmi.",
     "Tricolon: [Phrase 1], [Phrase 2], and [Phrase 3]\nAntithesis: Not because of [X], but because of [Y]\nAnaphora: We shall...",
     "We stand not to glorify previous triumphs, but to shoulder present trials, and to secure future liberties for all generations.",
     "TP 12.TL.4: Persuasive Oratory & Parliamentary Debate (Speaking & Writing)"),

    ("14", "Kelas XII TL", "Semester 5", "Parliamentary Signposting & Formal Rebuttal Formulae",
     "Menyampaikan interupsi resmi (Point of Information) dan membongkar cacat logika argumen pihak oposisi.",
     "Point of Information, Madam Speaker: The honorable opposition rests their contention upon a fundamentally flawed premise...",
     "Signposting: We on side Affirmative will demonstrate two mutually reinforcing burdens of proof...",
     "TP 12.TL.4: WUDC Parliamentary Debating (Speaking)"),

    # KELAS XII TL - SEMESTER 6
    ("15", "Kelas XII TL", "Semester 6", "Executive Register & Google X-Y-Z Accomplishment Syntax",
     "Merumuskan capaian profesional pada berkas dossier internasional dan resume berstandar ATS.",
     "Accomplished [X], as measured by [Y], by doing [Z] (Utilizing senior executive action verbs)",
     "Spearheaded a bilateral university research symposium, expanding student publication submissions by 42% via structured mentorship.",
     "TP 12.TL.5: Professional Admissions & ATS Resume (Writing)"),

    ("16", "Kelas XII TL", "Semester 6", "STAR Framework Syntactic Architecture in Spoken Interviews",
     "Menyusun respons lisan sistematis (Situation-Task-Action-Result) dalam wawancara beasiswa dunia dan viva voce.",
     "Situation (Past setting) -> Task (Obligation) -> Action (Targeted verbs) -> Result (Quantifiable outcome)",
     "When confronted with conflicting data during field trials (S), I reassessed sample integrity (T), standardized testing protocols (A), and restored data fidelity to 99% (R).",
     "TP 12.TL.6: Job Interview & Viva Voce Defense (Speaking)"),

    ("17", "Kelas XII TL", "Semester 6", "Dense Informational Packaging in IMRaD Abstracts",
     "Mengemas intisari riset ilmiah secara padat, efektif, dan presisi ke dalam abstrak terstruktur 250 kata.",
     "Passive voice + Prepositional phrases + Compound nominals (Background, Methods, Results, Conclusions)",
     "Spectroscopic analyses demonstrated notable polymer degradation following prolonged ultraviolet exposure, indicating compromised structural stability.",
     "TP 12.TL.7: Scientific Research Abstract & Poster (Writing & Reading)"),

    ("18", "Kelas XII TL", "Semester 6", "Grammatical Inversion for Formal Rhetorical Emphasis",
     "Mencapai kemahiran struktur tingkat mahir (CEFR B2+/C1) yang diuji pada standar internasional IELTS / TOEFL.",
     "Negative Adverbial (Seldom / Rarely / Under no circumstances / Not only) + Auxiliary + Subject + Main Verb",
     "Under no circumstances should research findings be disseminated prior to rigorous, independent peer verification.",
     "TP 12.TL.8: Standardized Proficiency Readiness (Writing & Reading)"),

    ("19", "Kelas XII TL", "Semester 6", "Cleft Sentences for Thematic Highlighting",
     "Memusatkan fokus pembaca pada variabel penelitian atau bukti inti dalam esai analisis tingkat lanjut.",
     "It is/was [Focused Element] that/who + Clause\nWhat [Subject] [Verb] is/was [Focused Element]",
     "It was the unprecedented velocity of glacial disintegration that compelled climatologists to revise their projections.",
     "TP 12.TL.8: IELTS Academic Writing Task 2 (Writing)"),

    ("20", "Kelas XII TL", "Semester 6", "Inverted Conditionals & Advanced Hypothetical Structures",
     "Mengemukakan skenario pengandaian formal tanpa kata hubung 'if' dalam analisis kebijakan dan esai terstandar.",
     "Had + Subject + V3, Subject + would have + V3\nWere + Subject + to-V, Subject + would + V1\nShould + Subject + V1, ...",
     "Had multilateral signatories enacted emission limits a decade earlier, severe coastal flooding would have been substantially averted.",
     "TP 12.TL.8: Advanced Grammar for International Tests (Writing & Reading)")
]


# ==============================================================================
# BUILD SHEET 1: RINGKASAN PETA MATERI (DIPERBARUI DENGAN MATERI GRAMMAR)
# ==============================================================================
def build_sheet_overview(ws):
    ws.title = "Ringkasan_Peta_Materi"
    apply_title(ws, "PETA MATERI LENGKAP BAHASA INGGRIS SMA (4 SKILLS + GRAMMAR)",
                "Ikhtisar Komprehensif Kurikulum Merdeka 2025/2026: Bahasa Inggris Wajib vs Bahasa Inggris Tingkat Lanjut", 8)

    # Section 1: Perbedaan Esensial
    r = 4
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws.cell(row=r, column=1, value="I. MATRIKS PERBEDAAN ESENSIAL: BAHASA INGGRIS WAJIB VS TINGKAT LANJUT").font = FONT_SEC_HEADER
    for c in range(1, 9):
        ws.cell(row=r, column=c).fill = FILL_NAVY
    ws.row_dimensions[r].height = 24

    comp_headers = ["Dimensi Kurikuler", "Bahasa Inggris Wajib (General English)", "Bahasa Inggris Tingkat Lanjut (Advanced English)", "Dampak pada Keterampilan & Tata Bahasa"]
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
        ("Kompleksitas Tata Bahasa (Grammar)",
         "Simple tenses, descriptive adjectives, modalitas saran umum, konjungsi kronologis/kausal, passive voice standar.",
         "Academic Nominalization, nominal clauses, klausa konsesif, inversi retoris, cleft sentences, conditional tanpa 'if', APA citation.",
         "Tata bahasa Wajib ditekankan pada ketepatan komunikatif; Tingkat Lanjut ditekankan pada kepadatan informasi akademik (syntactic density)."),
        ("Model Asesmen & Unjuk Kerja",
         "Storytelling dongeng, simulasi pembaca berita TV, presentasi brosur wisata, tes pilihan ganda fungsional.",
         "Debat parlemen resmi, pementasan drama monolog teatrikal, sidang pertahanan skripsi lisan (viva voce), esai 1.000 kata.",
         "Asesmen Tingkat Lanjut menguji ketahanan berpikir spontan lisan di bawah tekanan dan orisinalitas riset mandiri.")
    ]

    r += 1
    for r_data in comp_rows:
        ws.row_dimensions[r].height = 44
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

    r += 1
    # Section 3: Ringkasan Materi Grammar
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    ws.cell(row=r, column=1, value="III. PEMETAAN TATA BAHASA & STRUKTUR KALIMAT (GRAMMAR SCOPE & SEQUENCE)").font = FONT_SEC_HEADER
    for c in range(1, 9):
        ws.cell(row=r, column=c).fill = FILL_NAVY
    ws.row_dimensions[r].height = 24

    gr_headers = ["Jenjang / Fase", "Fokus Capaian Grammar", "Materi Tata Bahasa Bahasa Inggris Wajib", "Materi Tata Bahasa Bahasa Inggris Tingkat Lanjut"]
    r += 1
    ws.row_dimensions[r].height = 24
    for idx, (cs, ce) in enumerate(col_spans):
        ws.merge_cells(start_row=r, start_column=cs, end_row=r, end_column=ce)
        cell = ws.cell(row=r, column=cs, value=gr_headers[idx])
        cell.font = FONT_TH
        cell.alignment = ALIGN_CENTER
        for c in range(cs, ce + 1):
            ws.cell(row=r, column=c).fill = FILL_PURPLE
            ws.cell(row=r, column=c).border = CELL_BORDER

    gr_rows = [
        ("KELAS X\n(Fase E)",
         "Fondasi Gramatika Komunikatif & Struktur Dasar Teks Genre (CEFR A2+ to B1)",
         "• Simple Present & Subject-Verb Agreement\n• Descriptive Adjectives & OSASCOMP Noun Phrases\n• Simple Past (Regular & Irregular Verbs)\n• Past Continuous & Time Clauses (when/while)\n• Direct & Indirect Speech Basics\n• Imperative Sentences & Sequencing Words\n• Present Passive Voice & General Scientific Nouns",
         "✦ Belum ada peminatan Tingkat Lanjut di Kelas X (Siswa mengambil mata pelajaran Bahasa Inggris Umum/Wajib sebagai fondasi kemahiran)."),
        ("KELAS XI\n(Fase F)",
         "Pengembangan Argumen Logis, Klausa Majemuk & Wacana Ilmiah (CEFR B1 to B2)",
         "• Internal Conjunctions (Furthermore, Moreover)\n• Causal Connectors (Therefore, Due to, Consequently)\n• Modal Auxiliaries of Obligation (Should, Must)\n• Past Perfect & Relative Clauses (who/which/that)\n• Passive Voice across Perfect & Modal Tenses\n• Contrast & Concessive Markers (Although, However)\n• Evaluative Adjectives & Conditionals Type 1 & 2",
         "• Academic Nominalization (Verb/Adj -> Abstract Nouns)\n• Complex Nominal Clauses (That-clauses & Wh-clauses)\n• High Deontic Modality & Diplomatic Policy Registers\n• Literary Devices, Metaphor & Poetic Inversion\n• Multi-perspective Academic Signposting & Hedging\n• Complex Causal Passive Chains in Frontier Science\n• Historical Recount Synthesizers & In-text Quotations\n• Semiotic Lexis & Film Critique Subordinate Clauses"),
        ("KELAS XII\n(Fase F Akhir)",
         "Pemantapan Diskursus Profesional, Retorika & Kesiapan Global (CEFR B1+ to B2+)",
         "• Hedging Devices (tends to suggest, arguably)\n• Advanced Reported Speech (claimed, confirmed)\n• Prepositions of Data & Trend Movement Verbs\n• Formal Business Letter Register & Salutations\n• Dynamic Action Verbs in Professional CV\n• Spoken Tactical Fillers & Polite Job Interview Formulae\n• Participial Phrases in Multimodal Infographics",
         "• Bioethical Registers & Concessive Clauses (Whereas...)\n• APA 7th Edition In-Text Citation & Literature Synthesis\n• Dramatic Dialogue Conventions & Subtextual Ellipses\n• Classical Rhetoric Devices (Tricolon, Antithesis, Anaphora)\n• Parliamentary Debating Signposting & Points of Info\n• Executive Register & Google X-Y-Z Resume Syntax\n• STAR Architecture in Viva Voce Oral Defense\n• IMRaD Abstract Dense Informational Packaging\n• Grammatical Inversion (Seldom..., Under no circumstances...)\n• Cleft Sentences (It was [X] that...) & Inverted Conditionals")
    ]

    r += 1
    for r_data in gr_rows:
        ws.row_dimensions[r].height = 80
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
# BUILD SHEET 2: BAHASA INGGRIS WAJIB (TABEL MASTER 4 SKILLS + GRAMMAR)
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

    current_class = ""
    r += 1
    for row_data in DATA_WAJIB_4_SKILLS:
        cls_sem = row_data[1]
        cls_name = cls_sem.split(' - ')[0]
        if cls_name != current_class:
            current_class = cls_name
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
# BUILD SHEET 3: BAHASA INGGRIS TINGKAT LANJUT (TABEL MASTER 4 SKILLS + GRAMMAR)
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
# BUILD SHEET 8: MATERI KHUSUS GRAMMAR & STRUKTUR KALIMAT
# ==============================================================================
def build_sheet_grammar(ws):
    ws.title = "5_Grammar"
    apply_title(ws, "MATERI KHUSUS TATA BAHASA & STRUKTUR: GRAMMAR & SENTENCE STRUCTURE",
                "Scope & Sequence Tata Bahasa Bahasa Inggris SMA: Wajib (Fase E & F Umum) vs Tingkat Lanjut (Fase F Pilihan)", 8)

    gr_headers_wajib = [
        "No", "Jenjang / Kelas", "Semester", "Topik Grammar & Struktur Kalimat",
        "Fungsi Komunikatif (Why & When to Use)", "Pola Kalimat / Rumus / Formula",
        "Contoh Kalimat Kontekstual SMA", "Integrasi Teks Genre & Keterampilan"
    ]

    r = 4
    # Section Wajib
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    sec_wajib = ws.cell(row=r, column=1, value="BAGIAN I: BAHASA INGGRIS UMUM (WAJIB) - TATA BAHASA & STRUKTUR KELAS X, XI, XII")
    sec_wajib.font = FONT_SEC_HEADER
    for c in range(1, 9):
        ws.cell(row=r, column=c).fill = FILL_HEADER_WAJIB
    ws.row_dimensions[r].height = 24

    r += 1
    ws.row_dimensions[r].height = 24
    for col_idx, h in enumerate(gr_headers_wajib, 1):
        cell = ws.cell(row=r, column=col_idx, value=h)
        cell.font = FONT_TH
        cell.fill = FILL_NAVY
        cell.alignment = ALIGN_CENTER
        cell.border = CELL_BORDER

    r += 1
    for row_data in DATA_GRAMMAR_WAJIB:
        ws.row_dimensions[r].height = 48
        for c_idx, v in enumerate(row_data, 1):
            cell = ws.cell(row=r, column=c_idx, value=v)
            cell.font = FONT_DATA_BOLD if c_idx in [1, 2, 3, 4] else FONT_DATA
            cell.alignment = ALIGN_CENTER_TOP if c_idx in [1, 2, 3] else ALIGN_LEFT
            cell.border = CELL_BORDER
            if r % 2 == 1:
                cell.fill = FILL_ZEBRA
        r += 1

    r += 1
    # Section Tingkat Lanjut
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    sec_tl = ws.cell(row=r, column=1, value="BAGIAN II: BAHASA INGGRIS TINGKAT LANJUT (PILIHAN) - TATA BAHASA & STRUKTUR KELAS XI & XII")
    sec_tl.font = FONT_SEC_HEADER
    for c in range(1, 9):
        ws.cell(row=r, column=c).fill = FILL_HEADER_TL
    ws.row_dimensions[r].height = 24

    gr_headers_tl = [
        "No", "Jenjang / Kelas", "Semester", "Topik Grammar Akademik & Retorika",
        "Fungsi Retorika / Akademik (Discourse Purpose)", "Pola Sintaksis / Struktur Lanjutan",
        "Contoh Kalimat Kontekstual Akademik / Sastra", "Integrasi Teks Genre & Keterampilan"
    ]

    r += 1
    ws.row_dimensions[r].height = 24
    for col_idx, h in enumerate(gr_headers_tl, 1):
        cell = ws.cell(row=r, column=col_idx, value=h)
        cell.font = FONT_TH
        cell.fill = FILL_SLATE
        cell.alignment = ALIGN_CENTER
        cell.border = CELL_BORDER

    r += 1
    for row_data in DATA_GRAMMAR_TL:
        ws.row_dimensions[r].height = 52
        for c_idx, v in enumerate(row_data, 1):
            cell = ws.cell(row=r, column=c_idx, value=v)
            cell.font = FONT_DATA_BOLD if c_idx in [1, 2, 3, 4] else FONT_DATA
            cell.alignment = ALIGN_CENTER_TOP if c_idx in [1, 2, 3] else ALIGN_LEFT
            cell.border = CELL_BORDER
            if r % 2 == 1:
                cell.fill = FILL_ZEBRA
        r += 1

    auto_fit_columns(ws, min_w=8, max_w=46)
    ws.freeze_panes = "E6"


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

    # 8. Sheet 5_Grammar
    ws8 = wb.create_sheet(title="5_Grammar")
    build_sheet_grammar(ws8)

    wb.save(out_file)
    print("Workbook successfully created and saved with Grammar sheet!")

if __name__ == "__main__":
    main()
