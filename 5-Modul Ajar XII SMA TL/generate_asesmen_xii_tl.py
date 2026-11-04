# -*- coding: utf-8 -*-
"""
Generator script for complete Assessment Instruments (.docx) in 5-Modul Ajar XII SMA TL.
Covers TP 12.TL.1 through TP 12.TL.8 with complete components:
- Constructive Alignment Matrix
- Bagian I: Asesmen Diagnostik (Non-kognitif & Kognitif)
- Bagian II: Asesmen Formatif (Lembar Observasi & Rubrik Proses)
- Bagian III: Asesmen Sumatif (Kisi-kisi, 10 Soal PG HOTS, 5 Soal Esai, Performance Task, Rubrik Analitik, Kunci Jawaban & Pedoman Penskoran)
"""

import os
import sys
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
import docx_helper

TARGET_DIR = os.path.dirname(os.path.abspath(__file__))
os.makedirs(TARGET_DIR, exist_ok=True)

COLOR_NAVY = RGBColor(15, 41, 66)
COLOR_ROYAL = RGBColor(30, 64, 175)
COLOR_BODY = RGBColor(30, 41, 59)

def add_as_section(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(title)
    run.font.name = "Segoe UI"
    run.font.size = Pt(11.5)
    run.font.bold = True
    run.font.color.rgb = COLOR_NAVY

def add_as_sub(doc, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(title)
    run.font.name = "Segoe UI"
    run.font.size = Pt(10)
    run.font.bold = True
    run.font.color.rgb = COLOR_ROYAL

def add_obs_table(doc):
    obs_headers = ["No", "Nama Peserta Didik", "Partisipasi Kritis", "Argumentasi Logis", "Kesantunan Akademik", "Total", "Predikat"]
    obs_rows = [
        ["1", "Aditya Pratama", "4", "4", "4", "12", "Sangat Baik"],
        ["2", "Devina Anindya", "4", "3", "4", "11", "Sangat Baik"],
        ["3", "Farhan Ramadhan", "3", "3", "3", "9", "Baik"],
        ["4", "................................................", "....", "....", "....", "....", "..............."]
    ]
    docx_helper.add_styled_table(doc, obs_headers, obs_rows, [Inches(0.5), Inches(2.2), Inches(1.1), Inches(1.1), Inches(1.1), Inches(0.8), Inches(1.0)])


# ==============================================================================
# ASESMEN 1: TP 12.TL.1 - GLOBAL BIOETHICS DISCUSSION
# ==============================================================================
def create_asesmen_1():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "INSTRUMEN ASESMEN LENGKAP - BAHASA INGGRIS TINGKAT LANJUT KELAS XII",
        "Modul 1: Global Academic Discussion on Bioethics | TP 12.TL.1"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XII / Semester 5 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 12.TL.1 - Menganalisis wacana debat akademik tingkat tinggi (Advanced Discussion) mengenai bioetika medis, rekayasa genetika, eksplorasi antariksa, dan merumuskan sikap etis ilmiah secara komprehensif.",
        "Prinsip Utama": "Constructive Alignment (IKTP <-> Teknik & Bentuk Asesmen)",
        "Target Kurikulum": "Kurikulum Merdeka 2025/2026 (Permendikdasmen No. 13 Tahun 2025)"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_as_section(doc, "I. MATRIKS KESELARASAN INDIKATOR DENGAN TEKNIK ASESMEN")
    al_rows = [
        ["IKTP 1.1", "Mengidentifikasi argumen filosofis modifikasi genetik CRISPR.", "Tes Tertulis", "Pilihan Ganda & Isian LKPD"],
        ["IKTP 1.2", "Menerapkan stance markers dan klausa konsesif akademik.", "Tes Tertulis / Kinerja", "Soal Analisis Kalimat & Rubrik Esai"],
        ["IKTP 1.3", "Mengevaluasi risiko ketimpangan biologis dan eugenika baru.", "Tes Tertulis HOTS", "Soal Uraian Penalaran Kritis"],
        ["IKTP 1.4", "Menyusun Position Paper 500 kata & tampil dalam Meja Bundar.", "Penilaian Kinerja Produk", "Esai Posisi & Meja Bundar Resmi"]
    ]
    docx_helper.add_styled_table(doc, ["Kode IKTP", "Indikator Ketercapaian", "Teknik Asesmen", "Bentuk Instrumen"], al_rows, [Inches(1.2), Inches(3.0), Inches(1.3), Inches(1.5)])

    add_as_section(doc, "II. INSTRUMEN ASESMEN DIAGNOSTIK (AWAL PEMBELAJARAN)")
    add_as_sub(doc, "A. Angket Diagnostik Non-Kognitif (Kesiapan Belajar & Sikap Moral)")
    p_d = doc.add_paragraph()
    p_d.paragraph_format.left_indent = Inches(0.2)
    p_d.add_run("1. Sejauh mana kamu mengikuti perkembangan teknologi rekayasa genetika (CRISPR) dan kecerdasan buatan? (Sangat Tahu / Tahu Sedikit / Belum Tahu)\n")
    p_d.add_run("2. Menurut pandangan pribadimu, apakah sains boleh mengubah DNA embrio manusia demi mencegah penyakit bawaan?\n")
    p_d.add_run("3. Format pembelajaran apa yang paling memotivasimu: debat formal, penulisan esai mandiri, atau analisis studi kasus nyata?")

    add_as_sub(doc, "B. Tes Diagnostik Kognitif Awal (5 Butir Soal Prasyarat)")
    p_cog = doc.add_paragraph()
    p_cog.paragraph_format.left_indent = Inches(0.2)
    p_cog.add_run("1. What is the fundamental difference between somatic cell therapy and germline genetic modification?\n")
    p_cog.add_run("2. Identify the connector of concession: 'Furthermore', 'Notwithstanding', or 'Consequently'?\n")
    p_cog.add_run("3. Combine these two ideas using 'While': (A) Gene editing cures diseases. (B) It risks genetic inequality.\n")
    p_cog.add_run("4. What does the term 'moratorium' signify in scientific governance?\n")
    p_cog.add_run("5. Explain why hedging (e.g., 'arguably', 'evidence suggests') is essential in academic discourse.")

    add_as_section(doc, "III. INSTRUMEN ASESMEN FORMATIF (PROSES PEMBELAJARAN)")
    add_as_sub(doc, "Lembar Observasi Partisipasi Dialog Meja Bundar & Kolaborasi (Skala 1 - 4)")
    add_obs_table(doc)

    add_as_section(doc, "IV. INSTRUMEN ASESMEN SUMATIF (AKHIR MODUL)")
    add_as_sub(doc, "A. Kisi-Kisi Soal Asesmen Sumatif")
    kisi_rows = [
        ["1", "Menganalisis isu esensial perbedaan terapi somatik vs germline.", "C4 (Analisis)", "Pilihan Ganda", "1, 2"],
        ["2", "Menemukan argumen utilitarian pendukung rekayasa genetika.", "C4 (Analisis)", "Pilihan Ganda", "3, 4"],
        ["3", "Mengevaluasi argumen penentang terkait risiko eugenika & stratifikasi.", "C5 (Evaluasi)", "Pilihan Ganda", "5, 6"],
        ["4", "Menganalisis fungsi leksikal stance markers dan klausa konsesif.", "C3 (Aplikasi)", "Pilihan Ganda", "7, 8"],
        ["5", "Menyimpulkan rekomendasi jalan tengah tata kelola bioetika global.", "C5 (Evaluasi)", "Pilihan Ganda", "9, 10"],
        ["6", "Menguraikan dilema etis pendelegasian keputusan medis kepada AI.", "C5 (Evaluasi)", "Uraian / Esai", "1, 2, 3"],
        ["7", "Menyusun naskah Position Paper bioetika berimbang 500 kata.", "C6 (Kreasi)", "Tugas Kinerja", "Tugas Produk"]
    ]
    docx_helper.add_styled_table(doc, ["No", "Materi Pokok / Indikator Soal", "Level Kognitif", "Bentuk Soal", "Nomor Soal"], kisi_rows, [Inches(0.6), Inches(3.2), Inches(1.2), Inches(1.2), Inches(0.8)])

    add_as_sub(doc, "B. Naskah Soal Pilihan Ganda HOTS (10 Butir Soal)")
    pg_text = (
        "Stimulus: Read the passage to answer questions 1 – 5:\n"
        "'The advent of CRISPR-Cas9 genome editing represents one of the most transformative scientific breakthroughs in human history. By conferring the ability to excise deleterious genetic sequences and insert therapeutic variants, biotechnology stands at the threshold of eradicating hereditary pathologies. Nevertheless, the prospective transition from somatic gene therapy—which affects solely the treated individual—to germline modification, where alterations are inherited across subsequent generations, has precipitated intense moral deliberation. Proponents passionately articulate a humanitarian imperative: from a utilitarian standpoint, eradicating heritable afflictions spares families profound torment and alleviates public healthcare burdens. Conversely, critics caution that unrestricted germline editing will entrench biological stratification, bifurcating humanity into a genetically engineered ruling elite and an unenhanced underclass. In synthesis, human germline alteration demands a strict international moratorium until unequivocal consensus is forged.'\n\n"
        "1. What is the pivotal distinction between somatic therapy and germline modification highlighted in the text?\n"
        "   A. Somatic therapy is exponentially more expensive than germline alteration.\n"
        "   B. Somatic changes affect only the patient, whereas germline changes are inherited by future generations.\n"
        "   C. Germline modification is already legally practiced in every hospital worldwide.\n"
        "   D. Somatic therapy completely eliminates the risk of all future mutations.\n"
        "   E. Germline editing does not utilize the CRISPR-Cas9 enzyme system.\n\n"
        "2. Which philosophical perspective is employed by advocates to justify therapeutic gene editing?\n"
        "   A. Nihilism and the rejection of medical interventions.\n"
        "   B. Utilitarian beneficence aimed at minimizing suffering and lifetime healthcare costs.\n"
        "   C. Strict traditionalism demanding the preservation of all hereditary traits.\n"
        "   D. Free-market deregulation with zero governmental licensing.\n"
        "   E. Biological determinism favoring cosmetic enhancement.\n\n"
        "3. According to critics, what grave societal peril could emerge from commercialized germline modification?\n"
        "   A. Complete eradication of all pharmaceutical research facilities.\n"
        "   B. Severe biological stratification dividing society into genetically enhanced elites and unenhanced underclasses.\n"
        "   C. Immediate shutdown of all university research laboratories.\n"
        "   D. Overpopulation of hospital emergency triage wards.\n"
        "   E. Total collapse of international patent registries.\n\n"
        "4. The word 'bifurcating' in paragraph 2 is closest in meaning to...\n"
        "   A. Unifying and reconciling.\n"
        "   B. Dividing or splitting into two distinct branches.\n"
        "   C. Accelerating at high velocity.\n"
        "   D. Healing completely.\n"
        "   E. Eliminating through legal statutes.\n\n"
        "5. What balanced policy recommendation is articulated in the final synthesis?\n"
        "   A. Total, permanent abandonment of all genetic technologies including somatic therapies.\n"
        "   B. Unconditional immediate commercial release of germline editing to the highest bidder.\n"
        "   C. An international moratorium on germline alterations until universal consensus and safety protocols are achieved.\n"
        "   D. Allowing private clinics to operate without any ethical oversight.\n"
        "   E. Replicating space treaties on deep sea fishing vessels.\n\n"
        "6. Which sentence exemplifies an effective academic stance marker?\n"
        "   A. Obviously, anyone who opposes genetic editing is completely ignorant.\n"
        "   B. Arguably, the moral imperative to relieve suffering outweighs speculative sociological anxieties.\n"
        "   C. I believe that DNA modification is super cool and fun for all parents.\n"
        "   D. DNA editing will 100% cure every single disease by tomorrow morning.\n"
        "   E. It is stupid to regulate scientists.\n\n"
        "7. In clinical algorithmic triage (Text 2), what is the primary argument against autonomous AI decisions?\n"
        "   A. Computers consume excessive hospital electricity.\n"
        "   B. Automated algorithms fundamentally lack human empathy, intuitive moral judgment, and spiritual awareness.\n"
        "   C. Doctors enjoy doing routine paperwork and filling out billing sheets.\n"
        "   D. Hospital computers are unable to store patient blood pressure statistics.\n"
        "   E. Palliative care patients prefer talking exclusively to machines.\n\n"
        "8. In space frontier governance (Text 3), what does the 1967 Outer Space Treaty prohibit?\n"
        "   A. Scientific telescopes observing deep space galaxies.\n"
        "   B. National appropriation of celestial bodies by claim of sovereignty or occupation.\n"
        "   C. Rocket launches carrying civilian telecommunication satellites.\n"
        "   D. Astronauts conducting zero-gravity botanical experiments.\n"
        "   E. Solar power absorption on orbital stations.\n\n"
        "9. Which discourse marker correctly introduces a concession in formal academic debate?\n"
        "   A. Notwithstanding\n"
        "   B. And also\n"
        "   C. Because of\n"
        "   D. In addition to\n"
        "   E. As well as\n\n"
        "10. What is the quintessential purpose of a balanced conclusion in an Advanced Discussion Text?\n"
        "   A. To ridicule one side of the debate as morally bankrupt.\n"
        "   B. To synthesize opposing arguments and propose an equitable, regulated framework.\n"
        "   C. To avoid mentioning any solutions whatsoever.\n"
        "   D. To introduce ten completely unrelated scientific theories.\n"
        "   E. To summarize a fictional narrative story."
    )
    p_pg = doc.add_paragraph()
    p_pg.paragraph_format.left_indent = Inches(0.2)
    p_pg.add_run(pg_text)

    add_as_sub(doc, "C. Naskah Soal Uraian / Esai Analitis (5 Butir Soal HOTS)")
    p_es = doc.add_paragraph()
    p_es.paragraph_format.left_indent = Inches(0.2)
    p_es.add_run("1. Deconstruct the ethical tension between 'individual parental autonomy' and 'collective societal equality' regarding human germline genetic modification.\n")
    p_es.add_run("2. Explain why bioethicists fear that permitting genetic correction of diseases will create an inexorable 'slippery slope' toward cosmetic and cognitive eugenics.\n")
    p_es.add_run("3. In hospital triage, how can healthcare institutions ethically balance mathematical algorithmic objectivity with the physician's empathetic bedside judgment?\n")
    p_es.add_run("4. Contrast the legal status of outer space resources under the 1967 Outer Space Treaty with the economic incentives required by private aerospace enterprises.\n")
    p_es.add_run("5. Compose two complex academic sentences utilizing 'Notwithstanding' and 'arguably' to articulate your personal ethical stance on CRISPR regulation.")

    add_as_sub(doc, "D. Tugas Unjuk Kerja Autentik & Rubrik Analitik (Skala 1 - 4)")
    rub_rows = [
        ["1. Content & Depth of Argument (30%)", "Analisis isu sangat mendalam; argumen pro-kontra disajikan komprehensif berbasis data sains frontier.", "Analisis isu jelas; argumen pro-kontra disajikan seimbang didukung bukti relevan.", "Argumen kurang mendalam; salah satu sudut pandang sangat dominan.", "Pemahaman isu sangat dangkal; tidak ada keberimbangan perspektif."],
        ["2. Organization & Synthesis (25%)", "Alur berpikir logis mengalir sempurna; sintesis rekomendasi solutif sangat koheren.", "Struktur lengkap; transisi antargagasan mengalir cukup runtut.", "Struktur kurang teratur; perumusan rekomendasi terkesan memaksakan.", "Struktur acak dan tidak membentuk teks diskusi akademik yang utuh."],
        ["3. Academic Lexicon & Hedging (25%)", "Penggunaan stance markers, klausa konsesif, dan istilah bioetika presisi luar biasa.", "Penggunaan istilah akademik dan konjungsi kontras tepat dan sesuai konteks.", "Penggunaan istilah akademik terbatas; ada kesalahan penggunaan konjungsi.", "Pilihan kata informal/tidak baku; banyak kesalahan gramatikal mendasar."],
        ["4. Mechanics & Conventions (20%)", "Bebas dari kesalahan ejaan, tanda baca, dan konvensi penulisan akademik.", "Hanya sedikit kesalahan minor yang tidak mengganggu pemahaman isi.", "Cukup banyak kesalahan ejaan dan tanda baca di beberapa paragraf.", "Banyak kesalahan mekanik mendasar di setiap kalimat."]
    ]
    docx_helper.add_styled_table(doc, ["Kriteria Penilaian", "Skor 4 (Sangat Baik)", "Skor 3 (Baik)", "Skor 2 (Cukup)", "Skor 1 (Kurang)"], rub_rows, [Inches(1.5), Inches(1.5), Inches(1.5), Inches(1.5), Inches(1.5)])

    add_as_sub(doc, "E. Kunci Jawaban & Pedoman Penskoran")
    docx_helper.add_callout(
        doc,
        "Kunci Jawaban PG & Formula Nilai Akhir:",
        "Kunci Jawaban PG:\n"
        "1. B | 2. B | 3. B | 4. B | 5. C | 6. B | 7. B | 8. B | 9. A | 10. B\n\n"
        "Pedoman Penskoran:\n"
        "• Nilai Pilihan Ganda = (Benar / 10) x 40 poin (Maks: 40)\n"
        "• Nilai Soal Uraian = (Skor / 20) x 20 poin (Maks: 20)\n"
        "• Nilai Tugas Kinerja Position Paper = (Skor / 16) x 40 poin (Maks: 40)\n"
        "• NILAI AKHIR (NA) = Nilai PG + Nilai Uraian + Nilai Kinerja (Total Maks: 100)",
        bg="DCFCE7", border_color="16A34A"
    )

    fpath = os.path.join(TARGET_DIR, "Instrumen_Asesmen_TP_12.TL.1_Global_Bioethics_Discussion.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# ASESMEN 2: TP 12.TL.2 - ACADEMIC SYNTHESIS ESSAY
# ==============================================================================
def create_asesmen_2():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "INSTRUMEN ASESMEN LENGKAP - BAHASA INGGRIS TINGKAT LANJUT KELAS XII",
        "Modul 2: Academic Synthesis Essay & Research Integration | TP 12.TL.2"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XII / Semester 5 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 12.TL.2 - Menyusun esai penelitian akademik (Academic Synthesis Essay) 1.000 kata dengan memadukan multi-referensi, kutipan langsung/tidak langsung, dan daftar pustaka standar internasional (APA/MLA).",
        "Prinsip Utama": "Constructive Alignment (Integritas Akademik & Rigor Sitasi)"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_as_section(doc, "I. MATRIKS KESELARASAN ASESMEN")
    al_rows = [
        ["IKTP 2.1", "Merumuskan working thesis statement & menyintesis 3 jurnal.", "Tes Tertulis / Portofolio", "Matriks Sintesis Literatur"],
        ["IKTP 2.2", "Menerapkan teknik sitasi APA 7th Edition dan parafrasa etis.", "Tes Tertulis", "Soal Analisis Sitasi & Deteksi Plagiasi"],
        ["IKTP 2.3", "Mengevaluasi keselarasan temuan empiris & research gaps.", "Tes Tertulis HOTS", "Soal Analisis Komparasi Literatur"],
        ["IKTP 2.4", "Menyusun naskah esai sintesis penelitian akademik 1.000 kata.", "Penilaian Portofolio", "Makalah Riset & Daftar Pustaka APA"]
    ]
    docx_helper.add_styled_table(doc, ["Kode IKTP", "Indikator Ketercapaian", "Teknik Asesmen", "Bentuk Instrumen"], al_rows, [Inches(1.2), Inches(3.0), Inches(1.3), Inches(1.5)])

    add_as_section(doc, "II. KISI-KISI & SOAL ASESMEN SUMATIF")
    pg_text = (
        "Stimulus: Read the passage to answer questions 1 – 5:\n"
        "'Recent environmental toxicology studies confirm that microplastics are ubiquitously distributed across atmospheric, oceanic, and terrestrial ecosystems (Thompson et al., 2022). In a groundbreaking investigative cohort study, Leslie et al. (2022) identified quantifiable concentrations of polyethylene and PET in 77% of human whole-blood samples tested. The authors hypothesized that systemic translocation occurs via pulmonary inhalation and gastrointestinal mucosal absorption. These bio-accumulated polymer matrices induce inflammatory cascades, mitochondrial impairment, and oxidative cellular damage in human vascular endothelium (Prata et al., 2023). Nevertheless, synthesizing findings across contemporary literature reveals notable methodological discrepancies regarding chronic toxicity thresholds. While in vitro toxicology models demonstrate pronounced cytotoxic apoptotic pathways under high nanoparticle exposure (Vethaak & Legler, 2021), epidemiological studies evaluating real-world ambient human ingestion remain preliminary.'\n\n"
        "1. According to the synthesized literature, what percentage of tested human blood samples contained quantifiable polymers?\n"
        "   A. Exactly 100%\n"
        "   B. 77%\n"
        "   C. Less than 5%\n"
        "   D. 25%\n"
        "   E. 50%\n\n"
        "2. What biological mechanism is cited by Prata et al. (2023) regarding the impact of accumulated microplastics in human blood vessels?\n"
        "   A. Immediate eradication of all viral infections.\n"
        "   B. Inflammatory cascades, mitochondrial impairment, and oxidative cellular damage.\n"
        "   C. Rapid regeneration of heart valves.\n"
        "   D. Permanent immunity against airborne allergens.\n"
        "   E. Increased oxygen carrying capacity.\n\n"
        "3. Which parenthetical citation strictly adheres to the official APA 7th Edition style for three or more authors?\n"
        "   A. (Leslie and colleagues 2022)\n"
        "   B. (Leslie et al., 2022)\n"
        "   C. (Leslie, et. al., 2022.)\n"
        "   D. [Leslie et al: 2022]\n"
        "   E. (Leslie 2022 et al)\n\n"
        "4. What methodological discrepancy is identified between in vitro studies and real-world studies?\n"
        "   A. In vitro studies use real human participants while real-world studies use plastic mannequins.\n"
        "   B. In vitro tests show acute cellular toxicity under high doses, but real-world ambient human ingestion data remain preliminary.\n"
        "   C. Real-world studies prove that microplastics are completely non-existent in oceans.\n"
        "   D. Both methodologies have been officially banned by international health bodies.\n"
        "   E. In vitro tests are conducted exclusively in zero-gravity space laboratories.\n\n"
        "5. Which reporting verb implies that the cited author is proposing a tentative, reasoned explanation rather than an established fact?\n"
        "   A. Established\n"
        "   B. Hypothesized\n"
        "   C. Concluded\n"
        "   D. Proved\n"
        "   E. Measured\n\n"
        "6. In APA 7th Edition, what is the correct formatting for an entry in the References list for an online journal article?\n"
        "   A. Author, A. A. (Year). Title of article. Title of Periodical, volume(issue), pages. https://doi.org/xxx\n"
        "   B. Title of article. Author. Year. Journal.\n"
        "   C. Author, Title, Year, Page Number.\n"
        "   D. URL link only without author or title.\n"
        "   E. Journal name in ALL CAPS followed by student's name.\n\n"
        "7. Why is 'patchwriting' (copying phrases and merely substituting a few synonyms) considered plagiarism?\n"
        "   A. Because it uses too many advanced vocabulary words.\n"
        "   B. Because it mimics the original author's sentence structure without genuine conceptual processing.\n"
        "   C. Because dictionaries do not allow synonyms.\n"
        "   D. Because APA style requires all papers to be written in Latin.\n"
        "   E. Because teachers prefer direct quotes on every line.\n\n"
        "8. In energy grid synthesis (Text 2), why must renewable solar power be paired with storage technologies?\n"
        "   A. Because solar panels produce toxic radiation at night.\n"
        "   B. Because meteorological intermittency threatens electrical frequency regulation and grid stability.\n"
        "   C. Because fossil fuels are completely free of charge.\n"
        "   D. Because wind turbines cannot spin in normal atmospheric air.\n"
        "   E. Because governments forbid battery usage in cities.\n\n"
        "9. What did randomized trials in Finland and California discover regarding Universal Basic Income (Text 3)?\n"
        "   A. All participants immediately resigned from their jobs to sleep all day.\n"
        "   B. Recipients used financial security to reskill, seek better vocational paths, or care for family with negligible labor drop.\n"
        "   C. Basic income caused an instant 90% rise in inflation overnight.\n"
        "   D. Participants donated all their stipends back to corporate banks.\n"
        "   E. The government canceled the currency entirely.\n\n"
        "10. What defines a powerful academic thesis statement in a research synthesis essay?\n"
        "   A. A broad, undeniable statement of obvious fact (e.g., 'Pollution exists worldwide').\n"
        "   B. A specific, debatable, and evidence-grounded claim that synthesizes multiple perspectives.\n"
        "   C. A personal emotional rant without any cited literature.\n"
        "   D. A verbatim copy of the introductory textbook definition.\n"
        "   E. A question ending with multiple exclamation marks."
    )
    p_pg = doc.add_paragraph()
    p_pg.paragraph_format.left_indent = Inches(0.2)
    p_pg.add_run(pg_text)

    add_as_sub(doc, "C. Naskah Soal Uraian / Esai Sintesis")
    p_es = doc.add_paragraph()
    p_es.paragraph_format.left_indent = Inches(0.2)
    p_es.add_run("1. Explain the conceptual difference between a simple 'Literature Summary' and a 'Thematic Literature Synthesis'.\n")
    p_es.add_run("2. Paraphrase this passage without plagiarizing its sentence structure: 'Non-biodegradable synthetic polymers readily absorb hydrophobic organic pollutants, effectively transforming minute plastic particles into potent chemical vectors (Thompson et al., 2022).'\n")
    p_es.add_run("3. Synthesize the findings of Leslie et al. (2022) and Vethaak & Legler (2021) into one coherent academic sentence.\n")
    p_es.add_run("4. Why is the synthesis of contradictory empirical findings (such as in vitro vs epidemiological studies) vital in academic writing?\n")
    p_es.add_run("5. Draft a complete APA 7th Edition reference entry for a journal article published in 2024 by John Smith and Maria Garcia titled 'Renewable Microgrids' in Energy Policy, volume 45, pages 112-120, with DOI: 10.1016/j.enpol.2024.01.005.")

    add_as_sub(doc, "D. Tugas Portofolio Makalah Sintesis 1.000 Kata & Rubrik")
    docx_helper.add_callout(
        doc,
        "Kunci Jawaban PG & Formula Nilai Akhir:",
        "Kunci Jawaban PG:\n"
        "1. B | 2. B | 3. B | 4. B | 5. B | 6. A | 7. B | 8. B | 9. B | 10. B\n\n"
        "Pedoman Penskoran:\n"
        "• Nilai PG = (Benar / 10) x 40 poin\n"
        "• Nilai Soal Uraian = (Skor / 20) x 20 poin\n"
        "• Nilai Makalah Riset 1.000 Kata (Rubrik Internasional) = (Skor / 16) x 40 poin\n"
        "• NILAI AKHIR (NA) = Nilai PG + Nilai Uraian + Nilai Makalah (Total Maks: 100)",
        bg="DCFCE7", border_color="16A34A"
    )

    fpath = os.path.join(TARGET_DIR, "Instrumen_Asesmen_TP_12.TL.2_Academic_Synthesis_Essay.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# ASESMEN 3: TP 12.TL.3 - DRAMATIC SCRIPT & PERFORMANCE
# ==============================================================================
def create_asesmen_3():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "INSTRUMEN ASESMEN LENGKAP - BAHASA INGGRIS TINGKAT LANJUT KELAS XII",
        "Modul 3: Dramatic Literature, Subtext & Theatrical Performance | TP 12.TL.3"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XII / Semester 5 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 12.TL.3 - Menganalisis teks naskah drama klasik/modern, menafsirkan subteks psikologis, serta memperagakan adegan monolog/dialog dramatik dengan intonasi dan pelafalan standar profesional.",
        "Prinsip Utama": "Constructive Alignment (Apresiasi Sastra & Seni Peran)"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_as_section(doc, "I. MATRIKS KESELARASAN ASESMEN")
    al_rows = [
        ["IKTP 3.1", "Menganalisis petunjuk panggung, tensi dramatik, dan subteks.", "Tes Tertulis", "Pilihan Ganda & Analisis LKPD"],
        ["IKTP 3.2", "Menerapkan modulasi vokal, jeda dramatis, dan pelafalan presisi.", "Tes Kinerja", "Table Reading Rehearsal"],
        ["IKTP 3.3", "Mengevaluasi kritik sosial teks drama terhadap kemanusiaan.", "Tes Tertulis HOTS", "Soal Uraian Interpretasi Sastra"],
        ["IKTP 3.4", "Mementaskan monolog dramatik 3-4 menit dengan penghayatan.", "Unjuk Kerja Lisan", "Rubrik Pementasan Monolog / Sandiwara"]
    ]
    docx_helper.add_styled_table(doc, ["Kode IKTP", "Indikator Ketercapaian", "Teknik Asesmen", "Bentuk Instrumen"], al_rows, [Inches(1.2), Inches(3.0), Inches(1.3), Inches(1.5)])

    add_as_section(doc, "II. KISI-KISI & SOAL ASESMEN SUMATIF")
    pg_text = (
        "Stimulus: Read the dramatic excerpt to answer questions 1 – 5:\n"
        "'PROCTOR: (Snatching the parchment to his chest with sudden, trembling ferocity) No, no! I have signed it! You have witnessed me sign! You will not hang it on the church door!\n"
        "DANFORTH: If you have truly confessed before God and church, what matters where the village reads it?\n"
        "PROCTOR: (Breathing heavily, voice steadying into ringing clarity) Because it is my name! Because I cannot have another in my life! Because I lie and sign myself to lies! How may I live without my name? I have given you my soul; leave me my name!\n"
        "DANFORTH: (Harshly) Take him to the gallows!\n"
        "[PROCTOR stands tall, tearing the confession into shreds.]'\n\n"
        "1. What is the central psychological conflict experienced by John Proctor in this scene?\n"
        "   A. Fear of physical poverty and debt.\n"
        "   B. The agonizing choice between living a life anchored in public falsehood or preserving his moral integrity in death.\n"
        "   C. Anger at Danforth for not providing him enough ink and paper.\n"
        "   D. Regret for ever coming to Salem village.\n"
        "   E. An ambition to become the new chief judge of Salem.\n\n"
        "2. What is the function of the stage direction '(Snatching the parchment to his chest with sudden, trembling ferocity)'?\n"
        "   A. To show that Proctor is feeling cold in the jail cell.\n"
        "   B. To guide the actor to portray a sudden surge of protective, desperate defiance.\n"
        "   C. To indicate that the paper is soaking wet with rainwater.\n"
        "   D. To suggest that Proctor wants to hide the paper in his pocket forever.\n"
        "   E. To instruct the marshal to arrest Danforth.\n\n"
        "3. Which rhetorical device is deployed in Proctor's repetitive lines: 'Because it is my name! Because I cannot have another... Because I lie...'?\n"
        "   A. Onomatopoeia\n"
        "   B. Anaphora\n"
        "   C. Synecdoche\n"
        "   D. Euphemism\n"
        "   E. Hyperbole\n\n"
        "4. In theatrical acting, what does 'subtext' refer to?\n"
        "   A. The footnotes printed at the bottom of the script book.\n"
        "   B. The unspoken inner thoughts, emotional baggage, and real motivations behind the spoken lines.\n"
        "   C. The ticket prices sold at the theater box office.\n"
        "   D. The translation of the play into foreign languages.\n"
        "   E. The sound effects played by the backstage technician.\n\n"
        "5. When Proctor tears the signed confession into shreds, what does this physical action symbolize?\n"
        "   A. Total madness and loss of cognitive control.\n"
        "   B. The reclamation of his moral dignity and victory of conscience over authoritarian tyranny.\n"
        "   C. An accident because the paper was too brittle.\n"
        "   D. His hatred of church doors and architecture.\n"
        "   E. His desire to sign a cleaner piece of paper.\n\n"
        "6. In Lorraine Hansberry's A Raisin in the Sun (Text 2), why does Walter Lee refuse Mr. Lindner's buyout check?\n"
        "   A. Because Lindner offered counterfeit bank notes.\n"
        "   B. Because accepting money to abandon their home would betray his father's sacrificial labor and racial dignity.\n"
        "   C. Because the house in Clybourne Park had burned down.\n"
        "   D. Because his sister forbade him from touching money.\n"
        "   E. Because he wanted ten times more cash.\n\n"
        "7. In Hamlet's soliloquy 'To be or not to be' (Text 3), what does the 'undiscovered country from whose bourn no traveller returns' represent?\n"
        "   A. The mysterious realm of death and the afterlife.\n"
        "   B. The unexplored American continent.\n"
        "   C. The royal court of Norway.\n"
        "   D. Deep underwater ocean trenches.\n"
        "   E. An abandoned theatrical stage in London.\n\n"
        "8. Which vocal technique allows an actor to project intensity without straining or damaging vocal cords?\n"
        "   A. Screaming solely from the throat.\n"
        "   B. Diaphragmatic breath support and resonant placement in the chest and mask.\n"
        "   C. Whispering so quietly that no words can be heard.\n"
        "   D. Speaking as rapidly as humanly possible.\n"
        "   E. Coughing repeatedly before every line.\n\n"
        "9. What does 'blocking' signify in theatrical production?\n"
        "   A. Forbidding audience members from taking photographs.\n"
        "   B. The choreographed physical positioning and movement of actors across the stage space.\n"
        "   C. Painting the back wall of the theater black.\n"
        "   D. Forgetting your lines during live performance.\n"
        "   E. Canceling the show due to bad weather.\n\n"
        "10. Why is dramatic literature uniquely valuable in developing cultural empathy?\n"
        "   A. It requires students to memorize legal terminology verbatim.\n"
        "   B. It compels learners to inhabit the lived emotional experiences and moral struggles of diverse human characters.\n"
        "   C. It guarantees immediate financial wealth in Hollywood.\n"
        "   D. It eliminates the need for learning grammar.\n"
        "   E. It replaces all written examinations with pantomime."
    )
    p_pg = doc.add_paragraph()
    p_pg.paragraph_format.left_indent = Inches(0.2)
    p_pg.add_run(pg_text)

    add_as_sub(doc, "C. Naskah Soal Uraian & Studi Sastra")
    p_es = doc.add_paragraph()
    p_es.paragraph_format.left_indent = Inches(0.2)
    p_es.add_run("1. Analyze the historical allegory of Arthur Miller's The Crucible. How does Salem's 1692 witch trials mirror 1950s McCarthyism?\n")
    p_es.add_run("2. In A Raisin in the Sun, analyze the stage direction: 'Walter looks down at his son Travis. His hands shake, then ball into fists. He straightens his spine slowly.' What psychological shift does this physical transition signify?\n")
    p_es.add_run("3. Deconstruct Hamlet's statement: 'Thus conscience does make cowards of us all.' What philosophical paralysis is Shakespeare articulating?\n")
    p_es.add_run("4. Explain how vocal modulation (pitch, tempo, pause, volume) transforms the emotional impact of a dramatic monologue.\n")
    p_es.add_run("5. Choose a 3-line monologue excerpt and provide explicit stage directions for facial expression, eye focus, and physical gesture.")

    add_as_sub(doc, "D. Rubrik Pementasan Monolog / Sandiwara (Skala 1 - 4)")
    docx_helper.add_callout(
        doc,
        "Kunci Jawaban PG & Formula Nilai Akhir:",
        "Kunci Jawaban PG:\n"
        "1. B | 2. B | 3. B | 4. B | 5. B | 6. B | 7. A | 8. B | 9. B | 10. B\n\n"
        "Pedoman Penskoran:\n"
        "• Nilai PG = (Benar / 10) x 40 poin\n"
        "• Nilai Soal Uraian = (Skor / 20) x 20 poin\n"
        "• Nilai Pementasan Teatrikal = (Skor / 16) x 40 poin\n"
        "• NILAI AKHIR (NA) = Nilai PG + Nilai Uraian + Nilai Pementasan (Total Maks: 100)",
        bg="DCFCE7", border_color="16A34A"
    )

    fpath = os.path.join(TARGET_DIR, "Instrumen_Asesmen_TP_12.TL.3_Dramatic_Script_and_Performance.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# ASESMEN 4: TP 12.TL.4 - PERSUASIVE ORATORY & PARLIAMENTARY DEBATE
# ==============================================================================
def create_asesmen_4():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "INSTRUMEN ASESMEN LENGKAP - BAHASA INGGRIS TINGKAT LANJUT KELAS XII",
        "Modul 4: Persuasive Oratory & Parliamentary Debate | TP 12.TL.4"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XII / Semester 5 (Ganjil)",
        "Kode & Tujuan Pembelajaran": "TP 12.TL.4 - Merancang dan membawakan pidato persuasif tingkat tinggi (Persuasive Oratory) dan mempraktikkan debat formal gaya parlemen dengan retorika persuasif yang meyakinkan.",
        "Prinsip Utama": "Constructive Alignment (Retorika Klasik & Debat Formal)"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_as_section(doc, "I. MATRIKS KESELARASAN ASESMEN")
    al_rows = [
        ["IKTP 4.1", "Menerapkan retorika klasik (Ethos, Pathos, Logos).", "Tes Tertulis", "Pilihan Ganda & Analisis Pidato"],
        ["IKTP 4.2", "Menggunakan perangkat retoris (tricolon, parallelism, antithesis).", "Tes Tertulis / Kinerja", "Soal Analisis Gaya Bahasa Oratori"],
        ["IKTP 4.3", "Mengevaluasi falasi logika dan merumuskan sanggahan debat.", "Tes Tertulis HOTS", "Soal Kasus Rebuttal Konstruktif"],
        ["IKTP 4.4", "Membawakan pidato oratori persuasif 3-4 menit di podium.", "Unjuk Kerja Lisan", "Rubrik Public Speaking & Parliamentary Debate"]
    ]
    docx_helper.add_styled_table(doc, ["Kode IKTP", "Indikator Ketercapaian", "Teknik Asesmen", "Bentuk Instrumen"], al_rows, [Inches(1.2), Inches(3.0), Inches(1.3), Inches(1.5)])

    add_as_section(doc, "II. KISI-KISI & SOAL ASESMEN SUMATIF")
    pg_text = (
        "Stimulus: Read the oratorical excerpt to answer questions 1 – 5:\n"
        "'The laws of atmospheric physics do not compromise with political timetables; the melting point of permafrost does not wait for fiscal quarters; and the rising tides do not retreat for diplomatic platitudes. Let us examine the evidence with cold, unyielding precision (Logos). The IPCC has demonstrated that breaching 1.5°C is an irreversible planetary tipping point. Yet, while renewable investments expand, subsidies for fossil fuels reached an astronomical seven trillion dollars last year. As leaders, your legitimacy rests on your moral fidelity to protect the defenseless (Ethos). Consider the subsistence farmer whose crops withered in three years of drought (Pathos). The hour for hesitation has expired; the hour for courage is now.'\n\n"
        "1. Which classical rhetorical appeal is established by presenting IPCC data and the seven-trillion-dollar fossil subsidy figure?\n"
        "   A. Pathos\n"
        "   B. Logos\n"
        "   C. Mythos\n"
        "   D. Hubris\n"
        "   E. Kairos exclusively\n\n"
        "2. What rhetorical device is demonstrated in the opening sentence: 'laws of physics do not compromise... permafrost does not wait... tides do not retreat'?\n"
        "   A. Tricolon and parallelism\n"
        "   B. Hyperbole and spoonerism\n"
        "   C. Understatement and irony\n"
        "   D. Malapropism and slang\n"
        "   E. Passive inversion\n\n"
        "3. How does the orator deploy Pathos to evoke emotional moral responsibility?\n"
        "   A. By listing the legal articles of the United Nations Charter.\n"
        "   B. By invoking the vivid image of a subsistence farmer whose crops withered in prolonged drought.\n"
        "   C. By apologizing for the length of the speech.\n"
        "   D. By demanding an immediate financial refund for all delegates.\n"
        "   E. By joking about bad weather in Geneva.\n\n"
        "4. In parliamentary debating (Text 2), what is the function of 'signposting'?\n"
        "   A. Putting up wooden road signs outside the debate hall.\n"
        "   B. Verbally signaling the structural progression of arguments ('Firstly, Secondly, Thirdly') for adjudicators.\n"
        "   C. Holding up cards with written insults against the opposing team.\n"
        "   D. Whispering quietly to teammates during speech delivery.\n"
        "   E. Refusing to speak until the timer runs out.\n\n"
        "5. How did the Prime Minister in Text 2 refute the Opposition's claim that autonomous weapons reduce soldier casualties?\n"
        "   A. By denying that autonomous robots have computer chips.\n"
        "   B. By arguing that speed without ethical judgment is accelerated catastrophe that lowers the barrier to war.\n"
        "   C. By agreeing with the Opposition and resigning from the debate.\n"
        "   D. By stating that human soldiers enjoy danger.\n"
        "   E. By proposing that wars should be fought with bows and arrows.\n\n"
        "6. Which sentence represents an 'antithesis'?\n"
        "   A. To err is human; to forgive, divine.\n"
        "   B. The rain in Spain falls mainly on the plain.\n"
        "   C. She sells seashells by the seashore.\n"
        "   D. It is raining cats and dogs outside today.\n"
        "   E. I ran and ran and ran until I fell.\n\n"
        "7. What constitutes a 'Straw Man' logical fallacy in formal debate?\n"
        "   A. Building a wooden scarecrow to frighten the audience.\n"
        "   B. Misrepresenting or exaggerating an opponent's argument to make it easier to attack.\n"
        "   C. Citing official government statistics accurately.\n"
        "   D. Admitting that your evidence is incomplete.\n"
        "   E. Speaking with a calm, polite tone of voice.\n\n"
        "8. In competitive debate, what is an 'even-if' analysis?\n"
        "   A. An argument proving that two numbers are mathematically even.\n"
        "   B. A strategic argument demonstrating that even if the opponent's premise were true, their conclusion still fails.\n"
        "   C. Refusing to debate unless coin tosses end in even numbers.\n"
        "   D. Repeating the exact same sentence twice.\n"
        "   E. Quoting poetry instead of facts.\n\n"
        "9. What is the optimal stance for an orator delivering a formal address at a podium?\n"
        "   A. Slouching backward with hands stuffed deep into pockets.\n"
        "   B. Upright posture, balanced weight, expressive open gestures, and sustained eye contact across the auditorium.\n"
        "   C. Pacing frantically back and forth across the stage without stopping.\n"
        "   D. Staring exclusively at one's own shoes or reading from paper without looking up.\n"
        "   E. Leaning heavily on one elbow and yawning.\n\n"
        "10. What is the fundamental ethical responsibility of an orator regarding rhetorical persuasion?\n"
        "   A. To manipulate and deceive the public by fabricating emotional falsehoods.\n"
        "   B. To harness the beauty of language to illuminate truth and move citizens toward the common good.\n"
        "   C. To ensure that only wealthy delegates are allowed to speak.\n"
        "   D. To talk uninterrupted for ten continuous hours.\n"
        "   E. To avoid stating any clear opinions."
    )
    p_pg = doc.add_paragraph()
    p_pg.paragraph_format.left_indent = Inches(0.2)
    p_pg.add_run(pg_text)

    add_as_sub(doc, "C. Naskah Soal Uraian & Studi Retorika")
    p_es = doc.add_paragraph()
    p_es.paragraph_format.left_indent = Inches(0.2)
    p_es.add_run("1. Deconstruct how Winston Churchill's famous speech ('We shall fight on the beaches...') utilizes parallelism to transform military despair into national resolve.\n")
    p_es.add_run("2. Explain how a speaker establishes Ethos (moral credibility) before an audience that initially disagrees with their policy stance.\n")
    p_es.add_run("3. Construct a 3-part debate rebuttal against the motion: 'This House Would Ban Space Tourism'. Use signposting and 'even-if' logic.\n")
    p_es.add_run("4. Differentiate between an 'Ad Hominem' attack and a legitimate critique of an opponent's evidentiary credentials.\n")
    p_es.add_run("5. Compose an oratorical paragraph (60-80 words) championing educational equality utilizing both a tricolon and an antithesis.")

    add_as_sub(doc, "D. Rubrik Unjuk Kerja Pidato Oratori (Skala 1 - 4)")
    docx_helper.add_callout(
        doc,
        "Kunci Jawaban PG & Formula Nilai Akhir:",
        "Kunci Jawaban PG:\n"
        "1. B | 2. A | 3. B | 4. B | 5. B | 6. A | 7. B | 8. B | 9. B | 10. B\n\n"
        "Pedoman Penskoran:\n"
        "• Nilai PG = (Benar / 10) x 40 poin\n"
        "• Nilai Soal Uraian = (Skor / 20) x 20 poin\n"
        "• Nilai Performa Pidato Oratori di Podium = (Skor / 16) x 40 poin\n"
        "• NILAI AKHIR (NA) = Nilai PG + Nilai Uraian + Nilai Pidato (Total Maks: 100)",
        bg="DCFCE7", border_color="16A34A"
    )

    fpath = os.path.join(TARGET_DIR, "Instrumen_Asesmen_TP_12.TL.4_Persuasive_Oratory_and_Debate.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# ASESMEN 5: TP 12.TL.5 - ADMISSIONS & CAREER DOSSIER
# ==============================================================================
def create_asesmen_5():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "INSTRUMEN ASESMEN LENGKAP - BAHASA INGGRIS TINGKAT LANJUT KELAS XII",
        "Modul 5: Admissions & Career Dossier | TP 12.TL.5"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XII / Semester 6 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 12.TL.5 - Menyusun korespondensi kerja profesional dan akademik internasional: Surat Lamaran Kerja Eksekutif, Curriculum Vitae (CV) ATS-compliant, dan Pernyataan Motivasi Beasiswa/Kuliah (Personal Statement).",
        "Prinsip Utama": "Constructive Alignment (Korespondensi Eksekutif & Portofolio)"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_as_section(doc, "I. MATRIKS KESELARASAN ASESMEN")
    al_rows = [
        ["IKTP 5.1", "Menyusun Executive Cover Letter dengan register bisnis formal.", "Penilaian Produk", "Surat Lamaran Kerja Eksekutif"],
        ["IKTP 5.2", "Merancang Resume ATS-compliant berbasis Google X-Y-Z formula.", "Penilaian Produk", "Curriculum Vitae Standar ATS 1 Halaman"],
        ["IKTP 5.3", "Menulis Personal Statement 500-600 kata dengan narrative hook.", "Penilaian Produk", "Personal Statement Beasiswa Internasional"],
        ["IKTP 5.4", "Mengevaluasi keselarasan berkas terhadap standar seleksi global.", "Tes Tertulis HOTS", "Soal Analisis Evaluasi Portofolio Admisi"]
    ]
    docx_helper.add_styled_table(doc, ["Kode IKTP", "Indikator Ketercapaian", "Teknik Asesmen", "Bentuk Instrumen"], al_rows, [Inches(1.2), Inches(3.0), Inches(1.3), Inches(1.5)])

    add_as_section(doc, "II. KISI-KISI & SOAL ASESMEN SUMATIF")
    pg_text = (
        "Stimulus: Read the professional excerpts to answer questions 1 – 5:\n"
        "'Modern recruitment and university admissions utilize Applicant Tracking Systems (ATS)—automated software that parses, indexes, and ranks resumes based on semantic keywords and quantified achievements. To ensure high ranking: (1) Use clean, single-column formatting without tables or icons; (2) Precede every bullet with a dynamic action verb; (3) Deploy the Google X-Y-Z formula: Accomplished [X], as measured by [Y], by doing [Z]. For personal statements, top candidates establish an authentic narrative hook: connecting a childhood observation to tangible high school inquiry, demonstrating institutional fit, and outlining future social impact.'\n\n"
        "1. Why do modern companies and universities deploy Applicant Tracking Systems (ATS)?\n"
        "   A. To design colorful graphical banners for career fairs.\n"
        "   B. To automatically parse, index, and filter thousands of resumes based on keywords and competencies.\n"
        "   C. To replace all human workers with factory robots.\n"
        "   D. To conduct random background criminal investigations.\n"
        "   E. To prevent students from applying to universities.\n\n"
        "2. Which resume formatting choice is recommended to prevent ATS parsing errors?\n"
        "   A. Complex multi-column tables, colored text boxes, and embedded graphical icons.\n"
        "   B. Clean single-column layout with standard fonts (Segoe UI / Calibri) and clear section headings.\n"
        "   C. Saving the resume as an unreadable pixelated JPEG image.\n"
        "   D. Using decorative medieval gothic fonts and floral borders.\n"
        "   E. Submitting the resume written completely by hand on parchment.\n\n"
        "3. Which resume bullet point strictly implements the Google X-Y-Z formula?\n"
        "   A. I worked very hard in the school club and everyone liked me a lot.\n"
        "   B. Increased regional recycling participation by 35% across 4 districts by organizing a digital waste-bank app campaign.\n"
        "   C. Responsible for making social media posts whenever the president asked.\n"
        "   D. Helped with several clean energy projects during the summer vacation.\n"
        "   E. Member of the science club from 2023 to 2025.\n\n"
        "4. In a university personal statement, what is the role of the 'Narrative Hook'?\n"
        "   A. To boast about personal family wealth and social connections.\n"
        "   B. To capture the reader's curiosity through a compelling personal moment of intellectual awakening.\n"
        "   C. To copy and paste the entire course syllabus of the university.\n"
        "   D. To list every single examination grade since elementary school.\n"
        "   E. To beg the committee for financial aid.\n\n"
        "5. In Text 2 (Executive Cover Letter), what is the candidate's strategic objective in paragraph 2?\n"
        "   A. To complain about past employers and colleagues.\n"
        "   B. To demonstrate concrete technical capabilities (R, Python, econometric modeling) applied in high-impact projects.\n"
        "   C. To demand an immediate salary figure and vacation schedule.\n"
        "   D. To discuss personal hobbies such as playing video games.\n"
        "   E. To summarize a novel they read last week.\n\n"
        "6. Which valediction (closing phrase) is standard for a formal British/American business letter where the recipient's name is known?\n"
        "   A. Yours sincerely,\n"
        "   B. Cheers mate,\n"
        "   C. Talk to you soon,\n"
        "   D. Besties forever,\n"
        "   E. See ya later,\n\n"
        "7. What does 'institutional fit' mean in the context of admissions personal statements?\n"
        "   A. Buying clothes that match the university's official color palette.\n"
        "   B. Articulating specific professors, laboratories, and curricula that uniquely align with the applicant's research goals.\n"
        "   C. Proving that the applicant lives within walking distance of the campus.\n"
        "   D. Memorizing the university fight song.\n"
        "   E. Agreeing with every political opinion on social media.\n\n"
        "8. Why is using passive voice ('Was given the task of...') discouraged on executive resumes?\n"
        "   A. Because passive voice increases the word count of the resume.\n"
        "   B. Because it obscures individual initiative and projects a passive, reactive mindset rather than leadership.\n"
        "   C. Because computers cannot read passive verbs.\n"
        "   D. Because English grammar forbids passive voice in workplaces.\n"
        "   E. Because active voice is cheaper to print.\n\n"
        "9. What is a 'quantifiable metric' on a professional resume?\n"
        "   A. A vague feeling of satisfaction.\n"
        "   B. Concrete numbers, percentages, dollar values, or timeframes demonstrating measurable impact.\n"
        "   C. The physical weight of the printed paper.\n"
        "   D. The number of signatures on the bottom of the page.\n"
        "   E. The number of fonts used in the document.\n\n"
        "10. What is the overarching objective of the complete Admissions Dossier?\n"
        "   A. To confuse the admissions officers with contradictory information.\n"
        "   B. To present a unified, coherent narrative of intellectual rigor, character integrity, and future leadership potential.\n"
        "   C. To provide a fictional storybook about an imaginary superhero.\n"
        "   D. To submit as many pages as possible regardless of quality.\n"
        "   E. To replicate standard textbook exercises."
    )
    p_pg = doc.add_paragraph()
    p_pg.paragraph_format.left_indent = Inches(0.2)
    p_pg.add_run(pg_text)

    add_as_sub(doc, "C. Naskah Soal Uraian / Studi Kasus Admisi")
    p_es = doc.add_paragraph()
    p_es.paragraph_format.left_indent = Inches(0.2)
    p_es.add_run("1. Transform this weak resume bullet into a high-impact Google X-Y-Z bullet: 'I was in charge of school tree planting and we planted many trees.'\n")
    p_es.add_run("2. Analyze why opening a personal statement with a generic quotation (e.g., 'Nelson Mandela once said...') is considered ineffective by Ivy League admissions officers.\n")
    p_es.add_run("3. In an executive cover letter, how does a candidate demonstrate 'cultural diplomacy and institutional alignment'?\n")
    p_es.add_run("4. Explain the strategic rationale for tailoring resume keywords to the explicit job description or scholarship terms of reference.\n")
    p_es.add_run("5. Draft an opening paragraph (50-70 words) for a scholarship personal statement based on a real obstacle you overcame in STEM or the humanities.")

    add_as_sub(doc, "D. Tugas Portofolio Dossier Admisi & Rubrik Lengkap")
    docx_helper.add_callout(
        doc,
        "Kunci Jawaban PG & Formula Nilai Akhir:",
        "Kunci Jawaban PG:\n"
        "1. B | 2. B | 3. B | 4. B | 5. B | 6. A | 7. B | 8. B | 9. B | 10. B\n\n"
        "Pedoman Penskoran:\n"
        "• Nilai PG = (Benar / 10) x 40 poin\n"
        "• Nilai Soal Uraian = (Skor / 20) x 20 poin\n"
        "• Nilai Dossier Lengkap (Cover Letter + ATS CV + Personal Statement) = (Skor / 16) x 40 poin\n"
        "• NILAI AKHIR (NA) = Nilai PG + Nilai Uraian + Nilai Portofolio (Total Maks: 100)",
        bg="DCFCE7", border_color="16A34A"
    )

    fpath = os.path.join(TARGET_DIR, "Instrumen_Asesmen_TP_12.TL.5_Admissions_and_Career_Dossier.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# ASESMEN 6: TP 12.TL.6 - JOB INTERVIEW & ORAL DEFENSE
# ==============================================================================
def create_asesmen_6():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "INSTRUMEN ASESMEN LENGKAP - BAHASA INGGRIS TINGKAT LANJUT KELAS XII",
        "Modul 6: Executive Interview & Academic Oral Defense | TP 12.TL.6"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XII / Semester 6 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 12.TL.6 - Mempraktikkan wawancara kerja profesional (Job Interview Simulation) dan pertahanan karya ilmiah lisan (Oral Defense) dengan respons spontan, taktis, dan register bahasa formal.",
        "Prinsip Utama": "Constructive Alignment (Kemahiran Spontan Lisan Berbobot)"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_as_section(doc, "I. MATRIKS KESELARASAN ASESMEN")
    al_rows = [
        ["IKTP 6.1", "Menerapkan metode STAR dalam pertanyaan perilaku.", "Tes Lisan Kinerja", "Simulasi Wawancara Eksekutif STAR"],
        ["IKTP 6.2", "Mempertahankan kesimpulan riset di hadapan dewan penguji.", "Tes Lisan Kinerja", "Simulasi Sidang Viva Voce Tesis"],
        ["IKTP 6.3", "Menggunakan jeda taktis dan tata krama diplomatis lisan.", "Observasi Lisan", "Lembar Pengamatan Kelancaran Spontan"],
        ["IKTP 6.4", "Menampilkan simulasi wawancara / sidang mandiri 4-5 menit.", "Unjuk Kerja Lisan", "Rubrik Ujian Lisan Eksekutif & Akademik"]
    ]
    docx_helper.add_styled_table(doc, ["Kode IKTP", "Indikator Ketercapaian", "Teknik Asesmen", "Bentuk Instrumen"], al_rows, [Inches(1.2), Inches(3.0), Inches(1.3), Inches(1.5)])

    add_as_section(doc, "II. KISI-KISI & SOAL ASESMEN SUMATIF")
    pg_text = (
        "Stimulus: Read the interview dialogue to answer questions 1 – 5:\n"
        "'INTERVIEWER: Describe a situation where your initial methodology failed, and how you recovered.\n"
        "CANDIDATE: Certainly. During our regional robotics competition, our primary LiDAR navigation sensor malfunctioned 48 hours prior to field trials [Situation]. As captain, my responsibility was to either forfeit the sensor trial or engineer an emergency algorithmic workaround within our budget [Task]. Instead of panic, I convened an emergency triage meeting. I salvaged an optical camera from an old smartphone and reprogrammed our computer-vision script from LiDAR point-clouds to open-source edge-detection heuristics [Action]. Consequently, our robot completed the obstacle course in 3.4 minutes, securing second place and receiving the Judges' Innovation Award. It taught me that adaptability consistently outperforms rigid perfectionism under pressure [Result].'\n\n"
        "1. What proportion of response time in the STAR framework should be devoted to describing specific personal Actions?\n"
        "   A. Less than 5%\n"
        "   B. Approximately 60%\n"
        "   C. Exactly 100%\n"
        "   D. 10%\n"
        "   E. 95%\n\n"
        "2. Why is the candidate's use of first-person verbs ('I convened', 'I salvaged', 'I reprogrammed') effective in this behavioral interview?\n"
        "   A. It shows arrogance and disrespect toward teammates.\n"
        "   B. It clearly delineates the candidate's individual leadership contribution and problem-solving agency.\n"
        "   C. It hides the fact that other team members did all the work.\n"
        "   D. It allows the candidate to avoid answering the question.\n"
        "   E. It proves that no team existed.\n\n"
        "3. In the Result component, what two distinct outcomes did the candidate highlight?\n"
        "   A. A financial fine and an apology to the judges.\n"
        "   B. A quantifiable performance metric (2nd place, 3.4 min) and an introspective leadership lesson on adaptability.\n"
        "   C. A new smartphone purchase and a vacation.\n"
        "   D. Quitting robotics to pursue literature.\n"
        "   E. Blaming the manufacturer for the broken sensor.\n\n"
        "4. In an academic oral defense (Text 2), how should a candidate respond to an examiner identifying a research limitation?\n"
        "   A. Argue aggressively and accuse the examiner of bias.\n"
        "   B. Politely acknowledge the limitation (tactical validation), clarify the methodological scope, and highlight data triangulation.\n"
        "   C. Burst into tears and shred the research paper.\n"
        "   D. Pretend not to hear the question and change the subject.\n"
        "   E. Walk out of the examination room immediately.\n\n"
        "5. What is the psychological function of 'tactical fillers' (Text 3) like 'That is an incisive question; let me address its premise'?\n"
        "   A. To deceive the panel into thinking the interview is finished.\n"
        "   B. To purchase cognitive processing time without projecting verbal disfluency or nervousness.\n"
        "   C. To insult the interviewer's intelligence.\n"
        "   D. To fill up the entire interview time without giving any facts.\n"
        "   E. To test the room's acoustic microphone.\n\n"
        "6. Which phrase is an example of a disruptive vocalized disfluency to avoid in formal interviews?\n"
        "   A. 'Consequently, our findings indicate...'\n"
        "   B. 'Like, you know, err, basically, uhm...'\n"
        "   C. 'Furthermore, we observed...'\n"
        "   D. 'If I may clarify our epistemological framework...'\n"
        "   E. 'Certainly, Professor Hartanto.'\n\n"
        "7. Why is maintaining steady, respectful eye contact essential during oral interrogation?\n"
        "   A. It intimidates the examiners so they stop asking questions.\n"
        "   B. It projects intellectual gravitas, transparency, confidence, and interpersonal engagement.\n"
        "   C. It proves that the candidate is not wearing contact lenses.\n"
        "   D. It allows the candidate to read notes reflected in the examiner's glasses.\n"
        "   E. It is required by international traffic laws.\n\n"
        "8. In research defense terminology, what does 'triangulation' mean?\n"
        "   A. Drawing geometric triangles on the blackboard.\n"
        "   B. Cross-validating research findings through multiple data collection methods or data sources.\n"
        "   C. Having three examiners sitting in triangular formation.\n"
        "   D. Publishing three papers on unrelated subjects.\n"
        "   E. Dividing the thesis budget by three.\n\n"
        "9. What does the Latin term 'viva voce' literally translate to in academic examinations?\n"
        "   A. By written contract.\n"
        "   B. With the living voice (oral examination).\n"
        "   C. In total silence.\n"
        "   D. Through digital email.\n"
        "   E. Under judicial oath.\n\n"
        "10. Which candidate demonstrates executive composure under pressure?\n"
        "   A. The candidate who talks over the interviewer and shouts louder.\n"
        "   B. The candidate who pauses calmly, breathes deeply, validates the question, and delivers structured evidence.\n"
        "   C. The candidate who refuses to speak until their lawyer arrives.\n"
        "   D. The candidate who laughs hysterically at every question.\n"
        "   E. The candidate who reads memorized answers that do not address the question."
    )
    p_pg = doc.add_paragraph()
    p_pg.paragraph_format.left_indent = Inches(0.2)
    p_pg.add_run(pg_text)

    add_as_sub(doc, "C. Naskah Soal Uraian & Skenario Wawancara")
    p_es = doc.add_paragraph()
    p_es.paragraph_format.left_indent = Inches(0.2)
    p_es.add_run("1. Deconstruct why behavioral interview questions ('Tell me about a time when...') are superior predictors of candidate competence than hypothetical questions ('What would you do if...').\n")
    p_es.add_run("2. Script a 4-part STAR response to: 'Describe a situation where you had to collaborate with a difficult team member to achieve a tight deadline.'\n")
    p_es.add_run("3. In a thesis defense, what is the strategic difference between 'internal validity' and 'demographic generalizability'?\n")
    p_es.add_run("4. Explain how a deliberate 2-second silent pause can project authority rather than ignorance.\n")
    p_es.add_run("5. Formulate three tactical delay expressions you can deploy in high-stakes English interviews.")

    add_as_sub(doc, "D. Rubrik Unjuk Kerja Simulasi Wawancara & Oral Defense")
    docx_helper.add_callout(
        doc,
        "Kunci Jawaban PG & Formula Nilai Akhir:",
        "Kunci Jawaban PG:\n"
        "1. B | 2. B | 3. B | 4. B | 5. B | 6. B | 7. B | 8. B | 9. B | 10. B\n\n"
        "Pedoman Penskoran:\n"
        "• Nilai PG = (Benar / 10) x 40 poin\n"
        "• Nilai Soal Uraian = (Skor / 20) x 20 poin\n"
        "• Nilai Simulasi Ujian Lisan di Hadapan Panel = (Skor / 16) x 40 poin\n"
        "• NILAI AKHIR (NA) = Nilai PG + Nilai Uraian + Nilai Simulasi (Total Maks: 100)",
        bg="DCFCE7", border_color="16A34A"
    )

    fpath = os.path.join(TARGET_DIR, "Instrumen_Asesmen_TP_12.TL.6_Job_Interview_and_Oral_Defense.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# ASESMEN 7: TP 12.TL.7 - RESEARCH ABSTRACT & INFOGRAPHIC
# ==============================================================================
def create_asesmen_7():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "INSTRUMEN ASESMEN LENGKAP - BAHASA INGGRIS TINGKAT LANJUT KELAS XII",
        "Modul 7: Scientific Abstract & Infographic Symposium | TP 12.TL.7"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XII / Semester 6 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 12.TL.7 - Menyusun abstrak laporan penelitian ilmiah (Research Abstract) 250 kata dan mendesain presentasi infografis saintifik untuk diseminasi publik dalam simposium ilmiah siswa.",
        "Prinsip Utama": "Constructive Alignment (Presisi Saintifik & Literasi Multimodal)"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_as_section(doc, "I. MATRIKS KESELARASAN ASESMEN")
    al_rows = [
        ["IKTP 7.1", "Merangkum riset ke struktur IMRaD maksimal 250 kata.", "Penilaian Produk", "Naskah Abstrak Terstruktur 250 Kata"],
        ["IKTP 7.2", "Menerapkan nominalisasi ilmiah dan kalimat pasif objektif.", "Tes Tertulis", "Soal Analisis Transformasi Sintaksis"],
        ["IKTP 7.3", "Mendesain poster infografis tiga kolom (Ten-Second Rule).", "Penilaian Produk Multimodal", "Desain Poster Infografis Saintifik"],
        ["IKTP 7.4", "Menyajikan presentasi flash talk 2 menit dalam simposium.", "Unjuk Kerja Lisan", "Rubrik Gelar Simposium Ilmiah Siswa"]
    ]
    docx_helper.add_styled_table(doc, ["Kode IKTP", "Indikator Ketercapaian", "Teknik Asesmen", "Bentuk Instrumen"], al_rows, [Inches(1.2), Inches(3.0), Inches(1.3), Inches(1.5)])

    add_as_section(doc, "II. KISI-KISI & SOAL ASESMEN SUMATIF")
    pg_text = (
        "Stimulus: Read the model scientific abstract to answer questions 1 – 5:\n"
        "'[Background] Industrial wastewater contamination by toxic heavy metals represents an ecological hazard. This paper investigates the green synthesis of iron oxide magnetic nanoparticles (Fe3O4 NPs) utilizing discarded robusta coffee husk extracts for hexavalent chromium [Cr(VI)] adsorption from textile effluent.\n"
        "[Methodology] Fe3O4 NPs were characterized via XRD, TEM, and FTIR. Batch adsorption evaluated pH levels (2.0–9.0), contact durations (10–120 min), and dosages (0.2–1.5 g/L). Langmuir isotherm models determined adsorption thermodynamics.\n"
        "[Results] Optimal Cr(VI) sequestration occurred at pH 3.0 within 45 minutes, exhibiting a maximum adsorption capacity (qmax) of 84.6 mg/g. Equilibrium kinetics adhered strictly to the pseudo-second-order model (R² = 0.994). Following magnetic separation, the biosorbent retained 91.2% adsorption efficacy across five consecutive regeneration cycles.\n"
        "[Conclusion] Utilizing coffee husk waste provides an eco-friendly, cost-effective platform for industrial effluent treatment.'\n\n"
        "1. What is the maximum adsorption capacity (qmax) of the biosorbent established in the Results section?\n"
        "   A. 45 minutes\n"
        "   B. 84.6 mg/g\n"
        "   C. 91.2%\n"
        "   D. R² = 0.994\n"
        "   E. pH 3.0\n\n"
        "2. Which section of the IMRaD abstract receives the largest percentage of word allocation in top-tier journals?\n"
        "   A. Results (empirical findings and quantitative data)\n"
        "   B. Acknowledgments of family members\n"
        "   C. Background historical gossip\n"
        "   D. Appendix calculations\n"
        "   E. Title page borders\n\n"
        "3. In scientific writing, what is the primary benefit of nominalization (e.g., changing 'absorbed' to 'adsorption')?\n"
        "   A. It makes the paper sound poetic and artistic.\n"
        "   B. It compresses complex causal processes into tight grammatical noun subjects, maximizing lexical density.\n"
        "   C. It allows writers to avoid citing past researchers.\n"
        "   D. It guarantees that no numbers are needed.\n"
        "   E. It is required for children's books.\n\n"
        "4. According to conference design best practices (Text 2), what is the 'Ten-Second Rule' for scientific posters?\n"
        "   A. The presenter must finish speaking in exactly ten seconds.\n"
        "   B. An attendee walking past should comprehend the core research contribution within ten seconds.\n"
        "   C. The poster will fall off the wall after ten seconds.\n"
        "   D. Attendees are only permitted to look at posters for ten seconds.\n"
        "   E. Ten judges must evaluate each poster at once.\n\n"
        "5. Which headline is formulated as an active, conclusion-driven headline for an academic infographic?\n"
        "   A. A General Inquiry on Wastewater\n"
        "   B. Coffee Husk Nanoparticles Remove 91% of Toxic Chromium from Industrial Wastewater\n"
        "   C. Some Observations on Nanoparticles and Chemistry\n"
        "   D. Chemistry Project: Grade 12 Secondary School\n"
        "   E. Why We Like Green Science Experiments\n\n"
        "6. What is the recommended word limit for the body text of a scientific conference poster?\n"
        "   A. Over 3,000 words\n"
        "   B. Fewer than 400 words, utilizing diagrams and visual data charts\n"
        "   C. Exactly one single word\n"
        "   D. Zero words, only photographs\n"
        "   E. 10,000 words printed in 6pt font\n\n"
        "7. Which sentence correctly transforms this wordy sentence into a nominalized scientific sentence:\n"
        "   'When bacteria grew rapidly in the incubator, the temperature increased.'\n"
        "   A. Bacteria were happy and grew fast.\n"
        "   B. Rapid bacterial proliferation induced a temperature increase in the incubator.\n"
        "   C. The incubator was hot because of bacteria.\n"
        "   D. Because it grew, the temperature was high.\n"
        "   E. Bacteria made the machine warm.\n\n"
        "8. In research poster layout, what is the function of the three-column visual hierarchy?\n"
        "   A. To divide the poster into three different languages.\n"
        "   B. To guide the reader's eye logically from Problem (Left) to Methodology (Center) to Key Results and Practical Impact (Right).\n"
        "   C. To allow three students to stand in front of the poster simultaneously.\n"
        "   D. To save printing ink.\n"
        "   E. To replicate social media feeds.\n\n"
        "9. What is a 'Flash Talk' in an academic conference symposium?\n"
        "   A. Taking pictures with camera flash.\n"
        "   B. A concise, 2-to-3 minute oral pitch summarizing the core findings and impact of a research poster.\n"
        "   C. A debate where speakers interrupt each other.\n"
        "   D. An electrical engineering demonstration.\n"
        "   E. A silent mime performance.\n\n"
        "10. What is the strict word ceiling for standard international scientific abstracts (e.g., Scopus / Nature)?\n"
        "   A. 1,500 words\n"
        "   B. 250 words\n"
        "   C. 50 words\n"
        "   D. 5,000 words\n"
        "   E. Unlimited words"
    )
    p_pg = doc.add_paragraph()
    p_pg.paragraph_format.left_indent = Inches(0.2)
    p_pg.add_run(pg_text)

    add_as_sub(doc, "C. Naskah Soal Uraian & Rekayasa Abstrak")
    p_es = doc.add_paragraph()
    p_es.paragraph_format.left_indent = Inches(0.2)
    p_es.add_run("1. Deconstruct the 4 core sections of the IMRaD acronym and describe what specific information belongs in each section.\n")
    p_es.add_run("2. Transform these two sentences into one concise scientific sentence utilizing nominalization: 'The chemical solution precipitated rapidly. This occurred because thermal energy was absorbed during the reaction.'\n")
    p_es.add_run("3. Explain how visual workflow diagrams in the Methodology section of a poster communicate research rigor more efficiently than paragraphs of text.\n")
    p_es.add_run("4. Why must scientific abstracts report quantitative metrics (e.g., p-values, R², percentages) rather than vague qualitative adjectives (e.g., 'very good', 'huge')?\n")
    p_es.add_run("5. Draft a 40-word [Conclusion & Significance] section based on a study proving that banana peel biochar filters purify microplastics from well water.")

    add_as_sub(doc, "D. Tugas Produk Abstrak 250 Kata & Simposium Poster")
    docx_helper.add_callout(
        doc,
        "Kunci Jawaban PG & Formula Nilai Akhir:",
        "Kunci Jawaban PG:\n"
        "1. B | 2. A | 3. B | 4. B | 5. B | 6. B | 7. B | 8. B | 9. B | 10. B\n\n"
        "Pedoman Penskoran:\n"
        "• Nilai PG = (Benar / 10) x 40 poin\n"
        "• Nilai Soal Uraian = (Skor / 20) x 20 poin\n"
        "• Nilai Produk Poster & Abstrak Simposium = (Skor / 16) x 40 poin\n"
        "• NILAI AKHIR (NA) = Nilai PG + Nilai Uraian + Nilai Simposium (Total Maks: 100)",
        bg="DCFCE7", border_color="16A34A"
    )

    fpath = os.path.join(TARGET_DIR, "Instrumen_Asesmen_TP_12.TL.7_Research_Abstract_and_Infographic.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


# ==============================================================================
# ASESMEN 8: TP 12.TL.8 - STANDARDIZED READINESS (B2+)
# ==============================================================================
def create_asesmen_8():
    doc = docx_helper.create_styled_doc()
    docx_helper.add_header_banner(
        doc,
        "INSTRUMEN ASESMEN LENGKAP - BAHASA INGGRIS TINGKAT LANJUT KELAS XII",
        "Modul 8: Standardized English Proficiency (IELTS/TOEFL/CEFR B2+) | TP 12.TL.8"
    )
    info = {
        "Mata Pelajaran": "Bahasa Inggris Tingkat Lanjut (Pilihan)",
        "Fase / Kelas / Semester": "Fase F Tingkat Lanjut / Kelas XII / Semester 6 (Genap)",
        "Kode & Tujuan Pembelajaran": "TP 12.TL.8 - Mengembangkan kesiapan literasi bahasa Inggris standar internasional (International Academic Literacy: IELTS/TOEFL/CEFR B2+ reading & critical analytical writing).",
        "Prinsip Utama": "Constructive Alignment (Benchmark Standar Internasional)"
    }
    docx_helper.add_tp_info_box(doc, info)

    add_as_section(doc, "I. MATRIKS KESELARASAN ASESMEN")
    al_rows = [
        ["IKTP 8.1", "Menerapkan skimming & inferencing pada teks saintifik panjang.", "Tes Tertulis Standardized", "IELTS Academic Reading Module"],
        ["IKTP 8.2", "Menyusun laporan data grafik (Writing Task 1) dalam 20 menit.", "Tes Tertulis Standardized", "IELTS Academic Writing Task 1"],
        ["IKTP 8.3", "Menerapkan variasi kohesi dan struktur inversi/partisipial.", "Tes Analisis Bahasa", "Band Descriptors Lexical Scoring"],
        ["IKTP 8.4", "Menyelesaikan simulasi tes benchmark kemahiran global.", "Diagnostic Benchmark Test", "Skor Prediksi CEFR B2+ / IELTS Band"]
    ]
    docx_helper.add_styled_table(doc, ["Kode IKTP", "Indikator Ketercapaian", "Teknik Asesmen", "Bentuk Instrumen"], al_rows, [Inches(1.2), Inches(3.0), Inches(1.3), Inches(1.5)])

    add_as_section(doc, "II. KISI-KISI & SOAL ASESMEN SUMATIF")
    pg_text = (
        "Stimulus: Read the passage to answer questions 1 – 5:\n"
        "'Buried beneath the Antarctic ice sheet lies an atmospheric history spanning 800,000 years. By extracting cylindrical ice cores from Dome C, paleoclimatologists reconstruct ancient climate dynamics. Trapped air bubbles preserve primordial atmospheric compositions. Crucially, isotopic analysis of deuterium and oxygen-18 within ice water molecules enables scientists to calculate historic ambient surface temperatures. When correlated with carbon dioxide concentrations, the data reveal an unyielding correlation: greenhouse gas levels and global temperatures have moved in tight lockstep across nearly a million years. However, contemporary research reveals an alarming divergence: natural cycles fluctuated within 180 to 280 ppm, whereas post-industrial combustion has driven CO2 past 420 ppm.'\n\n"
        "1. How do paleoclimatologists reconstruct ancient surface temperatures from Antarctic ice cores?\n"
        "   A. By measuring the physical weight of polar snow.\n"
        "   B. By analyzing isotopic ratios of deuterium and oxygen-18 in trapped water molecules.\n"
        "   C. By looking at ancient paintings carved on the glaciers.\n"
        "   D. By measuring modern oceanic tides.\n"
        "   E. By asking local Antarctic wildlife.\n\n"
        "2. What historical fact is revealed by correlating trapped air bubbles with isotopic temperature data?\n"
        "   A. Greenhouse gases and global temperatures have moved in tight synchronous correlation across nearly a million years.\n"
        "   B. Carbon dioxide has no measurable effect on planetary temperatures.\n"
        "   C. The Earth was completely frozen until ten years ago.\n"
        "   D. Antarctica was once located on the equator.\n"
        "   E. Volcanoes are the sole cause of atmospheric carbon.\n\n"
        "3. What represents the 'alarming divergence' identified in modern climate records?\n"
        "   A. Atmospheric CO2 has plummeted to zero.\n"
        "   B. Modern CO2 levels have exceeded 420 ppm, far outpacing the 800,000-year natural ceiling of 280 ppm.\n"
        "   C. Glaciers are expanding rapidly across all seven continents.\n"
        "   D. Global temperatures are cooling at an exponential rate.\n"
        "   E. Polar ice sheets have turned into liquid oil.\n\n"
        "4. In IELTS Academic Writing Task 1 (Text 2), what is the most critical paragraph required to score Band 7.0 or higher?\n"
        "   A. A long personal anecdote about why the student likes clean energy.\n"
        "   B. A clear, comprehensive 'Overview' summarizing the main trends and significant differences without listing raw numbers.\n"
        "   C. A list of chemical formulas of solar panels.\n"
        "   D. A concluding political speech attacking fossil fuel executives.\n"
        "   E. Copying the entire test prompt word-for-word three times.\n\n"
        "5. In 2020, what significant milestone occurred between clean energy and fossil fuel investments according to Text 2?\n"
        "   A. Clean energy spending dropped to zero.\n"
        "   B. Clean energy investment reached financial parity with fossil fuels at $900 billion each.\n"
        "   C. Fossil fuel allocations grew tenfold.\n"
        "   D. All private banks stopped investing in technology.\n"
        "   E. The price of oil reached one million dollars per barrel.\n\n"
        "6. Which sentence demonstrates formal grammatical inversion to achieve stylistic emphasis?\n"
        "   A. Not only did clean energy investments accelerate, but they also surpassed fossil allocations by 2020.\n"
        "   B. Clean energy investments accelerated and went up.\n"
        "   C. They did accelerate because of policies.\n"
        "   D. Investments were going up and up very fast.\n"
        "   E. Clean energy is good, and fossil fuels are bad.\n\n"
        "7. In IELTS Band Descriptors, what does 'Lexical Resource' evaluate?\n"
        "   A. The neatness of the candidate's handwriting.\n"
        "   B. The range, precision, natural collocation, and stylistic sophistication of vocabulary used.\n"
        "   C. The number of dictionary pages the candidate has read.\n"
        "   D. Whether the candidate speaks with a British accent.\n"
        "   E. How loud the candidate reads the test paper.\n\n"
        "8. Which word represents a sophisticated academic synonym for 'stayed the same' when describing graph data?\n"
        "   A. Plateaued / Stabilized\n"
        "   B. Surged\n"
        "   C. Plummeted\n"
        "   D. Zigzagged\n"
        "   E. Bounced\n\n"
        "9. What is the standard time allocation recommended for IELTS Academic Writing Task 1?\n"
        "   A. 60 minutes\n"
        "   B. 20 minutes (writing at least 150 words)\n"
        "   C. 5 minutes\n"
        "   D. 2 hours\n"
        "   E. Unlimited time\n\n"
        "10. Achieving a CEFR B2+ level signifies that a secondary school graduate is capable of...\n"
        "   A. Memorizing simple nursery rhymes.\n"
        "   B. Understanding complex academic and technical texts, synthesizing multi-source data, and communicating with spontaneous fluency in global higher education.\n"
        "   C. Translating single vocabulary words using a pocket dictionary.\n"
        "   D. Passing a driver's license test.\n"
        "   E. Ordering food in a fast-food restaurant exclusively."
    )
    p_pg = doc.add_paragraph()
    p_pg.paragraph_format.left_indent = Inches(0.2)
    p_pg.add_run(pg_text)

    add_as_sub(doc, "C. Naskah Soal Uraian & Latihan Task 1")
    p_es = doc.add_paragraph()
    p_es.paragraph_format.left_indent = Inches(0.2)
    p_es.add_run("1. Explain the operational difference between 'skimming for the main idea' and 'scanning for specific numerical data' during a timed reading exam.\n")
    p_es.add_run("2. Write an official 'Overview' paragraph (30-40 words) for a graph showing that electric vehicle sales surged from 2 million in 2018 to 14 million in 2024 while diesel car sales dropped by 60%.\n")
    p_es.add_run("3. Identify and correct the grammatical/prepositional errors in: 'The investment increased in $200 million until $500 million between 2020 to 2024.'\n")
    p_es.add_run("4. Transform this sentence using formal inversion: 'Global emissions did not drop; they expanded rapidly.' (Start with: 'Not only...')\n")
    p_es.add_run("5. Explain why personal opinions (e.g., 'I think clean energy is wonderful') must NEVER be included in an IELTS Academic Writing Task 1 report.")

    add_as_sub(doc, "D. Tugas Simulasi Benchmark IELTS Academic Task 1 & Rubrik")
    docx_helper.add_callout(
        doc,
        "Kunci Jawaban PG & Formula Nilai Akhir:",
        "Kunci Jawaban PG:\n"
        "1. B | 2. A | 3. B | 4. B | 5. B | 6. A | 7. B | 8. A | 9. B | 10. B\n\n"
        "Pedoman Penskoran:\n"
        "• Nilai PG = (Benar / 10) x 40 poin\n"
        "• Nilai Soal Uraian = (Skor / 20) x 20 poin\n"
        "• Nilai Benchmark Writing Task 1 (IELTS Band Descriptors) = (Skor / 16) x 40 poin\n"
        "• NILAI AKHIR (NA) = Nilai PG + Nilai Uraian + Nilai Benchmark (Total Maks: 100)",
        bg="DCFCE7", border_color="16A34A"
    )

    fpath = os.path.join(TARGET_DIR, "Instrumen_Asesmen_TP_12.TL.8_Standardized_English_Proficiency_Readiness.docx")
    doc.save(fpath)
    print(f"Created: {fpath}")


def main():
    print("Generating comprehensive assessment instruments for 5-Modul Ajar XII SMA TL...")
    create_asesmen_1()
    create_asesmen_2()
    create_asesmen_3()
    create_asesmen_4()
    create_asesmen_5()
    create_asesmen_6()
    create_asesmen_7()
    create_asesmen_8()
    print("All 8 Asesmen documents successfully generated!")

if __name__ == "__main__":
    main()

