# Perangkat Pembelajaran Bahasa Inggris SMA (Kurikulum Merdeka)

Repositori ini memuat seluruh dokumen administrasi dan perangkat pembelajaran mata pelajaran **Bahasa Inggris SMA (Wajib & Tingkat Lanjut)** sesuai regulasi **Kurikulum Merdeka** terbaru (Permendikdasmen No. 13 Tahun 2025 dan Kepka BSKAP Kemendikdasmen No. 020 Tahun 2026).

---

## 📁 Struktur Repositori

```text
English-SMA/
├── 0-Peraturan (PDF)/               # Dokumen resmi regulasi, CP, dan panduan Kurikulum Merdeka
│
├── Perangkat SMA Umum/              # Perangkat Pembelajaran SMA Nasional Umum
│   ├── 1-Capaian Pembelajaran/      # CP Fase E & Fase F (Wajib & TL)
│   ├── 2-Alur dan Tujuan Pembelajaran/ # ATP Bahasa Inggris SMA (Wajib & TL)
│   ├── 3-Program Tahunan dan Program Semester/ # Prota & Prosem terpadu Excel
│   ├── 4-Silabus/                   # Silabus lengkap & Pemetaan Materi Bahasa Inggris SMA
│   ├── 5-Modul Ajar X SMA/          # Modul Ajar, Materi Otentik (CEFR A2+/B1) Kelas X
│   ├── 5-Modul Ajar XI SMA/         # Modul Ajar, Materi Otentik (CEFR B1/B1+) Kelas XI Wajib
│   ├── 5-Modul Ajar XI SMA TL/      # Modul Ajar, Materi Otentik (CEFR B1+/B2) Kelas XI TL
│   ├── 5-Modul Ajar XII SMA/        # Modul Ajar, LKPD, Asesmen & Materi Otentik Kelas XII Wajib
│   └── 5-Modul Ajar XII SMA TL/     # Modul Ajar, LKPD, Asesmen & Materi Otentik Kelas XII TL
│
└── Perangkat_SMAN_1_Bae/            # Perangkat Khusus Implementasi SMAN 1 Bae Kudus
    ├── 0-CP-ATP-Prota-Promes-Silabus/ # Workbook Excel terpadu SMAN 1 Bae
    ├── 0-Bahan Teks Otentik/        # Kumpulan teks otentik per kelas dan semester
    │   ├── 1-Kelas_X_Fase_E/
    │   ├── 2-Kelas_XI_Fase_F_Wajib/
    │   ├── 3-Kelas_XI_Fase_F_Tingkat_Lanjut/
    │   ├── 4-Kelas_XII_Fase_F_Wajib/
    │   └── 5-Kelas_XII_Fase_F_Tingkat_Lanjut/
    └── Materi Bahasa Inggris SMAN 1 Bae Kudus.jpg # Infografis pemetaan materi
```

---

## 🎯 Komponen Perangkat Pembelajaran

Setiap modul pembelajaran disusun dengan prinsip **Constructive Alignment** dan pendekatan **Genre-Based Approach (GBA)**:

1. **Materi Otentik (`Materi_Otentik_TP_*.docx`)**:
   - Teks otentik dari portal kredibel (*BBC, Reuters, World Bank, National Geographic, dsb.*).
   - Kalibrasi level kosakata CEFR yang sesuai fase perkembangan siswa.
   - Dilengkapi tabel atribusi sumber (portal, judul artikel, URL, tahun publikasi) dan glosarium kosakata.

2. **Modul Ajar (`Modul_Ajar_TP_*.docx`)**:
   - Format RPP Kurikulum Merdeka lengkap dengan 4 tahap GBA: *Building Knowledge of the Field (BKoF)*, *Modeling of the Text (MoT)*, *Joint Construction of the Text (JCoT)*, dan *Independent Construction of the Text (ICoT)*.
   - Integrasi 6 Dimensi Profil Pelajar Pancasila, Pemahaman Bermakna, Pertanyaan Pemantik, serta Strategi Pembelajaran Berdiferensiasi.

3. **Lembar Kerja Peserta Didik / LKPD (`LKPD_TP_*.docx`)**:
   - Lembar kerja siap cetak (*student-ready workbook*) dengan aktivitas terstruktur (*Activity 1 s.d. 4*).
   - Dilengkapi *graphic organizers*, panduan penulisan/debat, serta checklist penilaian mandiri/antar-teman.

4. **Instrumen Asesmen (`Instrumen_Asesmen_TP_*.docx`)**:
   - **Diagnostik**: Non-kognitif & kognitif.
   - **Formatif**: Lembar observasi diskusi/proses dan rubrik penilaian *peer-feedback*.
   - **Sumatif**: Matriks keselarasan (*Constructive Alignment*), 10 butir soal pilihan ganda HOTS (kunci + pembahasan analitis), 5 soal esai/studi kasus, dan tugas unjuk kerja autentik dengan rubrik analitik 4 skala (skor 1-4).

---

## 💻 Standar Teknis & Pembuatan

Dokumen-dokumen ini digenerate secara programatik menggunakan Python (`python-docx`, `openpyxl`) dengan tipografi modern (**Segoe UI**), palet warna profesional (*Deep Navy, Steel Blue, Cool Grey*), dan format tabel rapi berstandar administrasi sekolah modern.
