# PANDUAN PENGGUNAAN LATIHAN ASESMEN DI GOOGLE SHEETS
### SMA NEGERI 1 BAE KUDUS — TAHUN AJARAN 2025/2026

Dokumen ini memuat panduan praktis bagi Guru Bahasa Inggris SMA Negeri 1 Bae Kudus untuk memanfaatkan berkas latihan asesmen (STS 2 Topik & SAS 4 Topik) di ekosistem **Google Workspace for Education / Google Sheets**.

---

## 🚀 CARA 1: UPLOAD LANGSUNG KE GOOGLE DRIVE (SANGAT DIREKOMENDASIKAN)
Format berkas spreadsheet `.xlsx` yang telah kami buat dirancang dengan **100% kompatibilitas Google Sheets**. Seluruh formula (`IF`, `SUM`, `ROUND`, `AVERAGE`, `COUNTIF`), pewarnaan header, wrap-text, dan tata letak tabel akan terbuka sempurna tanpa distorsi.

### Langkah-langkah:
1. Buka browser dan masuk ke akun Google Anda di **[drive.google.com](https://drive.google.com/)**.
2. Buat folder baru (misal: `Asesmen Bahasa Inggris SMAN 1 Bae`).
3. Tarik (*drag and drop*) atau klik tombol **+ Baru > Upload File**, lalu pilih file dari komputer Anda:
   - `Latihan_Asesmen_Bahasa_Inggris_Kelas_10_SMAN_1_Bae.xlsx`
   - `Latihan_Asesmen_Bahasa_Inggris_Kelas_11_SMAN_1_Bae.xlsx`
   - `Latihan_Asesmen_Bahasa_Inggris_Kelas_12_SMAN_1_Bae.xlsx`
4. Klik dua kali (*double click*) pada berkas yang terunggah. Google Drive akan otomatis membukanya langsung di **Google Sheets**.
5. Untuk menyimpan sebagai dokumen native Google Sheets:
   - Klik menu **File > Simpan sebagai Google Spreadsheet**.

---

## 📲 CARA MEMBAGIKAN LATIHAN KE SISWA (GOOGLE CLASSROOM / WA)

### Opsi A: Setiap Siswa Mengerjakan Salinan Sendiri (*Make a Copy*)
Agar setiap siswa memiliki lembar latihan mandiri tanpa mengganggu lembar master:
1. Buka spreadsheet di Google Sheets.
2. Klik tombol **Bagikan (Share)** di pojok kanan atas.
3. Ubah akses umum menjadi: **"Siapa saja yang memiliki link dapat melihat (Viewer)"**.
4. Salin tautan (*copy link*). Tautan asli biasanya berakhiran `/edit?usp=sharing`.
5. Ubah kata `edit?usp=sharing` pada akhir link menjadi `copy`.
   - *Contoh Link Asli:* `https://docs.google.com/spreadsheets/d/1aBcDeF.../edit?usp=sharing`
   - *Ubah Menjadi:* `https://docs.google.com/spreadsheets/d/1aBcDeF.../copy`
6. Saat siswa mengklik link tersebut, Google Sheets akan langsung menampilkan tombol **"Buat Salinan (Make a copy)"**. Siswa dapat langsung mengetikkan jawaban A, B, C, D, atau E pada kolom H, dan skor mereka akan terkoreksi otomatis!

### Opsi B: Penugasan Melalui Google Classroom
1. Buka Google Classroom kelas Anda.
2. Buat **Tugas Baru (Create Assignment)**.
3. Lampirkan link Google Sheets tersebut dari Google Drive Anda.
4. Pada menu dropdown di sebelah lampiran file, pilih opsi: **"Buat salinan untuk setiap siswa (Make a copy for each student)"**.
5. Siswa dapat mengerjakan dan langsung mengklik tombol **Serahkan (Turn In)** setelah selesai. Guru dapat melihat total skor otomatis di sheet rekap!

---

## ⚙️ CARA 2: MENGGUNAKAN GOOGLE APPS SCRIPT (AUTO-GENERATOR)
Jika Anda ingin membuat dokumen baru secara otomatis via skrip:
1. Kunjungi **[script.google.com](https://script.google.com/)**.
2. Klik **Proyek Baru**.
3. Buka file `Generate_Google_Sheets_SMAN_1_Bae.js` yang ada di folder ini, salin seluruh kodenya, dan tempelkan ke editor Apps Script.
4. Pilih fungsi `createAllAssessmentsSMAN1Bae` dan klik **Run (Jalankan)**.
5. Berikan izin saat muncul dialog otorisasi Google Drive.
6. Folder dan dokumen latihan untuk Kelas 10, 11, dan 12 akan langsung tercipta di Google Drive Anda.

---

## 📊 FITUR FORMULA OTOMATIS DI SPREADSHEET
Di dalam spreadsheet telah ditanamkan formula cerdas:
1. **Koreksi Pilihan Ganda:**
   `=IF(ISBLANK(H12), "-", IF(UPPER(TRIM(H12))=F12, "BENAR", "SALAH"))`
   *Memeriksa jawaban siswa (tidak terpengaruh huruf besar/kecil atau spasi).*
2. **Kalkulasi Skor Butir:**
   `=IF(I12="BENAR", 10, 0)` *(pada STS)* atau `=IF(I12="BENAR", 3.5, 0)` *(pada SAS)*.
3. **Total Nilai Akhir Siswa:**
   `=ROUND(D8*0.4 + E8*0.6, 1)` *(Pembobotan 40% STS + 60% SAS)*.
4. **Status Ketercapaian KKTP:**
   `=IF(F8>=75, "Tuntas (Mencapai KKTP)", "Belum Tuntas (Perlu Remidial)")`.
5. **Statistik Kelas Otomatis:**
   Menghitung Rata-rata kelas (`AVERAGE`), Nilai Tertinggi (`MAX`), Nilai Terendah (`MIN`), dan Persentase Ketuntasan (`COUNTIF`).

---
*SMA Negeri 1 Bae Kudus — Mengembangkan Pembelajaran Bahasa Inggris Autentik & Berkualitas.*
