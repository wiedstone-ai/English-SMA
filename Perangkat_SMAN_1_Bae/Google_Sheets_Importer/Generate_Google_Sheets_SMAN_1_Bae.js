/**
 * =========================================================================================
 * GOOGLE APPS SCRIPT: AUTO-GENERATOR SPREADSHEET ASESMEN SMAN 1 BAE KUDUS
 * =========================================================================================
 * Script ini dapat dijalankan langsung di Google Apps Script (https://script.google.com/)
 * untuk secara otomatis membuat file Google Sheets baru di Google Drive Anda lengkap dengan
 * lembar STS (2 Topik), SAS (4 Topik), dan Rekap Nilai Siswa.
 *
 * CARA MENGGUNAKAN:
 * 1. Buka browser dan kunjungi https://script.google.com/
 * 2. Klik tombol "New project" (+ Proyek Baru).
 * 3. Hapus kode default, lalu salin dan tempel (copy-paste) seluruh isi skrip ini.
 * 4. Pada menu fungsi di bagian atas, pilih fungsi "createAllAssessmentsSMAN1Bae".
 * 5. Klik tombol "Run" (Jalankan) dan izinkan akses Google Drive saat diminta.
 * 6. Tiga file Google Sheets (Kelas 10, Kelas 11, dan Kelas 12) akan otomatis dibuat di Google Drive Anda!
 * =========================================================================================
 */

function createAllAssessmentsSMAN1Bae() {
  var folderName = "Perangkat Asesmen SMAN 1 Bae Kudus (2025-2026)";
  var folders = DriveApp.getFoldersByName(folderName);
  var targetFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  
  Logger.log("Membuat Spreadsheet di Folder: " + targetFolder.getName());
  
  createGradeSheet(targetFolder, "10", "Kelas X (Fase E)", "Descriptive Text", "Report Text", 
                   "Descriptive Text, Report Text, Simple Present & Passive Voice, Adjective Phrases & Intentions");
                   
  createGradeSheet(targetFolder, "11", "Kelas XI (Fase F)", "Narrative Text", "Analytical Exposition", 
                   "Narrative Text, Analytical Exposition, Opinions & Agreement, Conditional 1 & 2");
                   
  createGradeSheet(targetFolder, "12", "Kelas XII (Fase F)", "Advanced Narrative", "Discussion Text", 
                   "Advanced Narrative, Discussion Text, Complex Argumentation, Conditional 1, 2, & 3");
                   
  Logger.log("Semua spreadsheet berhasil dibuat di Google Drive Anda!");
}

function createGradeSheet(folder, gradeNum, gradeLabel, stsTop1, stsTop2, sasTopSummary) {
  var fileName = "Latihan Asesmen Bahasa Inggris Kelas " + gradeNum + " - SMAN 1 Bae";
  var ss = SpreadsheetApp.create(fileName);
  
  // Pindahkan ke folder tujuan
  var file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);
  
  // Sheet 1: STS
  var sheetSTS = ss.getActiveSheet();
  sheetSTS.setName("STS_2_Topik");
  setupHeader(sheetSTS, "LATIHAN ASESMEN TENGAH SEMESTER (STS) - 2 TOPIK", gradeLabel, "Gasal", stsTop1 + " & " + stsTop2);
  
  // Sheet 2: SAS
  var sheetSAS = ss.insertSheet("SAS_4_Topik");
  setupHeader(sheetSAS, "LATIHAN ASESMEN SEMESTER (SAS) - 4 TOPIK ESENSIAL", gradeLabel, "Gasal", sasTopSummary);
  
  // Sheet 3: Rekap Nilai
  var sheetRekap = ss.insertSheet("Rekap_Nilai_Siswa");
  setupGradebook(sheetRekap, gradeLabel);
  
  Logger.log("Selesai membuat file: " + fileName + " (URL: " + ss.getUrl() + ")");
}

function setupHeader(sheet, mainTitle, grade, semester, topicScope) {
  sheet.getRange("A1:K1").merge().setValue("SMA NEGERI 1 BAE KUDUS")
       .setFontSize(14).setFontWeight("bold").setFontColor("#1B365D");
  sheet.getRange("A2:K2").merge().setValue(mainTitle)
       .setFontSize(11).setFontWeight("bold").setFontColor("#334E68");
  sheet.getRange("A3:K3").merge().setValue("Mata Pelajaran: Bahasa Inggris | " + grade + " | Semester " + semester + " | Cakupan: " + topicScope)
       .setFontSize(9).setFontStyle("italic").setFontColor("#64748B");
       
  sheet.setFrozenRows(5);
}

function setupGradebook(sheet, gradeLabel) {
  sheet.getRange("A1:H1").merge().setValue("SMA NEGERI 1 BAE KUDUS").setFontSize(14).setFontWeight("bold").setFontColor("#1B365D");
  sheet.getRange("A2:H2").merge().setValue("REKAPITULASI NILAI LATIHAN ASESMEN - " + gradeLabel).setFontSize(11).setFontWeight("bold");
  
  var headers = ["No", "NIS", "Nama Lengkap Siswa", "Nilai STS (40%)", "Nilai SAS (60%)", "Nilai Akhir", "Status KKTP", "Tindak Lanjut"];
  sheet.getRange(5, 1, 1, headers.length).setValues([headers])
       .setBackground("#1B365D").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
       
  for (var i = 1; i <= 36; i++) {
    var row = 5 + i;
    sheet.getRange(row, 1).setValue(i);
    sheet.getRange(row, 2).setValue("2025" + i);
    sheet.getRange(row, 3).setValue("Siswa " + i);
    sheet.getRange(row, 4).setValue(75);
    sheet.getRange(row, 5).setValue(80);
    sheet.getRange(row, 6).setFormula("=ROUND(D" + row + "*0.4 + E" + row + "*0.6, 1)");
    sheet.getRange(row, 7).setFormula('=IF(F' + row + '>=75, "Tuntas (Mencapai KKTP)", "Belum Tuntas")');
    sheet.getRange(row, 8).setFormula('=IF(F' + row + '>=85, "Pengayaan Mandiri", IF(F' + row + '>=75, "Latihan Lanjutan", "Remidial"))');
  }
}
