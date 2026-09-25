/**
 * =========================================================================================
 * GOOGLE APPS SCRIPT: INTEGRASI SINKRONISASI METALVERSEFERY KE GOOGLE SPREADSHEET
 * =========================================================================================
 * 
 * Modul ini menerima data otomatis dari aplikasi web MetalVerseFery (Nama, No Absen, Kelas,
 * Asal Sekolah, Hasil Simulasi Siswa, dan Pengerjaan Tugas) lalu mencatatnya secara terstruktur
 * ke dalam 2 Sheet:
 *   1. "Rekap_Nilai_Siswa" -> Rekapitulasi nilai utama per No. Absen siswa (otomatis update baris)
 *   2. "Log_Aktivitas_Detail" -> Catatan riwayat setiap kali siswa menyelesaikan tugas/simulasi
 * 
 * -----------------------------------------------------------------------------------------
 * PANDUAN PENERAPAN CEPAT (HANYA BUTUH 1-2 MENIT):
 * -----------------------------------------------------------------------------------------
 * 1. Buka Google Drive Anda, buat Google Spreadsheet baru (beri judul: "Nilai MetalVerseFery").
 * 2. Di menu atas Spreadsheet, klik: "Ekstensi" (Extensions) > "Apps Script".
 * 3. Hapus semua teks di editor Apps Script, lalu COPY & PASTE seluruh kode dalam file ini.
 * 4. Klik tombol "Simpan" (ikon disket) di atas.
 * 5. Klik tombol biru "Terapkan" (Deploy) di kanan atas > pilih "Kelola Penerapan Baru" (New Deployment).
 * 6. Klik ikon gerigi (Settings) di sebelah "Pilih jenis" > pilih "Aplikasi Web" (Web App).
 * 7. Atur form deployment:
 *    - Deskripsi: MetalVerseFery Webhook
 *    - Jalankan sebagai: "Saya" (akun Google Anda)
 *    - Siapa yang memiliki akses: "Siapa saja" (Anyone) -> PENTING agar siswa bisa kirim data!
 * 8. Klik "Terapkan" (Deploy). Jika diminta izin akses Google ("Authorize Access"), izinkan.
 * 9. Salin "URL Aplikasi Web" (URL berakhiran /exec).
 * 10. Buka MetalVerseFery, klik tombol "⚙️ Google Sheets" di header/asesmen, lalu paste URL tersebut. Selesai!
 * =========================================================================================
 */

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'Spreadsheet aktif tidak ditemukan' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var data;
    if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        data = e.parameter;
      }
    } else if (e && e.parameter) {
      data = e.parameter;
    } else {
      data = {
        timestamp: Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss"),
        noAbsen: "00",
        nama: "Uji Coba Koneksi",
        kelas: "XI TPFL",
        asalSekolah: "SMKN 2 Yogyakarta",
        kategori: "Koneksi Tes",
        rincianTugas: "Pengujian webhook berhasil",
        skor: 100,
        xp: 50,
        status: "Berhasil"
      };
    }

    // Pastikan ada nilai default jika field kosong
    var timestamp = data.timestamp || Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    var noAbsen = data.noAbsen ? String(data.noAbsen).trim() : "-";
    var nama = data.nama ? String(data.nama).trim() : "Siswa";
    var kelas = data.kelas ? String(data.kelas).trim() : "-";
    var asalSekolah = data.asalSekolah ? String(data.asalSekolah).trim() : "-";
    var kategori = data.kategori || data.jenisAktivitas || "Simulasi";
    var rincianTugas = data.rincianTugas || data.task || "Pengerjaan Tugas";
    var skor = (data.skor !== undefined && data.skor !== null) ? data.skor : "-";
    var xp = data.xp || 0;
    var status = data.status || "Selesai";

    // -------------------------------------------------------------
    // SHEET 1: Log_Aktivitas_Detail (Catatan Riwayat Kronologis)
    // -------------------------------------------------------------
    var logSheet = sheet.getSheetByName("Log_Aktivitas_Detail");
    if (!logSheet) {
      logSheet = sheet.insertSheet("Log_Aktivitas_Detail");
      var logHeaders = [
        "Timestamp", "No Absen", "Nama Siswa", "Kelas", "Asal Sekolah", 
        "Kategori Aktivitas", "Rincian Tugas & Hasil", "Skor / Nilai", "XP", "Status"
      ];
      logSheet.appendRow(logHeaders);
      var headRange = logSheet.getRange(1, 1, 1, logHeaders.length);
      headRange.setBackground("#EA734F").setFontColor("#FFFFFF").setFontWeight("bold");
      logSheet.setFrozenRows(1);
    }
    logSheet.appendRow([timestamp, noAbsen, nama, kelas, asalSekolah, kategori, rincianTugas, skor, xp, status]);

    // -------------------------------------------------------------
    // SHEET 2: Rekap_Nilai_Siswa (Tabel Rekap Berdasarkan No Absen)
    // -------------------------------------------------------------
    var rekapSheet = sheet.getSheetByName("Rekap_Nilai_Siswa");
    if (!rekapSheet) {
      rekapSheet = sheet.insertSheet("Rekap_Nilai_Siswa");
      var rekapHeaders = [
        "No Absen", "Nama Siswa", "Kelas", "Asal Sekolah", 
        "Skor Guillotine (Potong)", "Skor SMAW Arc (Las)", "Skor QC Cacat (AWS)", 
        "Kuis Interaktif", "Job Sheet / Tugas", "Total XP", "Status Terakhir", "Waktu Update"
      ];
      rekapSheet.appendRow(rekapHeaders);
      var rHeadRange = rekapSheet.getRange(1, 1, 1, rekapHeaders.length);
      rHeadRange.setBackground("#568A7D").setFontColor("#FFFFFF").setFontWeight("bold");
      rekapSheet.setFrozenRows(1);
    }

    // Cari apakah No Absen siswa sudah ada di tabel rekap
    var lastRow = rekapSheet.getLastRow();
    var rowIndex = -1;
    if (lastRow > 1) {
      var absenRange = rekapSheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < absenRange.length; i++) {
        if (String(absenRange[i][0]).trim() === noAbsen && noAbsen !== "-") {
          rowIndex = i + 2; // baris di spreadsheet (1-indexed)
          break;
        }
      }
    }

    // Tentukan kolom mana yang diupdate berdasarkan kategori aktivitas
    var katLower = (kategori + " " + rincianTugas).toLowerCase();
    var colTarget = 9; // default: Job Sheet / Tugas (kolom 9)
    if (katLower.indexOf("guillotine") !== -1 || katLower.indexOf("potong") !== -1 || katLower.indexOf("shear") !== -1) {
      colTarget = 5; // Skor Guillotine
    } else if (katLower.indexOf("smaw") !== -1 || katLower.indexOf("arc") !== -1 || katLower.indexOf("las") !== -1) {
      colTarget = 6; // Skor SMAW Arc
    } else if (katLower.indexOf("cacat") !== -1 || katLower.indexOf("qc") !== -1 || katLower.indexOf("defect") !== -1) {
      colTarget = 7; // Skor QC Cacat
    } else if (katLower.indexOf("kuis") !== -1 || katLower.indexOf("quiz") !== -1 || katLower.indexOf("bab") !== -1) {
      colTarget = 8; // Kuis Interaktif
    }

    if (rowIndex !== -1) {
      // Baris siswa ditemukan -> Update baris yang sudah ada
      rekapSheet.getRange(rowIndex, 2).setValue(nama);
      rekapSheet.getRange(rowIndex, 3).setValue(kelas);
      rekapSheet.getRange(rowIndex, 4).setValue(asalSekolah);
      rekapSheet.getRange(rowIndex, colTarget).setValue(skor);
      
      // Akumulasi atau set XP
      var currentXp = Number(rekapSheet.getRange(rowIndex, 10).getValue()) || 0;
      rekapSheet.getRange(rowIndex, 10).setValue(currentXp + Number(xp || 0));
      rekapSheet.getRange(rowIndex, 11).setValue(status);
      rekapSheet.getRange(rowIndex, 12).setValue(timestamp);
    } else {
      // Siswa baru -> Tambah baris baru
      var newRow = [
        noAbsen, nama, kelas, asalSekolah,
        colTarget === 5 ? skor : "-",
        colTarget === 6 ? skor : "-",
        colTarget === 7 ? skor : "-",
        colTarget === 8 ? skor : "-",
        colTarget === 9 ? skor : "-",
        xp,
        status,
        timestamp
      ];
      rekapSheet.appendRow(newRow);
    }

    // Urutkan tabel rekap berdasarkan No Absen (numerik ascending)
    var updatedLastRow = rekapSheet.getLastRow();
    if (updatedLastRow > 2) {
      var sortRange = rekapSheet.getRange(2, 1, updatedLastRow - 1, 12);
      sortRange.sort({ column: 1, ascending: true });
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Data siswa No Absen ' + noAbsen + ' (' + nama + ') berhasil dicatat.',
      timestamp: timestamp
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(
    "✅ WEBHOOK METALVERSEFERY AKTIF!\n\n" +
    "Endpoint ini siap menerima kiriman data nilai simulasi & tugas siswa dari MetalVerseFery.id.\n" +
    "Gunakan metode POST dari aplikasi MetalVerseFery untuk menyimpan data otomatis."
  ).setMimeType(ContentService.MimeType.TEXT);
}
