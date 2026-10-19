const SHEET_SISWA = "DATA_SISWA";
const SHEET_JADWAL = "JADWAL";
const SHEET_ABSENSI = "ABSENSI";
const SHEET_PENGATURAN = "PENGATURAN";


/**
 * ================================
 * MENCARI JADWAL YANG AKTIF
 * ================================
 */
function getSesiAktif() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheet = ss.getSheetByName(SHEET_JADWAL);

  if (!sheet) {
    return {
      aktif: false,
      message: "Sheet JADWAL tidak ditemukan."
    };
  }

  const data = sheet.getDataRange().getValues();

  const sekarang = new Date();

  // =====================================
  // NAMA HARI INDONESIA
  // =====================================

  const daftarHari = [
    "Minggu",
    "Senin",
    "Selasa",
    "Rabu",
    "Kamis",
    "Jumat",
    "Sabtu"
  ];

  const hariSekarang =
    daftarHari[sekarang.getDay()];

  // =====================================
  // WAKTU INDONESIA
  // =====================================

  const timezone = "Asia/Jakarta";

  const jamSekarang =
    Utilities.formatDate(
      sekarang,
      timezone,
      "HH:mm"
    );

  Logger.log(
    "Hari sekarang: " +
    hariSekarang
  );

  Logger.log(
    "Jam sekarang: " +
    jamSekarang
  );


  // =====================================
  // MEMBACA DATA JADWAL
  // =====================================

  for (let i = 1; i < data.length; i++) {

    const idJadwal =
      String(data[i][0]).trim();

    const hariJadwal =
      String(data[i][1]).trim();

    const kelas =
      String(data[i][2]).trim();

    const mapel =
      String(data[i][3]).trim();

    const jamMasuk =
      formatJam(data[i][4]);

    const batasHadir =
      formatJam(data[i][5]);

    const batasAlpa =
      formatJam(data[i][6]);

    const aktif =
      String(data[i][7])
        .trim()
        .toUpperCase();


    // =====================================
    // CEK HARI
    // =====================================

    if (
      hariJadwal.toLowerCase()
      !==
      hariSekarang.toLowerCase()
    ) {
      continue;
    }


    // =====================================
    // CEK STATUS AKTIF
    // =====================================

    if (aktif !== "YA") {
      continue;
    }


    // =====================================
    // CEK JAM
    // =====================================

    if (
      jamSekarang >= jamMasuk &&
      jamSekarang <= batasAlpa
    ) {

      Logger.log(
        "SESI AKTIF: " +
        kelas +
        " | " +
        mapel
      );

      return {

        aktif: true,

        idJadwal:
          idJadwal,

        kelas:
          kelas,

        mapel:
          mapel,

        jamMasuk:
          jamMasuk,

        batasHadir:
          batasHadir,

        batasAlpa:
          batasAlpa,

        jamSekarang:
          jamSekarang

      };

    }

  }


  // =====================================
  // TIDAK ADA SESI
  // =====================================

  Logger.log(
    "Tidak ada sesi aktif."
  );

  return {

    aktif: false,

    message:
      "Tidak ada jadwal absensi aktif saat ini.",

    hari:
      hariSekarang,

    jamSekarang:
      jamSekarang

  };

}


/**
 * ==========================================
 * PROSES ABSENSI DARI QR CODE (OTOMATIS)
 * ==========================================
 */
function prosesAbsensi(kodeQR) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const siswaSheet = ss.getSheetByName(SHEET_SISWA);
  const absensiSheet = ss.getSheetByName(SHEET_ABSENSI);

  if (!siswaSheet) {
    return { success: false, message: "Sheet DATA_SISWA tidak ditemukan." };
  }

  if (!absensiSheet) {
    return { success: false, message: "Sheet ABSENSI tidak ditemukan." };
  }

  kodeQR = String(kodeQR).trim();

  if (!kodeQR) {
    return { success: false, message: "QR Code kosong." };
  }

  // ======================================
  // CARI DATA SISWA
  // ======================================
  const dataSiswa = siswaSheet.getDataRange().getValues();
  let siswa = null;

  for (let i = 1; i < dataSiswa.length; i++) {
    // DATA_SISWA:
    // A = ID, B = NAMA, C = NIPD, D = JK, E = NISN, F = KELAS, G = QR CODE, H = CHAT ID TELEGRAM
    if (String(dataSiswa[i][0]).trim() === kodeQR) {
      siswa = {
        id: String(dataSiswa[i][0]).trim(),
        nama: String(dataSiswa[i][1]).trim(),
        nipd: dataSiswa[i][2],
        jk: dataSiswa[i][3],
        nisn: dataSiswa[i][4],
        kelas: String(dataSiswa[i][5]).trim(),
        chatId: String(dataSiswa[i][7]).trim() // Kolom H untuk Chat ID Telegram Siswa/Orang Tua
      };
      break;
    }
  }

  if (!siswa) {
    return { success: false, message: "QR Code tidak terdaftar di DATA_SISWA." };
  }

  // ======================================
  // CARI SESI AKTIF & CEK KELAS
  // ======================================
  const sesi = getSesiAktif();

  if (!sesi.aktif) {
    return { success: false, message: sesi.message };
  }

  if (siswa.kelas.toUpperCase() !== sesi.kelas.toUpperCase()) {
    return {
      success: false,
      message: "Siswa " + siswa.nama + " berasal dari kelas " + siswa.kelas + ", bukan kelas " + sesi.kelas + "."
    };
  }

  // ======================================
  // WAKTU & CEK ABSENSI GANDA
  // ======================================
  const sekarang = new Date();
  const timezone = Session.getScriptTimeZone();
  const tanggal = Utilities.formatDate(sekarang, timezone, "dd/MM/yyyy");
  const jam = Utilities.formatDate(sekarang, timezone, "HH:mm:ss");

  const dataAbsensi = absensiSheet.getDataRange().getValues();

  for (let i = 1; i < dataAbsensi.length; i++) {
    const idData = String(dataAbsensi[i][2]).trim();
    const mapelData = String(dataAbsensi[i][3]).trim();
    const tanggalData = String(dataAbsensi[i][7]).trim();

    if (idData === siswa.id && mapelData === sesi.mapel && tanggalData === tanggal) {
      return {
        success: false,
        duplicate: true,
        nama: siswa.nama,
        kelas: siswa.kelas,
        jam: dataAbsensi[i][8],
        status: dataAbsensi[i][4],
        message: "Siswa sudah melakukan absensi."
      };
    }
  }

  // ======================================
  // MENENTUKAN STATUS & SIMPAN
  // ======================================
  let status = "HADIR";
  const jamSekarang = jam.substring(0, 5);

  if (jamSekarang > sesi.batasHadir) {
    status = "TERLAMBAT";
  }

  absensiSheet.appendRow([
    siswa.nama,       // A = NAMA
    siswa.kelas,      // B = KELAS
    siswa.id,         // C = ID
    sesi.mapel,       // D = MAPEL
    status,           // E = STATUS
    sekarang,         // F = TIMESTAMP
    tanggal,          // G = TANGGAL
    jam               // H = JAM
  ]);

  // ======================================
  // KIRIM NOTIFIKASI OTOMATIS KE TELEGRAM
  // ======================================
  var statusPesan = (status === "HADIR") ? "HADIR (Tepat Waktu)" : "TERLAMBAT";
  
  var pesanNotifikasi = "📢 PEMBERITAHUAN ABSENSI LAB\n\n" +
                        "Yth. Orang Tua/Wali dari:\n" +
                        "👤 *" + siswa.nama + "* (" + siswa.kelas + ")\n\n" +
                        "Telah tercatat melakukan absensi:\n" +
                        "Status: *" + statusPesan + "*\n" +
                        "Mata Pelajaran: " + sesi.mapel + "\n" +
                        "Tanggal: " + tanggal + "\n" +
                        "Pukul: " + jam + " WIB\n\n" +
                        "_Pesan otomatis dari Sistem Lab Komputer._";

  // 1. Kirim ke Orang Tua / Siswa (jika Kolom H terisi Chat ID)
  if (siswa.chatId && siswa.chatId !== "" && siswa.chatId !== "undefined") {
    kirimPesanTelegram(siswa.chatId, pesanNotifikasi);
  }

  // 2. Kirim ke Wali Kelas 
  // (Masukkan Chat ID Telegram Wali Kelas di bawah ini sesuai kelasnya)
  var chatIdWaliKelas = "";
  
  if (siswa.kelas === "X APHP 1") {
    chatIdWaliKelas = "MASUKKAN_CHAT_ID_WALI_KELAS_DISINI"; // Ganti dengan Chat ID Wali Kelas
  } 
  // Anda bisa tambahkan 'else if' untuk kelas lainnya jika diperlukan

  if (chatIdWaliKelas !== "") {
    var pesanWaliKelas = "📊 LAPORAN MASUK ABSENSI KELAS\n\n" + pesanNotifikasi;
    kirimPesanTelegram(chatIdWaliKelas, pesanWaliKelas);
  }

  // ======================================
  // HASIL
  // ======================================
  return {
    success: true,
    nama: siswa.nama,
    kelas: siswa.kelas,
    mapel: sesi.mapel,
    jam: jam,
    status: status,
    message: "Absensi berhasil disimpan."
  };
}

/**
 * ==========================================
 * MENGUBAH JAM MENJADI FORMAT HH:mm
 * ==========================================
 */
function formatJam(value) {

  if (value instanceof Date) {

    return Utilities.formatDate(

      value,

      Session.getScriptTimeZone(),

      "HH:mm"

    );

  }


  return String(value)
    .trim()
    .substring(0, 5);

}
// GANTI DENGAN TOKEN BOT TELEGRAM ANDA
var BOT_TOKEN = "8543899140:AAF_OFSBRIqZrprclYDyC7AHzItvU8adQdo"; 

var TELEGRAM_TOKEN = "8543899140:AAF_OFSBRIqZrprclYDyC7AHzItvU8adQdo";



/**
 * ==========================================
 * REKAP BULANAN PER KELAS & STATUS (HADIR, TERLAMBAT, ALPA)
 * ==========================================
 */
function buatRekapBulanan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetAbsensi = ss.getSheetByName(SHEET_ABSENSI);
  const sheetSiswa = ss.getSheetByName(SHEET_SISWA);
  
  let sheetRekap = ss.getSheetByName("REKAP");
  if (!sheetRekap) {
    sheetRekap = ss.insertSheet("REKAP");
  } else {
    sheetRekap.clear(); // Bersihkan data lama agar ter-update bersih
  }

  if (!sheetAbsensi || !sheetSiswa) {
    Logger.log("Sheet ABSENSI atau DATA_SISWA tidak ditemukan.");
    return;
  }

  const dataSiswa = sheetSiswa.getDataRange().getValues();
  const dataAbsensi = sheetAbsensi.getDataRange().getValues();

  // Mapping data siswa: Key = ID Siswa, Value = {nama, kelas}
  let petaSiswa = {};
  for (let i = 1; i < dataSiswa.length; i++) {
    let id = String(dataSiswa[i][0]).trim();     // Kolom A: ID
    let nama = String(dataSiswa[i][1]).trim();   // Kolom B: NAMA
    let kelas = String(dataSiswa[i][5]).trim();  // Kolom F: KELAS
    if (id) {
      petaSiswa[id] = { nama: nama, kelas: kelas };
    }
  }

  // Struktur penampung rekap: rekapData[kelas][idSiswa] = { nama, hadir, terlambat, alpa }
  let rekapData = {};

  // Proses data absensi (Lewati header baris 0)
  for (let i = 1; i < dataAbsensi.length; i++) {
    let namaAbsen = String(dataAbsensi[i][0]).trim();
    let kelasAbsen = String(dataAbsensi[i][1]).trim();
    let idAbsen = String(dataAbsensi[i][2]).trim();
    let status = String(dataAbsensi[i][4]).trim().toUpperCase(); // HADIR / TERLAMBAT

    if (!idAbsen) continue;

    // Jika kelas di absensi kosong, ambil dari data siswa
    if (!kelasAbsen && petaSiswa[idAbsen]) {
      kelasAbsen = petaSiswa[idAbsen].kelas;
    }
    if (!kelasAbsen) kelasAbsen = "TANPA KELAS";

    if (!rekapData[kelasAbsen]) {
      rekapData[kelasAbsen] = {};
    }

    if (!rekapData[kelasAbsen][idAbsen]) {
      rekapData[kelasAbsen][idAbsen] = {
        nama: petaSiswa[idAbsen] ? petaSiswa[idAbsen].nama : namaAbsen,
        hadir: 0,
        terlambat: 0,
        alpa: 0
      };
    }

    // Akumulasi status
    if (status === "HADIR") {
      rekapData[kelasAbsen][idAbsen].hadir += 1;
    } else if (status === "TERLAMBAT") {
      rekapData[kelasAbsen][idAbsen].terlambat += 1;
    }
  }

  // Tulis ke Sheet REKAP secara rapi per kelas
  let output = [["REKAPITULASI ABSENSI BULANAN SISWA"]];
  output.push([""]); // Diperbaiki: sudah ditutup dengan tanda kurung yang benar

  for (let kelas in rekapData) {
    output.push(["KELAS: " + kelas]);
    output.push(["No", "ID Siswa", "Nama Siswa", "Hadir", "Terlambat", "Alpa", "Total Masuk"]);
    
    let no = 1;
    for (let id in rekapData[kelas]) {
      let siswa = rekapData[kelas][id];
      let totalMasuk = siswa.hadir + siswa.terlambat;
      output.push([
        no++,
        id,
        siswa.nama,
        siswa.hadir, // Diperbaiki dari typo sebelumnya
        siswa.terlambat,
        siswa.alpa,
        totalMasuk
      ]);
    }
    output.push([""]); // Spasi antar kelas
  }

  sheetRekap.getRange(1, 1, output.length, output[0].length).setValues(output);
  Logger.log("Rekap bulanan berhasil dibuat per kelas!");
}
/**
 * ==========================================
 * FUNGSI KIRIM NOTIFIKASI KE TELEGRAM BOT
 * ==========================================
 */
function kirimPesanTelegram(chatId, pesan) {
  // Ganti token di bawah dengan Token yang Anda dapatkan dari @BotFather
  const tokenBot = "8543899140:AAF_OFSBRIqZrprclYDyC7AHzItvU8adQdo"; 
  const urlAPI = `https://api.telegram.org/bot${tokenBot}/sendMessage`;

  const payload = {
    "chat_id": chatId,
    "text": pesan,
    "parse_mode": "Markdown"
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  try {
    const response = UrlFetchApp.fetch(urlAPI, options);
    Logger.log("Respon Telegram: " + response.getContentText());
  } catch (error) {
    Logger.log("Gagal mengirim Telegram: " + error.message);
  }
}
var TELEGRAM_TOKEN = "8543899140:AAF_OFSBRIqZrprclYDyC7AHzItvU8adQdo";

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSiswa = ss.getSheetByName("DATA_SISWA");
  
  if (e && e.parameter && e.parameter.id) {
    var rawData = e.parameter.id.toString().trim();
    var idSiswa = rawData.split("|")[0].trim().toUpperCase(); // Ambil ID utamanya saja (misal LAB041)
    
    var data = sheetSiswa.getDataRange().getValues();
    var result = { status: "not_found" };
    
    for (var i = 1; i < data.length; i++) {
      var idDb = data[i][0] ? data[i][0].toString().trim().toUpperCase() : "";
      if (idDb === idSiswa) {
        var namaSiswa = data[i][1] || "Siswa";   // Kolom B: Nama
        var kelasSiswa = data[i][2] || "-";       // Kolom C: Kelas
        var chatId = data[i][7] || "";           // Kolom H: Telegram Chat ID
        
        // Tentukan Waktu & Status Kehadiran (Batas jam 08:00)
        var now = new Date();
        var jam = now.getHours();
        var menit = now.getMinutes();
        var statusAbsen = "HADIR";
        
        if (jam > 8 || (jam === 8 && menit > 0)) {
          statusAbsen = "TERLAMBAT";
        }
        
        var tanggal = Utilities.formatDate(now, "Asia/Jakarta", "dd-MM-yyyy");
        var waktuStr = Utilities.formatDate(now, "Asia/Jakarta", "HH:mm:ss");

        // 1. Simpan Rekap ke Google Sheets (Sheet: REKAP_ABSEN)
        try {
          var sheetAbsen = ss.getSheetByName("REKAP_ABSEN");
          if (!sheetAbsen) {
            sheetAbsen = ss.insertSheet("REKAP_ABSEN");
            sheetAbsen.appendRow(["Tanggal", "Waktu", "ID Siswa", "Nama", "Kelas", "Status"]);
          }
          sheetAbsen.appendRow([tanggal, waktuStr, idSiswa, namaSiswa, kelasSiswa, statusAbsen]);
        } catch (err) {
          console.error("Gagal simpan sheet: " + err.toString());
        }

        // 2. Kirim Notifikasi ke Telegram
        if (chatId) {
          var pesan = "📢 *PEMBERITAHUAN ABSENSI LAB*\n\n" +
                      "Yth. Orang Tua/Wali dari:\n" +
                      "👤 *" + namaSiswa + "* (" + kelasSiswa + ")\n\n" +
                      "Status Kehadiran: *" + statusAbsen + "*\n" +
                      "🕒 Waktu: " + waktuStr + " WIB\n" +
                      "📍 Lokasi: Lab Komputer";
          
          sendTelegramMessage(chatId, pesan);
        }

        // 3. Kirim balik data ke Web Scanner agar nama siswa muncul di layar
        result = {
          status: "success",
          nama: namaSiswa,
          kelas: kelasSiswa,
          absenStatus: statusAbsen
        };
        break;
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput("Bot Active");
}

function sendTelegramMessage(chatId, text) {
  var url = "https://api.telegram.org/bot" + TELEGRAM_TOKEN + "/sendMessage";
  var payload = {
    method: "post",
    payload: {
      chat_id: chatId,
      text: text,
      parse_mode: "Markdown"
    }
  };
  try {
    UrlFetchApp.fetch(url, payload);
  } catch (e) {
    console.error("Gagal kirim Telegram: " + e.toString());
  }
}
