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

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return responseOK();
    }

    var data = JSON.parse(e.postData.contents);
    var updateId = data.update_id;

    // 1. PENCEGAHAN SPAM (Cache 24 Jam)
    if (updateId) {
      var cache = CacheService.getScriptCache();
      if (cache.get("up_" + updateId)) {
        return responseOK(); // Abaikan jika update_id sudah pernah diproses
      }
      cache.put("up_" + updateId, "1", 86400); // Kunci selama 24 jam
    }

    // 2. PROSES PESAN TELEGRAM
    if (data.message && data.message.text) {
      var chatId = data.message.chat.id;
      var textPesan = String(data.message.text).trim();

      if (textPesan.indexOf("/start") === 0) {
        var parts = textPesan.split(/\s+/);
        
        // Memastikan ada parameter deep-link (misal: /start LAB096)
        if (parts.length > 1 && parts[1].trim() !== "" && parts[1].trim().toUpperCase() !== "START") {
          var kodeSiswa = parts[1].trim();
          
          // Lakukan pencarian & penyimpanan
          var hasil = simpanChatIdOtomatis(kodeSiswa, chatId);
          
          if (hasil === "BERHASIL") {
            kirimPesanTelegram(chatId, "✅ *Pendaftaran Berhasil!*\n\nChat ID Anda (" + chatId + ") telah terhubung dengan ID Siswa: *" + kodeSiswa + "*");
          } else if (hasil === "SUDAH_TERDAFTAR") {
            // Sudah terdaftar sebelumnya, tidak perlu kirim balasan ulang
          } else {
            kirimPesanTelegram(chatId, "❌ *Pendaftaran Gagal*\n\nKode Siswa *" + kodeSiswa + "* tidak ditemukan di database.");
          }
        }
      }
      return responseOK();
    }

    // 3. PROSES SCAN QR CODE (JIKA ADA)
    if (data.kodeQR) {
      prosesAbsensi(String(data.kodeQR).trim());
    }

  } catch (error) {
    Logger.log("Error doPost: " + error.message);
  }

  return responseOK();
}

function responseOK() {
  return ContentService.createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function simpanChatIdOtomatis(kodeSiswa, chatId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetSiswa = ss.getSheetByName("DATA_SISWA");
    if (!sheetSiswa) return "ERROR";

    var data = sheetSiswa.getDataRange().getValues();
    var targetKode = String(kodeSiswa).trim().toUpperCase();
    var strChatId = String(chatId).trim();

    // Loop mulai dari baris kedua (indeks 1) untuk melewati header
    for (var i = 1; i < data.length; i++) {
      var idDiSheet = String(data[i][0]).trim().toUpperCase(); // Kolom A
      var existingChatId = String(data[i][7]).trim();          // Kolom H (Index 7)

      if (idDiSheet === targetKode) {
        // Jika sudah terdaftar dengan Chat ID yang sama
        if (existingChatId === strChatId) {
          return "SUDAH_TERDAFTAR";
        }

        // KODE COCOK: Tulis Chat ID ke Kolom H (Baris i+1, Kolom 8)
        sheetSiswa.getRange(i + 1, 8).setValue(strChatId);
        SpreadsheetApp.flush();
        return "BERHASIL";
      }
    }
  } catch (err) {
    Logger.log("Error simpanChatIdOtomatis: " + err.message);
  }
  
  // HANYA mengembalikan TIDAK_DITEMUKAN dan TANPA MENGISI APAPUN di sheet
  return "TIDAK_DITEMUKAN";
}

function kirimPesanTelegram(chatId, text) {
  var url = "https://api.telegram.org/bot" + BOT_TOKEN + "/sendMessage";
  var payload = {
    chat_id: String(chatId),
    text: text,
    parse_mode: "Markdown"
  };
  
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  UrlFetchApp.fetch(url, options);
}

function prosesAbsensi(kodeQR) {
  // Fungsi penanganan absensi QR Code Anda
}

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
