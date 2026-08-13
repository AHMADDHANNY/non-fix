/**
 * ===================================================================
 * CODE.GS
 * File inti: konfigurasi, inisialisasi sheet & folder Drive,
 * entry point web app (doGet), dan helper umum yang dipakai
 * semua file .gs lain (Auth.gs, Absensi.gs, Izin.gs, Admin.gs).
 * ===================================================================
 */

/* ===================================================================
 * KONFIGURASI DASAR
 * =================================================================== */

const SHEET_USERS      = 'Users';
const SHEET_ABSENSI    = 'Absensi';
const SHEET_IZIN       = 'Izin';
const SHEET_PENGATURAN = 'Pengaturan';
const SHEET_SESSIONS   = 'Sessions';

const FOLDER_ABSENSI_NAME = 'Foto_Absensi';
const FOLDER_IZIN_NAME    = 'Lampiran_Izin';
const FOLDER_REKAP_NAME   = 'Rekap_PDF';

const CACHE_KEY_PENGATURAN = 'pengaturan_app';
const CACHE_DURATION_DETIK = 21600; // 6 jam (batas maksimal CacheService)

const SESSION_DURASI_JAM = 10; // token session berlaku 10 jam

/* ===================================================================
 * ENTRY POINT WEB APP
 * =================================================================== */

/**
 * Dipanggil otomatis saat web app diakses lewat browser.
 * Menyajikan halaman login sebagai halaman utama.
 * Nanti halaman lain (absen, admin) dimuat via include() atau
 * routing sisi client (?page=admin dsb) — sesuaikan dengan
 * struktur frontend yang akan kita buat.
 */
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'Login';

  let template;
  try {
    template = HtmlService.createTemplateFromFile(page);
  } catch (err) {
    // Fallback ke Login jika halaman tidak ditemukan
    template = HtmlService.createTemplateFromFile('Login');
  }

  return template.evaluate()
    .setTitle('Aplikasi Absensi Karyawan')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper untuk menyisipkan file HTML lain ke dalam template
 * (dipakai di file .html dengan <?!= include('NamaFile') ?>).
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ===================================================================
 * SETUP OTOMATIS (dijalankan sekali manual dari editor Apps Script,
 * atau otomatis saat pertama kali dibutuhkan)
 * =================================================================== */

/**
 * Jalankan fungsi ini SEKALI secara manual dari editor Apps Script
 * (pilih fungsi "setupAwal" lalu klik Run) untuk:
 * 1. Membuat semua sheet yang dibutuhkan beserta header-nya
 *    (jika belum ada / tidak menimpa yang sudah ada)
 * 2. Membuat folder Drive: Foto_Absensi, Lampiran_Izin, Rekap_PDF
 * 3. Mengisi baris default di sheet Pengaturan (jika masih kosong)
 * 4. Membuat 1 akun admin default agar bisa login pertama kali
 */
function setupAwal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  _pastikanSheet_(ss, SHEET_USERS, [
    'id_user', 'nama', 'username', 'password_hash', 'salt',
    'role', 'jabatan', 'foto_profil', 'status_aktif', 'created_at', 'updated_at'
  ]);

  _pastikanSheet_(ss, SHEET_ABSENSI, [
    'id_absen', 'id_user', 'tanggal', 'jam_masuk', 'foto_masuk',
    'lat_masuk', 'long_masuk', 'jam_pulang', 'foto_pulang',
    'lat_pulang', 'long_pulang', 'status', 'jarak_masuk_meter', 'jarak_pulang_meter'
  ]);

  _pastikanSheet_(ss, SHEET_IZIN, [
    'id_izin', 'id_user', 'tanggal_pengajuan', 'jenis_izin',
    'tanggal_mulai', 'tanggal_selesai', 'keterangan', 'lampiran',
    'status', 'catatan_admin', 'diproses_oleh', 'diproses_at'
  ]);

  _pastikanSheet_(ss, SHEET_PENGATURAN, [
    'nama_app', 'logo_url', 'lat_kantor', 'long_kantor',
    'radius_meter', 'jam_masuk', 'jam_pulang', 'toleransi_menit'
  ]);

  _pastikanSheet_(ss, SHEET_SESSIONS, [
    'token', 'id_user', 'created_at', 'expired_at'
  ]);

  // Isi baris default Pengaturan jika masih kosong (belum ada baris data)
  const sheetPengaturan = ss.getSheetByName(SHEET_PENGATURAN);
  if (sheetPengaturan.getLastRow() < 2) {
    sheetPengaturan.appendRow([
      'Aplikasi Absensi Karyawan', // nama_app
      '',                          // logo_url
      -6.200000,                   // lat_kantor (contoh, GANTI sesuai lokasi kantor asli)
      106.816666,                  // long_kantor (contoh, GANTI sesuai lokasi kantor asli)
      100,                         // radius_meter
      '08:00',                     // jam_masuk
      '17:00',                     // jam_pulang
      15                           // toleransi_menit
    ]);
  }

  // Buat folder Drive jika belum ada
  getOrCreateFolder_(FOLDER_ABSENSI_NAME);
  getOrCreateFolder_(FOLDER_IZIN_NAME);
  getOrCreateFolder_(FOLDER_REKAP_NAME);

  // Buat akun admin default jika sheet Users masih kosong
  const sheetUsers = ss.getSheetByName(SHEET_USERS);
  if (sheetUsers.getLastRow() < 2) {
    const salt = Utilities.getUuid();
    const passwordDefault = 'admin123'; // WAJIB diganti setelah login pertama kali!
    const hash = hashPassword(passwordDefault, salt);
    const now = new Date();

    sheetUsers.appendRow([
      'USR-000001',
      'Administrator',
      'admin',
      hash,
      salt,
      'admin',
      'Administrator',
      '',
      'aktif',
      now,
      now
    ]);

    Logger.log('Akun admin default dibuat -> username: admin | password: admin123');
    Logger.log('SEGERA login dan ganti password ini!');
  }

  Logger.log('Setup awal selesai.');
}

/**
 * Helper internal: buat sheet dengan header jika belum ada.
 * Tidak menimpa sheet yang sudah ada isinya.
 */
function _pastikanSheet_(ss, namaSheet, headerArray) {
  let sheet = ss.getSheetByName(namaSheet);
  if (!sheet) {
    sheet = ss.insertSheet(namaSheet);
    sheet.appendRow(headerArray);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headerArray.length).setFontWeight('bold');
  }
  return sheet;
}

/* ===================================================================
 * HELPER UMUM (dipakai semua file .gs lain)
 * =================================================================== */

/**
 * Ambil objek Sheet berdasarkan nama.
 */
function getSheet(namaSheet) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(namaSheet);
  if (!sheet) {
    throw new Error(`Sheet "${namaSheet}" tidak ditemukan. Jalankan fungsi setupAwal() terlebih dahulu.`);
  }
  return sheet;
}

/**
 * Bungkus hasil jadi response JSON untuk dikirim ke client
 * (dipanggil lewat google.script.run, jadi sebenarnya cukup
 * mengembalikan object biasa -- fungsi ini menjaga format
 * konsisten { success, data/message, ... } di semua fungsi).
 */
function jsonResponse(obj) {
  return obj;
}

/**
 * Hash password dengan SHA-256 + salt.
 */
function hashPassword(password, salt) {
  const rawHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password + salt,
    Utilities.Charset.UTF_8
  );
  return rawHash.map(byte => {
    const v = (byte < 0 ? byte + 256 : byte).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

/**
 * Generate ID berurutan (auto increment) dengan prefix tertentu.
 * Contoh: generateSequentialId('ABS', SHEET_ABSENSI, 0) -> 'ABS-000001'
 * @param {string} prefix     - misal 'ABS', 'USR', 'IZN'
 * @param {string} namaSheet
 * @param {number} kolomId    - index kolom ID (0 = kolom A)
 */
function generateSequentialId(prefix, namaSheet, kolomId) {
  const sheet = getSheet(namaSheet);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return `${prefix}-000001`;
  }

  const ids = sheet.getRange(2, kolomId + 1, lastRow - 1, 1).getValues().flat();
  let maxNum = 0;

  ids.forEach(id => {
    const match = String(id).match(new RegExp(`^${prefix}-(\\d+)$`));
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxNum) maxNum = num;
    }
  });

  const nextNum = maxNum + 1;
  return `${prefix}-${String(nextNum).padStart(6, '0')}`;
}

/**
 * Ambil data Pengaturan (1 baris) dengan cache (CacheService),
 * supaya tidak baca sheet setiap kali halaman absen dibuka.
 */
function getPengaturanCache() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY_PENGATURAN);

  if (cached) {
    return JSON.parse(cached);
  }

  const sheet = getSheet(SHEET_PENGATURAN);
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    throw new Error('Data Pengaturan kosong. Jalankan setupAwal() terlebih dahulu.');
  }

  const row = data[1];
  const pengaturan = {
    nama_app: row[0],
    logo_url: row[1],
    lat_kantor: row[2],
    long_kantor: row[3],
    radius_meter: row[4],
    jam_masuk: row[5],
    jam_pulang: row[6],
    toleransi_menit: row[7]
  };

  cache.put(CACHE_KEY_PENGATURAN, JSON.stringify(pengaturan), CACHE_DURATION_DETIK);
  return pengaturan;
}

/**
 * Panggil ini setiap kali data Pengaturan diubah dari halaman admin,
 * supaya cache lama tidak dipakai lagi.
 */
function bersihkanCachePengaturan() {
  CacheService.getScriptCache().remove(CACHE_KEY_PENGATURAN);
}

/**
 * Ambil folder Drive berdasarkan nama; buat jika belum ada.
 * ID folder di-cache di PropertiesService supaya tidak perlu
 * search folder berulang kali (mendukung performa).
 */
function getOrCreateFolder_(namaFolder) {
  const props = PropertiesService.getScriptProperties();
  const key = 'folder_' + namaFolder;
  const cachedId = props.getProperty(key);

  if (cachedId) {
    try {
      return DriveApp.getFolderById(cachedId);
    } catch (e) {
      // Folder mungkin sudah dihapus manual, lanjut cari/buat ulang
    }
  }

  const folders = DriveApp.getFoldersByName(namaFolder);
  let folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(namaFolder);
  }

  props.setProperty(key, folder.getId());
  return folder;
}

/**
 * Hitung jarak dua koordinat GPS dalam meter (rumus Haversine).
 * Dipakai di Absensi.gs untuk validasi radius kantor.
 */
function hitungJarakMeter(lat1, long1, lat2, long2) {
  const R = 6371000; // radius bumi dalam meter
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLong = (long2 - long1) * Math.PI / 180;
  const a =
    Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
    Math.sin(dLong / 2) * Math.sin(dLong / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return R * c;
}
