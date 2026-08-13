/**
 * ===================================================================
 * CODE.GS - FINAL
 * Konfigurasi, setup, doGet, helper umum.
 * Kompatibel dengan Auth.gs, Absensi.gs, Izin.gs, Admin.gs.
 * ===================================================================
 */

const SHEET_USERS      = 'Users';
const SHEET_ABSENSI    = 'Absensi';
const SHEET_IZIN       = 'Izin';
const SHEET_PENGATURAN = 'Pengaturan';
const SHEET_SESSIONS   = 'Sessions';

const FOLDER_ABSENSI_NAME = 'Foto_Absensi';
const FOLDER_IZIN_NAME    = 'Lampiran_Izin';
const FOLDER_REKAP_NAME   = 'Rekap_PDF';

const CACHE_KEY_PENGATURAN = 'pengaturan_app';
const CACHE_KEY_USER_PREFIX = 'user_auth_';
const CACHE_DURATION_DETIK = 21600;
const AUTH_CACHE_DETIK = 60;

const SESSION_DURASI_JAM = 10;


/* ===================================================================
 * WEB APP
 * =================================================================== */

function doGet(e) {
  var page = 'Login';

  if (
    e &&
    e.parameter &&
    e.parameter.page
  ) {
    page = String(e.parameter.page)
      .replace(/[^a-zA-Z0-9_-]/g, '');

    if (!page) page = 'Login';
  }

  try {
    return HtmlService
      .createTemplateFromFile(page)
      .evaluate()
      .setTitle(getNamaAppSafe_())
      .addMetaTag(
        'viewport',
        'width=device-width, initial-scale=1'
      )
      .setXFrameOptionsMode(
        HtmlService.XFrameOptionsMode.ALLOWALL
      );
  } catch (err) {
    return HtmlService
      .createTemplateFromFile('Login')
      .evaluate()
      .setTitle(getNamaAppSafe_())
      .addMetaTag(
        'viewport',
        'width=device-width, initial-scale=1'
      )
      .setXFrameOptionsMode(
        HtmlService.XFrameOptionsMode.ALLOWALL
      );
  }
}

function include(filename) {
  if (!filename) {
    throw new Error('Nama file HTML wajib diisi.');
  }

  filename = String(filename)
    .replace(/[^a-zA-Z0-9_-]/g, '');

  if (!filename) {
    throw new Error('Nama file HTML tidak valid.');
  }

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}

function getNamaAppSafe_() {
  try {
    var p = getPengaturanCache();
    return String(
      p.nama_app || 'Aplikasi Absensi Karyawan'
    );
  } catch (e) {
    return 'Aplikasi Absensi Karyawan';
  }
}


/* ===================================================================
 * SETUP
 * =================================================================== */

function setupAwal() {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(20000);

    var ss =
      SpreadsheetApp.getActiveSpreadsheet();

    _pastikanSheet_(
      ss,
      SHEET_USERS,
      [
        'id_user',
        'nama',
        'username',
        'password_hash',
        'salt',
        'role',
        'jabatan',
        'foto_profil',
        'status_aktif',
        'created_at',
        'updated_at'
      ]
    );

    _pastikanSheet_(
      ss,
      SHEET_ABSENSI,
      [
        'id_absen',
        'id_user',
        'tanggal',
        'jam_masuk',
        'foto_masuk',
        'lat_masuk',
        'long_masuk',
        'jam_pulang',
        'foto_pulang',
        'lat_pulang',
        'long_pulang',
        'status',
        'jarak_masuk_meter',
        'jarak_pulang_meter'
      ]
    );

    _pastikanSheet_(
      ss,
      SHEET_IZIN,
      [
        'id_izin',
        'id_user',
        'tanggal_pengajuan',
        'jenis_izin',
        'tanggal_mulai',
        'tanggal_selesai',
        'keterangan',
        'lampiran',
        'status',
        'catatan_admin',
        'diproses_oleh',
        'diproses_at'
      ]
    );

    _pastikanSheet_(
      ss,
      SHEET_PENGATURAN,
      [
        'nama_app',
        'logo_url',
        'lat_kantor',
        'long_kantor',
        'radius_meter',
        'jam_masuk',
        'jam_pulang',
        'toleransi_menit'
      ]
    );

    _pastikanSheet_(
      ss,
      SHEET_SESSIONS,
      [
        'token',
        'id_user',
        'created_at',
        'expired_at'
      ]
    );

    var sheetPengaturan =
      ss.getSheetByName(SHEET_PENGATURAN);

    if (
      sheetPengaturan.getLastRow() < 2
    ) {
      sheetPengaturan.appendRow([
        'Aplikasi Absensi Karyawan',
        '',
        -6.200000,
        106.816666,
        100,
        '08:00',
        '17:00',
        15
      ]);
    }

    getOrCreateFolder_(FOLDER_ABSENSI_NAME);
    getOrCreateFolder_(FOLDER_IZIN_NAME);
    getOrCreateFolder_(FOLDER_REKAP_NAME);

    /*
     * Jangan lagi membuat password admin statis.
     * Jika Users kosong, buat password random dan tampilkan
     * hanya pada hasil eksekusi/setup log.
     */
    var sheetUsers =
      ss.getSheetByName(SHEET_USERS);

    var hasilAdmin = null;

    if (sheetUsers.getLastRow() < 2) {
      var username = 'admin';
      var password =
        buatPasswordAwalRandom_();
      var salt = Utilities.getUuid();
      var hash =
        hashPassword(password, salt);
      var now = new Date();

      sheetUsers.appendRow([
        'USR-000001',
        'Administrator',
        username,
        hash,
        salt,
        'admin',
        'Administrator',
        '',
        'aktif',
        now,
        now
      ]);

      hasilAdmin = {
        username: username,
        password: password
      };

      Logger.log(
        'AKUN ADMIN AWAL: username=' +
        username +
        ' password=' +
        password
      );
      Logger.log(
        'Simpan password tersebut. Jangan dibagikan.'
      );
    }

    bersihkanCachePengaturan();

    return {
      success: true,
      message:
        'Setup awal selesai.',
      admin_awal:
        hasilAdmin
    };

  } catch (e) {
    return {
      success: false,
      message:
        e.message ||
        'Setup gagal.'
    };

  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}

function buatPasswordAwalRandom_() {
  var uuid =
    Utilities.getUuid()
      .replace(/-/g, '');

  return 'Adm!' +
    uuid.substring(0, 12) +
    '9x';
}

function _pastikanSheet_(
  ss,
  namaSheet,
  headerArray
) {
  var sheet =
    ss.getSheetByName(namaSheet);

  if (!sheet) {
    sheet =
      ss.insertSheet(namaSheet);

    sheet
      .getRange(
        1,
        1,
        1,
        headerArray.length
      )
      .setValues([headerArray]);

    sheet.setFrozenRows(1);

    sheet
      .getRange(
        1,
        1,
        1,
        headerArray.length
      )
      .setFontWeight('bold');

    return sheet;
  }

  /*
   * Jangan menimpa header/data lama.
   * Tetapi pastikan sheet kosong benar-benar memiliki header.
   */
  if (
    sheet.getLastRow() === 0
  ) {
    sheet
      .getRange(
        1,
        1,
        1,
        headerArray.length
      )
      .setValues([headerArray]);

    sheet.setFrozenRows(1);
  }

  return sheet;
}


/* ===================================================================
 * HELPER SHEET
 * =================================================================== */

function getSheet(namaSheet) {
  var sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(namaSheet);

  if (!sheet) {
    throw new Error(
      'Sheet "' +
      namaSheet +
      '" tidak ditemukan. Jalankan setupAwal() terlebih dahulu.'
    );
  }

  return sheet;
}

function jsonResponse(obj) {
  return obj;
}


/* ===================================================================
 * PASSWORD
 * =================================================================== */

function hashPassword(
  password,
  salt
) {
  password =
    String(password || '');

  salt =
    String(salt || '');

  if (!password || !salt) {
    throw new Error(
      'Password dan salt wajib diisi.'
    );
  }

  var rawHash =
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      password + salt,
      Utilities.Charset.UTF_8
    );

  return rawHash
    .map(function(byte) {
      var v =
        byte < 0
          ? byte + 256
          : byte;

      var hex =
        v.toString(16);

      return hex.length === 1
        ? '0' + hex
        : hex;
    })
    .join('');
}


/* ===================================================================
 * ID
 *
 * Pemanggil operasi tulis tetap WAJIB menggunakan ScriptLock.
 * Fungsi ini sendiri tidak mengambil lock agar tidak terjadi deadlock
 * ketika dipanggil dari fungsi yang sudah memegang ScriptLock.
 * =================================================================== */

function generateSequentialId(
  prefix,
  namaSheet,
  kolomId
) {
  prefix =
    String(prefix || '')
      .replace(/[^A-Z0-9_-]/gi, '');

  if (!prefix) {
    throw new Error(
      'Prefix ID tidak valid.'
    );
  }

  var sheet =
    getSheet(namaSheet);

  var lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return (
      prefix +
      '-000001'
    );
  }

  var jumlah =
    lastRow - 1;

  var ids =
    sheet
      .getRange(
        2,
        Number(kolomId) + 1,
        jumlah,
        1
      )
      .getValues();

  var regex =
    new RegExp(
      '^' +
      escapeRegex_(prefix) +
      '-(\\d+)$'
    );

  var maxNum = 0;

  for (
    var i = 0;
    i < ids.length;
    i++
  ) {
    var match =
      String(
        ids[i][0] || ''
      ).match(regex);

    if (match) {
      var num =
        parseInt(
          match[1],
          10
        );

      if (
        !isNaN(num) &&
        num > maxNum
      ) {
        maxNum = num;
      }
    }
  }

  return (
    prefix +
    '-' +
    String(maxNum + 1)
      .padStart(6, '0')
  );
}

function escapeRegex_(text) {
  return String(text)
    .replace(
      /[.*+?^${}()|[\]\\]/g,
      '\\$&'
    );
}


/* ===================================================================
 * PENGATURAN
 * =================================================================== */

function getPengaturanCache() {
  var cache =
    CacheService.getScriptCache();

  var cached =
    cache.get(
      CACHE_KEY_PENGATURAN
    );

  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      cache.remove(
        CACHE_KEY_PENGATURAN
      );
    }
  }

  var sheet =
    getSheet(
      SHEET_PENGATURAN
    );

  if (
    sheet.getLastRow() < 2
  ) {
    throw new Error(
      'Data Pengaturan kosong. Jalankan setupAwal() terlebih dahulu.'
    );
  }

  /*
   * Hanya membaca satu baris x 8 kolom.
   */
  var row =
    sheet
      .getRange(
        2,
        1,
        1,
        8
      )
      .getValues()[0];

  var pengaturan = {
    nama_app:
      row[0],
    logo_url:
      row[1],
    lat_kantor:
      row[2],
    long_kantor:
      row[3],
    radius_meter:
      row[4],
    jam_masuk:
      row[5],
    jam_pulang:
      row[6],
    toleransi_menit:
      row[7]
  };

  cache.put(
    CACHE_KEY_PENGATURAN,
    JSON.stringify(pengaturan),
    CACHE_DURATION_DETIK
  );

  return pengaturan;
}

function bersihkanCachePengaturan() {
  CacheService
    .getScriptCache()
    .remove(
      CACHE_KEY_PENGATURAN
    );
}


/* ===================================================================
 * DRIVE
 * =================================================================== */

function getOrCreateFolder_(
  namaFolder
) {
  var props =
    PropertiesService
      .getScriptProperties();

  var key =
    'folder_' +
    namaFolder;

  var cachedId =
    props.getProperty(key);

  if (cachedId) {
    try {
      return DriveApp
        .getFolderById(
          cachedId
        );
    } catch (e) {}
  }

  var folders =
    DriveApp
      .getFoldersByName(
        namaFolder
      );

  var folder;

  if (folders.hasNext()) {
    folder =
      folders.next();
  } else {
    folder =
      DriveApp.createFolder(
        namaFolder
      );
  }

  props.setProperty(
    key,
    folder.getId()
  );

  return folder;
}


/* ===================================================================
 * GPS
 * =================================================================== */

function hitungJarakMeter(
  lat1,
  long1,
  lat2,
  long2
) {
  lat1 = Number(lat1);
  long1 = Number(long1);
  lat2 = Number(lat2);
  long2 = Number(long2);

  if (
    !isFinite(lat1) ||
    !isFinite(long1) ||
    !isFinite(lat2) ||
    !isFinite(long2)
  ) {
    throw new Error(
      'Koordinat GPS tidak valid.'
    );
  }

  var R = 6371000;

  var dLat =
    (lat2 - lat1) *
    Math.PI / 180;

  var dLong =
    (long2 - long1) *
    Math.PI / 180;

  var a =
    Math.sin(dLat / 2) *
    Math.sin(dLat / 2) +
    Math.cos(
      lat1 * Math.PI / 180
    ) *
    Math.cos(
      lat2 * Math.PI / 180
    ) *
    Math.sin(dLong / 2) *
    Math.sin(dLong / 2);

  var c =
    2 *
    Math.atan2(
      Math.sqrt(a),
      Math.sqrt(1 - a)
    );

  return R * c;
}


/* ===================================================================
 * OPTIONAL MAINTENANCE
 * =================================================================== */

function buatTriggerPembersihanSession() {
  var triggers =
    ScriptApp.getProjectTriggers();

  for (
    var i = 0;
    i < triggers.length;
    i++
  ) {
    if (
      triggers[i]
        .getHandlerFunction() ===
      'bersihkanSessionExpired'
    ) {
      ScriptApp.deleteTrigger(
        triggers[i]
      );
    }
  }

  ScriptApp
    .newTrigger(
      'bersihkanSessionExpired'
    )
    .timeBased()
    .everyHours(1)
    .create();

  return {
    success: true,
    message:
      'Trigger pembersihan session dibuat.'
  };
}
