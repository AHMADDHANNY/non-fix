/**
 * ===================================================================
 * CODE.GS
 * APLIKASI ABSENSI KARYAWAN
 *
 * Fungsi:
 * - Konfigurasi aplikasi
 * - Nama sheet
 * - Entry point Web App
 * - include HTML
 * - setup database
 * - helper umum
 * - hashing password
 * - ID generator
 * - pengaturan/cache
 * - folder Google Drive
 * - perhitungan jarak GPS
 *
 * PENTING:
 * Web App SELALU membuka Login.html.
 *
 * TIDAK ADA:
 * ?page=DashboardAdmin
 * ?page=DashboardPegawai
 *
 * Role ditentukan setelah login berdasarkan session.
 * ===================================================================
 */


/* ===================================================================
 * KONFIGURASI SHEET
 * =================================================================== */

const SHEET_USERS      = 'Users';
const SHEET_ABSENSI    = 'Absensi';
const SHEET_IZIN       = 'Izin';
const SHEET_PENGATURAN = 'Pengaturan';
const SHEET_SESSIONS   = 'Sessions';


/* ===================================================================
 * KONFIGURASI FOLDER DRIVE
 * =================================================================== */

const FOLDER_ABSENSI_NAME = 'Foto_Absensi';
const FOLDER_IZIN_NAME    = 'Lampiran_Izin';
const FOLDER_REKAP_NAME   = 'Rekap_PDF';


/* ===================================================================
 * CACHE
 * =================================================================== */

const CACHE_KEY_PENGATURAN = 'pengaturan_app';
const CACHE_DURATION_DETIK = 21600;


/* ===================================================================
 * SESSION
 * =================================================================== */

const SESSION_DURASI_JAM = 10;


/* ===================================================================
 * ENTRY POINT WEB APP
 * ===================================================================
 *
 * SECURITY:
 *
 * Web App TIDAK BOLEH menerima parameter:
 *
 * ?page=DashboardAdmin
 * ?page=DashboardPegawai
 *
 * Semua user masuk melalui Login.html.
 *
 * Setelah login:
 *
 * admin   -> dashboard admin
 * pegawai -> dashboard pegawai
 *
 * Penentuan role dilakukan berdasarkan response login()
 * dan setiap fungsi backend tetap melakukan validasi session.
 * ===================================================================
 */

function doGet(e) {

  try {

    return HtmlService
      .createTemplateFromFile('Login')
      .evaluate()
      .setTitle(getNamaAppSafe_())
      .addMetaTag(
        'viewport',
        'width=device-width, initial-scale=1, maximum-scale=1'
      )
      .setXFrameOptionsMode(
        HtmlService.XFrameOptionsMode.ALLOWALL
      );

  } catch (err) {

    return HtmlService
      .createHtmlOutput(
        '<!DOCTYPE html>' +
        '<html lang="id">' +
        '<head>' +
        '<meta name="viewport" content="width=device-width,initial-scale=1">' +
        '<title>Aplikasi Absensi</title>' +
        '</head>' +
        '<body style="font-family:Arial;padding:30px">' +
        '<h2>Aplikasi tidak dapat dimuat</h2>' +
        '<p>File <b>Login.html</b> tidak ditemukan atau terjadi kesalahan pada aplikasi.</p>' +
        '</body>' +
        '</html>'
      )
      .setTitle('Aplikasi Absensi Karyawan');

  }

}


/* ===================================================================
 * INCLUDE HTML
 * ===================================================================
 *
 * Contoh di Login.html:
 *
 * <?!= include('CSS'); ?>
 * <?!= include('JS'); ?>
 *
 * Jika file tidak ada, error akan diberikan supaya mudah diketahui.
 * ===================================================================
 */

function include(filename) {

  if (!filename) {
    throw new Error('Nama file HTML tidak boleh kosong.');
  }

  return HtmlService
    .createHtmlOutputFromFile(String(filename))
    .getContent();

}


/* ===================================================================
 * NAMA APLIKASI
 * =================================================================== */

function getNamaAppSafe_() {

  try {

    const cache =
      CacheService.getScriptCache();

    const cached =
      cache.get(CACHE_KEY_PENGATURAN);

    if (cached) {

      const data =
        JSON.parse(cached);

      if (
        data &&
        data.nama_app
      ) {

        return String(
          data.nama_app
        );

      }

    }

    const sheet =
      getSheet(SHEET_PENGATURAN);

    if (
      sheet.getLastRow() < 2
    ) {

      return 'Aplikasi Absensi Karyawan';

    }

    const nama =
      sheet
        .getRange(2, 1)
        .getDisplayValue();

    return String(
      nama ||
      'Aplikasi Absensi Karyawan'
    );

  } catch (err) {

    return 'Aplikasi Absensi Karyawan';

  }

}


/* ===================================================================
 * SETUP AWAL
 * ===================================================================
 *
 * Jalankan SATU KALI dari Apps Script:
 *
 * setupAwal
 *
 * Fungsi ini:
 * - membuat Users
 * - membuat Absensi
 * - membuat Izin
 * - membuat Pengaturan
 * - membuat Sessions
 * - membuat folder Drive
 * - membuat akun admin pertama
 *
 * Admin default:
 *
 * username : admin
 * password : admin123
 *
 * SEGERA GANTI PASSWORD SETELAH LOGIN.
 * ===================================================================
 */

function setupAwal() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  /* ---------------------------------------------------------------
   * USERS
   * --------------------------------------------------------------- */

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


  /* ---------------------------------------------------------------
   * ABSENSI
   * --------------------------------------------------------------- */

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


  /* ---------------------------------------------------------------
   * IZIN
   * --------------------------------------------------------------- */

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


  /* ---------------------------------------------------------------
   * PENGATURAN
   * --------------------------------------------------------------- */

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


  /* ---------------------------------------------------------------
   * SESSIONS
   * --------------------------------------------------------------- */

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


  /* ---------------------------------------------------------------
   * DEFAULT PENGATURAN
   * --------------------------------------------------------------- */

  const sheetPengaturan =
    ss.getSheetByName(
      SHEET_PENGATURAN
    );

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


  /* ---------------------------------------------------------------
   * FOLDER DRIVE
   * --------------------------------------------------------------- */

  getOrCreateFolder_(
    FOLDER_ABSENSI_NAME
  );

  getOrCreateFolder_(
    FOLDER_IZIN_NAME
  );

  getOrCreateFolder_(
    FOLDER_REKAP_NAME
  );


  /* ---------------------------------------------------------------
   * ADMIN DEFAULT
   * --------------------------------------------------------------- */

  const sheetUsers =
    ss.getSheetByName(
      SHEET_USERS
    );

  if (
    sheetUsers.getLastRow() < 2
  ) {

    const salt =
      Utilities.getUuid();

    const passwordDefault =
      'admin123';

    const hash =
      hashPassword(
        passwordDefault,
        salt
      );

    const now =
      new Date();

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

    Logger.log(
      'ADMIN DEFAULT: admin / admin123'
    );

  }


  Logger.log(
    'Setup awal selesai.'
  );

}


/* ===================================================================
 * MEMBUAT SHEET
 * =================================================================== */

function _pastikanSheet_(
  ss,
  namaSheet,
  headerArray
) {

  let sheet =
    ss.getSheetByName(
      namaSheet
    );

  if (!sheet) {

    sheet =
      ss.insertSheet(
        namaSheet
      );

    sheet.appendRow(
      headerArray
    );

    sheet.setFrozenRows(1);

    sheet
      .getRange(
        1,
        1,
        1,
        headerArray.length
      )
      .setFontWeight('bold');

  }

  return sheet;

}


/* ===================================================================
 * GET SHEET
 * =================================================================== */

function getSheet(
  namaSheet
) {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        namaSheet
      );

  if (!sheet) {

    throw new Error(
      'Sheet "' +
      namaSheet +
      '" tidak ditemukan. ' +
      'Jalankan setupAwal() terlebih dahulu.'
    );

  }

  return sheet;

}


/* ===================================================================
 * RESPONSE
 * =================================================================== */

function jsonResponse(obj) {

  return obj;

}


/* ===================================================================
 * PASSWORD HASH
 * =================================================================== */

function hashPassword(
  password,
  salt
) {

  const rawHash =
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      String(password) +
      String(salt),
      Utilities.Charset.UTF_8
    );

  return rawHash
    .map(function(byte) {

      const value =
        byte < 0
          ? byte + 256
          : byte;

      const hex =
        value.toString(16);

      return hex.length === 1
        ? '0' + hex
        : hex;

    })
    .join('');

}


/* ===================================================================
 * GENERATE SEQUENTIAL ID
 * =================================================================== */

function generateSequentialId(
  prefix,
  namaSheet,
  kolomId
) {

  const sheet =
    getSheet(namaSheet);

  const lastRow =
    sheet.getLastRow();

  if (
    lastRow < 2
  ) {

    return (
      String(prefix) +
      '-000001'
    );

  }

  const ids =
    sheet
      .getRange(
        2,
        kolomId + 1,
        lastRow - 1,
        1
      )
      .getValues()
      .flat();

  let maxNum = 0;

  ids.forEach(function(id) {

    const match =
      String(id)
        .match(
          new RegExp(
            '^' +
            String(prefix) +
            '-(\\d+)$'
          )
        );

    if (match) {

      const number =
        parseInt(
          match[1],
          10
        );

      if (
        number > maxNum
      ) {

        maxNum =
          number;

      }

    }

  });

  return (
    String(prefix) +
    '-' +
    String(
      maxNum + 1
    ).padStart(
      6,
      '0'
    )
  );

}


/* ===================================================================
 * PENGATURAN + CACHE
 * =================================================================== */

function getPengaturanCache() {

  const cache =
    CacheService
      .getScriptCache();

  const cached =
    cache.get(
      CACHE_KEY_PENGATURAN
    );

  if (cached) {

    return JSON.parse(
      cached
    );

  }


  const sheet =
    getSheet(
      SHEET_PENGATURAN
    );

  if (
    sheet.getLastRow() < 2
  ) {

    throw new Error(
      'Data Pengaturan kosong. ' +
      'Jalankan setupAwal().'
    );

  }


  const row =
    sheet
      .getRange(
        2,
        1,
        1,
        8
      )
      .getValues()[0];


  const pengaturan = {

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
    JSON.stringify(
      pengaturan
    ),
    CACHE_DURATION_DETIK
  );


  return pengaturan;

}


/* ===================================================================
 * BERSIHKAN CACHE PENGATURAN
 * =================================================================== */

function bersihkanCachePengaturan() {

  CacheService
    .getScriptCache()
    .remove(
      CACHE_KEY_PENGATURAN
    );

}


/* ===================================================================
 * FOLDER DRIVE
 * =================================================================== */

function getOrCreateFolder_(
  namaFolder
) {

  const props =
    PropertiesService
      .getScriptProperties();

  const key =
    'folder_' +
    namaFolder;

  const cachedId =
    props.getProperty(
      key
    );


  if (cachedId) {

    try {

      return DriveApp
        .getFolderById(
          cachedId
        );

    } catch (err) {

      // Folder mungkin sudah dihapus.
      // Lanjut cari ulang.

    }

  }


  const folders =
    DriveApp
      .getFoldersByName(
        namaFolder
      );

  let folder;


  if (
    folders.hasNext()
  ) {

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
 * HITUNG JARAK GPS
 * =================================================================== */

function hitungJarakMeter(
  lat1,
  long1,
  lat2,
  long2
) {

  const R =
    6371000;


  const dLat =
    (
      Number(lat2) -
      Number(lat1)
    ) *
    Math.PI /
    180;


  const dLong =
    (
      Number(long2) -
      Number(long1)
    ) *
    Math.PI /
    180;


  const a =
    Math.sin(
      dLat / 2
    ) *
    Math.sin(
      dLat / 2
    ) +

    Math.cos(
      Number(lat1) *
      Math.PI /
      180
    ) *

    Math.cos(
      Number(lat2) *
      Math.PI /
      180
    ) *

    Math.sin(
      dLong / 2
    ) *
    Math.sin(
      dLong / 2
    );


  const c =
    2 *
    Math.atan2(
      Math.sqrt(a),
      Math.sqrt(1 - a)
    );


  return R * c;

}


/* ===================================================================
 * AUDIT DATABASE
 * ===================================================================
 *
 * Jalankan manual untuk memastikan struktur database benar.
 * =================================================================== */

function auditDatabase() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const expected = {

    Users: [
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
    ],

    Absensi: [
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
    ],

    Izin: [
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
    ],

    Pengaturan: [
      'nama_app',
      'logo_url',
      'lat_kantor',
      'long_kantor',
      'radius_meter',
      'jam_masuk',
      'jam_pulang',
      'toleransi_menit'
    ],

    Sessions: [
      'token',
      'id_user',
      'created_at',
      'expired_at'
    ]

  };


  const result = {
    success: true,
    timestamp:
      new Date().toISOString(),
    sheets: [],
    semua_valid: true
  };


  Object.keys(expected)
    .forEach(function(name) {

      const sheet =
        ss.getSheetByName(
          name
        );


      if (!sheet) {

        result.semua_valid =
          false;

        result.sheets.push({
          sheet: name,
          exists: false,
          valid: false,
          message:
            'Sheet tidak ditemukan.',
          expected_headers:
            expected[name]
        });

        return;

      }


      const headerCount =
        expected[name].length;

      let headers = [];

      if (
        sheet.getLastColumn() >=
        headerCount
      ) {

        headers =
          sheet
            .getRange(
              1,
              1,
              1,
              headerCount
            )
            .getValues()[0]
            .map(function(value) {
              return String(
                value || ''
              ).trim();
            });

      }


      const missing =
        expected[name]
          .filter(function(header) {
            return headers.indexOf(
              header
            ) === -1;
          });


      const extra =
        headers
          .filter(function(header) {
            return (
              header &&
              expected[name]
                .indexOf(header) === -1
            );
          });


      const valid =
        missing.length === 0 &&
        extra.length === 0 &&
        headers.length ===
          expected[name].length;


      if (!valid) {
        result.semua_valid =
          false;
      }


      result.sheets.push({

        sheet: name,

        exists: true,

        valid: valid,

        rows:
          Math.max(
            0,
            sheet.getLastRow() - 1
          ),

        columns:
          sheet.getLastColumn(),

        current_headers:
          headers,

        expected_headers:
          expected[name],

        missing_headers:
          missing,

        extra_headers:
          extra

      });

    });


  Logger.log(
    JSON.stringify(
      result,
      null,
      2
    )
  );


  return result;

}
