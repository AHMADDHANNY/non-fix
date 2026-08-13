/**
 * ============================================================
 * CODE.GS - CORE / ROUTING / HELPER
 * ============================================================
 *
 * Dipakai oleh:
 * - Auth.gs
 * - Admin.gs
 * - Absensi.gs
 * - Izin.gs
 *
 * Sheet yang digunakan:
 * - Users
 * - Sessions
 * - Absensi
 * - Izin
 * - Pengaturan
 * ============================================================
 */


/* ============================================================
 * KONFIGURASI DATABASE
 * ============================================================ */

var SHEET_USERS =
  'Users';

var SHEET_SESSIONS =
  'Sessions';

var SHEET_ABSENSI =
  'Absensi';

var SHEET_IZIN =
  'Izin';

var SHEET_PENGATURAN =
  'Pengaturan';


/* ============================================================
 * KONFIGURASI SESSION
 * ============================================================ */

var SESSION_DURASI_JAM =
  12;


/* ============================================================
 * KONFIGURASI FOLDER DRIVE
 * ============================================================ */

var FOLDER_ABSENSI_NAME =
  'Foto_Absensi';

var FOLDER_IZIN_NAME =
  'Lampiran_Izin';

var FOLDER_REKAP_NAME =
  'Rekap_Absensi';


/* ============================================================
 * ROUTING WEB APP
 * ============================================================ */

function doGet(e) {

  var page =
    'Login';

  try {

    var requestedPage =
      '';

    if (
      e &&
      e.parameter &&
      e.parameter.page
    ) {

      requestedPage =
        String(
          e.parameter.page
        )
          .trim()
          .replace(
            /[^a-zA-Z0-9_-]/g,
            ''
          );

    }


    var allowedPages = [
      'Login',
      'DashboardAdmin',
      'DashboardPegawai'
    ];


    if (
      allowedPages.indexOf(
        requestedPage
      ) !== -1
    ) {

      page =
        requestedPage;

    }


    return renderPage_(
      page
    );


  } catch (error) {

    return renderErrorPage_(
      page,
      error
    );

  }

}


/* ============================================================
 * RENDER PAGE
 * ============================================================ */

function renderPage_(
  page
) {

  var template =
    HtmlService
      .createTemplateFromFile(
        page
      );


  return template
    .evaluate()
    .setTitle(
      getNamaAppSafe_()
    )
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover'
    )
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/* ============================================================
 * INCLUDE HTML
 * ============================================================ */

function include(
  filename
) {

  if (!filename) {

    throw new Error(
      'include(): nama file kosong.'
    );

  }


  var name =
    String(
      filename
    )
      .trim()
      .replace(
        /\.html$/i,
        ''
      );


  if (
    !/^[a-zA-Z0-9_-]+$/.test(
      name
    )
  ) {

    throw new Error(
      'Nama file include tidak valid: ' +
      name
    );

  }


  try {

    return HtmlService
      .createHtmlOutputFromFile(
        name
      )
      .getContent();


  } catch (error) {

    throw new Error(
      'File HTML "' +
      name +
      '.html" tidak ditemukan. ' +
      error.message
    );

  }

}


/* ============================================================
 * GET SHEET
 * ============================================================ */

function getSheet(
  sheetName
) {

  if (!sheetName) {

    throw new Error(
      'Nama sheet kosong.'
    );

  }


  var ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  if (!ss) {

    throw new Error(
      'Spreadsheet aplikasi tidak ditemukan.'
    );

  }


  var sheet =
    ss.getSheetByName(
      String(
        sheetName
      )
    );


  if (!sheet) {

    throw new Error(
      'Sheet "' +
      sheetName +
      '" tidak ditemukan.'
    );

  }


  return sheet;

}


/* ============================================================
 * GENERATE ID BERURUTAN
 * ============================================================ */

function generateSequentialId(
  prefix,
  sheetName,
  idColumn
) {

  var sheet =
    getSheet(
      sheetName
    );


  var lastRow =
    sheet.getLastRow();


  if (
    lastRow < 2
  ) {

    return (
      prefix +
      '-000001'
    );

  }


  var values =
    sheet
      .getRange(
        2,
        idColumn + 1,
        lastRow - 1,
        1
      )
      .getDisplayValues();


  var maxNumber =
    0;


  for (
    var i = 0;
    i < values.length;
    i++
  ) {

    var value =
      String(
        values[i][0] || ''
      ).trim();


    if (!value) {
      continue;
    }


    var match =
      value.match(
        /(\d+)$/
      );


    if (!match) {
      continue;
    }


    var number =
      parseInt(
        match[1],
        10
      );


    if (
      !isNaN(number) &&
      number > maxNumber
    ) {

      maxNumber =
        number;

    }

  }


  return (
    String(prefix) +
    '-' +
    String(
      maxNumber + 1
    ).padStart(
      6,
      '0'
    )
  );

}


/* ============================================================
 * FOLDER DRIVE
 * ============================================================ */

function getOrCreateFolder_(
  folderName
) {

  if (!folderName) {

    throw new Error(
      'Nama folder kosong.'
    );

  }


  var folders =
    DriveApp
      .getFoldersByName(
        String(
          folderName
        )
      );


  if (
    folders.hasNext()
  ) {

    return folders.next();

  }


  return DriveApp
    .createFolder(
      String(
        folderName
      )
    );

}


/* ============================================================
 * PASSWORD HASH
 *
 * SHA-256(password + salt)
 * ============================================================ */

function hashPassword(
  password,
  salt
) {

  password =
    String(
      password || ''
    );

  salt =
    String(
      salt || ''
    );


  var bytes =
    Utilities
      .computeDigest(
        Utilities.DigestAlgorithm.SHA_256,
        password + salt,
        Utilities.Charset.UTF_8
      );


  return bytes
    .map(
      function(byte) {

        var value =
          byte < 0
            ? byte + 256
            : byte;

        return value
          .toString(16)
          .padStart(
            2,
            '0'
          );

      }
    )
    .join('');

}


/* ============================================================
 * VERIFY PASSWORD
 * ============================================================ */

function verifyPassword_(
  password,
  passwordHash,
  salt
) {

  if (
    !passwordHash ||
    !salt
  ) {

    return false;

  }


  return (
    hashPassword(
      password,
      salt
    ) ===
    String(
      passwordHash
    )
  );

}


/* ============================================================
 * COMPATIBILITY:
 * Auth.gs menggunakan createSession_()
 *
 * Auth.gs juga mempunyai buatSessionToken_()
 * ============================================================ */

function createSession_(
  idUser
) {

  return buatSessionToken_(
    idUser
  );

}


/* ============================================================
 * PENGATURAN APLIKASI
 * ============================================================ */

function getPengaturanCache() {

  var cache =
    CacheService
      .getScriptCache();


  var cached =
    cache.get(
      'pengaturan_aplikasi'
    );


  if (cached) {

    try {

      return JSON.parse(
        cached
      );

    } catch (ignore) {}

  }


  var sheet =
    getSheet(
      SHEET_PENGATURAN
    );


  var result = {

    nama_app:
      'Aplikasi Absensi',

    logo_url:
      '',

    lat_kantor:
      '',

    long_kantor:
      '',

    radius_meter:
      '',

    jam_masuk:
      '08:00',

    jam_pulang:
      '17:00',

    toleransi_menit:
      0

  };


  if (
    sheet.getLastRow() >= 2
  ) {

    var row =
      sheet
        .getRange(
          2,
          1,
          1,
          8
        )
        .getValues()[0];


    result.nama_app =
      String(
        row[0] ||
        result.nama_app
      );

    result.logo_url =
      String(
        row[1] ||
        ''
      );

    result.lat_kantor =
      row[2];

    result.long_kantor =
      row[3];

    result.radius_meter =
      row[4];

    result.jam_masuk =
      String(
        row[5] ||
        '08:00'
      );

    result.jam_pulang =
      String(
        row[6] ||
        '17:00'
      );

    result.toleransi_menit =
      Number(
        row[7] || 0
      );

  }


  cache.put(
    'pengaturan_aplikasi',
    JSON.stringify(
      result
    ),
    300
  );


  return result;

}


/* ============================================================
 * BERSIHKAN CACHE PENGATURAN
 * ============================================================ */

function bersihkanCachePengaturan() {

  CacheService
    .getScriptCache()
    .remove(
      'pengaturan_aplikasi'
    );

}


/* Alias kompatibilitas */
function bersihkanCachePengaturan_() {

  bersihkanCachePengaturan();

}


/* ============================================================
 * NAMA APLIKASI
 * ============================================================ */

function getNamaAppSafe_() {

  try {

    var pengaturan =
      getPengaturanCache();


    return (
      String(
        pengaturan.nama_app ||
        ''
      ).trim() ||
      'Aplikasi Absensi'
    );


  } catch (error) {

    return 'Aplikasi Absensi';

  }

}


/* ============================================================
 * JSON RESPONSE
 *
 * Mendukung DUA format:
 *
 * jsonResponse({
 *   success:true
 * })
 *
 * dan:
 *
 * jsonResponse(
 *   true,
 *   'Berhasil',
 *   data
 * )
 * ============================================================ */

function jsonResponse(
  success,
  message,
  data
) {

  /*
   * Format object.
   */
  if (
    success &&
    typeof success === 'object' &&
    !Array.isArray(success)
  ) {

    return success;

  }


  var result = {

    success:
      Boolean(
        success
      )

  };


  if (
    message !==
      undefined &&
    message !==
      null
  ) {

    result.message =
      String(
        message
      );

  }


  if (
    data !==
      undefined &&
    data !==
      null
  ) {

    result.data =
      data;

  }


  return result;

}


/* ============================================================
 * ERROR PAGE
 * ============================================================ */

function renderErrorPage_(
  page,
  error
) {

  var message =
    error &&
    error.message
      ? error.message
      : String(
          error
        );


  return HtmlService
    .createHtmlOutput(

      '<!DOCTYPE html>' +

      '<html lang="id">' +

      '<head>' +

      '<meta charset="UTF-8">' +

      '<meta name="viewport" content="width=device-width,initial-scale=1">' +

      '<title>Gagal Memuat Aplikasi</title>' +

      '<style>' +

      'body{' +
      'margin:0;' +
      'padding:30px;' +
      'font-family:Arial,sans-serif;' +
      'background:#f5f7fb;' +
      'color:#222' +
      '}' +

      '.box{' +
      'max-width:720px;' +
      'margin:40px auto;' +
      'background:#fff;' +
      'padding:30px;' +
      'border-radius:16px;' +
      'box-shadow:0 10px 30px rgba(0,0,0,.08)' +
      '}' +

      'h2{color:#d32f2f;margin-top:0}' +

      'pre{' +
      'background:#f1f1f1;' +
      'padding:15px;' +
      'border-radius:8px;' +
      'white-space:pre-wrap;' +
      'word-break:break-word' +
      '}' +

      '</style>' +

      '</head>' +

      '<body>' +

      '<div class="box">' +

      '<h2>Gagal memuat aplikasi</h2>' +

      '<p>Halaman yang dicoba:</p>' +

      '<pre>' +
      escapeHtml_(
        page
      ) +
      '</pre>' +

      '<p>Error asli:</p>' +

      '<pre>' +
      escapeHtml_(
        message
      ) +
      '</pre>' +

      '</div>' +

      '</body>' +

      '</html>'

    )
    .setTitle(
      'Gagal Memuat Aplikasi'
    );

}


/* ============================================================
 * ESCAPE HTML
 * ============================================================ */

function escapeHtml_(
  value
) {

  return String(
    value || ''
  )
    .replace(
      /&/g,
      '&amp;'
    )
    .replace(
      /</g,
      '&lt;'
    )
    .replace(
      />/g,
      '&gt;'
    )
    .replace(
      /"/g,
      '&quot;'
    )
    .replace(
      /'/g,
      '&#039;'
    );

}


/* ============================================================
 * TEST CORE
 *
 * Jalankan ini SEKALI dari Apps Script
 * setelah mengganti Code.gs.
 * ============================================================ */

function testCoreSystem() {

  var result = {
    success: true,
    checks: []
  };


  try {

    result.checks.push({
      name: 'Users',
      success:
        !!getSheet(
          SHEET_USERS
        )
    });


    result.checks.push({
      name: 'Sessions',
      success:
        !!getSheet(
          SHEET_SESSIONS
        )
    });


    result.checks.push({
      name: 'Absensi',
      success:
        !!getSheet(
          SHEET_ABSENSI
        )
    });


    result.checks.push({
      name: 'Izin',
      success:
        !!getSheet(
          SHEET_IZIN
        )
    });


    result.checks.push({
      name: 'Pengaturan',
      success:
        !!getSheet(
          SHEET_PENGATURAN
        )
    });


    result.checks.push({
      name: 'getPengaturanCache',
      success:
        !!getPengaturanCache()
    });


    return result;


  } catch (error) {

    return {

      success: false,

      message:
        error.message,

      checks:
        result.checks

    };

  }

}


/* ============================================================
 * TEST FILE
 * ============================================================ */

function testFileLogin() {

  try {

    HtmlService
      .createTemplateFromFile(
        'Login'
      );


    return {

      success: true,

      message:
        'Login.html berhasil ditemukan.'

    };


  } catch (error) {

    return {

      success: false,

      message:
        error.message

    };

  }

}


function testFileDashboardAdmin() {

  try {

    HtmlService
      .createTemplateFromFile(
        'DashboardAdmin'
      );


    return {

      success: true,

      message:
        'DashboardAdmin.html berhasil ditemukan.'

    };


  } catch (error) {

    return {

      success: false,

      message:
        error.message

    };

  }

}


function testFileDashboardPegawai() {

  try {

    HtmlService
      .createTemplateFromFile(
        'DashboardPegawai'
      );


    return {

      success: true,

      message:
        'DashboardPegawai.html berhasil ditemukan.'

    };


  } catch (error) {

    return {

      success: false,

      message:
        error.message

    };

  }

}
