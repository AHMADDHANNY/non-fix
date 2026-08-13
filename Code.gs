/**
 * ============================================================
 * CODE.GS - FINAL
 * ROUTING + DATABASE HELPER + APP HELPER
 * ============================================================
 */

var SHEET_USERS = 'Users';
var SHEET_ABSENSI = 'Absensi';
var SHEET_IZIN = 'Izin';
var SHEET_PENGATURAN = 'Pengaturan';
var SHEET_SESSIONS = 'Sessions';

var SESSION_DURASI_JAM = 12;

var FOLDER_REKAP_NAME = 'Absensi_Rekap';
var FOLDER_IZIN_NAME = 'Absensi_Lampiran_Izin';
var FOLDER_ABSENSI_NAME = 'Absensi_Foto';


/* ============================================================
 * WEB APP
 * ============================================================ */

function doGet(e) {

  var page = 'Login';

  try {

    var requestedPage = '';

    if (
      e &&
      e.parameter &&
      e.parameter.page
    ) {
      requestedPage =
        String(e.parameter.page)
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
      page = requestedPage;
    }

    return renderPage_(page);

  } catch (error) {

    return renderErrorPage_(
      page,
      error
    );

  }

}


/* ============================================================
 * RENDER
 * ============================================================ */

function renderPage_(page) {

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
 * INCLUDE
 * ============================================================ */

function include(filename) {

  if (!filename) {
    throw new Error(
      'include(): nama file kosong.'
    );
  }

  var name =
    String(filename)
      .trim()
      .replace(
        /\.html$/i,
        ''
      );

  if (
    !/^[a-zA-Z0-9_-]+$/.test(name)
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
      'Error asli: ' +
      error.message
    );

  }

}


/* ============================================================
 * NAMA APLIKASI
 *
 * PRIVATE:
 * getNamaAppSafe_()
 *
 * PUBLIC:
 * getNamaApp()
 *
 * Browser WAJIB menggunakan getNamaApp().
 * ============================================================ */

function getNamaAppSafe_() {

  var defaultName =
    'Aplikasi Absensi';

  try {

    var ss =
      SpreadsheetApp
        .getActiveSpreadsheet();

    if (!ss) {
      return defaultName;
    }

    var sheet =
      ss.getSheetByName(
        SHEET_PENGATURAN
      );

    if (!sheet) {
      return defaultName;
    }

    if (
      sheet.getLastRow() < 2
    ) {
      return defaultName;
    }

    var value =
      sheet
        .getRange(
          2,
          1
        )
        .getDisplayValue()
        .trim();

    return value ||
      defaultName;

  } catch (error) {

    return defaultName;

  }

}


/*
 * PUBLIC API.
 *
 * Bisa dipanggil:
 *
 * google.script.run.getNamaApp()
 *
 */
function getNamaApp() {

  return getNamaAppSafe_();

}


/* ============================================================
 * DATABASE
 * ============================================================ */

function getDatabase_() {

  var ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  if (!ss) {

    throw new Error(
      'Spreadsheet database tidak ditemukan.'
    );

  }

  return ss;

}


/* ============================================================
 * GET SHEET
 * ============================================================ */

function getSheet(sheetName) {

  var name =
    String(
      sheetName || ''
    ).trim();

  if (!name) {

    throw new Error(
      'Nama sheet tidak boleh kosong.'
    );

  }

  var ss =
    getDatabase_();

  var sheet =
    ss.getSheetByName(
      name
    );

  if (!sheet) {

    throw new Error(
      'Sheet "' +
      name +
      '" tidak ditemukan.'
    );

  }

  return sheet;

}


/* ============================================================
 * SAFE SHEET
 * ============================================================ */

function getSheetSafe_(sheetName) {

  var name =
    String(
      sheetName || ''
    ).trim();

  if (!name) {
    return null;
  }

  var ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  if (!ss) {
    return null;
  }

  return ss.getSheetByName(
    name
  );

}


/* ============================================================
 * JSON RESPONSE
 *
 * Ditaruh di Code.gs supaya SEMUA modul backend
 * mendapat helper yang sama.
 * ============================================================ */

function jsonResponse(
  success,
  message,
  data
) {

  if (
    arguments.length === 1 &&
    success &&
    typeof success === 'object' &&
    !Array.isArray(success)
  ) {

    return success;

  }

  var result = {
    success: Boolean(success)
  };

  if (
    message !== undefined &&
    message !== null
  ) {

    result.message =
      String(message);

  }

  if (
    data !== undefined &&
    data !== null
  ) {

    result.data =
      data;

  }

  return result;

}


/* ============================================================
 * HASH PASSWORD
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

  var digest =
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      password + salt,
      Utilities.Charset.UTF_8
    );

  return digest
    .map(
      function(byte) {

        var value =
          byte < 0
            ? byte + 256
            : byte;

        return (
          '0' +
          value.toString(16)
        ).slice(-2);

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

  password =
    String(
      password || ''
    );

  passwordHash =
    String(
      passwordHash || ''
    );

  salt =
    String(
      salt || ''
    );

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
    ) === passwordHash
  );

}


/* ============================================================
 * SEQUENTIAL ID
 *
 * Contoh:
 * USR-000001
 * ABS-000001
 * IZN-000001
 * ============================================================ */

function generateSequentialId(
  prefix,
  sheetName,
  idColumnIndex
) {

  prefix =
    String(
      prefix || 'ID'
    )
    .trim()
    .toUpperCase();

  var sheet =
    getSheet(
      sheetName
    );

  var lastRow =
    sheet.getLastRow();

  var maxNumber = 0;

  if (
    lastRow >= 2
  ) {

    var values =
      sheet
        .getRange(
          2,
          idColumnIndex + 1,
          lastRow - 1,
          1
        )
        .getDisplayValues();

    for (
      var i = 0;
      i < values.length;
      i++
    ) {

      var id =
        String(
          values[i][0] || ''
        ).trim();

      if (!id) {
        continue;
      }

      var match =
        id.match(
          /(\d+)$/
        );

      if (match) {

        var number =
          parseInt(
            match[1],
            10
          );

        if (
          number > maxNumber
        ) {
          maxNumber = number;
        }

      }

    }

  }

  return (
    prefix +
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
 * PENGATURAN CACHE
 * ============================================================ */

function getPengaturanCache() {

  var cache =
    CacheService
      .getScriptCache();

  var cached =
    cache.get(
      'ABSENSI_PENGATURAN'
    );

  if (cached) {

    try {

      return JSON.parse(
        cached
      );

    } catch (ignore) {}

  }

  var sheet =
    getSheetSafe_(
      SHEET_PENGATURAN
    );

  var result = {

    nama_app:
      'Aplikasi Absensi',

    logo_url:
      '',

    lat_kantor:
      0,

    long_kantor:
      0,

    radius_meter:
      100,

    jam_masuk:
      '08:00',

    jam_pulang:
      '17:00',

    toleransi_menit:
      15

  };

  if (
    !sheet ||
    sheet.getLastRow() < 2
  ) {

    cache.put(
      'ABSENSI_PENGATURAN',
      JSON.stringify(result),
      300
    );

    return result;

  }

  var row =
    sheet
      .getRange(
        2,
        1,
        1,
        Math.max(
          8,
          sheet.getLastColumn()
        )
      )
      .getValues()[0];

  result.nama_app =
    String(
      row[0] || result.nama_app
    );

  result.logo_url =
    String(
      row[1] || ''
    );

  result.lat_kantor =
    Number(
      row[2] || 0
    );

  result.long_kantor =
    Number(
      row[3] || 0
    );

  result.radius_meter =
    Number(
      row[4] || 100
    );

  result.jam_masuk =
    String(
      row[5] || '08:00'
    );

  result.jam_pulang =
    String(
      row[6] || '17:00'
    );

  result.toleransi_menit =
    Number(
      row[7] || 15
    );

  cache.put(
    'ABSENSI_PENGATURAN',
    JSON.stringify(result),
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
      'ABSENSI_PENGATURAN'
    );

  return true;

}


/* ============================================================
 * GOOGLE DRIVE FOLDER
 * ============================================================ */

function getOrCreateFolder_(
  folderName
) {

  folderName =
    String(
      folderName || ''
    ).trim();

  if (!folderName) {

    throw new Error(
      'Nama folder tidak boleh kosong.'
    );

  }

  var folders =
    DriveApp.getFoldersByName(
      folderName
    );

  if (
    folders.hasNext()
  ) {

    return folders.next();

  }

  return DriveApp.createFolder(
    folderName
  );

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
      : String(error);

  return HtmlService
    .createHtmlOutput(
      '<!DOCTYPE html>' +
      '<html lang="id">' +
      '<head>' +
      '<meta charset="UTF-8">' +
      '<meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<title>Gagal Memuat Aplikasi</title>' +
      '<style>' +
      'body{margin:0;padding:30px;font-family:Arial;background:#f5f7fb}' +
      '.box{max-width:720px;margin:40px auto;background:#fff;padding:30px;border-radius:16px}' +
      'h2{color:#d32f2f}' +
      'pre{background:#f1f1f1;padding:15px;white-space:pre-wrap;word-break:break-word}' +
      '</style>' +
      '</head>' +
      '<body>' +
      '<div class="box">' +
      '<h2>Gagal Memuat Aplikasi</h2>' +
      '<p>Halaman:</p>' +
      '<pre>' +
      escapeHtml_(
        page
      ) +
      '</pre>' +
      '<p>Error:</p>' +
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
 * CREATE SESSION COMPATIBILITY
 * ============================================================ */

function createSession_(
  idUser
) {

  if (
    typeof buatSessionToken_ ===
    'function'
  ) {

    return buatSessionToken_(
      idUser
    );

  }

  throw new Error(
    'Fungsi session belum tersedia.'
  );

}

function getWebAppUrl() {
  var url = ScriptApp.getService().getUrl();

  if (!url) {
    throw new Error(
      'URL Web App tidak tersedia. Pastikan project sudah di-deploy sebagai Web App.'
    );
  }

  return url;
}
/* ============================================================
 * TEST SYSTEM
 * ============================================================ */

function testSystem() {

  var result = {
    success: true,
    checks: {}
  };

  try {

    getSheet(
      SHEET_USERS
    );

    result.checks.Users =
      true;

  } catch (e) {

    result.success = false;

    result.checks.Users =
      e.message;

  }

  try {

    getSheet(
      SHEET_ABSENSI
    );

    result.checks.Absensi =
      true;

  } catch (e) {

    result.success = false;

    result.checks.Absensi =
      e.message;

  }

  try {

    getSheet(
      SHEET_IZIN
    );

    result.checks.Izin =
      true;

  } catch (e) {

    result.success = false;

    result.checks.Izin =
      e.message;

  }

  try {

    getSheet(
      SHEET_PENGATURAN
    );

    result.checks.Pengaturan =
      true;

  } catch (e) {

    result.success = false;

    result.checks.Pengaturan =
      e.message;

  }

  try {

    getSheet(
      SHEET_SESSIONS
    );

    result.checks.Sessions =
      true;

  } catch (e) {

    result.success = false;

    result.checks.Sessions =
      e.message;

  }

  return result;

}
