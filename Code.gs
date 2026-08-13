/**
 * ============================================================
 * CODE.GS
 * ROUTING UTAMA WEB APP
 * ============================================================
 *
 * FILE HTML YANG DIGUNAKAN:
 *
 * Login.html
 * DashboardAdmin.html
 * DashboardPegawai.html
 *
 * CATATAN KEAMANAN:
 *
 * parameter ?page= hanya menentukan halaman yang dirender.
 *
 * Keamanan dashboard TIDAK bergantung pada URL.
 *
 * DashboardAdmin melakukan:
 *   verifySession()
 *   role === admin
 *
 * DashboardPegawai melakukan:
 *   verifySession()
 *   role === pegawai
 *
 * ============================================================
 */


/**
 * ============================================================
 * WEB APP ENTRY
 * ============================================================
 */
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
        String(
          e.parameter.page
        )
        .trim()
        .replace(
          /[^a-zA-Z0-9_-]/g,
          ''
        );

    }


    /*
     * Halaman yang boleh dirender.
     */
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


    return renderPage_(page);

  } catch (error) {

    return renderErrorPage_(
      page,
      error
    );

  }

}


/**
 * ============================================================
 * RENDER PAGE
 * ============================================================
 */
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


/**
 * ============================================================
 * INCLUDE HTML
 * ============================================================
 *
 * Contoh:
 *
 * <?!= include('CSS'); ?>
 * <?!= include('JS'); ?>
 * <?!= include('AdminCSS'); ?>
 * <?!= include('AdminJS'); ?>
 *
 * ============================================================
 */
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
      '" tidak ditemukan. ' +
      'Pastikan file "' +
      name +
      '.html" ada di project Apps Script. ' +
      'Error asli: ' +
      error.message
    );

  }

}


/**
 * ============================================================
 * NAMA APLIKASI
 * ============================================================
 */
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
        'Pengaturan'
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


    return (
      value ||
      defaultName
    );

  } catch (error) {

    return defaultName;

  }

}


/**
 * ============================================================
 * ERROR PAGE
 * ============================================================
 */
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

      '<meta name="viewport" ' +
      'content="width=device-width,initial-scale=1">' +

      '<title>Gagal Memuat Aplikasi</title>' +

      '<style>' +

      'body{' +
      'margin:0;' +
      'padding:30px;' +
      'font-family:Arial,sans-serif;' +
      'background:#f5f7fb;' +
      'color:#222;' +
      '}' +

      '.box{' +
      'max-width:720px;' +
      'margin:40px auto;' +
      'background:#fff;' +
      'padding:30px;' +
      'border-radius:16px;' +
      'box-shadow:0 10px 30px rgba(0,0,0,.08);' +
      '}' +

      'h2{' +
      'color:#d32f2f;' +
      'margin-top:0;' +
      '}' +

      'pre{' +
      'background:#f1f1f1;' +
      'padding:15px;' +
      'border-radius:8px;' +
      'white-space:pre-wrap;' +
      'word-break:break-word;' +
      '}' +

      '</style>' +

      '</head>' +

      '<body>' +

      '<div class="box">' +

      '<h2>Gagal Memuat Aplikasi</h2>' +

      '<p>Halaman yang dicoba:</p>' +

      '<pre>' +
      escapeHtml_(page) +
      '</pre>' +

      '<p>Error asli:</p>' +

      '<pre>' +
      escapeHtml_(message) +
      '</pre>' +

      '</div>' +

      '</body>' +

      '</html>'
    )
    .setTitle(
      'Gagal Memuat Aplikasi'
    );

}


/**
 * ============================================================
 * ESCAPE HTML
 * ============================================================
 */
function escapeHtml_(value) {

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


/**
 * ============================================================
 * TEST LOGIN
 * ============================================================
 */
function testFileLogin() {

  try {

    HtmlService
      .createTemplateFromFile(
        'Login'
      );


    Logger.log(
      'BERHASIL: Login.html ditemukan.'
    );


    return {
      success: true,
      message:
        'Login.html berhasil ditemukan.'
    };

  } catch (error) {

    Logger.log(
      'GAGAL: ' +
      error.message
    );


    return {
      success: false,
      message:
        error.message
    };

  }

}


/**
 * ============================================================
 * TEST DASHBOARD ADMIN
 * ============================================================
 */
function testFileDashboardAdmin() {

  try {

    HtmlService
      .createTemplateFromFile(
        'DashboardAdmin'
      );


    Logger.log(
      'BERHASIL: DashboardAdmin.html ditemukan.'
    );


    return {
      success: true,
      message:
        'DashboardAdmin.html berhasil ditemukan.'
    };

  } catch (error) {

    Logger.log(
      'GAGAL: ' +
      error.message
    );


    return {
      success: false,
      message:
        error.message
    };

  }

}


/**
 * ============================================================
 * TEST DASHBOARD PEGAWAI
 * ============================================================
 */
function testFileDashboardPegawai() {

  try {

    HtmlService
      .createTemplateFromFile(
        'DashboardPegawai'
      );


    Logger.log(
      'BERHASIL: DashboardPegawai.html ditemukan.'
    );


    return {
      success: true,
      message:
        'DashboardPegawai.html berhasil ditemukan.'
    };

  } catch (error) {

    Logger.log(
      'GAGAL: ' +
      error.message
    );


    return {
      success: false,
      message:
        error.message
    };

  }

}

/**
 * ============================================================
 * DATABASE SHEET HELPER
 * ============================================================
 *
 * Semua modul backend menggunakan helper ini:
 *
 * - Auth.gs
 * - Absensi.gs
 * - Izin.gs
 * - Admin.gs
 *
 * Nama sheet mengikuti struktur database:
 *
 * Users
 * Absensi
 * Izin
 * Pengaturan
 * Sessions
 */

/**
 * Nama sheet database.
 */
var SHEET_USERS = 'Users';
var SHEET_ABSENSI = 'Absensi';
var SHEET_IZIN = 'Izin';
var SHEET_PENGATURAN = 'Pengaturan';
var SHEET_SESSIONS = 'Sessions';


/**
 * Durasi session dalam jam.
 */
var SESSION_DURASI_JAM = 12;


/**
 * ============================================================
 * GET SPREADSHEET
 * ============================================================
 */
function getDatabase_() {

  var ss =
    SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {

    throw new Error(
      'Spreadsheet database tidak ditemukan.'
    );

  }

  return ss;

}


/**
 * ============================================================
 * GET SHEET
 * ============================================================
 *
 * Contoh:
 *
 * var sheet = getSheet('Users');
 *
 * Jika sheet tidak ada, fungsi akan memberikan error
 * yang jelas daripada:
 *
 * TypeError: Cannot read properties of null
 *
 */
function getSheet(
  sheetName
) {

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


/**
 * ============================================================
 * GET SHEET AMAN
 * ============================================================
 *
 * Digunakan jika pemanggil memang ingin menerima null
 * ketika sheet belum tersedia.
 */
function getSheetSafe_(
  sheetName
) {

  var name =
    String(
      sheetName || ''
    ).trim();


  if (!name) {
    return null;
  }


  var ss =
    SpreadsheetApp.getActiveSpreadsheet();


  if (!ss) {
    return null;
  }


  return ss.getSheetByName(
    name
  );

}


/**
 * ============================================================
 * HASH PASSWORD
 * ============================================================
 */
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
      function (byte) {

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


/**
 * ============================================================
 * VERIFY PASSWORD
 * ============================================================
 */
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


  var calculatedHash =
    hashPassword(
      password,
      salt
    );


  return calculatedHash ===
    passwordHash;

}


/**
 * ============================================================
 * CREATE SESSION COMPATIBILITY
 * ============================================================
 *
 * Auth.gs memanggil:
 *
 * createSession_()
 *
 * Implementasi utama berada pada:
 *
 * buatSessionToken_()
 */
function createSession_(
  idUser
) {

  return buatSessionToken_(
    idUser
  );

}
