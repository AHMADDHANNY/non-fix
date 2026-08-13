/**
 * ============================================================
 * CODE.GS
 * ROUTING UTAMA + HELPER DATABASE
 * ============================================================
 */

var SHEET_USERS = 'Users';
var SHEET_ABSENSI = 'Absensi';
var SHEET_IZIN = 'Izin';
var SHEET_PENGATURAN = 'Pengaturan';
var SHEET_SESSIONS = 'Sessions';

var SESSION_DURASI_JAM = 12;


/**
 * ============================================================
 * WEB APP
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
 *
 * SEMUA halaman HTML dirender sebagai TEMPLATE.
 *
 * Jangan menggunakan createHtmlOutputFromFile()
 * untuk halaman utama karena halaman seperti Login
 * dan DashboardAdmin menggunakan <?!= include(...) ?>.
 */
function renderPage_(page) {

  var template =
    HtmlService
      .createTemplateFromFile(
        page
      );


  /*
   * Nilai ini tersedia di HTML:
   *
   * window.ABSENSI_WEB_APP_URL
   */
  template.webAppUrl =
    getWebAppUrl();


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
 * INCLUDE
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
      'Pastikan file tersebut ada di project Apps Script. ' +
      'Error asli: ' +
      error.message
    );

  }

}


/**
 * ============================================================
 * DATABASE
 * ============================================================
 */
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


/**
 * ============================================================
 * GET SHEET
 * ============================================================
 */
function getSheet(sheetName) {

  var name =
    String(
      sheetName || ''
    )
    .trim();


  if (!name) {

    throw new Error(
      'Nama sheet tidak boleh kosong.'
    );

  }


  var sheet =
    getDatabase_()
      .getSheetByName(
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
 * GET SHEET SAFE
 * ============================================================
 */
function getSheetSafe_(sheetName) {

  var name =
    String(
      sheetName || ''
    )
    .trim();


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


/**
 * ============================================================
 * NAMA APLIKASI
 * ============================================================
 */
function getNamaAppSafe_() {

  var defaultName =
    'Aplikasi Absensi';


  try {

    var sheet =
      getSheetSafe_(
        SHEET_PENGATURAN
      );


    if (
      !sheet ||
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
 * PUBLIC.
 *
 * google.script.run TIDAK boleh memanggil
 * getNamaAppSafe_() karena underscore dianggap private.
 */
function getNamaApp() {

  return getNamaAppSafe_();

}


/**
 * ============================================================
 * URL WEB APP
 * ============================================================
 */
function getWebAppUrl() {

  var url =
    ScriptApp
      .getService()
      .getUrl();


  if (!url) {

    throw new Error(
      'URL Web App tidak tersedia. ' +
      'Pastikan project sudah di-deploy sebagai Web App.'
    );

  }


  return url;

}


/**
 * ============================================================
 * TEMPLATE ERROR
 * ============================================================
 */
function renderErrorPage_(page, error) {

  return HtmlService
    .createHtmlOutput(
      '<!DOCTYPE html>' +
      '<html lang="id">' +
      '<head>' +
      '<meta charset="UTF-8">' +
      '<meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<title>Gagal memuat aplikasi</title>' +
      '<style>' +
      'body{' +
      'font-family:Arial,sans-serif;' +
      'background:#f5f7fb;' +
      'padding:30px;' +
      '}' +
      '.box{' +
      'max-width:800px;' +
      'margin:40px auto;' +
      'background:#fff;' +
      'padding:25px;' +
      'border-radius:14px;' +
      'box-shadow:0 8px 30px rgba(0,0,0,.08);' +
      '}' +
      'h2{color:#d11a2a;}' +
      'pre{' +
      'background:#f1f3f5;' +
      'padding:15px;' +
      'white-space:pre-wrap;' +
      'word-break:break-word;' +
      'border-radius:8px;' +
      '}' +
      '</style>' +
      '</head>' +
      '<body>' +
      '<div class="box">' +
      '<h2>Gagal memuat aplikasi</h2>' +
      '<p>Halaman:</p>' +
      '<pre>' +
      escapeHtml_(
        page
      ) +
      '</pre>' +
      '<p>Error:</p>' +
      '<pre>' +
      escapeHtml_(
        String(error)
      ) +
      '</pre>' +
      '</div>' +
      '</body>' +
      '</html>'
    )
    .setTitle(
      'Gagal memuat aplikasi'
    )
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/**
 * ============================================================
 * DASHBOARD PEGAWAI
 * ============================================================
 */
function getDashboardPegawai(token) {

  try {

    var session =
      verifySession(
        token
      );


    if (
      !session ||
      session.valid !== true
    ) {

      throw new Error(
        session &&
        session.message
          ? session.message
          : 'Sesi tidak valid.'
      );

    }


    var user =
      session.user ||
      {};


    var role =
      String(
        user.role || ''
      )
      .trim()
      .toLowerCase();


    if (
      role !== 'pegawai'
    ) {

      throw new Error(
        'Akses hanya untuk pegawai.'
      );

    }


    var status =
      getStatusAbsensiHariIni(
        token
      );


    if (
      status &&
      status.success === true
    ) {

      return status;

    }


    return {
      success: true,
      data: {
        tanggal: '',
        status: '',
        jam_masuk: '',
        jam_pulang: ''
      }
    };


  } catch (error) {

    return {
      success: false,
      message:
        error &&
        error.message
          ? error.message
          : String(error)
    };

  }

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
