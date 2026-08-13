/**
 * ============================================================
 * CODE.GS
 * ROUTING + HELPER UTAMA
 * ============================================================
 */

var SHEET_USERS = 'Users';
var SHEET_ABSENSI = 'Absensi';
var SHEET_IZIN = 'Izin';
var SHEET_PENGATURAN = 'Pengaturan';
var SHEET_SESSIONS = 'Sessions';

/*
 * Durasi session: 12 jam
 */
var SESSION_DURASI_JAM = 12;


/**
 * ============================================================
 * WEB APP
 * ============================================================
 */
function doGet(e) {

  var page = 'Login';

  try {

    if (
      e &&
      e.parameter &&
      e.parameter.page
    ) {

      var requestedPage =
        String(e.parameter.page)
          .trim()
          .replace(
            /[^a-zA-Z0-9_-]/g,
            ''
          );

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
 * DashboardAdmin.html adalah FRAGMENT.
 * Karena itu harus dibungkus CSS + AdminCSS + AdminJS.
 */
function renderPage_(page) {

  if (
    page ===
    'DashboardAdmin'
  ) {

    return renderAdminPage_();

  }


  /*
   * DashboardPegawai.html saat ini
   * merupakan halaman HTML lengkap.
   */
  if (
    page ===
    'DashboardPegawai'
  ) {

    return HtmlService
      .createTemplateFromFile(
        'DashboardPegawai'
      )
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


  /*
   * Login.html
   */
  return HtmlService
    .createTemplateFromFile(
      'Login'
    )
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
 * RENDER DASHBOARD ADMIN
 * ============================================================
 */
function renderAdminPage_() {

  var html =
    '<!DOCTYPE html>' +
    '<html lang="id">' +
    '<head>' +
    '<meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover">' +
    '<title>' +
    escapeHtml_(
      getNamaAppSafe_()
    ) +
    '</title>' +

    include('CSS') +
    include('AdminCSS') +

    '</head>' +
    '<body>' +

    include(
      'DashboardAdmin'
    ) +

    include(
      'AdminJS'
    ) +

    '</body>' +
    '</html>';


  return HtmlService
    .createHtmlOutput(
      html
    )
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
      error.message
    );

  }

}


/**
 * ============================================================
 * GET SHEET
 * ============================================================
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
      'Nama sheet kosong.'
    );

  }


  var ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  if (!ss) {

    throw new Error(
      'Spreadsheet aktif tidak ditemukan.'
    );

  }


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
        SHEET_PENGATURAN
      );


    if (!sheet) {

      return defaultName;

    }


    if (
      sheet.getLastRow() <
      2
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
      '<meta name="viewport" content="width=device-width,initial-scale=1">' +
      '<title>Gagal Memuat Aplikasi</title>' +
      '<style>' +
      'body{margin:0;padding:30px;font-family:Arial;background:#f5f7fb}' +
      '.box{max-width:720px;margin:40px auto;background:#fff;padding:30px;border-radius:16px;box-shadow:0 10px 30px rgba(0,0,0,.08)}' +
      'h2{color:#d32f2f}' +
      'pre{background:#f1f1f1;padding:15px;border-radius:8px;white-space:pre-wrap;word-break:break-word}' +
      '</style>' +
      '</head>' +
      '<body>' +
      '<div class="box">' +
      '<h2>Gagal Memuat Aplikasi</h2>' +
      '<p>Halaman:</p>' +
      '<pre>' +
      escapeHtml_(page) +
      '</pre>' +
      '<p>Error:</p>' +
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


/**
 * ============================================================
 * TEST
 * ============================================================
 */
function testFileLogin() {

  try {

    HtmlService
      .createTemplateFromFile(
        'Login'
      );

    return {
      success: true,
      message:
        'Login.html ditemukan.'
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
      .createHtmlOutputFromFile(
        'DashboardAdmin'
      );

    return {
      success: true,
      message:
        'DashboardAdmin.html ditemukan.'
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
        'DashboardPegawai.html ditemukan.'
    };

  } catch (error) {

    return {
      success: false,
      message:
        error.message
    };

  }

}
