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
