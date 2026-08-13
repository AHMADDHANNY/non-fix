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
 */
function renderPage_(page) {

  /*
   * DashboardAdmin adalah FRAGMENT.
   *
   * Jadi jangan memakai:
   *
   * createTemplateFromFile('DashboardAdmin')
   *
   * secara langsung.
   *
   * Kita bungkus dengan HTML lengkap,
   * CSS, AdminCSS dan AdminJS.
   */
  if (
    page === 'DashboardAdmin'
  ) {

    return renderDashboardAdmin_();

  }


  /*
   * Login dan DashboardPegawai
   * adalah halaman HTML lengkap.
   */
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
 * RENDER DASHBOARD ADMIN
 * ============================================================
 */
function renderDashboardAdmin_() {

  var html =
    '<!DOCTYPE html>' +

    '<html lang="id">' +

    '<head>' +

    '<base target="_top">' +

    '<meta charset="UTF-8">' +

    '<meta name="viewport" ' +
    'content="width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover">' +

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

    /*
     * AdminJS mendefinisikan:
     *
     * window.AdminApp
     *
     * tetapi tidak otomatis menjalankan init().
     */
    '<script>' +

    'document.addEventListener("DOMContentLoaded", function () {' +

    '  if (' +
    '    window.AdminApp && ' +
    '    typeof window.AdminApp.init === "function"' +
    '  ) {' +

    '    window.AdminApp.init();' +

    '  }' +

    '});' +

    '</script>' +

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
 * INCLUDE HTML
 * ============================================================
 */
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
function getSheet(
  sheetName
) {

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
 * GET SHEET SAFE
 * ============================================================
 */
function getSheetSafe_(
  sheetName
) {

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
 * NAMA APLIKASI - PRIVATE
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
 * NAMA APLIKASI - PUBLIC
 *
 * JS.html tidak memanggil fungsi yang berakhiran "_".
 * ============================================================
 */
function getNamaApp() {

  return getNamaAppSafe_();

}


/**
 * ============================================================
 * DASHBOARD PEGAWAI
 * ============================================================
 *
 * Endpoint ini dipakai oleh:
 *
 * DashboardPegawai.html
 *
 * Endpoint sebelumnya belum tersedia.
 */
function getDashboardPegawai(
  token
) {

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


    /*
     * Gunakan endpoint absensi
     * yang sudah tersedia.
     */
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
 * TEST FILE LOGIN
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


/**
 * ============================================================
 * TEST FILE DASHBOARD ADMIN
 * ============================================================
 */
function testFileDashboardAdmin() {

  try {

    HtmlService
      .createHtmlOutputFromFile(
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


/**
 * ============================================================
 * TEST FILE DASHBOARD PEGAWAI
 * ============================================================
 */
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

function getWebAppUrl() {
  var url = ScriptApp.getService().getUrl();

  if (!url) {
    throw new Error(
      'URL Web App belum tersedia. Deploy project sebagai Web App terlebih dahulu.'
    );
  }

  return url;
}
