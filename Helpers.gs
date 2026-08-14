/**
 * ===================================================================
 * HELPERS.GS
 * Shared helpers used by Absensi.gs, Izin.gs and Admin.gs.
 * ===================================================================
 */

var FOLDER_ABSENSI_NAME = 'Foto_Absensi';
var FOLDER_IZIN_NAME = 'Lampiran_Izin';
var FOLDER_REKAP_NAME = 'Rekap_Absensi';

var CACHE_PENGATURAN_KEY = 'absensi_pengaturan_v1';
var CACHE_PENGATURAN_TTL = 300;


/* ===================================================================
 * PENGATURAN
 * =================================================================== */

function getPengaturanCache() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(CACHE_PENGATURAN_KEY);

  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (ignore) {
      cache.remove(CACHE_PENGATURAN_KEY);
    }
  }

  var sheet = getSheet(SHEET_PENGATURAN);

  if (sheet.getLastRow() < 2) {
    throw new Error('Pengaturan aplikasi belum dikonfigurasi.');
  }

  var values = sheet.getRange(2, 1, 1, 8).getDisplayValues()[0];

  var result = {
    nama_app: String(values[0] || '').trim(),
    logo_url: String(values[1] || '').trim(),
    lat_kantor: String(values[2] || '').trim(),
    long_kantor: String(values[3] || '').trim(),
    radius_meter: String(values[4] || '').trim(),
    jam_masuk: String(values[5] || '').trim(),
    jam_pulang: String(values[6] || '').trim(),
    toleransi_menit: String(values[7] || '').trim()
  };

  cache.put(
    CACHE_PENGATURAN_KEY,
    JSON.stringify(result),
    CACHE_PENGATURAN_TTL
  );

  return result;
}


function bersihkanCachePengaturan() {
  CacheService.getScriptCache().remove(CACHE_PENGATURAN_KEY);
}


/* ===================================================================
 * ID SEQUENTIAL
 * =================================================================== */

/**
 * Generates IDs such as ABS-000001.
 *
 * The previous implementation scanned the sheet and returned max+1.
 * That is unsafe when several IDs are generated before their rows are
 * appended (for example tandaiAlpaHarian), because every call can see
 * the same max and return the same ID.
 *
 * A script-property counter is therefore used as the authoritative
 * sequence after it has been synchronized with existing data.
 * Calls made inside an existing ScriptLock remain serialized.
 */
function generateSequentialId(prefix, sheetName, columnIndex) {
  prefix = String(prefix || '').trim().toUpperCase();

  if (!prefix) {
    throw new Error('Prefix ID tidak boleh kosong.');
  }

  var sheet = getSheet(sheetName);
  var properties = PropertiesService.getScriptProperties();
  var key = 'SEQ_' + prefix + '_' + String(sheetName).replace(/[^A-Za-z0-9_]/g, '_').toUpperCase();
  var stored = parseInt(properties.getProperty(key) || '0', 10);
  if (!isFinite(stored) || stored < 0) {
    stored = 0;
  }

  /* Sinkronkan sekali/ketika counter tertinggal dari data yang sudah ada. */
  var lastRow = sheet.getLastRow();
  var maxNumber = stored;

  if (lastRow >= 2) {
    var values = sheet
      .getRange(2, Number(columnIndex) + 1, lastRow - 1, 1)
      .getDisplayValues();

    var pattern = new RegExp(
      '^' + escapeRegExp_(prefix) + '[-_]?(\\d+)$',
      'i'
    );

    for (var i = 0; i < values.length; i++) {
      var text = String(values[i][0] || '').trim();
      var match = text.match(pattern);

      if (!match) {
        continue;
      }

      var number = parseInt(match[1], 10);
      if (isFinite(number) && number > maxNumber) {
        maxNumber = number;
      }
    }
  }

  var nextNumber = maxNumber + 1;
  properties.setProperty(key, String(nextNumber));

  return prefix + '-' + padNumber_(nextNumber, 6);
}


function escapeRegExp_(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}


function padNumber_(value, width) {
  var text = String(value);

  while (text.length < width) {
    text = '0' + text;
  }

  return text;
}


/* ===================================================================
 * GEOLOCATION
 * =================================================================== */

function hitungJarakMeter(lat1, lon1, lat2, lon2) {
  var a = Number(lat1);
  var b = Number(lon1);
  var c = Number(lat2);
  var d = Number(lon2);

  if (!isFinite(a) || !isFinite(b) || !isFinite(c) || !isFinite(d)) {
    throw new Error('Koordinat untuk perhitungan jarak tidak valid.');
  }

  if (a < -90 || a > 90 || c < -90 || c > 90 || b < -180 || b > 180 || d < -180 || d > 180) {
    throw new Error('Koordinat untuk perhitungan jarak berada di luar batas.');
  }

  var rad = Math.PI / 180;
  var dLat = (c - a) * rad;
  var dLon = (d - b) * rad;
  var sinLat = Math.sin(dLat / 2);
  var sinLon = Math.sin(dLon / 2);

  var h =
    sinLat * sinLat +
    Math.cos(a * rad) *
    Math.cos(c * rad) *
    sinLon * sinLon;

  h = Math.max(0, Math.min(1, h));

  return 6371000 * 2 * Math.atan2(
    Math.sqrt(h),
    Math.sqrt(1 - h)
  );
}


/* ===================================================================
 * DRIVE FOLDERS
 * =================================================================== */

function getOrCreateFolder_(folderName) {
  var name = String(folderName || '').trim();

  if (!name) {
    throw new Error('Nama folder Drive tidak boleh kosong.');
  }

  var properties = PropertiesService.getScriptProperties();
  var key = 'FOLDER_ID_' + name.replace(/[^A-Za-z0-9_]/g, '_').toUpperCase();
  var savedId = properties.getProperty(key);

  if (savedId) {
    try {
      var savedFolder = DriveApp.getFolderById(savedId);

      if (savedFolder && savedFolder.getName() === name) {
        return savedFolder;
      }
    } catch (ignore) {
      properties.deleteProperty(key);
    }
  }

  var folders = DriveApp.getFoldersByName(name);

  if (folders.hasNext()) {
    var existing = folders.next();
    properties.setProperty(key, existing.getId());
    return existing;
  }

  var created = DriveApp.createFolder(name);
  properties.setProperty(key, created.getId());

  return created;
}
