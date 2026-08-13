/**
 * ===================================================================
 * ABSENSI.GS
 * Modul absen masuk & absen pulang, termasuk validasi GPS (radius
 * kantor), upload foto selfie ke Drive, dan penentuan status
 * (Hadir/Terlambat/Alpa) berdasarkan Pengaturan.
 *
 * Sudah disinkronkan dengan Code.gs asli (repo non-fix):
 * - Kolom sheet "Absensi" (setupAwal di Code.gs):
 *   0 id_absen, 1 id_user, 2 tanggal, 3 jam_masuk, 4 foto_masuk,
 *   5 lat_masuk, 6 long_masuk, 7 jam_pulang, 8 foto_pulang,
 *   9 lat_pulang, 10 long_pulang, 11 status,
 *   12 jarak_masuk_meter, 13 jarak_pulang_meter
 * - Folder Drive: FOLDER_ABSENSI_NAME ('Foto_Absensi')
 * - generateSequentialId(prefix, namaSheet, kolomId)
 * - getPengaturanCache() -> { nama_app, logo_url, lat_kantor,
 *   long_kantor, radius_meter, jam_masuk, jam_pulang, toleransi_menit }
 *   dengan jam_masuk/jam_pulang format "HH:mm"
 * ===================================================================
 */

var KOLOM_ABSENSI = {
  ID: 0, ID_USER: 1, TANGGAL: 2, JAM_MASUK: 3, FOTO_MASUK: 4,
  LAT_MASUK: 5, LONG_MASUK: 6, JAM_PULANG: 7, FOTO_PULANG: 8,
  LAT_PULANG: 9, LONG_PULANG: 10, STATUS: 11,
  JARAK_MASUK: 12, JARAK_PULANG: 13
};

/**
 * Absen masuk.
 * @param {string} token
 * @param {number} lat
 * @param {number} long
 * @param {string} fotoBase64 - data URL / base64 hasil kompres di client (canvas)
 */
function absenMasuk(token, lat, long, fotoBase64) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // tunggu maks 15 detik kalau ada user lain sedang menulis

    var session = verifySession(token);
    if (!session.valid) throw new Error(session.message || 'Sesi tidak valid.');
    var user = session.user;

    // --- Validasi ulang di backend (jangan percaya client 100%) ---
    validasiKoordinat_(lat, long);
    if (!fotoBase64) throw new Error('Foto selfie wajib diambil.');

    var pengaturan = getPengaturanCache();
    var jarak = hitungJarakMeter(
      parseFloat(lat), parseFloat(long),
      parseFloat(pengaturan.lat_kantor), parseFloat(pengaturan.long_kantor)
    );
    if (jarak > parseFloat(pengaturan.radius_meter)) {
      throw new Error(
        'Anda berada di luar radius kantor (' + Math.round(jarak) +
        ' m dari kantor, maksimal ' + pengaturan.radius_meter + ' m). Absen ditolak.'
      );
    }

    var now = new Date();
    var tanggalStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var jamStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');

    var sheet = getSheet(SHEET_ABSENSI);
    var data = sheet.getDataRange().getValues();

    // Cek sudah absen masuk hari ini?
    for (var i = 1; i < data.length; i++) {
      if (data[i][KOLOM_ABSENSI.ID_USER] === user.id_user &&
          data[i][KOLOM_ABSENSI.TANGGAL] === tanggalStr) {
        throw new Error('Anda sudah melakukan absen masuk hari ini.');
      }
    }

    // Upload foto ke Drive
    var urlFoto = simpanFotoAbsen_(fotoBase64, user.id_user + '_masuk_' + tanggalStr);

    // Tentukan status: Hadir / Terlambat
    var status = tentukanStatusMasuk_(jamStr, pengaturan);

    var idBaru = generateSequentialId('ABS', SHEET_ABSENSI, KOLOM_ABSENSI.ID);
    var rowBaru = [];
    rowBaru[KOLOM_ABSENSI.ID] = idBaru;
    rowBaru[KOLOM_ABSENSI.ID_USER] = user.id_user;
    rowBaru[KOLOM_ABSENSI.TANGGAL] = tanggalStr;
    rowBaru[KOLOM_ABSENSI.JAM_MASUK] = jamStr;
    rowBaru[KOLOM_ABSENSI.FOTO_MASUK] = urlFoto;
    rowBaru[KOLOM_ABSENSI.LAT_MASUK] = lat;
    rowBaru[KOLOM_ABSENSI.LONG_MASUK] = long;
    rowBaru[KOLOM_ABSENSI.JAM_PULANG] = '';
    rowBaru[KOLOM_ABSENSI.FOTO_PULANG] = '';
    rowBaru[KOLOM_ABSENSI.LAT_PULANG] = '';
    rowBaru[KOLOM_ABSENSI.LONG_PULANG] = '';
    rowBaru[KOLOM_ABSENSI.STATUS] = status;
    rowBaru[KOLOM_ABSENSI.JARAK_MASUK] = Math.round(jarak);
    rowBaru[KOLOM_ABSENSI.JARAK_PULANG] = '';

    sheet.appendRow(rowBaru);

    return jsonResponse({
      success: true,
      message: 'Absen masuk berhasil (' + status + ').',
      data: { jam_masuk: jamStr, status: status }
    });

  } catch (e) {
    return jsonResponse({ success: false, message: e.message });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Absen pulang.
 */
function absenPulang(token, lat, long, fotoBase64) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    var session = verifySession(token);
    if (!session.valid) throw new Error(session.message || 'Sesi tidak valid.');
    var user = session.user;

    validasiKoordinat_(lat, long);
    if (!fotoBase64) throw new Error('Foto selfie wajib diambil.');

    var pengaturan = getPengaturanCache();
    var jarak = hitungJarakMeter(
      parseFloat(lat), parseFloat(long),
      parseFloat(pengaturan.lat_kantor), parseFloat(pengaturan.long_kantor)
    );
    if (jarak > parseFloat(pengaturan.radius_meter)) {
      throw new Error(
        'Anda berada di luar radius kantor (' + Math.round(jarak) +
        ' m dari kantor, maksimal ' + pengaturan.radius_meter + ' m). Absen ditolak.'
      );
    }

    var now = new Date();
    var tanggalStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var jamStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');

    var sheet = getSheet(SHEET_ABSENSI);
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < data.length; i++) {
      if (data[i][KOLOM_ABSENSI.ID_USER] === user.id_user &&
          data[i][KOLOM_ABSENSI.TANGGAL] === tanggalStr) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error('Anda belum melakukan absen masuk hari ini.');
    }
    if (data[rowIndex][KOLOM_ABSENSI.JAM_PULANG]) {
      throw new Error('Anda sudah melakukan absen pulang hari ini.');
    }

    var urlFoto = simpanFotoAbsen_(fotoBase64, user.id_user + '_pulang_' + tanggalStr);

    var statusLama = data[rowIndex][KOLOM_ABSENSI.STATUS];
    var statusBaru = tentukanStatusPulang_(jamStr, pengaturan, statusLama);

    var row = rowIndex + 1; // 1-indexed di Sheets
    sheet.getRange(row, KOLOM_ABSENSI.JAM_PULANG + 1).setValue(jamStr);
    sheet.getRange(row, KOLOM_ABSENSI.FOTO_PULANG + 1).setValue(urlFoto);
    sheet.getRange(row, KOLOM_ABSENSI.LAT_PULANG + 1).setValue(lat);
    sheet.getRange(row, KOLOM_ABSENSI.LONG_PULANG + 1).setValue(long);
    sheet.getRange(row, KOLOM_ABSENSI.STATUS + 1).setValue(statusBaru);
    sheet.getRange(row, KOLOM_ABSENSI.JARAK_PULANG + 1).setValue(Math.round(jarak));

    return jsonResponse({
      success: true,
      message: 'Absen pulang berhasil (' + statusBaru + ').',
      data: { jam_pulang: jamStr, status: statusBaru }
    });

  } catch (e) {
    return jsonResponse({ success: false, message: e.message });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Riwayat absensi pribadi milik user yang login, dengan paginasi.
 * @param {string} token
 * @param {number} page      - mulai dari 1
 * @param {number} pageSize  - default 20
 * @param {string} bulanFilter - opsional, format "yyyy-MM"
 */
function getHistoryAbsensi(token, page, pageSize, bulanFilter) {
  try {
    var session = verifySession(token);
    if (!session.valid) throw new Error(session.message || 'Sesi tidak valid.');

    page = page && page > 0 ? page : 1;
    pageSize = pageSize && pageSize > 0 ? pageSize : 20;

    var sheet = getSheet(SHEET_ABSENSI);
    var data = sheet.getDataRange().getValues();

    var hasil = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[KOLOM_ABSENSI.ID_USER] !== session.user.id_user) continue;
      if (bulanFilter && String(row[KOLOM_ABSENSI.TANGGAL]).indexOf(bulanFilter) !== 0) continue;

      hasil.push({
        id: row[KOLOM_ABSENSI.ID],
        tanggal: row[KOLOM_ABSENSI.TANGGAL],
        jam_masuk: row[KOLOM_ABSENSI.JAM_MASUK],
        jam_pulang: row[KOLOM_ABSENSI.JAM_PULANG],
        status: row[KOLOM_ABSENSI.STATUS],
        lokasi_masuk: { lat: row[KOLOM_ABSENSI.LAT_MASUK], long: row[KOLOM_ABSENSI.LONG_MASUK] },
        lokasi_pulang: { lat: row[KOLOM_ABSENSI.LAT_PULANG], long: row[KOLOM_ABSENSI.LONG_PULANG] }
      });
    }

    // Terbaru dulu
    hasil.sort(function (a, b) { return a.tanggal < b.tanggal ? 1 : -1; });

    var total = hasil.length;
    var start = (page - 1) * pageSize;
    var potongan = hasil.slice(start, start + pageSize);

    return jsonResponse({
      success: true,
      data: potongan,
      pagination: { page: page, pageSize: pageSize, total: total, totalPage: Math.ceil(total / pageSize) }
    });
  } catch (e) {
    return jsonResponse({ success: false, message: e.message });
  }
}

/* ===================================================================
 * HELPER INTERNAL
 * =================================================================== */

function validasiKoordinat_(lat, long) {
  if (lat === undefined || long === undefined || lat === null || long === null || lat === '' || long === '') {
    throw new Error('Lokasi GPS tidak terdeteksi. Aktifkan GPS dan coba lagi.');
  }
  var latNum = parseFloat(lat), longNum = parseFloat(long);
  if (isNaN(latNum) || isNaN(longNum) || latNum < -90 || latNum > 90 || longNum < -180 || longNum > 180) {
    throw new Error('Koordinat GPS tidak valid.');
  }
}

/**
 * Simpan foto base64 ke folder Drive khusus foto absen, return URL.
 * Batas ukuran dicek di sini juga (jaga-jaga kalau kompresi client gagal).
 */
function simpanFotoAbsen_(fotoBase64, namaFileTanpaExt) {
  var MAKS_BYTES = 2 * 1024 * 1024; // 2MB, sesuaikan kalau perlu

  var bagian = fotoBase64.split(',');
  var meta = bagian[0]; // contoh: "data:image/jpeg;base64"
  var base64Data = bagian.length > 1 ? bagian[1] : bagian[0];

  var contentType = 'image/jpeg';
  var match = meta.match(/data:(image\/\w+);base64/);
  if (match) contentType = match[1];

  var bytes = Utilities.base64Decode(base64Data);
  if (bytes.length > MAKS_BYTES) {
    throw new Error('Ukuran foto terlalu besar. Maksimal 2MB.');
  }

  var ekstensi = contentType.split('/')[1] || 'jpg';
  var blob = Utilities.newBlob(bytes, contentType, namaFileTanpaExt + '.' + ekstensi);

  var folder = getOrCreateFolder_(FOLDER_ABSENSI_NAME);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return 'https://drive.google.com/uc?id=' + file.getId();
}

/**
 * Bandingkan jam masuk aktual dengan jam_masuk + toleransi_menit
 * dari Pengaturan -> "Hadir" atau "Terlambat".
 */
function tentukanStatusMasuk_(jamStr, pengaturan) {
  var batasToleransi = tambahMenitKeJam_(pengaturan.jam_masuk, pengaturan.toleransi_menit);
  return jamStr <= batasToleransi ? 'Hadir' : 'Terlambat';
}

/**
 * Kalau pulang sebelum jam_pulang resmi -> tandai "Pulang Cepat",
 * tapi tidak menimpa status "Terlambat" dari sesi masuk.
 */
function tentukanStatusPulang_(jamStr, pengaturan, statusMasuk) {
  if (jamStr < pengaturan.jam_pulang) {
    return statusMasuk === 'Terlambat' ? 'Terlambat & Pulang Cepat' : 'Pulang Cepat';
  }
  return statusMasuk; // tetap "Hadir" atau "Terlambat"
}

/**
 * Tambah sejumlah menit ke string jam "HH:mm", return "HH:mm:ss".
 */
function tambahMenitKeJam_(jamHHmm, menit) {
  var parts = jamHHmm.split(':');
  var totalMenit = parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10) + parseInt(menit, 10);
  var jam = Math.floor(totalMenit / 60) % 24;
  var mnt = totalMenit % 60;
  return (jam < 10 ? '0' + jam : jam) + ':' + (mnt < 10 ? '0' + mnt : mnt) + ':59';
}
