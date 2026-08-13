/**
 * ===================================================================
 * IZIN.GS
 * Modul pengajuan izin/cuti/sakit oleh pegawai, riwayat pengajuan
 * pribadi, daftar izin untuk admin (dengan filter & paginasi),
 * dan approve/reject oleh admin.
 *
 * Disinkronkan dengan Code.gs & Auth.gs asli (repo non-fix):
 * - Kolom sheet "Izin" (setupAwal di Code.gs):
 *   0 id_izin, 1 id_user, 2 tanggal_pengajuan, 3 jenis_izin,
 *   4 tanggal_mulai, 5 tanggal_selesai, 6 keterangan, 7 lampiran,
 *   8 status, 9 catatan_admin, 10 diproses_oleh, 11 diproses_at
 * - Folder Drive: FOLDER_IZIN_NAME ('Lampiran_Izin')
 * - generateSequentialId(prefix, namaSheet, kolomId)
 * - verifySession(token) -> { valid, user: { id_user, nama, username,
 *   role, jabatan, foto_profil } }  (dari Auth,gs)
 * - Pola LockService.getScriptLock() dipakai di semua fungsi tulis,
 *   sama seperti di Absensi.gs.
 * ===================================================================
 */

var KOLOM_IZIN = {
  ID: 0, ID_USER: 1, TANGGAL_PENGAJUAN: 2, JENIS_IZIN: 3,
  TANGGAL_MULAI: 4, TANGGAL_SELESAI: 5, KETERANGAN: 6, LAMPIRAN: 7,
  STATUS: 8, CATATAN_ADMIN: 9, DIPROSES_OLEH: 10, DIPROSES_AT: 11
};

var JENIS_IZIN_VALID = ['Izin', 'Cuti', 'Sakit'];
var STATUS_IZIN_PENDING = 'Menunggu';
var STATUS_IZIN_DISETUJUI = 'Disetujui';
var STATUS_IZIN_DITOLAK = 'Ditolak';

/* ===================================================================
 * FITUR PEGAWAI
 * =================================================================== */

/**
 * Ajukan izin/cuti/sakit baru.
 * @param {string} token
 * @param {string} jenisIzin - 'Izin' | 'Cuti' | 'Sakit'
 * @param {string} tanggalMulai - format 'yyyy-MM-dd'
 * @param {string} tanggalSelesai - format 'yyyy-MM-dd'
 * @param {string} keterangan
 * @param {string} [lampiranBase64] - opsional, data URL base64 (gambar/PDF)
 */
function ajukanIzin(token, jenisIzin, tanggalMulai, tanggalSelesai, keterangan, lampiranBase64) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    var session = verifySession(token);
    if (!session.valid) throw new Error(session.message || 'Sesi tidak valid.');
    var user = session.user;

    // --- Validasi input di backend (jangan percaya client 100%) ---
    if (JENIS_IZIN_VALID.indexOf(jenisIzin) === -1) {
      throw new Error('Jenis izin tidak valid. Pilih Izin, Cuti, atau Sakit.');
    }
    if (!tanggalMulai || !tanggalSelesai) {
      throw new Error('Tanggal mulai dan tanggal selesai wajib diisi.');
    }
    if (tanggalSelesai < tanggalMulai) {
      throw new Error('Tanggal selesai tidak boleh sebelum tanggal mulai.');
    }
    if (!keterangan || String(keterangan).trim().length < 3) {
      throw new Error('Keterangan wajib diisi (minimal 3 karakter).');
    }

    var sheet = getSheet(SHEET_IZIN);
    var data = sheet.getDataRange().getValues();

    // Cegah pengajuan tumpang tindih (tanggal beririsan) yang statusnya
    // masih Menunggu atau sudah Disetujui, supaya tidak dobel.
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[KOLOM_IZIN.ID_USER] !== user.id_user) continue;
      var statusLama = row[KOLOM_IZIN.STATUS];
      if (statusLama !== STATUS_IZIN_PENDING && statusLama !== STATUS_IZIN_DISETUJUI) continue;

      var mulaiLama = String(row[KOLOM_IZIN.TANGGAL_MULAI]);
      var selesaiLama = String(row[KOLOM_IZIN.TANGGAL_SELESAI]);
      var tumpangTindih = tanggalMulai <= selesaiLama && tanggalSelesai >= mulaiLama;
      if (tumpangTindih) {
        throw new Error(
          'Anda sudah punya pengajuan izin (' + statusLama + ') yang tanggalnya beririsan ' +
          '(' + mulaiLama + ' s/d ' + selesaiLama + ').'
        );
      }
    }

    var now = new Date();
    var tanggalPengajuanStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    var urlLampiran = '';
    if (lampiranBase64) {
      urlLampiran = simpanLampiranIzin_(lampiranBase64, user.id_user + '_izin_' + tanggalPengajuanStr);
    }

    var idBaru = generateSequentialId('IZN', SHEET_IZIN, KOLOM_IZIN.ID);
    var rowBaru = [];
    rowBaru[KOLOM_IZIN.ID] = idBaru;
    rowBaru[KOLOM_IZIN.ID_USER] = user.id_user;
    rowBaru[KOLOM_IZIN.TANGGAL_PENGAJUAN] = tanggalPengajuanStr;
    rowBaru[KOLOM_IZIN.JENIS_IZIN] = jenisIzin;
    rowBaru[KOLOM_IZIN.TANGGAL_MULAI] = tanggalMulai;
    rowBaru[KOLOM_IZIN.TANGGAL_SELESAI] = tanggalSelesai;
    rowBaru[KOLOM_IZIN.KETERANGAN] = keterangan;
    rowBaru[KOLOM_IZIN.LAMPIRAN] = urlLampiran;
    rowBaru[KOLOM_IZIN.STATUS] = STATUS_IZIN_PENDING;
    rowBaru[KOLOM_IZIN.CATATAN_ADMIN] = '';
    rowBaru[KOLOM_IZIN.DIPROSES_OLEH] = '';
    rowBaru[KOLOM_IZIN.DIPROSES_AT] = '';

    sheet.appendRow(rowBaru);

    return jsonResponse({
      success: true,
      message: 'Pengajuan ' + jenisIzin.toLowerCase() + ' berhasil dikirim. Menunggu persetujuan admin.',
      data: { id_izin: idBaru, status: STATUS_IZIN_PENDING }
    });

  } catch (e) {
    return jsonResponse({ success: false, message: e.message });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Riwayat pengajuan izin milik user yang login, dengan paginasi
 * dan filter status opsional.
 * @param {string} token
 * @param {number} page - mulai dari 1
 * @param {number} pageSize - default 20
 * @param {string} [statusFilter] - 'Menunggu' | 'Disetujui' | 'Ditolak'
 */
function getHistoryIzin(token, page, pageSize, statusFilter) {
  try {
    var session = verifySession(token);
    if (!session.valid) throw new Error(session.message || 'Sesi tidak valid.');

    page = page && page > 0 ? page : 1;
    pageSize = pageSize && pageSize > 0 ? pageSize : 20;

    var sheet = getSheet(SHEET_IZIN);
    var data = sheet.getDataRange().getValues();
    var hasil = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row[KOLOM_IZIN.ID_USER] !== session.user.id_user) continue;
      if (statusFilter && row[KOLOM_IZIN.STATUS] !== statusFilter) continue;

      hasil.push(_baristIzinKeObjek_(row));
    }

    // Terbaru dulu (berdasarkan tanggal pengajuan)
    hasil.sort(function (a, b) { return a.tanggal_pengajuan < b.tanggal_pengajuan ? 1 : -1; });

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
 * FITUR ADMIN
 * =================================================================== */

/**
 * Daftar seluruh pengajuan izin (semua pegawai), dengan paginasi
 * dan filter status / bulan / pegawai. Khusus admin.
 * @param {string} token
 * @param {number} page
 * @param {number} pageSize
 * @param {string} [statusFilter]
 * @param {string} [bulanFilter] - format 'yyyy-MM', dicek pada tanggal_mulai
 * @param {string} [idUserFilter]
 */
function getDaftarIzinAdmin(token, page, pageSize, statusFilter, bulanFilter, idUserFilter) {
  try {
    var session = verifySession(token);
    if (!session.valid) throw new Error(session.message || 'Sesi tidak valid.');
    if (session.user.role !== 'admin') throw new Error('Akses ditolak. Hanya admin.');

    page = page && page > 0 ? page : 1;
    pageSize = pageSize && pageSize > 0 ? pageSize : 20;

    var petaNama = _petaNamaUser_();

    var sheet = getSheet(SHEET_IZIN);
    var data = sheet.getDataRange().getValues();
    var hasil = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (statusFilter && row[KOLOM_IZIN.STATUS] !== statusFilter) continue;
      if (idUserFilter && row[KOLOM_IZIN.ID_USER] !== idUserFilter) continue;
      if (bulanFilter && String(row[KOLOM_IZIN.TANGGAL_MULAI]).indexOf(bulanFilter) !== 0) continue;

      var obj = _baristIzinAdminKeObjek_(row, petaNama);
      hasil.push(obj);
    }

    hasil.sort(function (a, b) { return a.tanggal_pengajuan < b.tanggal_pengajuan ? 1 : -1; });

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

/**
 * Approve atau reject pengajuan izin. Khusus admin.
 * @param {string} token
 * @param {string} idIzin
 * @param {string} statusBaru - 'Disetujui' | 'Ditolak'
 * @param {string} [catatanAdmin]
 */
function prosesIzin(token, idIzin, statusBaru, catatanAdmin) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    var session = verifySession(token);
    if (!session.valid) throw new Error(session.message || 'Sesi tidak valid.');
    if (session.user.role !== 'admin') throw new Error('Akses ditolak. Hanya admin.');

    if (statusBaru !== STATUS_IZIN_DISETUJUI && statusBaru !== STATUS_IZIN_DITOLAK) {
      throw new Error('Status baru tidak valid. Gunakan "Disetujui" atau "Ditolak".');
    }

    var sheet = getSheet(SHEET_IZIN);
    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < data.length; i++) {
      if (data[i][KOLOM_IZIN.ID] === idIzin) {
        rowIndex = i;
        break;
      }
    }
    if (rowIndex === -1) throw new Error('Pengajuan izin tidak ditemukan.');

    var statusSaatIni = data[rowIndex][KOLOM_IZIN.STATUS];
    if (statusSaatIni !== STATUS_IZIN_PENDING) {
      throw new Error('Pengajuan ini sudah diproses sebelumnya (' + statusSaatIni + ').');
    }

    var now = new Date();
    var row = rowIndex + 1; // 1-indexed di Sheets

    sheet.getRange(row, KOLOM_IZIN.STATUS + 1).setValue(statusBaru);
    sheet.getRange(row, KOLOM_IZIN.CATATAN_ADMIN + 1).setValue(catatanAdmin || '');
    sheet.getRange(row, KOLOM_IZIN.DIPROSES_OLEH + 1).setValue(session.user.id_user);
    sheet.getRange(row, KOLOM_IZIN.DIPROSES_AT + 1).setValue(now);

    return jsonResponse({
      success: true,
      message: 'Pengajuan izin berhasil ' + (statusBaru === STATUS_IZIN_DISETUJUI ? 'disetujui.' : 'ditolak.')
    });

  } catch (e) {
    return jsonResponse({ success: false, message: e.message });
  } finally {
    lock.releaseLock();
  }
}

/* ===================================================================
 * HELPER INTERNAL
 * =================================================================== */

/**
 * Simpan lampiran (gambar atau PDF) ke folder Drive khusus lampiran izin.
 * Batas ukuran 5MB (lebih longgar dari foto absen karena bisa berupa
 * scan surat dokter / PDF).
 */
function simpanLampiranIzin_(fileBase64, namaFileTanpaExt) {
  var MAKS_BYTES = 5 * 1024 * 1024; // 5MB
  var bagian = fileBase64.split(',');
  var meta = bagian[0]; // contoh: "data:application/pdf;base64" atau "data:image/jpeg;base64"
  var base64Data = bagian.length > 1 ? bagian[1] : bagian[0];

  var contentType = 'application/octet-stream';
  var match = meta.match(/data:([^;]+);base64/);
  if (match) contentType = match[1];

  var tipeDiizinkan = ['image/jpeg', 'image/png', 'image/webp', 'application/pdf'];
  if (tipeDiizinkan.indexOf(contentType) === -1) {
    throw new Error('Format lampiran tidak didukung. Gunakan JPG, PNG, atau PDF.');
  }

  var bytes = Utilities.base64Decode(base64Data);
  if (bytes.length > MAKS_BYTES) {
    throw new Error('Ukuran lampiran terlalu besar. Maksimal 5MB.');
  }

  var ekstensi = contentType.split('/')[1] || 'bin';
  if (ekstensi === 'jpeg') ekstensi = 'jpg';

  var blob = Utilities.newBlob(bytes, contentType, namaFileTanpaExt + '.' + ekstensi);
  var folder = getOrCreateFolder_(FOLDER_IZIN_NAME);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return 'https://drive.google.com/uc?id=' + file.getId();
}

/**
 * Ubah 1 baris sheet Izin menjadi object untuk riwayat pribadi pegawai.
 */
function _baristIzinKeObjek_(row) {
  return {
    id: row[KOLOM_IZIN.ID],
    tanggal_pengajuan: row[KOLOM_IZIN.TANGGAL_PENGAJUAN],
    jenis_izin: row[KOLOM_IZIN.JENIS_IZIN],
    tanggal_mulai: row[KOLOM_IZIN.TANGGAL_MULAI],
    tanggal_selesai: row[KOLOM_IZIN.TANGGAL_SELESAI],
    keterangan: row[KOLOM_IZIN.KETERANGAN],
    lampiran: row[KOLOM_IZIN.LAMPIRAN],
    status: row[KOLOM_IZIN.STATUS],
    catatan_admin: row[KOLOM_IZIN.CATATAN_ADMIN]
  };
}
// Alias supaya nama fungsi konsisten dipakai di getHistoryIzin di atas
// (dijaga tetap ada dua-duanya kalau ada pemanggilan lama ke nama lain).
function _baristIzinAdminKeObjek_(row, petaNama) {
  var obj = _baristIzinKeObjek_(row);
  obj.id_user = row[KOLOM_IZIN.ID_USER];
  obj.nama_pegawai = petaNama[row[KOLOM_IZIN.ID_USER]] || '(tidak diketahui)';
  obj.diproses_oleh = row[KOLOM_IZIN.DIPROSES_OLEH] ? (petaNama[row[KOLOM_IZIN.DIPROSES_OLEH]] || row[KOLOM_IZIN.DIPROSES_OLEH]) : '';
  obj.diproses_at = row[KOLOM_IZIN.DIPROSES_AT];
  return obj;
}

/**
 * Bangun peta { id_user: nama } dari sheet Users, dipakai admin
 * supaya daftar izin menampilkan nama pegawai, bukan hanya id_user.
 */
function _petaNamaUser_() {
  var sheetUsers = getSheet(SHEET_USERS);
  var data = sheetUsers.getDataRange().getValues();
  var peta = {};
  for (var i = 1; i < data.length; i++) {
    peta[data[i][0]] = data[i][1]; // id_user -> nama
  }
  return peta;
}
