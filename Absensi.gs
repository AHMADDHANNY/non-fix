/**
 * ===================================================================
 * ABSENSI.GS - VERSI PERBAIKAN
 *
 * Modul:
 * - Absen masuk
 * - Absen pulang
 * - Validasi GPS/radius kantor di backend
 * - Validasi waktu server
 * - Upload foto selfie
 * - History absensi pribadi + paginasi
 * - Penentuan Hadir/Terlambat
 * - Penandaan Alpa otomatis
 * - Trigger otomatis untuk Alpa
 *
 * Dependency:
 * - Code.gs
 * - Auth.gs
 * - Izin.gs
 * - Admin.gs
 * ===================================================================
 */

var KOLOM_ABSENSI = {
  ID: 0,
  ID_USER: 1,
  TANGGAL: 2,
  JAM_MASUK: 3,
  FOTO_MASUK: 4,
  LAT_MASUK: 5,
  LONG_MASUK: 6,
  JAM_PULANG: 7,
  FOTO_PULANG: 8,
  LAT_PULANG: 9,
  LONG_PULANG: 10,
  STATUS: 11,
  JARAK_MASUK: 12,
  JARAK_PULANG: 13
};


/* ===================================================================
 * 1. ABSEN MASUK
 * =================================================================== */

/**
 * @param {string} token
 * @param {number|string} lat
 * @param {number|string} long
 * @param {string} fotoBase64
 */
function absenMasuk(token, lat, long, fotoBase64) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    var session = verifySession(token);
    if (!session.valid) {
      throw new Error(
        session.message || 'Sesi tidak valid. Silakan login ulang.'
      );
    }

    var user = session.user;

    /*
     * Validasi backend. Jangan percaya data dari browser.
     */
    var koordinat = validasiKoordinat_(lat, long);

    if (!fotoBase64) {
      throw new Error('Foto selfie wajib diambil.');
    }

    var pengaturan = getPengaturanCache();
    validasiPengaturanAbsensi_(pengaturan);

    var jarak = hitungJarakMeter(
      koordinat.lat,
      koordinat.long,
      parseFloat(pengaturan.lat_kantor),
      parseFloat(pengaturan.long_kantor)
    );

    var radius = parseFloat(pengaturan.radius_meter);

    if (jarak > radius) {
      throw new Error(
        'Anda berada di luar radius kantor (' +
        Math.round(jarak) +
        ' m dari kantor, maksimal ' +
        radius +
        ' m). Absen ditolak.'
      );
    }

    /*
     * Waktu selalu berasal dari server Apps Script.
     */
    var sekarang = new Date();
    var tanggalStr = formatTanggalAbsensi_(sekarang);
    var jamStr = formatJamAbsensi_(sekarang);

    var sheet = getSheet(SHEET_ABSENSI);
    var lastRow = sheet.getLastRow();

    var data = [];
    if (lastRow >= 2) {
      data = sheet
        .getRange(2, 1, lastRow - 1, 14)
        .getValues();
    }

    var rowIndex = -1;

    /*
     * Cari record pegawai pada tanggal tersebut.
     */
    for (var i = 0; i < data.length; i++) {
      var idUserRow = String(
        data[i][KOLOM_ABSENSI.ID_USER] || ''
      );

      var tanggalRow = normalizeTanggal_(data[i][KOLOM_ABSENSI.TANGGAL]);

      if (
        idUserRow === String(user.id_user) &&
        tanggalRow === tanggalStr
      ) {
        rowIndex = i;
        break;
      }
    }

    /*
     * Jika sudah ada record:
     *
     * - Bila sudah punya jam masuk -> tolak.
     * - Bila record dibuat otomatis sebagai ALPA tetapi belum
     *   punya jam masuk -> kita isi record tersebut kembali.
     *
     * Dengan demikian trigger Alpa tidak menghalangi pegawai
     * yang ternyata melakukan absensi setelah record Alpa dibuat.
     */
    if (rowIndex !== -1) {
      var existing = data[rowIndex];

      var jamMasukLama = String(
        existing[KOLOM_ABSENSI.JAM_MASUK] || ''
      );

      if (jamMasukLama) {
        throw new Error(
          'Anda sudah melakukan absen masuk hari ini.'
        );
      }

      var statusLama = String(
        existing[KOLOM_ABSENSI.STATUS] || ''
      );

      if (
        statusLama !== 'Alpa' &&
        statusLama !== ''
      ) {
        throw new Error(
          'Data absensi hari ini sudah ada dan tidak dapat diubah.'
        );
      }

      var urlFoto = simpanFotoAbsen_(
        fotoBase64,
        user.id_user + '_masuk_' + tanggalStr
      );

      var status = tentukanStatusMasuk_(
        jamStr,
        pengaturan
      );

      var sheetRow = rowIndex + 2;

      sheet.getRange(
        sheetRow,
        KOLOM_ABSENSI.JAM_MASUK + 1,
        1,
        4
      ).setValues([[
        jamStr,
        urlFoto,
        koordinat.lat,
        koordinat.long
      ]]);

      sheet.getRange(
        sheetRow,
        KOLOM_ABSENSI.STATUS + 1
      ).setValue(status);

      sheet.getRange(
        sheetRow,
        KOLOM_ABSENSI.JARAK_MASUK + 1
      ).setValue(Math.round(jarak));

      /*
       * Pastikan field pulang tetap kosong.
       */
      return jsonResponse({
        success: true,
        message:
          'Absen masuk berhasil (' + status + ').',
        data: {
          id_absen:
            existing[KOLOM_ABSENSI.ID],
          jam_masuk: jamStr,
          status: status,
          jarak_meter: Math.round(jarak)
        }
      });
    }

    /*
     * Record baru.
     */
    var urlFotoBaru = simpanFotoAbsen_(
      fotoBase64,
      user.id_user + '_masuk_' + tanggalStr
    );

    var statusBaru = tentukanStatusMasuk_(
      jamStr,
      pengaturan
    );

    var idBaru = generateSequentialId(
      'ABS',
      SHEET_ABSENSI,
      KOLOM_ABSENSI.ID
    );

    var rowBaru = new Array(14);

    rowBaru[KOLOM_ABSENSI.ID] = idBaru;
    rowBaru[KOLOM_ABSENSI.ID_USER] = user.id_user;
    rowBaru[KOLOM_ABSENSI.TANGGAL] = tanggalStr;
    rowBaru[KOLOM_ABSENSI.JAM_MASUK] = jamStr;
    rowBaru[KOLOM_ABSENSI.FOTO_MASUK] = urlFotoBaru;
    rowBaru[KOLOM_ABSENSI.LAT_MASUK] = koordinat.lat;
    rowBaru[KOLOM_ABSENSI.LONG_MASUK] = koordinat.long;
    rowBaru[KOLOM_ABSENSI.JAM_PULANG] = '';
    rowBaru[KOLOM_ABSENSI.FOTO_PULANG] = '';
    rowBaru[KOLOM_ABSENSI.LAT_PULANG] = '';
    rowBaru[KOLOM_ABSENSI.LONG_PULANG] = '';
    rowBaru[KOLOM_ABSENSI.STATUS] = statusBaru;
    rowBaru[KOLOM_ABSENSI.JARAK_MASUK] = Math.round(jarak);
    rowBaru[KOLOM_ABSENSI.JARAK_PULANG] = '';

    sheet.appendRow(rowBaru);

    return jsonResponse({
      success: true,
      message:
        'Absen masuk berhasil (' + statusBaru + ').',
      data: {
        id_absen: idBaru,
        jam_masuk: jamStr,
        status: statusBaru,
        jarak_meter: Math.round(jarak)
      }
    });

  } catch (e) {
    return jsonResponse({
      success: false,
      message: e.message || 'Gagal melakukan absen masuk.'
    });

  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}


/* ===================================================================
 * 2. ABSEN PULANG
 * =================================================================== */

/**
 * @param {string} token
 * @param {number|string} lat
 * @param {number|string} long
 * @param {string} fotoBase64
 */
function absenPulang(token, lat, long, fotoBase64) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    var session = verifySession(token);
    if (!session.valid) {
      throw new Error(
        session.message || 'Sesi tidak valid. Silakan login ulang.'
      );
    }

    var user = session.user;

    var koordinat = validasiKoordinat_(lat, long);

    if (!fotoBase64) {
      throw new Error('Foto selfie wajib diambil.');
    }

    var pengaturan = getPengaturanCache();
    validasiPengaturanAbsensi_(pengaturan);

    var jarak = hitungJarakMeter(
      koordinat.lat,
      koordinat.long,
      parseFloat(pengaturan.lat_kantor),
      parseFloat(pengaturan.long_kantor)
    );

    var radius = parseFloat(pengaturan.radius_meter);

    if (jarak > radius) {
      throw new Error(
        'Anda berada di luar radius kantor (' +
        Math.round(jarak) +
        ' m dari kantor, maksimal ' +
        radius +
        ' m). Absen ditolak.'
      );
    }

    var sekarang = new Date();
    var tanggalStr = formatTanggalAbsensi_(sekarang);
    var jamStr = formatJamAbsensi_(sekarang);

    var sheet = getSheet(SHEET_ABSENSI);
    var lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      throw new Error(
        'Anda belum melakukan absen masuk hari ini.'
      );
    }

    var data = sheet
      .getRange(2, 1, lastRow - 1, 14)
      .getValues();

    var rowIndex = -1;

    for (var i = 0; i < data.length; i++) {
      var idUserRow = String(
        data[i][KOLOM_ABSENSI.ID_USER] || ''
      );

      var tanggalRow = normalizeTanggal_(
        data[i][KOLOM_ABSENSI.TANGGAL]
      );

      if (
        idUserRow === String(user.id_user) &&
        tanggalRow === tanggalStr
      ) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error(
        'Anda belum melakukan absen masuk hari ini.'
      );
    }

    var rowData = data[rowIndex];

    var jamMasuk = String(
      rowData[KOLOM_ABSENSI.JAM_MASUK] || ''
    );

    if (!jamMasuk) {
      throw new Error(
        'Absen masuk belum tercatat. Silakan absen masuk terlebih dahulu.'
      );
    }

    var jamPulangLama = String(
      rowData[KOLOM_ABSENSI.JAM_PULANG] || ''
    );

    if (jamPulangLama) {
      throw new Error(
        'Anda sudah melakukan absen pulang hari ini.'
      );
    }

    /*
     * Secara logika jam pulang tidak boleh lebih awal dari
     * jam masuk pada hari yang sama.
     */
    if (jamStr < jamMasuk) {
      throw new Error(
        'Jam pulang tidak boleh lebih awal dari jam masuk.'
      );
    }

    var urlFoto = simpanFotoAbsen_(
      fotoBase64,
      user.id_user + '_pulang_' + tanggalStr
    );

    var statusLama = String(
      rowData[KOLOM_ABSENSI.STATUS] || ''
    );

    var statusBaru = tentukanStatusPulang_(
      jamStr,
      pengaturan,
      statusLama
    );

    var sheetRow = rowIndex + 2;

    /*
     * Satu operasi range untuk field pulang.
     */
    sheet.getRange(
      sheetRow,
      KOLOM_ABSENSI.JAM_PULANG + 1,
      1,
      4
    ).setValues([[
      jamStr,
      urlFoto,
      koordinat.lat,
      koordinat.long
    ]]);

    sheet.getRange(
      sheetRow,
      KOLOM_ABSENSI.STATUS + 1
    ).setValue(statusBaru);

    sheet.getRange(
      sheetRow,
      KOLOM_ABSENSI.JARAK_PULANG + 1
    ).setValue(Math.round(jarak));

    return jsonResponse({
      success: true,
      message:
        'Absen pulang berhasil (' + statusBaru + ').',
      data: {
        id_absen:
          rowData[KOLOM_ABSENSI.ID],
        jam_pulang: jamStr,
        status: statusBaru,
        jarak_meter: Math.round(jarak)
      }
    });

  } catch (e) {
    return jsonResponse({
      success: false,
      message: e.message || 'Gagal melakukan absen pulang.'
    });

  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}


/* ===================================================================
 * 3. HISTORY ABSENSI PEGAWAI
 * =================================================================== */

/**
 * @param {string} token
 * @param {number} page
 * @param {number} pageSize
 * @param {string} bulanFilter yyyy-MM
 */
function getHistoryAbsensi(
  token,
  page,
  pageSize,
  bulanFilter
) {
  try {
    var session = verifySession(token);

    if (!session.valid) {
      throw new Error(
        session.message || 'Sesi tidak valid.'
      );
    }

    page = Number(page) || 1;
    pageSize = Number(pageSize) || 20;

    if (page < 1) page = 1;
    if (pageSize < 1) pageSize = 20;
    if (pageSize > 50) pageSize = 50;

    bulanFilter = String(
      bulanFilter || ''
    ).trim();

    if (
      bulanFilter &&
      !/^\d{4}-\d{2}$/.test(bulanFilter)
    ) {
      throw new Error(
        'Format filter bulan tidak valid.'
      );
    }

    var sheet = getSheet(SHEET_ABSENSI);
    var lastRow = sheet.getLastRow();

    var hasil = [];

    if (lastRow >= 2) {
      var data = sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          14
        )
        .getValues();

      for (var i = 0; i < data.length; i++) {
        var row = data[i];

        if (
          String(
            row[KOLOM_ABSENSI.ID_USER] || ''
          ) !==
          String(session.user.id_user)
        ) {
          continue;
        }

        var tanggal =
          normalizeTanggal_(
            row[KOLOM_ABSENSI.TANGGAL]
          );

        if (
          bulanFilter &&
          tanggal.indexOf(bulanFilter) !== 0
        ) {
          continue;
        }

        hasil.push(
          absensiRowKeObjek_(row)
        );
      }
    }

    hasil.sort(function(a, b) {
      if (a.tanggal === b.tanggal) {
        return String(
          b.jam_masuk || ''
        ).localeCompare(
          String(a.jam_masuk || '')
        );
      }

      return a.tanggal < b.tanggal
        ? 1
        : -1;
    });

    var total = hasil.length;
    var totalPage =
      Math.ceil(total / pageSize);

    var start =
      (page - 1) * pageSize;

    return jsonResponse({
      success: true,
      data: hasil.slice(
        start,
        start + pageSize
      ),
      pagination: {
        page: page,
        pageSize: pageSize,
        total: total,
        totalPage: totalPage
      }
    });

  } catch (e) {
    return jsonResponse({
      success: false,
      message: e.message
    });
  }
}


/* ===================================================================
 * 4. STATUS ABSENSI HARI INI
 * =================================================================== */

/**
 * Dipakai dashboard pegawai untuk mengetahui status hari ini.
 */
function getStatusAbsensiHariIni(token) {
  try {
    var session = verifySession(token);

    if (!session.valid) {
      throw new Error(
        session.message || 'Sesi tidak valid.'
      );
    }

    var tanggalHariIni =
      formatTanggalAbsensi_(
        new Date()
      );

    var sheet =
      getSheet(SHEET_ABSENSI);

    var lastRow =
      sheet.getLastRow();

    if (lastRow < 2) {
      return jsonResponse({
        success: true,
        data: {
          tanggal: tanggalHariIni,
          sudah_masuk: false,
          sudah_pulang: false,
          jam_masuk: '',
          jam_pulang: '',
          status: ''
        }
      });
    }

    var data =
      sheet.getRange(
        2,
        1,
        lastRow - 1,
        14
      ).getValues();

    for (var i = 0; i < data.length; i++) {
      var row = data[i];

      if (
        String(
          row[KOLOM_ABSENSI.ID_USER] || ''
        ) !==
        String(session.user.id_user)
      ) {
        continue;
      }

      if (
        normalizeTanggal_(
          row[KOLOM_ABSENSI.TANGGAL]
        ) !== tanggalHariIni
      ) {
        continue;
      }

      return jsonResponse({
        success: true,
        data: {
          tanggal: tanggalHariIni,
          sudah_masuk:
            !!row[KOLOM_ABSENSI.JAM_MASUK],
          sudah_pulang:
            !!row[KOLOM_ABSENSI.JAM_PULANG],
          jam_masuk:
            row[KOLOM_ABSENSI.JAM_MASUK] || '',
          jam_pulang:
            row[KOLOM_ABSENSI.JAM_PULANG] || '',
          status:
            row[KOLOM_ABSENSI.STATUS] || '',
          jarak_masuk_meter:
            row[KOLOM_ABSENSI.JARAK_MASUK] || '',
          jarak_pulang_meter:
            row[KOLOM_ABSENSI.JARAK_PULANG] || ''
        }
      });
    }

    return jsonResponse({
      success: true,
      data: {
        tanggal: tanggalHariIni,
        sudah_masuk: false,
        sudah_pulang: false,
        jam_masuk: '',
        jam_pulang: '',
        status: ''
      }
    });

  } catch (e) {
    return jsonResponse({
      success: false,
      message: e.message
    });
  }
}


/* ===================================================================
 * 5. PENANDAAN ALPA OTOMATIS
 * =================================================================== */

/**
 * Menandai pegawai aktif yang tidak memiliki absensi masuk
 * pada hari kerja yang diproses.
 *
 * Fungsi ini aman dipanggil berulang kali karena tidak akan
 * membuat record Alpa duplikat.
 *
 * Catatan:
 * - Pengajuan izin/cuti/sakit yang sudah DISETUJUI tidak diberi Alpa.
 * - Pengajuan Menunggu belum dianggap izin resmi.
 * - Hanya role "pegawai" aktif yang diproses.
 */
function tandaiAlpaHarian() {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(20000);

    var sekarang = new Date();
    var tanggalHariIni =
      formatTanggalAbsensi_(
        sekarang
      );

    var sheetUsers =
      getSheet(SHEET_USERS);

    var sheetAbsensi =
      getSheet(SHEET_ABSENSI);

    var usersLastRow =
      sheetUsers.getLastRow();

    if (usersLastRow < 2) {
      return jsonResponse({
        success: true,
        message: 'Tidak ada user.',
        data: {
          tanggal: tanggalHariIni,
          dibuat: 0,
          dilewati: 0
        }
      });
    }

    var users =
      sheetUsers
        .getRange(
          2,
          1,
          usersLastRow - 1,
          11
        )
        .getValues();

    /*
     * Ambil absensi yang sudah ada.
     */
    var absensiMap = {};
    var absensiLastRow =
      sheetAbsensi.getLastRow();

    if (absensiLastRow >= 2) {
      var absensi =
        sheetAbsensi
          .getRange(
            2,
            1,
            absensiLastRow - 1,
            14
          )
          .getValues();

      for (
        var i = 0;
        i < absensi.length;
        i++
      ) {
        var idUser =
          String(
            absensi[i][KOLOM_ABSENSI.ID_USER] || ''
          );

        var tanggal =
          normalizeTanggal_(
            absensi[i][KOLOM_ABSENSI.TANGGAL]
          );

        if (
          idUser &&
          tanggal
        ) {
          absensiMap[
            idUser + '|' + tanggal
          ] = true;
        }
      }
    }

    /*
     * Cari izin yang sudah disetujui pada tanggal tersebut.
     */
    var izinMap =
      getPetaIzinDisetujuiUntukTanggal_(
        tanggalHariIni
      );

    var rowsBaru = [];
    var dibuat = 0;
    var dilewati = 0;

    for (
      var u = 0;
      u < users.length;
      u++
    ) {
      var user = users[u];

      var idUser =
        String(
          user[0] || ''
        );

      var role =
        String(
          user[5] || ''
        ).toLowerCase();

      var statusAktif =
        String(
          user[8] || ''
        ).toLowerCase();

      if (
        !idUser ||
        role !== 'pegawai' ||
        statusAktif !== 'aktif'
      ) {
        continue;
      }

      var key =
        idUser + '|' + tanggalHariIni;

      /*
       * Sudah punya record absensi.
       */
      if (absensiMap[key]) {
        dilewati++;
        continue;
      }

      /*
       * Sudah memiliki izin resmi yang disetujui.
       */
      if (izinMap[idUser]) {
        dilewati++;
        continue;
      }

      var idBaru =
        generateSequentialId(
          'ABS',
          SHEET_ABSENSI,
          KOLOM_ABSENSI.ID
        );

      var rowBaru =
        new Array(14);

      rowBaru[
        KOLOM_ABSENSI.ID
      ] = idBaru;

      rowBaru[
        KOLOM_ABSENSI.ID_USER
      ] = idUser;

      rowBaru[
        KOLOM_ABSENSI.TANGGAL
      ] = tanggalHariIni;

      rowBaru[
        KOLOM_ABSENSI.JAM_MASUK
      ] = '';

      rowBaru[
        KOLOM_ABSENSI.FOTO_MASUK
      ] = '';

      rowBaru[
        KOLOM_ABSENSI.LAT_MASUK
      ] = '';

      rowBaru[
        KOLOM_ABSENSI.LONG_MASUK
      ] = '';

      rowBaru[
        KOLOM_ABSENSI.JAM_PULANG
      ] = '';

      rowBaru[
        KOLOM_ABSENSI.FOTO_PULANG
      ] = '';

      rowBaru[
        KOLOM_ABSENSI.LAT_PULANG
      ] = '';

      rowBaru[
        KOLOM_ABSENSI.LONG_PULANG
      ] = '';

      rowBaru[
        KOLOM_ABSENSI.STATUS
      ] = 'Alpa';

      rowBaru[
        KOLOM_ABSENSI.JARAK_MASUK
      ] = '';

      rowBaru[
        KOLOM_ABSENSI.JARAK_PULANG
      ] = '';

      rowsBaru.push(rowBaru);

      /*
       * Masukkan ke map agar tidak dibuat lagi
       * jika terjadi duplikasi di data user.
       */
      absensiMap[key] = true;
      dibuat++;
    }

    if (rowsBaru.length > 0) {
      var firstRow =
        sheetAbsensi.getLastRow() + 1;

      sheetAbsensi
        .getRange(
          firstRow,
          1,
          rowsBaru.length,
          14
        )
        .setValues(rowsBaru);
    }

    return jsonResponse({
      success: true,
      message:
        'Penandaan Alpa selesai.',
      data: {
        tanggal: tanggalHariIni,
        dibuat: dibuat,
        dilewati: dilewati
      }
    });

  } catch (e) {
    return jsonResponse({
      success: false,
      message: e.message
    });

  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}


/* ===================================================================
 * 6. TRIGGER ALPA
 * =================================================================== */

/**
 * Membuat trigger harian untuk tandaiAlpaHarian().
 *
 * Jalankan manual SATU KALI dari Apps Script.
 *
 * Trigger menggunakan sekitar pukul 23:00-00:00.
 * Karena Apps Script menggunakan window waktu, fungsi tetap
 * mengambil tanggal server saat dieksekusi.
 */
function buatTriggerAlpaHarian() {
  var lock =
    LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    /*
     * Hapus trigger lama milik fungsi yang sama.
     */
    var triggers =
      ScriptApp.getProjectTriggers();

    for (
      var i = 0;
      i < triggers.length;
      i++
    ) {
      if (
        triggers[i].getHandlerFunction() ===
        'tandaiAlpaHarian'
      ) {
        ScriptApp.deleteTrigger(
          triggers[i]
        );
      }
    }

    /*
     * Jalankan setiap hari pada sekitar jam 23.
     */
    ScriptApp
      .newTrigger(
        'tandaiAlpaHarian'
      )
      .timeBased()
      .everyDays(1)
      .atHour(23)
      .create();

    return jsonResponse({
      success: true,
      message:
        'Trigger Alpa harian berhasil dibuat.'
    });

  } catch (e) {
    return jsonResponse({
      success: false,
      message: e.message
    });

  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}


/**
 * Menghapus trigger Alpa harian.
 */
function hapusTriggerAlpaHarian() {
  var lock =
    LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    var triggers =
      ScriptApp.getProjectTriggers();

    var jumlah = 0;

    for (
      var i = triggers.length - 1;
      i >= 0;
      i--
    ) {
      if (
        triggers[i].getHandlerFunction() ===
        'tandaiAlpaHarian'
      ) {
        ScriptApp.deleteTrigger(
          triggers[i]
        );
        jumlah++;
      }
    }

    return jsonResponse({
      success: true,
      message:
        'Trigger Alpa dihapus.',
      data: {
        jumlah_dihapus: jumlah
      }
    });

  } catch (e) {
    return jsonResponse({
      success: false,
      message: e.message
    });

  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}


/* ===================================================================
 * 7. HELPER VALIDASI GPS
 * =================================================================== */

function validasiKoordinat_(lat, long) {
  if (
    lat === undefined ||
    lat === null ||
    lat === '' ||
    long === undefined ||
    long === null ||
    long === ''
  ) {
    throw new Error(
      'Lokasi GPS tidak terdeteksi. Aktifkan GPS dan coba lagi.'
    );
  }

  var latNum =
    Number(lat);

  var longNum =
    Number(long);

  if (
    !isFinite(latNum) ||
    !isFinite(longNum)
  ) {
    throw new Error(
      'Koordinat GPS tidak valid.'
    );
  }

  if (
    latNum < -90 ||
    latNum > 90
  ) {
    throw new Error(
      'Lintang GPS tidak valid.'
    );
  }

  if (
    longNum < -180 ||
    longNum > 180
  ) {
    throw new Error(
      'Bujur GPS tidak valid.'
    );
  }

  return {
    lat: latNum,
    long: longNum
  };
}


/* ===================================================================
 * 8. HELPER VALIDASI PENGATURAN
 * =================================================================== */

function validasiPengaturanAbsensi_(pengaturan) {
  if (!pengaturan) {
    throw new Error(
      'Pengaturan aplikasi tidak tersedia.'
    );
  }

  var latKantor =
    Number(
      pengaturan.lat_kantor
    );

  var longKantor =
    Number(
      pengaturan.long_kantor
    );

  var radius =
    Number(
      pengaturan.radius_meter
    );

  if (
    !isFinite(latKantor) ||
    latKantor < -90 ||
    latKantor > 90
  ) {
    throw new Error(
      'Koordinat lintang kantor belum dikonfigurasi dengan benar.'
    );
  }

  if (
    !isFinite(longKantor) ||
    longKantor < -180 ||
    longKantor > 180
  ) {
    throw new Error(
      'Koordinat bujur kantor belum dikonfigurasi dengan benar.'
    );
  }

  if (
    !isFinite(radius) ||
    radius <= 0
  ) {
    throw new Error(
      'Radius kantor belum dikonfigurasi dengan benar.'
    );
  }

  if (
    radius > 10000
  ) {
    throw new Error(
      'Radius kantor terlalu besar. Maksimal 10.000 meter.'
    );
  }

  if (
    !/^\d{2}:\d{2}$/.test(
      String(
        pengaturan.jam_masuk || ''
      )
    )
  ) {
    throw new Error(
      'Format jam masuk pada Pengaturan tidak valid.'
    );
  }

  if (
    !/^\d{2}:\d{2}$/.test(
      String(
        pengaturan.jam_pulang || ''
      )
    )
  ) {
    throw new Error(
      'Format jam pulang pada Pengaturan tidak valid.'
    );
  }

  var toleransi =
    Number(
      pengaturan.toleransi_menit
    );

  if (
    !isFinite(toleransi) ||
    toleransi < 0 ||
    toleransi > 1440
  ) {
    throw new Error(
      'Toleransi keterlambatan tidak valid.'
    );
  }
}


/* ===================================================================
 * 9. HELPER FOTO ABSENSI
 * =================================================================== */

/**
 * Menyimpan foto selfie ke folder Foto_Absensi.
 *
 * Maksimum 2 MB.
 *
 * Foto sebaiknya sudah diperkecil di client menggunakan canvas.
 *
 * Catatan keamanan:
 * File tetap menggunakan sharing link agar URL foto yang disimpan
 * pada sheet dapat dipakai oleh frontend Admin tanpa endpoint file
 * tambahan. Jika nanti kita ingin Drive benar-benar private,
 * frontend Admin perlu memakai endpoint proxy/base64 khusus.
 */
function simpanFotoAbsen_(
  fotoBase64,
  namaFileTanpaExt
) {
  if (
    typeof fotoBase64 !== 'string' ||
    !fotoBase64
  ) {
    throw new Error(
      'Data foto tidak valid.'
    );
  }

  var MAKS_BYTES =
    2 * 1024 * 1024;

  var bagian =
    fotoBase64.split(',');

  var meta =
    bagian.length > 1
      ? bagian[0]
      : '';

  var base64Data =
    bagian.length > 1
      ? bagian[1]
      : bagian[0];

  if (!base64Data) {
    throw new Error(
      'Data foto kosong.'
    );
  }

  var match =
    meta.match(
      /^data:(image\/(?:jpeg|jpg|png|webp));base64$/i
    );

  if (!match) {
    throw new Error(
      'Format foto tidak didukung. Gunakan JPG, PNG, atau WEBP.'
    );
  }

  var contentType =
    match[1].toLowerCase();

  if (contentType === 'image/jpg') {
    contentType = 'image/jpeg';
  }

  var bytes;

  try {
    bytes =
      Utilities.base64Decode(
        base64Data
      );
  } catch (e) {
    throw new Error(
      'Data foto tidak dapat diproses.'
    );
  }

  if (
    bytes.length > MAKS_BYTES
  ) {
    throw new Error(
      'Ukuran foto terlalu besar. Maksimal 2MB.'
    );
  }

  var ekstensi =
    contentType === 'image/png'
      ? 'png'
      : contentType === 'image/webp'
        ? 'webp'
        : 'jpg';

  var namaAman =
    sanitasiNamaFile_(
      namaFileTanpaExt
    );

  var blob =
    Utilities.newBlob(
      bytes,
      contentType,
      namaAman + '.' + ekstensi
    );

  var folder =
    getOrCreateFolder_(
      FOLDER_ABSENSI_NAME
    );

  var file =
    folder.createFile(
      blob
    );

  /*
   * Dipertahankan kompatibel dengan frontend yang akan kita buat.
   * Jangan mengirim URL Drive mentah dari client sebagai sumber foto.
   */
  file.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  return (
    'https://drive.google.com/uc?id=' +
    file.getId()
  );
}


/* ===================================================================
 * 10. STATUS MASUK/PULANG
 * =================================================================== */

function tentukanStatusMasuk_(
  jamStr,
  pengaturan
) {
  var batasToleransi =
    tambahMenitKeJam_(
      pengaturan.jam_masuk,
      pengaturan.toleransi_menit
    );

  return jamStr <= batasToleransi
    ? 'Hadir'
    : 'Terlambat';
}


function tentukanStatusPulang_(
  jamStr,
  pengaturan,
  statusMasuk
) {
  var status =
    String(
      statusMasuk || ''
    );

  if (
    jamStr <
    String(
      pengaturan.jam_pulang
    ) +
    ':00'
  ) {
    if (
      status === 'Terlambat'
    ) {
      return 'Terlambat & Pulang Cepat';
    }

    return 'Pulang Cepat';
  }

  /*
   * Jika status lama Alpa karena trigger dan pegawai kemudian
   * melakukan absen pulang setelah absen masuk, jangan pertahankan
   * Alpa.
   */
  if (
    status === 'Alpa' ||
    !status
  ) {
    return tentukanStatusMasuk_(
      String(
        pengaturan.jam_masuk
      ) + ':00',
      pengaturan
    );
  }

  return status;
}


function tambahMenitKeJam_(
  jamHHmm,
  menit
) {
  var parts =
    String(
      jamHHmm || ''
    ).split(':');

  if (
    parts.length !== 2
  ) {
    throw new Error(
      'Format jam Pengaturan tidak valid.'
    );
  }

  var jam =
    parseInt(
      parts[0],
      10
    );

  var mnt =
    parseInt(
      parts[1],
      10
    );

  var tambahan =
    parseInt(
      menit,
      10
    );

  if (
    !isFinite(jam) ||
    !isFinite(mnt) ||
    !isFinite(tambahan) ||
    jam < 0 ||
    jam > 23 ||
    mnt < 0 ||
    mnt > 59 ||
    tambahan < 0
  ) {
    throw new Error(
      'Nilai waktu Pengaturan tidak valid.'
    );
  }

  var totalMenit =
    jam * 60 +
    mnt +
    tambahan;

  /*
   * Batas toleransi tidak boleh melewati 24 jam.
   */
  totalMenit =
    Math.min(
      totalMenit,
      1439
    );

  var jamHasil =
    Math.floor(
      totalMenit / 60
    );

  var menitHasil =
    totalMenit % 60;

  return (
    pad2_(jamHasil) +
    ':' +
    pad2_(menitHasil) +
    ':59'
  );
}


/* ===================================================================
 * 11. HELPER DATA ABSENSI
 * =================================================================== */

function absensiRowKeObjek_(
  row
) {
  return {
    id:
      row[KOLOM_ABSENSI.ID],
    id_absen:
      row[KOLOM_ABSENSI.ID],
    id_user:
      row[KOLOM_ABSENSI.ID_USER],
    tanggal:
      normalizeTanggal_(
        row[KOLOM_ABSENSI.TANGGAL]
      ),
    jam_masuk:
      row[KOLOM_ABSENSI.JAM_MASUK] || '',
    foto_masuk:
      row[KOLOM_ABSENSI.FOTO_MASUK] || '',
    jam_pulang:
      row[KOLOM_ABSENSI.JAM_PULANG] || '',
    foto_pulang:
      row[KOLOM_ABSENSI.FOTO_PULANG] || '',
    status:
      row[KOLOM_ABSENSI.STATUS] || '',
    lokasi_masuk: {
      lat:
        row[KOLOM_ABSENSI.LAT_MASUK] || '',
      long:
        row[KOLOM_ABSENSI.LONG_MASUK] || ''
    },
    lokasi_pulang: {
      lat:
        row[KOLOM_ABSENSI.LAT_PULANG] || '',
      long:
        row[KOLOM_ABSENSI.LONG_PULANG] || ''
    },
    jarak_masuk_meter:
      row[KOLOM_ABSENSI.JARAK_MASUK] || '',
    jarak_pulang_meter:
      row[KOLOM_ABSENSI.JARAK_PULANG] || ''
  };
}


/* ===================================================================
 * 12. HELPER IZIN UNTUK ALPA
 * =================================================================== */

/**
 * Mengembalikan map:
 *
 * {
 *   "USR-000002": true
 * }
 *
 * untuk izin yang statusnya Disetujui dan mencakup tanggal
 * yang sedang diproses.
 */
function getPetaIzinDisetujuiUntukTanggal_(
  tanggalTarget
) {
  var peta = {};

  /*
   * Bila file Izin.gs belum terpasang, fungsi ini jangan
   * membuat sistem absensi gagal total.
   */
  if (
    typeof KOLOM_IZIN === 'undefined' ||
    typeof STATUS_IZIN_DISETUJUI === 'undefined'
  ) {
    return peta;
  }

  var sheetIzin =
    getSheet(SHEET_IZIN);

  var lastRow =
    sheetIzin.getLastRow();

  if (lastRow < 2) {
    return peta;
  }

  var data =
    sheetIzin
      .getRange(
        2,
        1,
        lastRow - 1,
        12
      )
      .getValues();

  for (
    var i = 0;
    i < data.length;
    i++
  ) {
    var row =
      data[i];

    var status =
      String(
        row[KOLOM_IZIN.STATUS] || ''
      );

    if (
      status !==
      STATUS_IZIN_DISETUJUI
    ) {
      continue;
    }

    var idUser =
      String(
        row[KOLOM_IZIN.ID_USER] || ''
      );

    var mulai =
      normalizeTanggal_(
        row[KOLOM_IZIN.TANGGAL_MULAI]
      );

    var selesai =
      normalizeTanggal_(
        row[KOLOM_IZIN.TANGGAL_SELESAI]
      );

    if (
      idUser &&
      mulai &&
      selesai &&
      tanggalTarget >= mulai &&
      tanggalTarget <= selesai
    ) {
      peta[idUser] = true;
    }
  }

  return peta;
}


/* ===================================================================
 * 13. HELPER TANGGAL/WAKTU
 * =================================================================== */

function formatTanggalAbsensi_(
  tanggal
) {
  return Utilities.formatDate(
    tanggal,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}


function formatJamAbsensi_(
  tanggal
) {
  return Utilities.formatDate(
    tanggal,
    Session.getScriptTimeZone(),
    'HH:mm:ss'
  );
}


function normalizeTanggal_(
  value
) {
  if (
    value instanceof Date &&
    !isNaN(value.getTime())
  ) {
    return formatTanggalAbsensi_(
      value
    );
  }

  var text =
    String(
      value || ''
    ).trim();

  if (
    /^\d{4}-\d{2}-\d{2}$/.test(
      text
    )
  ) {
    return text;
  }

  /*
   * Beberapa data lama mungkin tersimpan sebagai string
   * dengan format Date. Coba parse secara aman.
   */
  var parsed =
    new Date(text);

  if (
    !isNaN(
      parsed.getTime()
    )
  ) {
    return formatTanggalAbsensi_(
      parsed
    );
  }

  return text;
}


function pad2_(
  number
) {
  return number < 10
    ? '0' + number
    : String(number);
}


function sanitasiNamaFile_(
  nama
) {
  var hasil =
    String(
      nama || 'file'
    )
    .replace(
      /[^a-zA-Z0-9._-]/g,
      '_'
    )
    .replace(
      /_+/g,
      '_'
    );

  if (!hasil) {
    hasil = 'file';
  }

  return hasil.substring(
    0,
    100
  );
}
