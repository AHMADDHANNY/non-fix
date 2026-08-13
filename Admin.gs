/**
 * ===================================================================
 * ADMIN.GS
 * Modul khusus administrator:
 *
 * 1. Dashboard admin
 * 2. Rekap absensi seluruh pegawai
 * 3. Detail absensi + foto
 * 4. Manajemen user
 * 5. Reset password user
 * 6. Pengaturan aplikasi
 * 7. Export rekap absensi ke PDF
 *
 * Bergantung pada helper dari:
 * - Code.gs
 * - Auth,gs
 * - Absensi.gs
 * - Izin.gs
 * ===================================================================
 */


/* ===================================================================
 * KONFIGURASI KOLOM USERS
 * =================================================================== */

var KOLOM_USERS_ADMIN = {
  ID: 0,
  NAMA: 1,
  USERNAME: 2,
  PASSWORD_HASH: 3,
  SALT: 4,
  ROLE: 5,
  JABATAN: 6,
  FOTO_PROFIL: 7,
  STATUS_AKTIF: 8,
  CREATED_AT: 9,
  UPDATED_AT: 10
};


/* ===================================================================
 * HELPER AUTORISASI ADMIN
 * =================================================================== */

/**
 * Pastikan token valid dan user adalah admin.
 *
 * @param {string} token
 * @returns {Object} session
 */
function _verifyAdmin_(token) {
  var session = verifySession(token);

  if (!session.valid) {
    throw new Error(session.message || 'Sesi tidak valid. Silakan login ulang.');
  }

  if (session.user.role !== 'admin') {
    throw new Error('Akses ditolak. Hanya administrator yang dapat mengakses fitur ini.');
  }

  return session;
}


/* ===================================================================
 * 1. DASHBOARD ADMIN
 * =================================================================== */

/**
 * Mengambil ringkasan dashboard admin.
 *
 * Data:
 * - total pegawai aktif
 * - total admin aktif
 * - total user aktif
 * - hadir hari ini
 * - terlambat hari ini
 * - belum absen hari ini
 * - izin pending
 * - izin disetujui hari ini
 * - izin ditolak hari ini
 */
function getDashboardAdmin(token) {
  try {
    _verifyAdmin_(token);

    var sheetUsers = getSheet(SHEET_USERS);
    var sheetAbsensi = getSheet(SHEET_ABSENSI);
    var sheetIzin = getSheet(SHEET_IZIN);

    var users = sheetUsers.getDataRange().getValues();
    var absensi = sheetAbsensi.getDataRange().getValues();
    var izin = sheetIzin.getDataRange().getValues();

    var totalPegawai = 0;
    var totalAdmin = 0;
    var totalAktif = 0;

    var pegawaiAktif = {};

    for (var i = 1; i < users.length; i++) {
      var role = String(users[i][KOLOM_USERS_ADMIN.ROLE] || '').toLowerCase();
      var status = String(users[i][KOLOM_USERS_ADMIN.STATUS_AKTIF] || '').toLowerCase();
      var idUser = String(users[i][KOLOM_USERS_ADMIN.ID] || '');

      if (status === 'aktif') {
        totalAktif++;

        if (role === 'pegawai') {
          totalPegawai++;
          pegawaiAktif[idUser] = true;
        }

        if (role === 'admin') {
          totalAdmin++;
        }
      }
    }

    var sekarang = new Date();

    var tanggalHariIni = Utilities.formatDate(
      sekarang,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );

    var hadirHariIni = 0;
    var terlambatHariIni = 0;
    var sudahAbsen = {};

    for (var j = 1; j < absensi.length; j++) {
      var tanggalAbsensi = String(absensi[j][KOLOM_ABSENSI.TANGGAL] || '');

      if (tanggalAbsensi !== tanggalHariIni) {
        continue;
      }

      var idUserAbsensi = String(
        absensi[j][KOLOM_ABSENSI.ID_USER] || ''
      );

      if (!pegawaiAktif[idUserAbsensi]) {
        continue;
      }

      sudahAbsen[idUserAbsensi] = true;

      var statusAbsensi = String(
        absensi[j][KOLOM_ABSENSI.STATUS] || ''
      );

      if (statusAbsensi === 'Terlambat') {
        terlambatHariIni++;
      } else {
        hadirHariIni++;
      }
    }

    var belumAbsenHariIni = Math.max(
      0,
      totalPegawai - Object.keys(sudahAbsen).length
    );

    var izinPending = 0;
    var izinDisetujuiHariIni = 0;
    var izinDitolakHariIni = 0;

    for (var k = 1; k < izin.length; k++) {
      var statusIzin = String(izin[k][KOLOM_IZIN.STATUS] || '');

      if (statusIzin === STATUS_IZIN_PENDING) {
        izinPending++;
      }

      var diprosesAt = izin[k][KOLOM_IZIN.DIPROSES_AT];

      if (diprosesAt) {
        var tanggalProses = Utilities.formatDate(
          new Date(diprosesAt),
          Session.getScriptTimeZone(),
          'yyyy-MM-dd'
        );

        if (tanggalProses === tanggalHariIni) {
          if (statusIzin === STATUS_IZIN_DISETUJUI) {
            izinDisetujuiHariIni++;
          }

          if (statusIzin === STATUS_IZIN_DITOLAK) {
            izinDitolakHariIni++;
          }
        }
      }
    }

    return jsonResponse({
      success: true,
      data: {
        tanggal: tanggalHariIni,
        total_pegawai: totalPegawai,
        total_admin: totalAdmin,
        total_user_aktif: totalAktif,
        hadir_hari_ini: hadirHariIni,
        terlambat_hari_ini: terlambatHariIni,
        belum_absen_hari_ini: belumAbsenHariIni,
        izin_pending: izinPending,
        izin_disetujui_hari_ini: izinDisetujuiHariIni,
        izin_ditolak_hari_ini: izinDitolakHariIni
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
 * 2. REKAP ABSENSI ADMIN
 * =================================================================== */

/**
 * Mengambil rekap absensi seluruh pegawai.
 *
 * @param {string} token
 * @param {number} page
 * @param {number} pageSize
 * @param {string} tanggalFilter - yyyy-MM-dd
 * @param {string} bulanFilter - yyyy-MM
 * @param {string} idUserFilter
 * @param {string} statusFilter
 */
function getRekapAbsensiAdmin(
  token,
  page,
  pageSize,
  tanggalFilter,
  bulanFilter,
  idUserFilter,
  statusFilter
) {
  try {
    _verifyAdmin_(token);

    page = page && page > 0 ? Number(page) : 1;
    pageSize = pageSize && pageSize > 0 ? Number(pageSize) : 20;

    if (pageSize > 100) {
      pageSize = 100;
    }

    var sheetUsers = getSheet(SHEET_USERS);
    var sheetAbsensi = getSheet(SHEET_ABSENSI);

    var users = sheetUsers.getDataRange().getValues();
    var absensi = sheetAbsensi.getDataRange().getValues();

    var petaUser = {};

    for (var i = 1; i < users.length; i++) {
      var idUser = String(users[i][KOLOM_USERS_ADMIN.ID] || '');

      petaUser[idUser] = {
        id_user: idUser,
        nama: users[i][KOLOM_USERS_ADMIN.NAMA] || '',
        username: users[i][KOLOM_USERS_ADMIN.USERNAME] || '',
        role: users[i][KOLOM_USERS_ADMIN.ROLE] || '',
        jabatan: users[i][KOLOM_USERS_ADMIN.JABATAN] || '',
        foto_profil: users[i][KOLOM_USERS_ADMIN.FOTO_PROFIL] || '',
        status_aktif: users[i][KOLOM_USERS_ADMIN.STATUS_AKTIF] || ''
      };
    }

    var hasil = [];

    for (var j = 1; j < absensi.length; j++) {
      var row = absensi[j];

      var tanggal = String(
        row[KOLOM_ABSENSI.TANGGAL] || ''
      );

      var userId = String(
        row[KOLOM_ABSENSI.ID_USER] || ''
      );

      var status = String(
        row[KOLOM_ABSENSI.STATUS] || ''
      );

      if (tanggalFilter && tanggal !== tanggalFilter) {
        continue;
      }

      if (
        bulanFilter &&
        tanggal.indexOf(bulanFilter) !== 0
      ) {
        continue;
      }

      if (
        idUserFilter &&
        userId !== idUserFilter
      ) {
        continue;
      }

      if (
        statusFilter &&
        status !== statusFilter
      ) {
        continue;
      }

      var user = petaUser[userId] || {};

      hasil.push({
        id_absen: row[KOLOM_ABSENSI.ID],
        id_user: userId,
        nama: user.nama || 'User tidak ditemukan',
        username: user.username || '',
        jabatan: user.jabatan || '',
        tanggal: tanggal,
        jam_masuk: row[KOLOM_ABSENSI.JAM_MASUK] || '',
        foto_masuk: row[KOLOM_ABSENSI.FOTO_MASUK] || '',
        lat_masuk: row[KOLOM_ABSENSI.LAT_MASUK] || '',
        long_masuk: row[KOLOM_ABSENSI.LONG_MASUK] || '',
        jarak_masuk_meter: row[KOLOM_ABSENSI.JARAK_MASUK] || '',
        jam_pulang: row[KOLOM_ABSENSI.JAM_PULANG] || '',
        foto_pulang: row[KOLOM_ABSENSI.FOTO_PULANG] || '',
        lat_pulang: row[KOLOM_ABSENSI.LAT_PULANG] || '',
        long_pulang: row[KOLOM_ABSENSI.LONG_PULANG] || '',
        jarak_pulang_meter: row[KOLOM_ABSENSI.JARAK_PULANG] || '',
        status: status
      });
    }

    // Terbaru dulu
    hasil.sort(function(a, b) {
      if (a.tanggal === b.tanggal) {
        return String(b.jam_masuk).localeCompare(
          String(a.jam_masuk)
        );
      }

      return a.tanggal < b.tanggal ? 1 : -1;
    });

    var total = hasil.length;

    var start = (page - 1) * pageSize;

    var potongan = hasil.slice(
      start,
      start + pageSize
    );

    return jsonResponse({
      success: true,
      data: potongan,
      pagination: {
        page: page,
        pageSize: pageSize,
        total: total,
        totalPage: Math.ceil(total / pageSize)
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
 * 3. DETAIL ABSENSI
 * =================================================================== */

/**
 * Mengambil satu detail absensi berdasarkan ID.
 */
function getDetailAbsensiAdmin(token, idAbsen) {
  try {
    _verifyAdmin_(token);

    if (!idAbsen) {
      throw new Error('ID absensi wajib diisi.');
    }

    var sheetUsers = getSheet(SHEET_USERS);
    var sheetAbsensi = getSheet(SHEET_ABSENSI);

    var users = sheetUsers.getDataRange().getValues();
    var absensi = sheetAbsensi.getDataRange().getValues();

    var user = null;

    for (var i = 1; i < users.length; i++) {
      if (
        String(users[i][KOLOM_USERS_ADMIN.ID]) ===
        String(
          absensi.length > 0 ? '' : ''
        )
      ) {
        // Tidak digunakan.
      }
    }

    for (var j = 1; j < absensi.length; j++) {
      var row = absensi[j];

      if (
        String(row[KOLOM_ABSENSI.ID]) !==
        String(idAbsen)
      ) {
        continue;
      }

      var idUser = String(
        row[KOLOM_ABSENSI.ID_USER] || ''
      );

      for (var k = 1; k < users.length; k++) {
        if (
          String(users[k][KOLOM_USERS_ADMIN.ID]) ===
          idUser
        ) {
          user = {
            id_user: users[k][KOLOM_USERS_ADMIN.ID],
            nama: users[k][KOLOM_USERS_ADMIN.NAMA],
            username: users[k][KOLOM_USERS_ADMIN.USERNAME],
            jabatan: users[k][KOLOM_USERS_ADMIN.JABATAN],
            foto_profil: users[k][KOLOM_USERS_ADMIN.FOTO_PROFIL]
          };
          break;
        }
      }

      return jsonResponse({
        success: true,
        data: {
          id_absen: row[KOLOM_ABSENSI.ID],
          user: user,
          tanggal: row[KOLOM_ABSENSI.TANGGAL],

          masuk: {
            jam: row[KOLOM_ABSENSI.JAM_MASUK],
            foto: row[KOLOM_ABSENSI.FOTO_MASUK],
            lat: row[KOLOM_ABSENSI.LAT_MASUK],
            long: row[KOLOM_ABSENSI.LONG_MASUK],
            jarak_meter: row[KOLOM_ABSENSI.JARAK_MASUK]
          },

          pulang: {
            jam: row[KOLOM_ABSENSI.JAM_PULANG],
            foto: row[KOLOM_ABSENSI.FOTO_PULANG],
            lat: row[KOLOM_ABSENSI.LAT_PULANG],
            long: row[KOLOM_ABSENSI.LONG_PULANG],
            jarak_meter: row[KOLOM_ABSENSI.JARAK_PULANG]
          },

          status: row[KOLOM_ABSENSI.STATUS]
        }
      });
    }

    throw new Error('Data absensi tidak ditemukan.');

  } catch (e) {
    return jsonResponse({
      success: false,
      message: e.message
    });
  }
}


/* ===================================================================
 * 4. DAFTAR USER
 * =================================================================== */

/**
 * Mengambil daftar user untuk halaman manajemen user.
 *
 * Password dan salt TIDAK pernah dikirim ke frontend.
 */
function getDaftarUsers(token, page, pageSize, search, roleFilter, statusFilter) {
  try {
    _verifyAdmin_(token);

    page = page && page > 0 ? Number(page) : 1;
    pageSize = pageSize && pageSize > 0 ? Number(pageSize) : 20;

    if (pageSize > 100) {
      pageSize = 100;
    }

    search = String(search || '').trim().toLowerCase();

    var sheet = getSheet(SHEET_USERS);
    var data = sheet.getDataRange().getValues();

    var hasil = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];

      var id = String(row[KOLOM_USERS_ADMIN.ID] || '');
      var nama = String(row[KOLOM_USERS_ADMIN.NAMA] || '');
      var username = String(row[KOLOM_USERS_ADMIN.USERNAME] || '');
      var role = String(row[KOLOM_USERS_ADMIN.ROLE] || '');
      var jabatan = String(row[KOLOM_USERS_ADMIN.JABATAN] || '');
      var status = String(row[KOLOM_USERS_ADMIN.STATUS_AKTIF] || '');

      if (search) {
        var teksCari = (
          id + ' ' +
          nama + ' ' +
          username + ' ' +
          jabatan
        ).toLowerCase();

        if (teksCari.indexOf(search) === -1) {
          continue;
        }
      }

      if (
        roleFilter &&
        role !== roleFilter
      ) {
        continue;
      }

      if (
        statusFilter &&
        status !== statusFilter
      ) {
        continue;
      }

      hasil.push({
        id_user: id,
        nama: nama,
        username: username,
        role: role,
        jabatan: jabatan,
        foto_profil: row[KOLOM_USERS_ADMIN.FOTO_PROFIL] || '',
        status_aktif: status,
        created_at: row[KOLOM_USERS_ADMIN.CREATED_AT] || '',
        updated_at: row[KOLOM_USERS_ADMIN.UPDATED_AT] || ''
      });
    }

    hasil.sort(function(a, b) {
      return String(a.nama).localeCompare(
        String(b.nama)
      );
    });

    var total = hasil.length;
    var start = (page - 1) * pageSize;

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
        totalPage: Math.ceil(total / pageSize)
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
 * 5. TAMBAH USER
 * =================================================================== */

/**
 * Tambah user baru.
 *
 * @param {string} token
 * @param {string} nama
 * @param {string} username
 * @param {string} password
 * @param {string} role
 * @param {string} jabatan
 * @param {string} fotoProfil
 */
function tambahUser(
  token,
  nama,
  username,
  password,
  role,
  jabatan,
  fotoProfil
) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    _verifyAdmin_(token);

    nama = String(nama || '').trim();
    username = String(username || '').trim();
    password = String(password || '');
    role = String(role || '').trim().toLowerCase();
    jabatan = String(jabatan || '').trim();
    fotoProfil = String(fotoProfil || '').trim();

    if (!nama) {
      throw new Error('Nama wajib diisi.');
    }

    if (!username) {
      throw new Error('Username wajib diisi.');
    }

    if (!/^[a-zA-Z0-9._-]{3,30}$/.test(username)) {
      throw new Error(
        'Username hanya boleh berisi huruf, angka, titik, garis bawah, atau tanda minus (3-30 karakter).'
      );
    }

    if (!password || password.length < 6) {
      throw new Error(
        'Password minimal 6 karakter.'
      );
    }

    if (
      role !== 'admin' &&
      role !== 'pegawai'
    ) {
      throw new Error(
        'Role hanya boleh admin atau pegawai.'
      );
    }

    var sheet = getSheet(SHEET_USERS);
    var data = sheet.getDataRange().getValues();

    // Cek username unik
    for (var i = 1; i < data.length; i++) {
      var usernameLama = String(
        data[i][KOLOM_USERS_ADMIN.USERNAME] || ''
      );

      if (
        usernameLama.toLowerCase() ===
        username.toLowerCase()
      ) {
        throw new Error(
          'Username tersebut sudah digunakan.'
        );
      }
    }

    var salt = Utilities.getUuid();
    var hash = hashPassword(
      password,
      salt
    );

    var idBaru = generateSequentialId(
      'USR',
      SHEET_USERS,
      KOLOM_USERS_ADMIN.ID
    );

    var now = new Date();

    sheet.appendRow([
      idBaru,
      nama,
      username,
      hash,
      salt,
      role,
      jabatan,
      fotoProfil,
      'aktif',
      now,
      now
    ]);

    return jsonResponse({
      success: true,
      message: 'User berhasil ditambahkan.',
      data: {
        id_user: idBaru,
        nama: nama,
        username: username,
        role: role,
        jabatan: jabatan,
        status_aktif: 'aktif'
      }
    });

  } catch (e) {
    return jsonResponse({
      success: false,
      message: e.message
    });

  } finally {
    lock.releaseLock();
  }
}


/* ===================================================================
 * 6. EDIT USER
 * =================================================================== */

/**
 * Edit data user.
 *
 * Password tidak diubah melalui fungsi ini.
 * Gunakan resetPasswordUser().
 */
function editUser(
  token,
  idUser,
  nama,
  username,
  role,
  jabatan,
  fotoProfil,
  statusAktif
) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    var session = _verifyAdmin_(token);

    idUser = String(idUser || '').trim();
    nama = String(nama || '').trim();
    username = String(username || '').trim();
    role = String(role || '').trim().toLowerCase();
    jabatan = String(jabatan || '').trim();
    fotoProfil = String(fotoProfil || '').trim();
    statusAktif = String(statusAktif || '').trim().toLowerCase();

    if (!idUser) {
      throw new Error('ID user wajib diisi.');
    }

    if (!nama) {
      throw new Error('Nama wajib diisi.');
    }

    if (!username) {
      throw new Error('Username wajib diisi.');
    }

    if (
      role !== 'admin' &&
      role !== 'pegawai'
    ) {
      throw new Error(
        'Role hanya boleh admin atau pegawai.'
      );
    }

    if (
      statusAktif !== 'aktif' &&
      statusAktif !== 'nonaktif'
    ) {
      throw new Error(
        'Status hanya boleh aktif atau nonaktif.'
      );
    }

    // Jangan izinkan admin menonaktifkan dirinya sendiri
    if (
      idUser === String(session.user.id_user) &&
      statusAktif !== 'aktif'
    ) {
      throw new Error(
        'Anda tidak dapat menonaktifkan akun admin yang sedang digunakan.'
      );
    }

    var sheet = getSheet(SHEET_USERS);
    var data = sheet.getDataRange().getValues();

    var rowIndex = -1;

    for (var i = 1; i < data.length; i++) {
      if (
        String(data[i][KOLOM_USERS_ADMIN.ID]) ===
        idUser
      ) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error('User tidak ditemukan.');
    }

    // Cek username dipakai user lain
    for (var j = 1; j < data.length; j++) {
      if (j === rowIndex) {
        continue;
      }

      var usernameLama = String(
        data[j][KOLOM_USERS_ADMIN.USERNAME] || ''
      );

      if (
        usernameLama.toLowerCase() ===
        username.toLowerCase()
      ) {
        throw new Error(
          'Username tersebut sudah digunakan user lain.'
        );
      }
    }

    var row = rowIndex + 1;
    var now = new Date();

    sheet.getRange(
      row,
      KOLOM_USERS_ADMIN.NAMA + 1
    ).setValue(nama);

    sheet.getRange(
      row,
      KOLOM_USERS_ADMIN.USERNAME + 1
    ).setValue(username);

    sheet.getRange(
      row,
      KOLOM_USERS_ADMIN.ROLE + 1
    ).setValue(role);

    sheet.getRange(
      row,
      KOLOM_USERS_ADMIN.JABATAN + 1
    ).setValue(jabatan);

    sheet.getRange(
      row,
      KOLOM_USERS_ADMIN.FOTO_PROFIL + 1
    ).setValue(fotoProfil);

    sheet.getRange(
      row,
      KOLOM_USERS_ADMIN.STATUS_AKTIF + 1
    ).setValue(statusAktif);

    sheet.getRange(
      row,
      KOLOM_USERS_ADMIN.UPDATED_AT + 1
    ).setValue(now);

    return jsonResponse({
      success: true,
      message: 'Data user berhasil diperbarui.'
    });

  } catch (e) {
    return jsonResponse({
      success: false,
      message: e.message
    });

  } finally {
    lock.releaseLock();
  }
}


/* ===================================================================
 * 7. HAPUS USER
 * =================================================================== */

/**
 * "Hapus" user dengan cara soft delete:
 * status_aktif = nonaktif.
 *
 * Data absensi dan izin tetap aman sehingga history
 * tidak kehilangan relasi id_user.
 */
function hapusUser(token, idUser) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    var session = _verifyAdmin_(token);

    idUser = String(idUser || '').trim();

    if (!idUser) {
      throw new Error('ID user wajib diisi.');
    }

    if (
      idUser ===
      String(session.user.id_user)
    ) {
      throw new Error(
        'Anda tidak dapat menghapus akun yang sedang digunakan.'
      );
    }

    var sheet = getSheet(SHEET_USERS);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (
        String(data[i][KOLOM_USERS_ADMIN.ID]) ===
        idUser
      ) {
        sheet.getRange(
          i + 1,
          KOLOM_USERS_ADMIN.STATUS_AKTIF + 1
        ).setValue('nonaktif');

        sheet.getRange(
          i + 1,
          KOLOM_USERS_ADMIN.UPDATED_AT + 1
        ).setValue(new Date());

        return jsonResponse({
          success: true,
          message:
            'User berhasil dinonaktifkan.'
        });
      }
    }

    throw new Error('User tidak ditemukan.');

  } catch (e) {
    return jsonResponse({
      success: false,
      message: e.message
    });

  } finally {
    lock.releaseLock();
  }
}


/* ===================================================================
 * 8. RESET PASSWORD USER
 * =================================================================== */

/**
 * Admin mengatur password baru milik user lain.
 */
function resetPasswordUser(
  token,
  idUser,
  passwordBaru
) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    _verifyAdmin_(token);

    idUser = String(idUser || '').trim();
    passwordBaru = String(passwordBaru || '');

    if (!idUser) {
      throw new Error('ID user wajib diisi.');
    }

    if (
      !passwordBaru ||
      passwordBaru.length < 6
    ) {
      throw new Error(
        'Password baru minimal 6 karakter.'
      );
    }

    var sheet = getSheet(SHEET_USERS);
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (
        String(data[i][KOLOM_USERS_ADMIN.ID]) ===
        idUser
      ) {
        var saltBaru = Utilities.getUuid();

        var hashBaru = hashPassword(
          passwordBaru,
          saltBaru
        );

        var row = i + 1;

        sheet.getRange(
          row,
          KOLOM_USERS_ADMIN.PASSWORD_HASH + 1
        ).setValue(hashBaru);

        sheet.getRange(
          row,
          KOLOM_USERS_ADMIN.SALT + 1
        ).setValue(saltBaru);

        sheet.getRange(
          row,
          KOLOM_USERS_ADMIN.UPDATED_AT + 1
        ).setValue(new Date());

        return jsonResponse({
          success: true,
          message:
            'Password user berhasil direset.'
        });
      }
    }

    throw new Error('User tidak ditemukan.');

  } catch (e) {
    return jsonResponse({
      success: false,
      message: e.message
    });

  } finally {
    lock.releaseLock();
  }
}


/* ===================================================================
 * 9. PENGATURAN APLIKASI
 * =================================================================== */

/**
 * Mengambil pengaturan aplikasi.
 */
function getPengaturanAdmin(token) {
  try {
    _verifyAdmin_(token);

    var pengaturan = getPengaturanCache();

    return jsonResponse({
      success: true,
      data: pengaturan
    });

  } catch (e) {
    return jsonResponse({
      success: false,
      message: e.message
    });
  }
}


/**
 * Simpan pengaturan aplikasi.
 */
function simpanPengaturan(
  token,
  namaApp,
  logoUrl,
  latKantor,
  longKantor,
  radiusMeter,
  jamMasuk,
  jamPulang,
  toleransiMenit
) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    _verifyAdmin_(token);

    namaApp = String(namaApp || '').trim();
    logoUrl = String(logoUrl || '').trim();

    latKantor = parseFloat(latKantor);
    longKantor = parseFloat(longKantor);
    radiusMeter = parseFloat(radiusMeter);
    toleransiMenit = parseInt(
      toleransiMenit,
      10
    );

    jamMasuk = String(jamMasuk || '').trim();
    jamPulang = String(jamPulang || '').trim();

    if (!namaApp) {
      throw new Error(
        'Nama aplikasi wajib diisi.'
      );
    }

    if (
      isNaN(latKantor) ||
      latKantor < -90 ||
      latKantor > 90
    ) {
      throw new Error(
        'Latitude kantor tidak valid.'
      );
    }

    if (
      isNaN(longKantor) ||
      longKantor < -180 ||
      longKantor > 180
    ) {
      throw new Error(
        'Longitude kantor tidak valid.'
      );
    }

    if (
      isNaN(radiusMeter) ||
      radiusMeter <= 0 ||
      radiusMeter > 10000
    ) {
      throw new Error(
        'Radius harus lebih dari 0 dan maksimal 10.000 meter.'
      );
    }

    if (
      !/^(?:[01]\d|2[0-3]):[0-5]\d$/.test(
        jamMasuk
      )
    ) {
      throw new Error(
        'Format jam masuk harus HH:mm.'
      );
    }

    if (
      !/^(?:[01]\d|2[0-3]):[0-5]\d$/.test(
        jamPulang
      )
    ) {
      throw new Error(
        'Format jam pulang harus HH:mm.'
      );
    }

    if (
      isNaN(toleransiMenit) ||
      toleransiMenit < 0 ||
      toleransiMenit > 300
    ) {
      throw new Error(
        'Toleransi keterlambatan harus 0-300 menit.'
      );
    }

    var sheet = getSheet(
      SHEET_PENGATURAN
    );

    if (sheet.getLastRow() < 2) {
      sheet.appendRow([
        namaApp,
        logoUrl,
        latKantor,
        longKantor,
        radiusMeter,
        jamMasuk,
        jamPulang,
        toleransiMenit
      ]);
    } else {
      sheet.getRange(2, 1, 1, 8).setValues([
        [
          namaApp,
          logoUrl,
          latKantor,
          longKantor,
          radiusMeter,
          jamMasuk,
          jamPulang,
          toleransiMenit
        ]
      ]);
    }

    // Sangat penting:
    // hapus cache supaya Absensi.gs langsung menggunakan
    // pengaturan terbaru.
    bersihkanCachePengaturan();

    return jsonResponse({
      success: true,
      message:
        'Pengaturan berhasil disimpan.',
      data: {
        nama_app: namaApp,
        logo_url: logoUrl,
        lat_kantor: latKantor,
        long_kantor: longKantor,
        radius_meter: radiusMeter,
        jam_masuk: jamMasuk,
        jam_pulang: jamPulang,
        toleransi_menit: toleransiMenit
      }
    });

  } catch (e) {
    return jsonResponse({
      success: false,
      message: e.message
    });

  } finally {
    lock.releaseLock();
  }
}


/* ===================================================================
 * 10. EXPORT REKAP ABSENSI KE PDF
 * =================================================================== */

/**
 * Export rekap absensi ke PDF menggunakan Google Docs sementara.
 *
 * Filter:
 * - tanggalFilter
 * - bulanFilter
 * - idUserFilter
 *
 * Return:
 * {
 *   success: true,
 *   data: {
 *      fileName,
 *      url
 *   }
 * }
 */
function exportRekapAbsensiPDF(
  token,
  tanggalFilter,
  bulanFilter,
  idUserFilter
) {
  try {
    _verifyAdmin_(token);

    var sheetUsers = getSheet(SHEET_USERS);
    var sheetAbsensi = getSheet(SHEET_ABSENSI);

    var users = sheetUsers.getDataRange().getValues();
    var absensi = sheetAbsensi.getDataRange().getValues();

    var petaUser = {};

    for (var i = 1; i < users.length; i++) {
      petaUser[
        String(users[i][KOLOM_USERS_ADMIN.ID])
      ] = {
        nama: users[i][KOLOM_USERS_ADMIN.NAMA] || '',
        jabatan: users[i][KOLOM_USERS_ADMIN.JABATAN] || ''
      };
    }

    var hasil = [];

    for (var j = 1; j < absensi.length; j++) {
      var row = absensi[j];

      var tanggal = String(
        row[KOLOM_ABSENSI.TANGGAL] || ''
      );

      var idUser = String(
        row[KOLOM_ABSENSI.ID_USER] || ''
      );

      if (
        tanggalFilter &&
        tanggal !== tanggalFilter
      ) {
        continue;
      }

      if (
        bulanFilter &&
        tanggal.indexOf(bulanFilter) !== 0
      ) {
        continue;
      }

      if (
        idUserFilter &&
        idUser !== idUserFilter
      ) {
        continue;
      }

      var user = petaUser[idUser] || {
        nama: 'User tidak ditemukan',
        jabatan: ''
      };

      hasil.push({
        id_absen: row[KOLOM_ABSENSI.ID],
        nama: user.nama,
        jabatan: user.jabatan,
        tanggal: tanggal,
        jam_masuk: row[KOLOM_ABSENSI.JAM_MASUK] || '-',
        jam_pulang: row[KOLOM_ABSENSI.JAM_PULANG] || '-',
        status: row[KOLOM_ABSENSI.STATUS] || '-',
        jarak_masuk: row[KOLOM_ABSENSI.JARAK_MASUK] || '-',
        jarak_pulang: row[KOLOM_ABSENSI.JARAK_PULANG] || '-'
      });
    }

    hasil.sort(function(a, b) {
      if (a.tanggal === b.tanggal) {
        return String(a.nama).localeCompare(
          String(b.nama)
        );
      }

      return a.tanggal.localeCompare(
        b.tanggal
      );
    });

    var namaPeriode = 'Semua_Data';

    if (tanggalFilter) {
      namaPeriode = tanggalFilter;
    } else if (bulanFilter) {
      namaPeriode = bulanFilter;
    }

    var namaUser = '';

    if (idUserFilter) {
      namaUser =
        petaUser[idUserFilter] ?
        '_' + petaUser[idUserFilter].nama :
        '_' + idUserFilter;
    }

    var fileName =
      'Rekap_Absensi_' +
      namaPeriode +
      namaUser +
      '_' +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        'yyyyMMdd_HHmmss'
      ) +
      '.pdf';

    /*
     * Buat Google Docs sementara
     */
    var doc = DocumentApp.create(
      fileName.replace('.pdf', '')
    );

    var body = doc.getBody();

    body.appendParagraph(
      'REKAP ABSENSI KARYAWAN'
    )
      .setHeading(
        DocumentApp.ParagraphHeading.TITLE
      );

    var pengaturan = getPengaturanCache();

    body.appendParagraph(
      String(
        pengaturan.nama_app ||
        'Aplikasi Absensi Karyawan'
      )
    )
      .setHeading(
        DocumentApp.ParagraphHeading.HEADING2
      );

    var filterText = 'Periode: ';

    if (tanggalFilter) {
      filterText += tanggalFilter;
    } else if (bulanFilter) {
      filterText += bulanFilter;
    } else {
      filterText += 'Semua';
    }

    body.appendParagraph(filterText);

    if (idUserFilter) {
      body.appendParagraph(
        'Pegawai: ' +
        (
          petaUser[idUserFilter] ?
          petaUser[idUserFilter].nama :
          idUserFilter
        )
      );
    }

    body.appendParagraph(
      'Dicetak: ' +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        'dd-MM-yyyy HH:mm:ss'
      )
    );

    body.appendParagraph('');

    body.appendParagraph(
      'Total data: ' +
      hasil.length
    );

    if (hasil.length === 0) {
      body.appendParagraph(
        'Tidak ada data absensi untuk filter yang dipilih.'
      );
    } else {
      var tableData = [
        [
          'No',
          'Tanggal',
          'Nama',
          'Jabatan',
          'Masuk',
          'Pulang',
          'Status',
          'Jarak Masuk',
          'Jarak Pulang'
        ]
      ];

      for (var k = 0; k < hasil.length; k++) {
        tableData.push([
          k + 1,
          hasil[k].tanggal,
          hasil[k].nama,
          hasil[k].jabatan,
          hasil[k].jam_masuk,
          hasil[k].jam_pulang,
          hasil[k].status,
          hasil[k].jarak_masuk === '' ?
            '-' :
            String(hasil[k].jarak_masuk) + ' m',
          hasil[k].jarak_pulang === '' ?
            '-' :
            String(hasil[k].jarak_pulang) + ' m'
        ]);
      }

      var table = body.appendTable(
        tableData
      );

      // Bold header
      var headerRow = table.getRow(0);

      for (
        var col = 0;
        col < headerRow.getNumCells();
        col++
      ) {
        headerRow
          .getCell(col)
          .editAsText()
          .setBold(true);
      }
    }

    doc.saveAndClose();

    /*
     * Ambil file Google Docs sementara,
     * convert ke PDF.
     */
    var docFile = DriveApp.getFileById(
      doc.getId()
    );

    var pdfBlob = docFile
      .getAs(MimeType.PDF)
      .setName(fileName);

    var folder = getOrCreateFolder_(
      FOLDER_REKAP_NAME
    );

    var pdfFile = folder.createFile(
      pdfBlob
    );

    pdfFile.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );

    /*
     * Hapus Google Docs sementara.
     * Yang dipertahankan hanya PDF.
     */
    docFile.setTrashed(true);

    return jsonResponse({
      success: true,
      message:
        'Rekap berhasil diekspor ke PDF.',
      data: {
        fileName: pdfFile.getName(),
        fileId: pdfFile.getId(),
        url:
          'https://drive.google.com/uc?id=' +
          pdfFile.getId()
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
 * 11. DAFTAR PEGAWAI UNTUK DROPDOWN ADMIN
 * =================================================================== */

/**
 * Mengambil daftar pegawai aktif.
 *
 * Dipakai frontend untuk:
 * - filter rekap
 * - filter izin
 * - dropdown user
 */
function getDaftarPegawaiAktif(token) {
  try {
    _verifyAdmin_(token);

    var sheet = getSheet(SHEET_USERS);
    var data = sheet.getDataRange().getValues();

    var hasil = [];

    for (var i = 1; i < data.length; i++) {
      var role = String(
        data[i][KOLOM_USERS_ADMIN.ROLE] || ''
      ).toLowerCase();

      var status = String(
        data[i][KOLOM_USERS_ADMIN.STATUS_AKTIF] || ''
      ).toLowerCase();

      if (
        role !== 'pegawai' ||
        status !== 'aktif'
      ) {
        continue;
      }

      hasil.push({
        id_user: data[i][KOLOM_USERS_ADMIN.ID],
        nama: data[i][KOLOM_USERS_ADMIN.NAMA],
        username: data[i][KOLOM_USERS_ADMIN.USERNAME],
        jabatan: data[i][KOLOM_USERS_ADMIN.JABATAN]
      });
    }

    hasil.sort(function(a, b) {
      return String(a.nama).localeCompare(
        String(b.nama)
      );
    });

    return jsonResponse({
      success: true,
      data: hasil
    });

  } catch (e) {
    return jsonResponse({
      success: false,
      message: e.message
    });
  }
}
