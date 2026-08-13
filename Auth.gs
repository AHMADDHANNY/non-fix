/**
 * ===================================================================
 * AUTH.GS
 * Autentikasi, session token, logout, dan perubahan password.
 *
 * Dependency:
 * - Code.gs
 *
 * Sheet:
 * Users
 * Sessions
 *
 * Catatan:
 * - Password disimpan SHA-256 + salt.
 * - Session disimpan di sheet Sessions.
 * - Session berlaku sesuai SESSION_DURASI_JAM di Code.gs.
 * - Operasi tulis menggunakan LockService agar aman saat
 *   banyak pengguna melakukan request bersamaan.
 * ===================================================================
 */


/* ===================================================================
 * KONSTANTA INDEX USERS
 * =================================================================== */

const AUTH_USER = {
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
 * LOGIN
 * =================================================================== */

/**
 * Login menggunakan username + password.
 *
 * @param {string} username
 * @param {string} password
 * @returns {Object}
 */
function login(username, password) {
  try {
    username = String(username || '').trim();
    password = String(password || '');

    if (!username || !password) {
      throw new Error(
        'Username dan password wajib diisi.'
      );
    }

    const sheet = getSheet(SHEET_USERS);
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      throw new Error(
        'Belum ada akun pengguna.'
      );
    }

    /*
     * Kita hanya membaca kolom Users yang diperlukan.
     * Tidak perlu mengambil kolom di luar A:K.
     */
    const data = sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        11
      )
      .getValues();

    for (let i = 0; i < data.length; i++) {
      const row = data[i];

      const rowUsername =
        String(
          row[AUTH_USER.USERNAME] || ''
        ).trim();

      if (
        rowUsername.toLowerCase() !==
        username.toLowerCase()
      ) {
        continue;
      }

      /*
       * Cek status akun terlebih dahulu.
       */
      const statusAktif =
        String(
          row[AUTH_USER.STATUS_AKTIF] || ''
        ).toLowerCase();

      if (statusAktif !== 'aktif') {
        throw new Error(
          'Akun Anda tidak aktif. Hubungi admin.'
        );
      }

      const salt =
        String(
          row[AUTH_USER.SALT] || ''
        );

      const hashTersimpan =
        String(
          row[AUTH_USER.PASSWORD_HASH] || ''
        );

      if (!salt || !hashTersimpan) {
        throw new Error(
          'Data autentikasi akun tidak valid. Hubungi admin.'
        );
      }

      const hashInput =
        hashPassword(
          password,
          salt
        );

      if (hashInput !== hashTersimpan) {
        throw new Error(
          'Username atau password salah.'
        );
      }

      /*
       * Jangan pernah mengirim password hash/salt
       * ke frontend.
       */
      const user = {
        id_user:
          row[AUTH_USER.ID],
        nama:
          row[AUTH_USER.NAMA],
        username:
          row[AUTH_USER.USERNAME],
        role:
          row[AUTH_USER.ROLE],
        jabatan:
          row[AUTH_USER.JABATAN],
        foto_profil:
          row[AUTH_USER.FOTO_PROFIL]
      };

      /*
       * Session baru dibuat setelah autentikasi berhasil.
       */
      const token =
        buatSessionToken_(
          user.id_user
        );

      return jsonResponse({
        success: true,
        token: token,
        user: user
      });
    }

    /*
     * Jangan membedakan username tidak ada
     * dengan password salah.
     */
    throw new Error(
      'Username atau password salah.'
    );

  } catch (e) {
    return jsonResponse({
      success: false,
      message: e.message
    });
  }
}


/* ===================================================================
 * LOGOUT
 * =================================================================== */

/**
 * Logout dengan menghapus token dari Sessions.
 *
 * Menggunakan LockService karena beberapa request bisa terjadi
 * bersamaan, terutama ketika browser melakukan retry.
 */
function logout(token) {
  const lock =
    LockService.getScriptLock();

  try {
    token = String(token || '').trim();

    /*
     * Tidak ada token = sudah dianggap logout.
     */
    if (!token) {
      return jsonResponse({
        success: true,
        message: 'Sudah logout.'
      });
    }

    lock.waitLock(10000);

    const sheet =
      getSheet(SHEET_SESSIONS);

    const lastRow =
      sheet.getLastRow();

    if (lastRow >= 2) {
      const data =
        sheet
          .getRange(
            2,
            1,
            lastRow - 1,
            4
          )
          .getValues();

      /*
       * Cari dari bawah ke atas supaya aman
       * jika ada lebih dari satu operasi.
       */
      for (
        let i = data.length - 1;
        i >= 0;
        i--
      ) {
        if (
          String(data[i][0]) === token
        ) {
          sheet.deleteRow(i + 2);
          break;
        }
      }
    }

    return jsonResponse({
      success: true,
      message: 'Berhasil logout.'
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
 * CHANGE PASSWORD
 * =================================================================== */

/**
 * Ganti password milik user yang sedang login.
 *
 * Admin reset password user lain menggunakan:
 * Admin.gs -> resetPasswordUser()
 */
function changePassword(
  token,
  passwordLama,
  passwordBaru
) {
  const lock =
    LockService.getScriptLock();

  try {
    /*
     * Verifikasi session terlebih dahulu.
     */
    const session =
      verifySession(token);

    if (!session.valid) {
      throw new Error(
        session.message ||
        'Sesi tidak valid. Silakan login ulang.'
      );
    }

    passwordLama =
      String(passwordLama || '');

    passwordBaru =
      String(passwordBaru || '');

    if (!passwordLama) {
      throw new Error(
        'Password lama wajib diisi.'
      );
    }

    if (
      !passwordBaru ||
      passwordBaru.length < 6
    ) {
      throw new Error(
        'Password baru minimal 6 karakter.'
      );
    }

    if (
      passwordLama === passwordBaru
    ) {
      throw new Error(
        'Password baru harus berbeda dengan password lama.'
      );
    }

    lock.waitLock(10000);

    const sheet =
      getSheet(SHEET_USERS);

    const lastRow =
      sheet.getLastRow();

    if (lastRow < 2) {
      throw new Error(
        'Data pengguna kosong.'
      );
    }

    const data =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          11
        )
        .getValues();

    for (let i = 0; i < data.length; i++) {
      const row = data[i];

      if (
        String(
          row[AUTH_USER.ID]
        ) !==
        String(
          session.user.id_user
        )
      ) {
        continue;
      }

      /*
       * Pastikan akun masih aktif.
       */
      if (
        String(
          row[AUTH_USER.STATUS_AKTIF]
        ).toLowerCase() !== 'aktif'
      ) {
        throw new Error(
          'Akun tidak aktif.'
        );
      }

      const saltLama =
        String(
          row[AUTH_USER.SALT] || ''
        );

      const hashTersimpan =
        String(
          row[AUTH_USER.PASSWORD_HASH] || ''
        );

      const hashLama =
        hashPassword(
          passwordLama,
          saltLama
        );

      if (
        hashLama !== hashTersimpan
      ) {
        throw new Error(
          'Password lama tidak sesuai.'
        );
      }

      /*
       * Salt baru untuk password baru.
       */
      const saltBaru =
        Utilities.getUuid();

      const hashBaru =
        hashPassword(
          passwordBaru,
          saltBaru
        );

      const nomorBaris = i + 2;

      /*
       * Tulis hash + salt + updated_at
       * secara berurutan dalam satu operasi range.
       *
       * Kolom:
       * D = password_hash
       * E = salt
       * F-I tidak boleh disentuh.
       * K = updated_at
       */
      sheet
        .getRange(
          nomorBaris,
          4,
          1,
          2
        )
        .setValues([
          [
            hashBaru,
            saltBaru
          ]
        ]);

      sheet
        .getRange(
          nomorBaris,
          11
        )
        .setValue(
          new Date()
        );

      return jsonResponse({
        success: true,
        message:
          'Password berhasil diubah.'
      });
    }

    throw new Error(
      'User tidak ditemukan.'
    );

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
 * BUAT SESSION TOKEN
 * =================================================================== */

/**
 * Membuat token session baru.
 *
 * Session:
 * token
 * id_user
 * created_at
 * expired_at
 */
function buatSessionToken_(id_user) {
  const lock =
    LockService.getScriptLock();

  try {
    if (!id_user) {
      throw new Error(
        'ID user tidak valid.'
      );
    }

    lock.waitLock(10000);

    const sheet =
      getSheet(SHEET_SESSIONS);

    /*
     * Bersihkan session lama milik user ini
     * agar tidak menumpuk.
     *
     * Kita batasi agar user tidak mempunyai
     * puluhan token aktif karena login berulang.
     */
    const lastRow =
      sheet.getLastRow();

    if (lastRow >= 2) {
      const data =
        sheet
          .getRange(
            2,
            1,
            lastRow - 1,
            4
          )
          .getValues();

      const now =
        new Date();

      for (
        let i = data.length - 1;
        i >= 0;
        i--
      ) {
        const sessionUser =
          String(
            data[i][1] || ''
          );

        const expiredAt =
          new Date(
            data[i][3]
          );

        /*
         * Hapus:
         * 1. session milik user yang sama
         *    (supaya login baru menggantikan token lama)
         * 2. session yang sudah expired.
         */
        if (
          sessionUser ===
            String(id_user) ||
          expiredAt < now
        ) {
          sheet.deleteRow(
            i + 2
          );
        }
      }
    }

    const token =
      Utilities.getUuid();

    const now =
      new Date();

    const expired =
      new Date(
        now.getTime() +
        SESSION_DURASI_JAM *
        60 *
        60 *
        1000
      );

    sheet.appendRow([
      token,
      id_user,
      now,
      expired
    ]);

    return token;

  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}


/* ===================================================================
 * VERIFY SESSION
 * =================================================================== */

/**
 * Memvalidasi session token.
 *
 * Return:
 *
 * {
 *   valid: true,
 *   user: {...}
 * }
 *
 * atau:
 *
 * {
 *   valid: false,
 *   message: '...'
 * }
 */
function verifySession(token) {
  try {
    token =
      String(token || '').trim();

    if (!token) {
      return {
        valid: false,
        message:
          'Token tidak ada.'
      };
    }

    const sheetSession =
      getSheet(
        SHEET_SESSIONS
      );

    const lastRow =
      sheetSession.getLastRow();

    if (lastRow < 2) {
      return {
        valid: false,
        message:
          'Sesi tidak ditemukan. Silakan login ulang.'
      };
    }

    const dataSession =
      sheetSession
        .getRange(
          2,
          1,
          lastRow - 1,
          4
        )
        .getValues();

    let idUser = null;
    let rowIndexSession = -1;

    const now =
      new Date();

    for (
      let i = 0;
      i < dataSession.length;
      i++
    ) {
      if (
        String(
          dataSession[i][0]
        ) !== token
      ) {
        continue;
      }

      const expiredAt =
        new Date(
          dataSession[i][3]
        );

      if (
        isNaN(
          expiredAt.getTime()
        )
      ) {
        return {
          valid: false,
          message:
            'Session rusak. Silakan login ulang.'
        };
      }

      if (
        expiredAt < now
      ) {
        /*
         * Jangan delete di sini tanpa lock.
         *
         * verifySession dipanggil sangat sering.
         * Session expired akan dibersihkan oleh
         * bersihkanSessionExpired().
         */
        return {
          valid: false,
          message:
            'Sesi sudah kedaluwarsa. Silakan login ulang.'
        };
      }

      idUser =
        dataSession[i][1];

      rowIndexSession =
        i;

      break;
    }

    if (!idUser) {
      return {
        valid: false,
        message:
          'Sesi tidak ditemukan. Silakan login ulang.'
      };
    }

    /*
     * Ambil user terbaru.
     *
     * Ini sengaja tidak menggunakan data user
     * yang disimpan dalam session supaya perubahan:
     *
     * role
     * status_aktif
     * nama
     * jabatan
     *
     * langsung berlaku.
     */
    const sheetUsers =
      getSheet(
        SHEET_USERS
      );

    const lastUserRow =
      sheetUsers.getLastRow();

    if (lastUserRow < 2) {
      return {
        valid: false,
        message:
          'User tidak ditemukan.'
      };
    }

    const dataUsers =
      sheetUsers
        .getRange(
          2,
          1,
          lastUserRow - 1,
          11
        )
        .getValues();

    for (
      let i = 0;
      i < dataUsers.length;
      i++
    ) {
      const row =
        dataUsers[i];

      if (
        String(
          row[AUTH_USER.ID]
        ) !==
        String(idUser)
      ) {
        continue;
      }

      /*
       * Jika admin menonaktifkan akun,
       * session langsung dianggap tidak valid.
       */
      if (
        String(
          row[AUTH_USER.STATUS_AKTIF] ||
          ''
        ).toLowerCase() !==
        'aktif'
      ) {
        return {
          valid: false,
          message:
            'Akun tidak aktif.'
        };
      }

      return {
        valid: true,
        user: {
          id_user:
            row[AUTH_USER.ID],
          nama:
            row[AUTH_USER.NAMA],
          username:
            row[AUTH_USER.USERNAME],
          role:
            row[AUTH_USER.ROLE],
          jabatan:
            row[AUTH_USER.JABATAN],
          foto_profil:
            row[AUTH_USER.FOTO_PROFIL]
        }
      };
    }

    return {
      valid: false,
      message:
        'User tidak ditemukan.'
    };

  } catch (e) {
    /*
     * verifySession dipakai sebagai helper internal.
     * Jangan melempar error mentah ke frontend.
     */
    return {
      valid: false,
      message:
        e.message ||
        'Gagal memvalidasi session.'
    };
  }
}


/* ===================================================================
 * BERSIHKAN SESSION EXPIRED
 * =================================================================== */

/**
 * Menghapus session yang sudah kedaluwarsa.
 *
 * Fungsi ini idealnya dijalankan menggunakan
 * time-driven trigger, misalnya setiap hari.
 */
function bersihkanSessionExpired() {
  const lock =
    LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    const sheet =
      getSheet(
        SHEET_SESSIONS
      );

    const lastRow =
      sheet.getLastRow();

    if (lastRow < 2) {
      Logger.log(
        'Tidak ada session untuk dibersihkan.'
      );
      return;
    }

    const data =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          4
        )
        .getValues();

    const now =
      new Date();

    /*
     * Hapus dari bawah ke atas.
     */
    let jumlahDihapus = 0;

    for (
      let i = data.length - 1;
      i >= 0;
      i--
    ) {
      const expiredAt =
        new Date(
          data[i][3]
        );

      if (
        isNaN(
          expiredAt.getTime()
        )
      ) {
        /*
         * Session dengan tanggal rusak
         * dianggap invalid dan dibersihkan.
         */
        sheet.deleteRow(
          i + 2
        );

        jumlahDihapus++;
        continue;
      }

      if (
        expiredAt < now
      ) {
        sheet.deleteRow(
          i + 2
        );

        jumlahDihapus++;
      }
    }

    Logger.log(
      'Pembersihan session selesai. ' +
      'Dihapus: ' +
      jumlahDihapus
    );

  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}
