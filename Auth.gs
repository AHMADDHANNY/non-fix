/**
 * ===================================================================
 * AUTH.GS - FINAL
 * Login, session, logout, change password.
 * Dependency: Code.gs
 * ===================================================================
 */

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

function login(
  username,
  password
) {
  try {
    username =
      String(
        username || ''
      ).trim();

    password =
      String(
        password || ''
      );

    if (
      !username ||
      !password
    ) {
      throw new Error(
        'Username dan password wajib diisi.'
      );
    }

    var sheet =
      getSheet(
        SHEET_USERS
      );

    var lastRow =
      sheet.getLastRow();

    if (lastRow < 2) {
      throw new Error(
        'Belum ada akun pengguna.'
      );
    }

    /*
     * TextFinder menghindari pembacaan seluruh Users.
     * Username tetap dicari pada kolom C saja.
     */
    var finder =
      sheet
        .getRange(
          2,
          AUTH_USER.USERNAME + 1,
          lastRow - 1,
          1
        )
        .createTextFinder(
          username
        )
        .matchCase(false)
        .matchEntireCell(true);

    var cell =
      finder.findNext();

    if (!cell) {
      throw new Error(
        'Username atau password salah.'
      );
    }

    var rowNumber =
      cell.getRow();

    var row =
      sheet
        .getRange(
          rowNumber,
          1,
          1,
          11
        )
        .getValues()[0];

    var status =
      String(
        row[
          AUTH_USER.STATUS_AKTIF
        ] || ''
      ).toLowerCase();

    if (
      status !== 'aktif'
    ) {
      throw new Error(
        'Akun Anda tidak aktif. Hubungi admin.'
      );
    }

    var salt =
      String(
        row[
          AUTH_USER.SALT
        ] || ''
      );

    var storedHash =
      String(
        row[
          AUTH_USER.PASSWORD_HASH
        ] || ''
      );

    if (
      !salt ||
      !storedHash
    ) {
      throw new Error(
        'Data autentikasi akun tidak valid. Hubungi admin.'
      );
    }

    var inputHash =
      hashPassword(
        password,
        salt
      );

    if (
      inputHash !== storedHash
    ) {
      throw new Error(
        'Username atau password salah.'
      );
    }

    var user = {
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

    var token =
      buatSessionToken_(
        user.id_user
      );

    return jsonResponse({
      success: true,
      token: token,
      user: user
    });

  } catch (e) {
    return jsonResponse({
      success: false,
      message:
        e.message ||
        'Login gagal.'
    });
  }
}


/* ===================================================================
 * LOGOUT
 * =================================================================== */

function logout(token) {
  var lock =
    LockService.getScriptLock();

  try {
    token =
      String(
        token || ''
      ).trim();

    if (!token) {
      return jsonResponse({
        success: true,
        message:
          'Sudah logout.'
      });
    }

    lock.waitLock(10000);

    var sheet =
      getSheet(
        SHEET_SESSIONS
      );

    var lastRow =
      sheet.getLastRow();

    if (lastRow >= 2) {
      var data =
        sheet
          .getRange(
            2,
            1,
            lastRow - 1,
            4
          )
          .getValues();

      for (
        var i =
          data.length - 1;
        i >= 0;
        i--
      ) {
        if (
          String(
            data[i][0] || ''
          ) === token
        ) {
          sheet.deleteRow(
            i + 2
          );
          break;
        }
      }
    }

    return jsonResponse({
      success: true,
      message:
        'Berhasil logout.'
    });

  } catch (e) {
    return jsonResponse({
      success: false,
      message:
        e.message ||
        'Logout gagal.'
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

function changePassword(
  token,
  passwordLama,
  passwordBaru
) {
  var lock =
    LockService.getScriptLock();

  try {
    var session =
      verifySession(token);

    if (!session.valid) {
      throw new Error(
        session.message ||
        'Sesi tidak valid.'
      );
    }

    passwordLama =
      String(
        passwordLama || ''
      );

    passwordBaru =
      String(
        passwordBaru || ''
      );

    validasiPasswordBaru_(
      passwordBaru
    );

    if (
      passwordLama ===
      passwordBaru
    ) {
      throw new Error(
        'Password baru harus berbeda dengan password lama.'
      );
    }

    lock.waitLock(10000);

    var sheet =
      getSheet(
        SHEET_USERS
      );

    var lastRow =
      sheet.getLastRow();

    if (lastRow < 2) {
      throw new Error(
        'Data pengguna kosong.'
      );
    }

    var finder =
      sheet
        .getRange(
          2,
          AUTH_USER.ID + 1,
          lastRow - 1,
          1
        )
        .createTextFinder(
          String(
            session.user.id_user
          )
        )
        .matchCase(true)
        .matchEntireCell(true);

    var cell =
      finder.findNext();

    if (!cell) {
      throw new Error(
        'User tidak ditemukan.'
      );
    }

    var rowNumber =
      cell.getRow();

    var row =
      sheet
        .getRange(
          rowNumber,
          1,
          1,
          11
        )
        .getValues()[0];

    if (
      String(
        row[
          AUTH_USER.STATUS_AKTIF
        ] || ''
      ).toLowerCase() !==
      'aktif'
    ) {
      throw new Error(
        'Akun tidak aktif.'
      );
    }

    var saltLama =
      String(
        row[
          AUTH_USER.SALT
        ] || ''
      );

    var hashLama =
      String(
        row[
          AUTH_USER.PASSWORD_HASH
        ] || ''
      );

    if (
      hashPassword(
        passwordLama,
        saltLama
      ) !== hashLama
    ) {
      throw new Error(
        'Password lama tidak sesuai.'
      );
    }

    var saltBaru =
      Utilities.getUuid();

    var hashBaru =
      hashPassword(
        passwordBaru,
        saltBaru
      );

    sheet
      .getRange(
        rowNumber,
        4,
        1,
        2
      )
      .setValues([[
        hashBaru,
        saltBaru
      ]]);

    sheet
      .getRange(
        rowNumber,
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

  } catch (e) {
    return jsonResponse({
      success: false,
      message:
        e.message ||
        'Gagal mengubah password.'
    });

  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}

function validasiPasswordBaru_(
  password
) {
  if (
    !password ||
    password.length < 8
  ) {
    throw new Error(
      'Password baru minimal 8 karakter.'
    );
  }

  if (
    password.length > 128
  ) {
    throw new Error(
      'Password baru maksimal 128 karakter.'
    );
  }
}


/* ===================================================================
 * SESSION TOKEN
 * =================================================================== */

function buatSessionToken_(
  id_user
) {
  var lock =
    LockService.getScriptLock();

  try {
    if (!id_user) {
      throw new Error(
        'ID user tidak valid.'
      );
    }

    lock.waitLock(10000);

    var sheet =
      getSheet(
        SHEET_SESSIONS
      );

    var lastRow =
      sheet.getLastRow();

    var data = [];

    if (lastRow >= 2) {
      data =
        sheet
          .getRange(
            2,
            1,
            lastRow - 1,
            4
          )
          .getValues();
    }

    var now =
      new Date();

    /*
     * Login baru menggantikan token user lama.
     * Session expired juga dibersihkan.
     *
     * Karena Sessions tidak menyimpan formula/data lain,
     * kita rebuild isi data sekaligus daripada deleteRow
     * berulang-ulang.
     */
    var keep = [];

    for (
      var i = 0;
      i < data.length;
      i++
    ) {
      var sessionUser =
        String(
          data[i][1] || ''
        );

      var expired =
        data[i][3] instanceof Date
          ? data[i][3]
          : new Date(
              data[i][3]
            );

      if (
        sessionUser ===
        String(id_user)
      ) {
        continue;
      }

      if (
        isNaN(
          expired.getTime()
        ) ||
        expired <= now
      ) {
        continue;
      }

      keep.push(data[i]);
    }

    if (lastRow >= 2) {
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          4
        )
        .clearContent();
    }

    if (keep.length) {
      sheet
        .getRange(
          2,
          1,
          keep.length,
          4
        )
        .setValues(keep);
    }

    /*
     * UUID Apps Script cukup untuk session token.
     * Tambahkan timestamp agar token tidak monoton.
     */
    var token =
      Utilities.getUuid() +
      '-' +
      Utilities.getUuid();

    var expiredAt =
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
      expiredAt
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
 *
 * Tidak lagi membaca seluruh Sessions dan Users.
 * Token dicari dengan TextFinder pada kolom A, lalu user ID
 * dicari dengan TextFinder pada kolom A Users.
 * =================================================================== */

function verifySession(
  token
) {
  try {
    token =
      String(
        token || ''
      ).trim();

    if (!token) {
      return {
        valid: false,
        message:
          'Token tidak ada.'
      };
    }

    var sheetSession =
      getSheet(
        SHEET_SESSIONS
      );

    var lastSessionRow =
      sheetSession.getLastRow();

    if (
      lastSessionRow < 2
    ) {
      return {
        valid: false,
        message:
          'Sesi tidak ditemukan. Silakan login ulang.'
      };
    }

    var tokenCell =
      sheetSession
        .getRange(
          2,
          1,
          lastSessionRow - 1,
          1
        )
        .createTextFinder(
          token
        )
        .matchCase(true)
        .matchEntireCell(true)
        .findNext();

    if (!tokenCell) {
      return {
        valid: false,
        message:
          'Sesi tidak ditemukan. Silakan login ulang.'
      };
    }

    var sessionRow =
      tokenCell.getRow();

    var sessionData =
      sheetSession
        .getRange(
          sessionRow,
          1,
          1,
          4
        )
        .getValues()[0];

    var expiredAt =
      sessionData[3] instanceof Date
        ? sessionData[3]
        : new Date(
            sessionData[3]
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
      expiredAt <= new Date()
    ) {
      return {
        valid: false,
        message:
          'Sesi sudah kedaluwarsa. Silakan login ulang.'
      };
    }

    var idUser =
      String(
        sessionData[1] || ''
      ).trim();

    if (!idUser) {
      return {
        valid: false,
        message:
          'Session tidak memiliki user.'
      };
    }

    var sheetUsers =
      getSheet(
        SHEET_USERS
      );

    var lastUserRow =
      sheetUsers.getLastRow();

    if (
      lastUserRow < 2
    ) {
      return {
        valid: false,
        message:
          'User tidak ditemukan.'
      };
    }

    var userCell =
      sheetUsers
        .getRange(
          2,
          1,
          lastUserRow - 1,
          1
        )
        .createTextFinder(
          idUser
        )
        .matchCase(true)
        .matchEntireCell(true)
        .findNext();

    if (!userCell) {
      return {
        valid: false,
        message:
          'User tidak ditemukan.'
      };
    }

    var userRowNumber =
      userCell.getRow();

    var row =
      sheetUsers
        .getRange(
          userRowNumber,
          1,
          1,
          11
        )
        .getValues()[0];

    if (
      String(
        row[
          AUTH_USER.STATUS_AKTIF
        ] || ''
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

  } catch (e) {
    return {
      valid: false,
      message:
        e.message ||
        'Gagal memvalidasi session.'
    };
  }
}


/* ===================================================================
 * CLEAN EXPIRED SESSIONS
 * =================================================================== */

function bersihkanSessionExpired() {
  var lock =
    LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    var sheet =
      getSheet(
        SHEET_SESSIONS
      );

    var lastRow =
      sheet.getLastRow();

    if (
      lastRow < 2
    ) {
      return {
        success: true,
        deleted: 0
      };
    }

    var data =
      sheet
        .getRange(
          2,
          1,
          lastRow - 1,
          4
        )
        .getValues();

    var now =
      new Date();

    var keep = [];
    var deleted = 0;

    for (
      var i = 0;
      i < data.length;
      i++
    ) {
      var expired =
        data[i][3] instanceof Date
          ? data[i][3]
          : new Date(
              data[i][3]
            );

      if (
        isNaN(
          expired.getTime()
        ) ||
        expired <= now
      ) {
        deleted++;
      } else {
        keep.push(data[i]);
      }
    }

    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        4
      )
      .clearContent();

    if (keep.length) {
      sheet
        .getRange(
          2,
          1,
          keep.length,
          4
        )
        .setValues(keep);
    }

    return {
      success: true,
      deleted: deleted
    };

  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}
