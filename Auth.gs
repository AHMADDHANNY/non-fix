function login(username, password) {
  try {
    if (!username || !password) {
      throw new Error('Username dan password wajib diisi.');
    }

    const sheet = getSheet(SHEET_USERS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowUsername = String(row[2]);

      if (rowUsername.toLowerCase() === username.toLowerCase()) {
        // Cek status aktif dulu sebelum proses hash (hemat, dan pesan lebih jelas)
        if (row[8] !== 'aktif') {
          throw new Error('Akun Anda tidak aktif. Hubungi admin.');
        }

        const salt = row[4];
        const hashInput = hashPassword(password, salt);
        const hashTersimpan = row[3];

        if (hashInput !== hashTersimpan) {
          throw new Error('Username atau password salah.');
        }

        // Login berhasil -> buat session token baru
        const user = {
          id_user: row[0],
          nama: row[1],
          username: row[2],
          role: row[5],
          jabatan: row[6],
          foto_profil: row[7]
        };

        const token = buatSessionToken_(user.id_user);

        return jsonResponse({ success: true, token: token, user: user });
      }
    }

    // Tidak ketemu username sama sekali -> pesan digeneralisasi
    // (sengaja tidak bedakan "username tidak ada" vs "password salah"
    // supaya tidak membocorkan username mana yang terdaftar)
    throw new Error('Username atau password salah.');
  } catch (e) {
    return jsonResponse({ success: false, message: e.message });
  }
}

/**
 * Logout: hapus token session dari sheet Sessions.
 */
function logout(token) {
  try {
    if (!token) return jsonResponse({ success: true });

    const sheet = getSheet(SHEET_SESSIONS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === token) {
        sheet.deleteRow(i + 1);
        break;
      }
    }

    return jsonResponse({ success: true, message: 'Berhasil logout.' });
  } catch (e) {
    return jsonResponse({ success: false, message: e.message });
  }
}

/**
 * Ganti password (dipanggil user yang sedang login, bukan admin
 * reset-kan punya orang lain -- itu ada di Admin.gs).
 */
function changePassword(token, passwordLama, passwordBaru) {
  try {
    const session = verifySession(token);
    if (!session.valid) {
      throw new Error('Sesi tidak valid. Silakan login ulang.');
    }
    if (!passwordBaru || passwordBaru.length < 6) {
      throw new Error('Password baru minimal 6 karakter.');
    }

    const sheet = getSheet(SHEET_USERS);
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === session.user.id_user) {
        const row = i + 1;
        const salt = data[i][4];
        const hashLama = hashPassword(passwordLama, salt);

        if (hashLama !== data[i][3]) {
          throw new Error('Password lama tidak sesuai.');
        }

        const saltBaru = Utilities.getUuid();
        const hashBaru = hashPassword(passwordBaru, saltBaru);

        sheet.getRange(row, 4).setValue(hashBaru);   // password_hash
        sheet.getRange(row, 5).setValue(saltBaru);   // salt
        sheet.getRange(row, 11).setValue(new Date()); // updated_at

        return jsonResponse({ success: true, message: 'Password berhasil diubah.' });
      }
    }

    throw new Error('User tidak ditemukan.');
  } catch (e) {
    return jsonResponse({ success: false, message: e.message });
  }
}

/* ===================================================================
 * SESSION TOKEN (disimpan di sheet Sessions, sesuai kesepakatan
 * awal -- lebih andal daripada CacheService yang max 6 jam)
 * =================================================================== */

/**
 * Buat token session baru untuk user, simpan ke sheet Sessions.
 * @returns {string} token
 */
function buatSessionToken_(id_user) {
  const sheet = getSheet(SHEET_SESSIONS);
  const token = Utilities.getUuid();
  const now = new Date();
  const expired = new Date(now.getTime() + SESSION_DURASI_JAM * 60 * 60 * 1000);

  sheet.appendRow([token, id_user, now, expired]);

  return token;
}

/**
 * Verifikasi token session: cek ada, belum expired, dan ambil data user-nya.
 * Dipakai oleh SEMUA fungsi backend lain (Absensi.gs, Izin.gs, Admin.gs)
 * sebelum memproses request.
 *
 * @returns {Object} { valid: boolean, user?: {...}, message?: string }
 */
function verifySession(token) {
  if (!token) {
    return { valid: false, message: 'Token tidak ada.' };
  }

  const sheetSession = getSheet(SHEET_SESSIONS);
  const dataSession = sheetSession.getDataRange().getValues();

  let idUser = null;
  let rowIndexSession = -1;

  for (let i = 1; i < dataSession.length; i++) {
    if (dataSession[i][0] === token) {
      const expiredAt = new Date(dataSession[i][3]);
      if (expiredAt < new Date()) {
        // Token sudah kedaluwarsa -> hapus dan anggap tidak valid
        sheetSession.deleteRow(i + 1);
        return { valid: false, message: 'Sesi sudah kedaluwarsa. Silakan login ulang.' };
      }
      idUser = dataSession[i][1];
      rowIndexSession = i;
      break;
    }
  }

  if (!idUser) {
    return { valid: false, message: 'Sesi tidak ditemukan. Silakan login ulang.' };
  }

  // Ambil data user terbaru (bukan dari cache token, supaya kalau role
  // atau status_aktif berubah di tengah sesi, langsung ke-detect)
  const sheetUsers = getSheet(SHEET_USERS);
  const dataUsers = sheetUsers.getDataRange().getValues();

  for (let i = 1; i < dataUsers.length; i++) {
    if (dataUsers[i][0] === idUser) {
      if (dataUsers[i][8] !== 'aktif') {
        return { valid: false, message: 'Akun tidak aktif.' };
      }
      return {
        valid: true,
        user: {
          id_user: dataUsers[i][0],
          nama: dataUsers[i][1],
          username: dataUsers[i][2],
          role: dataUsers[i][5],
          jabatan: dataUsers[i][6],
          foto_profil: dataUsers[i][7]
        }
      };
    }
  }

  return { valid: false, message: 'User tidak ditemukan.' };
}

/**
 * Bersihkan token session yang sudah kedaluwarsa dari sheet Sessions.
 * Sebaiknya dijadwalkan jalan otomatis tiap malam lewat Trigger
 * (Edit > Pemicu proyek saat ini > Tambahkan Pemicu > time-driven, harian).
 */
function bersihkanSessionExpired() {
  const sheet = getSheet(SHEET_SESSIONS);
  const data = sheet.getDataRange().getValues();
  const now = new Date();

  // Hapus dari bawah ke atas supaya index baris tidak bergeser
  for (let i = data.length - 1; i >= 1; i--) {
    const expiredAt = new Date(data[i][3]);
    if (expiredAt < now) {
      sheet.deleteRow(i + 1);
    }
  }

  Logger.log('Pembersihan session selesai.');
}
