/**
 * ===================================================================
 * IZIN.GS - VERSI PERBAIKAN
 *
 * Fitur:
 * - Pengajuan Izin/Cuti/Sakit
 * - Lampiran opsional
 * - Riwayat pegawai dengan paginasi
 * - Daftar izin admin dengan filter & paginasi
 * - Approve / Reject
 * - Validasi overlap tanggal
 * - LockService untuk seluruh operasi tulis
 * - Validasi file lampiran di backend
 *
 * Struktur Sheet Izin:
 * 0 id
 * 1 id_user
 * 2 tanggal_pengajuan
 * 3 jenis_izin
 * 4 tanggal_mulai
 * 5 tanggal_selesai
 * 6 keterangan
 * 7 lampiran
 * 8 status
 * 9 catatan_admin
 * 10 diproses_oleh
 * 11 diproses_at
 * ===================================================================
 */

var KOLOM_IZIN = {
  ID: 0,
  ID_USER: 1,
  TANGGAL_PENGAJUAN: 2,
  JENIS_IZIN: 3,
  TANGGAL_MULAI: 4,
  TANGGAL_SELESAI: 5,
  KETERANGAN: 6,
  LAMPIRAN: 7,
  STATUS: 8,
  CATATAN_ADMIN: 9,
  DIPROSES_OLEH: 10,
  DIPROSES_AT: 11
};

var JENIS_IZIN_VALID = [
  'Izin',
  'Cuti',
  'Sakit'
];

var STATUS_IZIN_PENDING = 'Menunggu';
var STATUS_IZIN_DISETUJUI = 'Disetujui';
var STATUS_IZIN_DITOLAK = 'Ditolak';

var MAKS_HARI_PENGAJUAN_IZIN = 366;
var MAKS_KETERANGAN_IZIN = 2000;
var MAKS_CATATAN_ADMIN_IZIN = 2000;
var MAKS_LAMPIRAN_IZIN_BYTES = 5 * 1024 * 1024;


/* ===================================================================
 * 1. PEGAWAI - AJUKAN IZIN
 * =================================================================== */

/**
 * @param {string} token
 * @param {string} jenisIzin
 * @param {string} tanggalMulai yyyy-MM-dd
 * @param {string} tanggalSelesai yyyy-MM-dd
 * @param {string} keterangan
 * @param {string} lampiranBase64 optional data URL
 */
function ajukanIzin(
  token,
  jenisIzin,
  tanggalMulai,
  tanggalSelesai,
  keterangan,
  lampiranBase64
) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    var session = verifySession(token);

    if (!session.valid) {
      throw new Error(
        session.message || 'Sesi tidak valid. Silakan login ulang.'
      );
    }

    if (
      String(session.user.role || '').toLowerCase() !==
      'pegawai'
    ) {
      throw new Error(
        'Hanya pegawai yang dapat mengajukan izin.'
      );
    }

    jenisIzin =
      String(jenisIzin || '').trim();

    if (
      JENIS_IZIN_VALID.indexOf(jenisIzin) === -1
    ) {
      throw new Error(
        'Jenis izin tidak valid. Pilih Izin, Cuti, atau Sakit.'
      );
    }

    var mulai =
      validasiTanggalInputIzin_(
        tanggalMulai,
        'Tanggal mulai'
      );

    var selesai =
      validasiTanggalInputIzin_(
        tanggalSelesai,
        'Tanggal selesai'
      );

    if (selesai < mulai) {
      throw new Error(
        'Tanggal selesai tidak boleh sebelum tanggal mulai.'
      );
    }

    var jumlahHari =
      hitungSelisihHariIzin_(
        mulai,
        selesai
      );

    if (
      jumlahHari >
      MAKS_HARI_PENGAJUAN_IZIN
    ) {
      throw new Error(
        'Rentang izin terlalu panjang. Maksimal ' +
        MAKS_HARI_PENGAJUAN_IZIN +
        ' hari.'
      );
    }

    keterangan =
      String(keterangan || '').trim();

    if (keterangan.length < 3) {
      throw new Error(
        'Keterangan wajib diisi minimal 3 karakter.'
      );
    }

    if (
      keterangan.length >
      MAKS_KETERANGAN_IZIN
    ) {
      throw new Error(
        'Keterangan terlalu panjang. Maksimal ' +
        MAKS_KETERANGAN_IZIN +
        ' karakter.'
      );
    }

    var tanggalHariIni =
      formatTanggalIzin_(
        new Date()
      );

    /*
     * Jangan izinkan pengajuan sangat jauh ke masa lalu.
     * Pengajuan untuk tanggal hari ini dan masa depan diperbolehkan.
     */
    if (mulai < tanggalHariIni) {
      throw new Error(
        'Tanggal mulai izin tidak boleh sebelum hari ini.'
      );
    }

    var sheet =
      getSheet(SHEET_IZIN);

    var data =
      sheet.getDataRange().getValues();

    /*
     * Cegah overlap dengan pengajuan Menunggu/Disetujui.
     */
    for (
      var i = 1;
      i < data.length;
      i++
    ) {
      var row = data[i];

      if (
        String(
          row[KOLOM_IZIN.ID_USER] || ''
        ) !==
        String(session.user.id_user)
      ) {
        continue;
      }

      var statusLama =
        String(
          row[KOLOM_IZIN.STATUS] || ''
        );

      if (
        statusLama !==
          STATUS_IZIN_PENDING &&
        statusLama !==
          STATUS_IZIN_DISETUJUI
      ) {
        continue;
      }

      var mulaiLama =
        normalizeTanggalIzin_(
          row[KOLOM_IZIN.TANGGAL_MULAI]
        );

      var selesaiLama =
        normalizeTanggalIzin_(
          row[KOLOM_IZIN.TANGGAL_SELESAI]
        );

      if (
        !mulaiLama ||
        !selesaiLama
      ) {
        continue;
      }

      var overlap =
        mulai <= selesaiLama &&
        selesai >= mulaiLama;

      if (overlap) {
        throw new Error(
          'Pengajuan Anda bertumpang tindih dengan izin ' +
          statusLama.toLowerCase() +
          ' (' +
          mulaiLama +
          ' s/d ' +
          selesaiLama +
          ').'
        );
      }
    }

    var now = new Date();

    var tanggalPengajuan =
      formatTanggalIzin_(now);

    var urlLampiran = '';

    if (
      lampiranBase64 &&
      String(lampiranBase64).trim()
    ) {
      urlLampiran =
        simpanLampiranIzin_(
          lampiranBase64,
          session.user.id_user +
            '_izin_' +
            tanggalPengajuan +
            '_' +
            new Date().getTime()
        );
    }

    var idBaru =
      generateSequentialId(
        'IZN',
        SHEET_IZIN,
        KOLOM_IZIN.ID
      );

    var rowBaru =
      new Array(12);

    rowBaru[
      KOLOM_IZIN.ID
    ] = idBaru;

    rowBaru[
      KOLOM_IZIN.ID_USER
    ] = session.user.id_user;

    rowBaru[
      KOLOM_IZIN.TANGGAL_PENGAJUAN
    ] = tanggalPengajuan;

    rowBaru[
      KOLOM_IZIN.JENIS_IZIN
    ] = jenisIzin;

    rowBaru[
      KOLOM_IZIN.TANGGAL_MULAI
    ] = mulai;

    rowBaru[
      KOLOM_IZIN.TANGGAL_SELESAI
    ] = selesai;

    rowBaru[
      KOLOM_IZIN.KETERANGAN
    ] = keterangan;

    rowBaru[
      KOLOM_IZIN.LAMPIRAN
    ] = urlLampiran;

    rowBaru[
      KOLOM_IZIN.STATUS
    ] = STATUS_IZIN_PENDING;

    rowBaru[
      KOLOM_IZIN.CATATAN_ADMIN
    ] = '';

    rowBaru[
      KOLOM_IZIN.DIPROSES_OLEH
    ] = '';

    rowBaru[
      KOLOM_IZIN.DIPROSES_AT
    ] = '';

    sheet.appendRow(rowBaru);

    return jsonResponse({
      success: true,
      message:
        'Pengajuan ' +
        jenisIzin.toLowerCase() +
        ' berhasil dikirim. Menunggu persetujuan admin.',
      data: {
        id_izin: idBaru,
        status: STATUS_IZIN_PENDING,
        tanggal_mulai: mulai,
        tanggal_selesai: selesai
      }
    });

  } catch (e) {
    return jsonResponse({
      success: false,
      message:
        e.message ||
        'Gagal mengajukan izin.'
    });

  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}


/* ===================================================================
 * 2. PEGAWAI - RIWAYAT IZIN
 * =================================================================== */

/**
 * @param {string} token
 * @param {number} page
 * @param {number} pageSize
 * @param {string} statusFilter
 */
function getHistoryIzin(
  token,
  page,
  pageSize,
  statusFilter
) {
  try {
    var session =
      verifySession(token);

    if (!session.valid) {
      throw new Error(
        session.message ||
        'Sesi tidak valid.'
      );
    }

    page =
      parseInt(page, 10) || 1;

    pageSize =
      parseInt(pageSize, 10) || 20;

    if (page < 1) page = 1;
    if (pageSize < 1) pageSize = 20;
    if (pageSize > 50) pageSize = 50;

    statusFilter =
      String(
        statusFilter || ''
      ).trim();

    if (
      statusFilter &&
      !isStatusIzinValid_(
        statusFilter
      )
    ) {
      throw new Error(
        'Filter status izin tidak valid.'
      );
    }

    var sheet =
      getSheet(SHEET_IZIN);

    var lastRow =
      sheet.getLastRow();

    var hasil = [];

    if (lastRow >= 2) {
      var data =
        sheet.getRange(
          2,
          1,
          lastRow - 1,
          12
        ).getValues();

      for (
        var i = 0;
        i < data.length;
        i++
      ) {
        var row = data[i];

        if (
          String(
            row[KOLOM_IZIN.ID_USER] || ''
          ) !==
          String(session.user.id_user)
        ) {
          continue;
        }

        if (
          statusFilter &&
          String(
            row[KOLOM_IZIN.STATUS] || ''
          ) !== statusFilter
        ) {
          continue;
        }

        hasil.push(
          _barisIzinKeObjek_(row)
        );
      }
    }

    hasil.sort(
      function(a, b) {
        if (
          a.tanggal_pengajuan ===
          b.tanggal_pengajuan
        ) {
          return String(b.id || '')
            .localeCompare(
              String(a.id || '')
            );
        }

        return a.tanggal_pengajuan <
          b.tanggal_pengajuan
          ? 1
          : -1;
      }
    );

    var total =
      hasil.length;

    var totalPage =
      total === 0
        ? 0
        : Math.ceil(
            total / pageSize
          );

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
 * 3. ADMIN - DAFTAR IZIN
 * =================================================================== */

/**
 * @param {string} token
 * @param {number} page
 * @param {number} pageSize
 * @param {string} statusFilter
 * @param {string} bulanFilter yyyy-MM
 * @param {string} idUserFilter
 */
function getDaftarIzinAdmin(
  token,
  page,
  pageSize,
  statusFilter,
  bulanFilter,
  idUserFilter
) {
  try {
    var session =
      verifySession(token);

    if (!session.valid) {
      throw new Error(
        session.message ||
        'Sesi tidak valid.'
      );
    }

    pastikanAdminIzin_(
      session
    );

    page =
      parseInt(page, 10) || 1;

    pageSize =
      parseInt(pageSize, 10) || 20;

    if (page < 1) page = 1;
    if (pageSize < 1) pageSize = 20;
    if (pageSize > 50) pageSize = 50;

    statusFilter =
      String(
        statusFilter || ''
      ).trim();

    bulanFilter =
      String(
        bulanFilter || ''
      ).trim();

    idUserFilter =
      String(
        idUserFilter || ''
      ).trim();

    if (
      statusFilter &&
      !isStatusIzinValid_(
        statusFilter
      )
    ) {
      throw new Error(
        'Filter status izin tidak valid.'
      );
    }

    if (
      bulanFilter &&
      !/^\d{4}-\d{2}$/.test(
        bulanFilter
      )
    ) {
      throw new Error(
        'Format bulan tidak valid. Gunakan yyyy-MM.'
      );
    }

    var petaNama =
      _petaNamaUser_();

    var sheet =
      getSheet(SHEET_IZIN);

    var lastRow =
      sheet.getLastRow();

    var hasil = [];

    if (lastRow >= 2) {
      var data =
        sheet.getRange(
          2,
          1,
          lastRow - 1,
          12
        ).getValues();

      for (
        var i = 0;
        i < data.length;
        i++
      ) {
        var row = data[i];

        if (
          statusFilter &&
          String(
            row[KOLOM_IZIN.STATUS] || ''
          ) !== statusFilter
        ) {
          continue;
        }

        if (
          idUserFilter &&
          String(
            row[KOLOM_IZIN.ID_USER] || ''
          ) !== idUserFilter
        ) {
          continue;
        }

        var tanggalMulai =
          normalizeTanggalIzin_(
            row[KOLOM_IZIN.TANGGAL_MULAI]
          );

        if (
          bulanFilter &&
          tanggalMulai.indexOf(
            bulanFilter
          ) !== 0
        ) {
          continue;
        }

        hasil.push(
          _barisIzinAdminKeObjek_(
            row,
            petaNama
          )
        );
      }
    }

    hasil.sort(
      function(a, b) {
        if (
          a.tanggal_pengajuan ===
          b.tanggal_pengajuan
        ) {
          return String(b.id || '')
            .localeCompare(
              String(a.id || '')
            );
        }

        return a.tanggal_pengajuan <
          b.tanggal_pengajuan
          ? 1
          : -1;
      }
    );

    var total =
      hasil.length;

    var totalPage =
      total === 0
        ? 0
        : Math.ceil(
            total / pageSize
          );

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
 * 4. ADMIN - PROSES IZIN
 * =================================================================== */

/**
 * @param {string} token
 * @param {string} idIzin
 * @param {string} statusBaru
 * @param {string} catatanAdmin
 */
function prosesIzin(
  token,
  idIzin,
  statusBaru,
  catatanAdmin
) {
  var lock =
    LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    var session =
      verifySession(token);

    if (!session.valid) {
      throw new Error(
        session.message ||
        'Sesi tidak valid.'
      );
    }

    pastikanAdminIzin_(
      session
    );

    idIzin =
      String(
        idIzin || ''
      ).trim();

    statusBaru =
      String(
        statusBaru || ''
      ).trim();

    catatanAdmin =
      String(
        catatanAdmin || ''
      ).trim();

    if (!idIzin) {
      throw new Error(
        'ID pengajuan izin wajib diisi.'
      );
    }

    if (
      statusBaru !==
        STATUS_IZIN_DISETUJUI &&
      statusBaru !==
        STATUS_IZIN_DITOLAK
    ) {
      throw new Error(
        'Status baru tidak valid.'
      );
    }

    if (
      catatanAdmin.length >
      MAKS_CATATAN_ADMIN_IZIN
    ) {
      throw new Error(
        'Catatan admin terlalu panjang. Maksimal ' +
        MAKS_CATATAN_ADMIN_IZIN +
        ' karakter.'
      );
    }

    var sheet =
      getSheet(SHEET_IZIN);

    var lastRow =
      sheet.getLastRow();

    if (lastRow < 2) {
      throw new Error(
        'Pengajuan izin tidak ditemukan.'
      );
    }

    var data =
      sheet.getRange(
        2,
        1,
        lastRow - 1,
        12
      ).getValues();

    var rowIndex = -1;

    for (
      var i = 0;
      i < data.length;
      i++
    ) {
      if (
        String(
          data[i][KOLOM_IZIN.ID] || ''
        ) === idIzin
      ) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error(
        'Pengajuan izin tidak ditemukan.'
      );
    }

    var rowData =
      data[rowIndex];

    var statusSaatIni =
      String(
        rowData[KOLOM_IZIN.STATUS] || ''
      );

    if (
      statusSaatIni !==
      STATUS_IZIN_PENDING
    ) {
      throw new Error(
        'Pengajuan ini sudah diproses sebelumnya (' +
        statusSaatIni +
        ').'
      );
    }

    /*
     * Race-condition protection:
     * Lock sudah aktif, jadi dua admin tidak dapat memproses
     * pengajuan yang sama secara bersamaan di script ini.
     */

    /*
     * Jika approve, lakukan pengecekan overlap lagi.
     * Ini penting karena selama pengajuan masih Menunggu,
     * pengajuan lain bisa saja baru saja berubah menjadi
     * Disetujui.
     */
    if (
      statusBaru ===
      STATUS_IZIN_DISETUJUI
    ) {
      pastikanTidakAdaIzinBentrokSaatApprove_(
        data,
        rowIndex
      );
    }

    var sheetRow =
      rowIndex + 2;

    var now =
      new Date();

    sheet
      .getRange(
        sheetRow,
        KOLOM_IZIN.STATUS + 1,
        1,
        4
      )
      .setValues([[
        statusBaru,
        catatanAdmin,
        session.user.id_user,
        now
      ]]);

    return jsonResponse({
      success: true,
      message:
        statusBaru ===
        STATUS_IZIN_DISETUJUI
          ? 'Pengajuan izin berhasil disetujui.'
          : 'Pengajuan izin berhasil ditolak.',
      data: {
        id_izin: idIzin,
        status: statusBaru
      }
    });

  } catch (e) {
    return jsonResponse({
      success: false,
      message:
        e.message ||
        'Gagal memproses izin.'
    });

  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}


/* ===================================================================
 * 5. ADMIN - DETAIL IZIN
 * =================================================================== */

function getDetailIzinAdmin(
  token,
  idIzin
) {
  try {
    var session =
      verifySession(token);

    if (!session.valid) {
      throw new Error(
        session.message ||
        'Sesi tidak valid.'
      );
    }

    pastikanAdminIzin_(
      session
    );

    idIzin =
      String(
        idIzin || ''
      ).trim();

    if (!idIzin) {
      throw new Error(
        'ID izin wajib diisi.'
      );
    }

    var sheet =
      getSheet(SHEET_IZIN);

    var lastRow =
      sheet.getLastRow();

    if (lastRow < 2) {
      throw new Error(
        'Pengajuan izin tidak ditemukan.'
      );
    }

    var data =
      sheet.getRange(
        2,
        1,
        lastRow - 1,
        12
      ).getValues();

    var petaNama =
      _petaNamaUser_();

    for (
      var i = 0;
      i < data.length;
      i++
    ) {
      if (
        String(
          data[i][KOLOM_IZIN.ID] || ''
        ) !== idIzin
      ) {
        continue;
      }

      return jsonResponse({
        success: true,
        data:
          _barisIzinAdminKeObjek_(
            data[i],
            petaNama
          )
      });
    }

    throw new Error(
      'Pengajuan izin tidak ditemukan.'
    );

  } catch (e) {
    return jsonResponse({
      success: false,
      message: e.message
    });
  }
}


/* ===================================================================
 * 6. HELPER - VALIDASI ADMIN
 * =================================================================== */

function pastikanAdminIzin_(
  session
) {
  if (
    !session ||
    !session.user ||
    String(
      session.user.role || ''
    ).toLowerCase() !==
      'admin'
  ) {
    throw new Error(
      'Akses ditolak. Hanya admin.'
    );
  }
}


/* ===================================================================
 * 7. HELPER - OVERLAP SAAT APPROVE
 * =================================================================== */

function pastikanTidakAdaIzinBentrokSaatApprove_(
  data,
  targetIndex
) {
  var target =
    data[targetIndex];

  var targetUser =
    String(
      target[KOLOM_IZIN.ID_USER] || ''
    );

  var targetMulai =
    normalizeTanggalIzin_(
      target[KOLOM_IZIN.TANGGAL_MULAI]
    );

  var targetSelesai =
    normalizeTanggalIzin_(
      target[KOLOM_IZIN.TANGGAL_SELESAI]
    );

  for (
    var i = 0;
    i < data.length;
    i++
  ) {
    if (i === targetIndex) {
      continue;
    }

    var row =
      data[i];

    if (
      String(
        row[KOLOM_IZIN.ID_USER] || ''
      ) !== targetUser
    ) {
      continue;
    }

    var status =
      String(
        row[KOLOM_IZIN.STATUS] || ''
      );

    /*
     * Menunggu juga dianggap konflik saat approve.
     * Admin sebaiknya memproses pengajuan yang bentrok
     * secara eksplisit, bukan membuat dua izin aktif.
     */
    if (
      status !==
        STATUS_IZIN_PENDING &&
      status !==
        STATUS_IZIN_DISETUJUI
    ) {
      continue;
    }

    var mulai =
      normalizeTanggalIzin_(
        row[KOLOM_IZIN.TANGGAL_MULAI]
      );

    var selesai =
      normalizeTanggalIzin_(
        row[KOLOM_IZIN.TANGGAL_SELESAI]
      );

    if (
      targetMulai <= selesai &&
      targetSelesai >= mulai
    ) {
      throw new Error(
        'Pengajuan tidak dapat disetujui karena tanggalnya ' +
        'beririsan dengan pengajuan lain yang masih aktif (' +
        status +
        ').'
      );
    }
  }
}


/* ===================================================================
 * 8. HELPER - VALIDASI TANGGAL
 * =================================================================== */

function validasiTanggalInputIzin_(
  value,
  label
) {
  var text =
    String(
      value || ''
    ).trim();

  if (
    !/^\d{4}-\d{2}-\d{2}$/.test(
      text
    )
  ) {
    throw new Error(
      label +
      ' harus menggunakan format yyyy-MM-dd.'
    );
  }

  /*
   * Validasi kalender secara ketat.
   */
  var parts =
    text.split('-');

  var year =
    parseInt(parts[0], 10);

  var month =
    parseInt(parts[1], 10);

  var day =
    parseInt(parts[2], 10);

  var date =
    new Date(
      year,
      month - 1,
      day
    );

  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    throw new Error(
      label +
      ' tidak valid.'
    );
  }

  return text;
}


function normalizeTanggalIzin_(
  value
) {
  if (
    value instanceof Date &&
    !isNaN(
      value.getTime()
    )
  ) {
    return formatTanggalIzin_(
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
   * Dukungan untuk format tanggal lama dari Sheets.
   */
  var parsed =
    new Date(text);

  if (
    !isNaN(
      parsed.getTime()
    )
  ) {
    return formatTanggalIzin_(
      parsed
    );
  }

  return text;
}


function formatTanggalIzin_(
  date
) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}


function hitungSelisihHariIzin_(
  mulai,
  selesai
) {
  var a =
    parseTanggalIzinUTC_(
      mulai
    );

  var b =
    parseTanggalIzinUTC_(
      selesai
    );

  return (
    Math.floor(
      (b - a) /
      86400000
    ) + 1
  );
}


function parseTanggalIzinUTC_(
  text
) {
  var parts =
    String(text).split('-');

  return new Date(
    Date.UTC(
      Number(parts[0]),
      Number(parts[1]) - 1,
      Number(parts[2])
    )
  );
}


/* ===================================================================
 * 9. HELPER - STATUS
 * =================================================================== */

function isStatusIzinValid_(
  status
) {
  return (
    status === STATUS_IZIN_PENDING ||
    status === STATUS_IZIN_DISETUJUI ||
    status === STATUS_IZIN_DITOLAK
  );
}


/* ===================================================================
 * 10. HELPER - UPLOAD LAMPIRAN
 * =================================================================== */

function simpanLampiranIzin_(
  fileBase64,
  namaFileTanpaExt
) {
  if (
    typeof fileBase64 !==
    'string'
  ) {
    throw new Error(
      'Lampiran tidak valid.'
    );
  }

  var bagian =
    fileBase64.split(',');

  var meta =
    bagian.length > 1
      ? bagian[0]
      : '';

  var base64Data =
    bagian.length > 1
      ? bagian[1]
      : bagian[0];

  if (
    !base64Data
  ) {
    throw new Error(
      'Data lampiran kosong.'
    );
  }

  var match =
    meta.match(
      /^data:(image\/(?:jpeg|jpg|png|webp)|application\/pdf);base64$/i
    );

  if (!match) {
    throw new Error(
      'Format lampiran tidak didukung. Gunakan JPG, PNG, WEBP, atau PDF.'
    );
  }

  var contentType =
    match[1].toLowerCase();

  if (
    contentType ===
    'image/jpg'
  ) {
    contentType =
      'image/jpeg';
  }

  var bytes;

  try {
    bytes =
      Utilities.base64Decode(
        base64Data
      );
  } catch (e) {
    throw new Error(
      'Data lampiran tidak dapat diproses.'
    );
  }

  if (
    bytes.length >
    MAKS_LAMPIRAN_IZIN_BYTES
  ) {
    throw new Error(
      'Ukuran lampiran terlalu besar. Maksimal 5MB.'
    );
  }

  var ekstensi;

  if (
    contentType ===
    'application/pdf'
  ) {
    ekstensi = 'pdf';
  } else if (
    contentType ===
    'image/png'
  ) {
    ekstensi = 'png';
  } else if (
    contentType ===
    'image/webp'
  ) {
    ekstensi = 'webp';
  } else {
    ekstensi = 'jpg';
  }

  var namaAman =
    sanitasiNamaFileIzin_(
      namaFileTanpaExt
    );

  var blob =
    Utilities.newBlob(
      bytes,
      contentType,
      namaAman +
        '.' +
        ekstensi
    );

  var folder =
    getOrCreateFolder_(
      FOLDER_IZIN_NAME
    );

  var file =
    folder.createFile(
      blob
    );

  /*
   * Kompatibel dengan frontend yang menampilkan link lampiran.
   * Pada tahap keamanan lanjutan, ini bisa dipindahkan ke endpoint
   * file privat yang memvalidasi session.
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


function sanitasiNamaFileIzin_(
  nama
) {
  var hasil =
    String(
      nama || 'lampiran_izin'
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
    hasil =
      'lampiran_izin';
  }

  return hasil.substring(
    0,
    100
  );
}


/* ===================================================================
 * 11. HELPER - OBJECT RESPONSE
 * =================================================================== */

function _barisIzinKeObjek_(
  row
) {
  return {
    id:
      row[KOLOM_IZIN.ID],
    id_izin:
      row[KOLOM_IZIN.ID],
    tanggal_pengajuan:
      normalizeTanggalIzin_(
        row[
          KOLOM_IZIN
            .TANGGAL_PENGAJUAN
        ]
      ),
    jenis_izin:
      row[
        KOLOM_IZIN.JENIS_IZIN
      ],
    tanggal_mulai:
      normalizeTanggalIzin_(
        row[
          KOLOM_IZIN.TANGGAL_MULAI
        ]
      ),
    tanggal_selesai:
      normalizeTanggalIzin_(
        row[
          KOLOM_IZIN.TANGGAL_SELESAI
        ]
      ),
    keterangan:
      row[
        KOLOM_IZIN.KETERANGAN
      ] || '',
    lampiran:
      row[
        KOLOM_IZIN.LAMPIRAN
      ] || '',
    status:
      row[
        KOLOM_IZIN.STATUS
      ] || '',
    catatan_admin:
      row[
        KOLOM_IZIN.CATATAN_ADMIN
      ] || ''
  };
}


function _barisIzinAdminKeObjek_(
  row,
  petaNama
) {
  var obj =
    _barisIzinKeObjek_(
      row
    );

  var idUser =
    String(
      row[KOLOM_IZIN.ID_USER] || ''
    );

  var diprosesOleh =
    String(
      row[
        KOLOM_IZIN.DIPROSES_OLEH
      ] || ''
    );

  obj.id_user =
    idUser;

  obj.nama_pegawai =
    petaNama[idUser] ||
    '(tidak diketahui)';

  obj.diproses_oleh =
    diprosesOleh
      ? (
          petaNama[
            diprosesOleh
          ] ||
          diprosesOleh
        )
      : '';

  obj.diproses_at =
    row[
      KOLOM_IZIN.DIPROSES_AT
    ] || '';

  return obj;
}


/* ===================================================================
 * 12. HELPER - PETA NAMA USER
 * =================================================================== */

function _petaNamaUser_() {
  var sheetUsers =
    getSheet(SHEET_USERS);

  var lastRow =
    sheetUsers.getLastRow();

  var peta = {};

  if (lastRow < 2) {
    return peta;
  }

  /*
   * Hanya ambil dua kolom pertama:
   * id dan nama.
   */
  var data =
    sheetUsers.getRange(
      2,
      1,
      lastRow - 1,
      2
    ).getValues();

  for (
    var i = 0;
    i < data.length;
    i++
  ) {
    var id =
      String(
        data[i][0] || ''
      );

    if (id) {
      peta[id] =
        String(
          data[i][1] || ''
        );
    }
  }

  return peta;
}
