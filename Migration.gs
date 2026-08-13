/**
 * ===================================================================
 * MIGRATION.GS
 * Pemeriksaan dan migrasi database secara NON-DESTRUKTIF.
 *
 * Tujuan:
 * 1. Mengecek struktur seluruh sheet.
 * 2. Membuat backup sebelum migrasi.
 * 3. Memetakan kolom berdasarkan NAMA HEADER, bukan nomor kolom.
 * 4. Tidak menghapus data lama.
 * 5. Tidak menjalankan migrasi otomatis tanpa perintah eksplisit.
 *
 * Jalankan:
 *   auditDatabase()
 *
 * Jika hasil audit menyatakan perlu migrasi:
 *   migrasiDatabaseAman()
 *
 * Setelah migrasi:
 *   verifikasiDatabase()
 * ===================================================================
 */


/* ===================================================================
 * STRUKTUR DATABASE FINAL
 * =================================================================== */

var DB_SCHEMA_FINAL = {

  Users: [
    'id_user',
    'nama',
    'username',
    'password_hash',
    'salt',
    'role',
    'jabatan',
    'foto_profil',
    'status_aktif',
    'created_at',
    'updated_at'
  ],

  Absensi: [
    'id_absen',
    'id_user',
    'tanggal',
    'jam_masuk',
    'foto_masuk',
    'lat_masuk',
    'long_masuk',
    'jam_pulang',
    'foto_pulang',
    'lat_pulang',
    'long_pulang',
    'status',
    'jarak_masuk_meter',
    'jarak_pulang_meter'
  ],

  Izin: [
    'id_izin',
    'id_user',
    'tanggal_pengajuan',
    'jenis_izin',
    'tanggal_mulai',
    'tanggal_selesai',
    'keterangan',
    'lampiran',
    'status',
    'catatan_admin',
    'diproses_oleh',
    'diproses_at'
  ],

  Pengaturan: [
    'nama_app',
    'logo_url',
    'lat_kantor',
    'long_kantor',
    'radius_meter',
    'jam_masuk',
    'jam_pulang',
    'toleransi_menit'
  ],

  Sessions: [
    'token',
    'id_user',
    'created_at',
    'expired_at'
  ]
};


/* ===================================================================
 * AUDIT
 * =================================================================== */

function auditDatabase() {

  var ss =
    SpreadsheetApp.getActiveSpreadsheet();

  var hasil = {
    success: true,
    timestamp: new Date(),
    sheets: [],
    semua_valid: true
  };

  Object.keys(DB_SCHEMA_FINAL).forEach(
    function(namaSheet) {

      var expected =
        DB_SCHEMA_FINAL[namaSheet];

      var sheet =
        ss.getSheetByName(namaSheet);

      if (!sheet) {

        hasil.semua_valid = false;

        hasil.sheets.push({
          sheet: namaSheet,
          exists: false,
          valid: false,
          message:
            'Sheet tidak ditemukan.',
          expected_headers:
            expected
        });

        return;
      }

      var lastColumn =
        sheet.getLastColumn();

      var lastRow =
        sheet.getLastRow();

      if (
        lastColumn === 0
      ) {

        hasil.semua_valid = false;

        hasil.sheets.push({
          sheet: namaSheet,
          exists: true,
          valid: false,
          rows: 0,
          columns: 0,
          message:
            'Sheet kosong.',
          expected_headers:
            expected
        });

        return;
      }

      var headers =
        sheet
          .getRange(
            1,
            1,
            1,
            lastColumn
          )
          .getValues()[0]
          .map(
            function(v) {
              return String(
                v || ''
              ).trim();
            }
          );

      var exact =
        headers.length ===
        expected.length;

      if (exact) {

        for (
          var i = 0;
          i < expected.length;
          i++
        ) {

          if (
            headers[i] !==
            expected[i]
          ) {
            exact = false;
            break;
          }
        }
      }

      var missing = [];

      expected.forEach(
        function(header) {

          if (
            headers.indexOf(
              header
            ) === -1
          ) {
            missing.push(header);
          }

        }
      );

      var extra = [];

      headers.forEach(
        function(header) {

          if (
            header &&
            expected.indexOf(
              header
            ) === -1
          ) {
            extra.push(header);
          }

        }
      );

      if (!exact) {
        hasil.semua_valid = false;
      }

      hasil.sheets.push({
        sheet: namaSheet,
        exists: true,
        valid: exact,
        rows:
          Math.max(
            0,
            lastRow - 1
          ),
        columns:
          lastColumn,
        current_headers:
          headers,
        expected_headers:
          expected,
        missing_headers:
          missing,
        extra_headers:
          extra
      });

    }
  );

  Logger.log(
    JSON.stringify(
      hasil,
      null,
      2
    )
  );

  return hasil;
}


/* ===================================================================
 * BACKUP
 * =================================================================== */

function backupDatabase() {

  var lock =
    LockService.getScriptLock();

  try {

    lock.waitLock(30000);

    var ss =
      SpreadsheetApp
        .getActiveSpreadsheet();

    var timestamp =
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        'yyyyMMdd_HHmmss'
      );

    var backupSpreadsheet =
      SpreadsheetApp.create(
        'BACKUP_ABSENSI_' +
        timestamp
      );

    var backupId =
      backupSpreadsheet.getId();

    Object.keys(DB_SCHEMA_FINAL)
      .forEach(
        function(namaSheet) {

          var source =
            ss.getSheetByName(
              namaSheet
            );

          if (!source) {
            return;
          }

          var sourceRows =
            source.getLastRow();

          var sourceCols =
            source.getLastColumn();

          var target =
            backupSpreadsheet
              .getSheets()[0];

          if (
            target.getName() ===
            'Sheet1'
          ) {

            target.setName(
              namaSheet
            );

          } else {

            target =
              backupSpreadsheet
                .insertSheet(
                  namaSheet
                );
          }

          if (
            sourceRows > 0 &&
            sourceCols > 0
          ) {

            var values =
              source
                .getRange(
                  1,
                  1,
                  sourceRows,
                  sourceCols
                )
                .getValues();

            target
              .getRange(
                1,
                1,
                values.length,
                values[0].length
              )
              .setValues(values);

          }

          target.setFrozenRows(1);

        }
      );

    return {
      success: true,
      backup_spreadsheet_id:
        backupId,
      backup_url:
        backupSpreadsheet.getUrl(),
      message:
        'Backup berhasil dibuat.'
    };

  } finally {

    try {
      lock.releaseLock();
    } catch (ignore) {}

  }
}


/* ===================================================================
 * MIGRASI AMAN
 *
 * HANYA bekerja jika header tidak sesuai.
 *
 * Mekanisme:
 * OLD:
 *   Users
 *
 * menjadi:
 *   Users_BACKUP_...
 *   Users
 *
 * Data baru dipetakan berdasarkan nama header.
 *
 * Kolom yang tidak ada pada database lama diisi kosong.
 * =================================================================== */

function migrasiDatabaseAman() {

  var lock =
    LockService.getScriptLock();

  try {

    lock.waitLock(30000);

    var audit =
      auditDatabase();

    if (
      audit.semua_valid
    ) {

      return {
        success: true,
        migrated: false,
        message:
          'Database sudah menggunakan struktur final. Tidak ada migrasi yang diperlukan.'
      };

    }

    /*
     * Backup terlebih dahulu.
     */
    var backup =
      backupDatabase();

    var ss =
      SpreadsheetApp
        .getActiveSpreadsheet();

    var timestamp =
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone(),
        'yyyyMMdd_HHmmss'
      );

    var migratedSheets = [];

    Object.keys(DB_SCHEMA_FINAL)
      .forEach(
        function(namaSheet) {

          var expected =
            DB_SCHEMA_FINAL[namaSheet];

          var oldSheet =
            ss.getSheetByName(
              namaSheet
            );

          /*
           * Jika sheet belum ada, buat baru.
           */
          if (!oldSheet) {

            var newSheet =
              ss.insertSheet(
                namaSheet
              );

            newSheet
              .getRange(
                1,
                1,
                1,
                expected.length
              )
              .setValues([
                expected
              ]);

            newSheet.setFrozenRows(1);

            migratedSheets.push({
              sheet: namaSheet,
              action: 'created'
            });

            return;
          }

          var oldLastColumn =
            oldSheet.getLastColumn();

          var oldLastRow =
            oldSheet.getLastRow();

          var oldHeaders = [];

          if (
            oldLastColumn > 0
          ) {

            oldHeaders =
              oldSheet
                .getRange(
                  1,
                  1,
                  1,
                  oldLastColumn
                )
                .getValues()[0]
                .map(
                  function(v) {
                    return String(
                      v || ''
                    ).trim();
                  }
                );
          }

          var isExact =
            headersSamaPersis_(
              oldHeaders,
              expected
            );

          if (isExact) {

            return;

          }

          /*
           * Rename sheet lama sebagai backup lokal.
           */
          var oldName =
            namaSheet +
            '_BACKUP_' +
            timestamp;

          /*
           * Hindari nama > 100 karakter
           */
          oldName =
            oldName.substring(
              0,
              99
            );

          oldSheet.setName(
            oldName
          );

          /*
           * Buat sheet baru dengan nama asli.
           */
          var newSheet =
            ss.insertSheet(
              namaSheet
            );

          newSheet
            .getRange(
              1,
              1,
              1,
              expected.length
            )
            .setValues([
              expected
            ]);

          newSheet.setFrozenRows(1);

          /*
           * Jika tidak ada data lama,
           * selesai.
           */
          if (
            oldLastRow < 2 ||
            oldLastColumn < 1
          ) {

            migratedSheets.push({
              sheet: namaSheet,
              action: 'recreated_empty'
            });

            return;

          }

          var oldValues =
            oldSheet
              .getRange(
                2,
                1,
                oldLastRow - 1,
                oldLastColumn
              )
              .getValues();

          /*
           * Mapping:
           *
           * old header → index
           */
          var oldMap = {};

          oldHeaders.forEach(
            function(header, index) {

              if (
                header
              ) {
                oldMap[
                  header
                ] = index;
              }

            }
          );

          /*
           * Buat data baru berdasarkan
           * header final.
           */
          var newValues =
            [];

          for (
            var r = 0;
            r < oldValues.length;
            r++
          ) {

            var newRow = [];

            for (
              var c = 0;
              c < expected.length;
              c++
            ) {

              var header =
                expected[c];

              if (
                Object.prototype
                  .hasOwnProperty
                  .call(
                    oldMap,
                    header
                  )
              ) {

                newRow.push(
                  oldValues[r][
                    oldMap[header]
                  ]
                );

              } else {

                newRow.push('');

              }

            }

            newValues.push(
              newRow
            );
          }

          /*
           * Tulis batch.
           */
          if (
            newValues.length
          ) {

            newSheet
              .getRange(
                2,
                1,
                newValues.length,
                expected.length
              )
              .setValues(
                newValues
              );

          }

          migratedSheets.push({
            sheet: namaSheet,
            action: 'migrated',
            old_backup_sheet:
              oldName,
            rows:
              newValues.length
          });

        }
      );

    /*
     * Bersihkan cache pengaturan
     * setelah struktur database berubah.
     */
    try {
      bersihkanCachePengaturan();
    } catch (ignore) {}

    return {
      success: true,
      migrated: true,
      backup:
        backup,
      sheets:
        migratedSheets,
      message:
        'Migrasi database selesai. Sheet lama dipertahankan sebagai backup.'
    };

  } finally {

    try {
      lock.releaseLock();
    } catch (ignore) {}

  }
}


/* ===================================================================
 * VERIFIKASI SETELAH MIGRASI
 * =================================================================== */

function verifikasiDatabase() {

  var hasil =
    auditDatabase();

  if (
    hasil.semua_valid
  ) {

    return {
      success: true,
      valid: true,
      message:
        'Database VALID dan sesuai struktur final.',
      audit: hasil
    };

  }

  return {
    success: false,
    valid: false,
    message:
      'Database masih memiliki struktur yang belum sesuai.',
    audit: hasil
  };
}


/* ===================================================================
 * HELPER
 * =================================================================== */

function headersSamaPersis_(
  actual,
  expected
) {

  if (
    !Array.isArray(actual) ||
    !Array.isArray(expected)
  ) {
    return false;
  }

  if (
    actual.length !==
    expected.length
  ) {
    return false;
  }

  for (
    var i = 0;
    i < expected.length;
    i++
  ) {

    if (
      String(
        actual[i] || ''
      ).trim() !==
      String(
        expected[i] || ''
      ).trim()
    ) {
      return false;
    }

  }

  return true;
}


/* ===================================================================
 * CEK DATA KRITIS
 * =================================================================== */

function cekIntegritasDatabase() {

  var ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  var hasil = {
    success: true,
    errors: [],
    warnings: []
  };

  /*
   * USERS
   */
  var users =
    ss.getSheetByName(
      SHEET_USERS
    );

  if (users) {

    var rows =
      users.getLastRow();

    if (rows >= 2) {

      var data =
        users
          .getRange(
            2,
            1,
            rows - 1,
            11
          )
          .getValues();

      var ids = {};
      var usernames = {};

      data.forEach(
        function(row, index) {

          var nomor =
            index + 2;

          var id =
            String(
              row[0] || ''
            ).trim();

          var username =
            String(
              row[2] || ''
            )
              .trim()
              .toLowerCase();

          if (!id) {

            hasil.errors.push(
              'Users baris ' +
              nomor +
              ': id_user kosong.'
            );

          } else if (
            ids[id]
          ) {

            hasil.errors.push(
              'Users baris ' +
              nomor +
              ': id_user duplikat ' +
              id
            );

          } else {

            ids[id] = true;

          }

          if (!username) {

            hasil.errors.push(
              'Users baris ' +
              nomor +
              ': username kosong.'
            );

          } else if (
            usernames[username]
          ) {

            hasil.errors.push(
              'Users baris ' +
              nomor +
              ': username duplikat ' +
              username
            );

          } else {

            usernames[username] =
              true;

          }

        }
      );

    }

  }


  /*
   * ABSENSI
   */
  var absensi =
    ss.getSheetByName(
      SHEET_ABSENSI
    );

  if (absensi) {

    var lastRow =
      absensi.getLastRow();

    if (lastRow >= 2) {

      var dataAbsensi =
        absensi
          .getRange(
            2,
            1,
            lastRow - 1,
            14
          )
          .getValues();

      dataAbsensi.forEach(
        function(row, index) {

          var nomor =
            index + 2;

          if (
            !String(
              row[0] || ''
            ).trim()
          ) {

            hasil.errors.push(
              'Absensi baris ' +
              nomor +
              ': id_absen kosong.'
            );

          }

          if (
            !String(
              row[1] || ''
            ).trim()
          ) {

            hasil.errors.push(
              'Absensi baris ' +
              nomor +
              ': id_user kosong.'
            );

          }

        }
      );

    }

  }


  /*
   * IZIN
   */
  var izin =
    ss.getSheetByName(
      SHEET_IZIN
    );

  if (izin) {

    var lastRowIzin =
      izin.getLastRow();

    if (lastRowIzin >= 2) {

      var dataIzin =
        izin
          .getRange(
            2,
            1,
            lastRowIzin - 1,
            12
          )
          .getValues();

      dataIzin.forEach(
        function(row, index) {

          var nomor =
            index + 2;

          if (
            !String(
              row[0] || ''
            ).trim()
          ) {

            hasil.errors.push(
              'Izin baris ' +
              nomor +
              ': id_izin kosong.'
            );

          }

          if (
            !String(
              row[1] || ''
            ).trim()
          ) {

            hasil.errors.push(
              'Izin baris ' +
              nomor +
              ': id_user kosong.'
            );

          }

        }
      );

    }

  }


  hasil.success =
    hasil.errors.length === 0;

  hasil.message =
    hasil.success
      ? 'Tidak ditemukan error integritas kritis.'
      : 'Ditemukan ' +
        hasil.errors.length +
        ' error integritas.';

  Logger.log(
    JSON.stringify(
      hasil,
      null,
      2
    )
  );

  return hasil;
}
