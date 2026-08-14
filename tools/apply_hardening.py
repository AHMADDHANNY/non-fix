from pathlib import Path

ROOT = Path('.')


def patch(path, replacements):
    p = ROOT / path
    text = p.read_text(encoding='utf-8')
    original = text
    for old, new in replacements:
        if old not in text:
            raise SystemExit(f'Patch target not found in {path}: {old[:120]!r}')
        text = text.replace(old, new, 1)
    if text != original:
        p.write_text(text, encoding='utf-8')
        print('patched', path)


patch('Auth.gs', [
    ("rowUsername === username &&\n        statusAktif === 'aktif'", "rowUsername.toLowerCase() === username.toLowerCase() &&\n        statusAktif === 'aktif'"),
    ("message:\n        'Terjadi kesalahan server: ' +\n        error.message", "message:\n        'Terjadi kesalahan server. Silakan coba lagi.'"),
])

patch('Admin.gs', [
    ("fotoProfil = String(fotoProfil || '').trim();\n    statusAktif = String(statusAktif || '').trim().toLowerCase();", "fotoProfil = String(fotoProfil || '').trim();\n    statusAktif = String(statusAktif || '').trim().toLowerCase();"),
    ("    sheet.getRange(\n      row,\n      KOLOM_USERS_ADMIN.FOTO_PROFIL + 1\n    ).setValue(fotoProfil);", "    /* Foto kosong berarti pertahankan foto lama. */\n    if (fotoProfil) {\n      sheet.getRange(\n        row,\n        KOLOM_USERS_ADMIN.FOTO_PROFIL + 1\n      ).setValue(fotoProfil);\n    }"),
])

patch('DashboardPegawai.html', [
    ("    var TOKEN_KEY =\n      'absensi_session_token';", "    var TOKEN_KEY =\n      'absensi_session_token';\n\n    var logoutBusy = false;"),
    ("        .withFailureHandler(\n          function (error) {\n\n            showMessage(\n              normalizeError(error),\n              true\n            );\n\n          }\n        )\n\n        .verifySession(\n          token\n        );", "        .withFailureHandler(\n          function (error) {\n            clearSession();\n            showMessage(\n              'Sesi tidak dapat diverifikasi. Silakan login kembali.',\n              true\n            );\n            setTimeout(redirectLogin, 500);\n          }\n        )\n\n        .verifySession(\n          token\n        );"),
    ("        .withFailureHandler(\n          function () {\n\n            /*\n             * Dashboard dasar tetap tampil\n             * walaupun endpoint absensi pegawai\n             * belum tersedia.\n             */\n\n          }\n        )", "        .withFailureHandler(\n          function (error) {\n            var message = normalizeError(error);\n            if (/sesi|session|token|akses|unauthorized|expired/i.test(message)) {\n              clearSession();\n              redirectLogin();\n            } else {\n              showMessage(message, true);\n            }\n          }\n        )"),
    ("    document\n      .getElementById('logoutButton')\n      .addEventListener(\n        'click',\n        function () {\n\n          var token =\n            getToken();", "    document\n      .getElementById('logoutButton')\n      .addEventListener(\n        'click',\n        function () {\n          if (logoutBusy) return;\n          logoutBusy = true;\n          var button = document.getElementById('logoutButton');\n          if (button) { button.disabled = true; button.textContent = 'Keluar...'; }\n\n          var token =\n            getToken();"),
    ("      var base =\n        window.location.href\n          .split('?')[0];\n\n\n      window.top.location.href =\n        base + '?page=Login';", "      var base = String(window.ABSENSI_WEB_APP_URL || '').trim();\n      if (!base) {\n        base = window.location.href.split('?')[0];\n      }\n      window.top.location.href = base + (base.indexOf('?') >= 0 ? '&' : '?') + 'page=Login';"),
])

patch('AdminJS.html', [
    ("  var loading =\n    false;", "  var loading =\n    false;\n\n  var logoutBusy = false;"),
    ("          showToast(\n            normalizeError(error),\n            'error'\n          );\n\n        }\n      )\n\n      .getDashboardAdmin(", "          var message = normalizeError(error);\n          if (/sesi|session|token|akses|unauthorized|expired/i.test(message)) {\n            clearSession();\n            redirectLogin();\n            return;\n          }\n          showToast(message, 'error');\n\n        }\n      )\n\n      .getDashboardAdmin("),
    ("  function performLogout() {\n\n  var token =", "  function performLogout() {\n\n  if (logoutBusy) return;\n  logoutBusy = true;\n\n  var confirmButton = document.getElementById('confirmLogout');\n  if (confirmButton) { confirmButton.disabled = true; confirmButton.textContent = 'Keluar...'; }\n\n  var token ="),
])

# Migration: avoid re-entering the same ScriptLock from migrasiDatabaseAman -> backupDatabase.
p = ROOT / 'Migration.gs'
text = p.read_text(encoding='utf-8')
old = "function backupDatabase() {\n\n  var lock =\n    LockService.getScriptLock();\n\n  try {\n\n    lock.waitLock(30000);"
new = "function backupDatabase(skipLock) {\n\n  var lock = skipLock ? null : LockService.getScriptLock();\n\n  try {\n\n    if (lock) lock.waitLock(30000);"
if old not in text:
    raise SystemExit('backupDatabase lock block not found')
text = text.replace(old, new, 1)
old2 = "  } finally {\n\n    try {\n      lock.releaseLock();\n    } catch (ignore) {}\n\n  }\n}\n\n\n/* ===================================================================\n * MIGRASI AMAN"
new2 = "  } finally {\n\n    if (lock) {\n      try {\n        lock.releaseLock();\n      } catch (ignore) {}\n    }\n\n  }\n}\n\n\n/* ===================================================================\n * MIGRASI AMAN"
if old2 not in text:
    raise SystemExit('backupDatabase finally block not found')
text = text.replace(old2, new2, 1)
old3 = "    var backup =\n      backupDatabase();"
new3 = "    var backup =\n      backupDatabase(true);"
if old3 not in text:
    raise SystemExit('migration backup call not found')
text = text.replace(old3, new3, 1)
p.write_text(text, encoding='utf-8')
print('patched Migration.gs')

print('hardening patch complete')
