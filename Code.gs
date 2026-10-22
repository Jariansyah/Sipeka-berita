/**
 * ============================================================================
 * SiPeKA BERITA RRI - SISTEM INFORMASI PELAPORAN BERITA & PROGRAM LPP RRI
 * Arsitektur : Single Page Application (SPA) - Modular 4 Files
 * Backend    : Google Apps Script (V8 Engine) + Google Sheets Database
 * Keamanan   : SHA-256 Password Hashing, Session Token, & ScriptLock
 * ============================================================================
 */

const APP_TITLE = 'SiPeKA BERITA RRI - Sistem Informasi Pelaporan Kinerja dan Administrasi RRI Jambi';
const SESSION_EXPIRY_HOURS = 12;

/**
 * Entry Point Web App (Mendukung Embed Iframe Super Apps RRI)
 */
function doGet(e) {
  initDatabase();
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle(APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
}

/**
 * Helper untuk mengimpor file CSS dan JS modular ke dalam Index.html
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Helper Enkripsi SHA-256 untuk Password & Token
 */
function hashSHA256(input) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(input), Utilities.Charset.UTF_8);
  let txtHash = '';
  for (let i = 0; i < rawHash.length; i++) {
    let hashVal = rawHash[i];
    if (hashVal < 0) hashVal += 256;
    if (hashVal.toString(16).length === 1) txtHash += '0';
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

/**
 * Helper Format Tanggal ke String YYYY-MM-DD sesuai Timezone Script
 */
function formatDateYMD(dateInput) {
  if (!dateInput) return '';
  if (typeof dateInput === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(dateInput.trim())) {
    return dateInput.trim();
  }
  const d = new Date(dateInput);
  if (isNaN(d.getTime())) return String(dateInput);
  const tz = Session.getScriptTimeZone() || 'Asia/Jakarta';
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

/**
 * Helper Format Timestamp Lengkap
 */
function formatTimestamp(dateInput) {
  if (!dateInput) return '';
  const d = new Date(dateInput);
  if (isNaN(d.getTime())) return String(dateInput);
  const tz = Session.getScriptTimeZone() || 'Asia/Jakarta';
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm:ss');
}

/**
 * ============================================================================
 * INISIALISASI OTOMATIS DATABASE GOOGLE SPREADSHEET & AUTO-MIGRASI DIALOG
 * ============================================================================
 */
function initDatabase() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Sheet DATABASE_BERITA
    let shBerita = ss.getSheetByName('DATABASE_BERITA');
    if (!shBerita) {
      shBerita = ss.insertSheet('DATABASE_BERITA');
      shBerita.appendRow([
        'ID_Berita',
        'Timestamp_Input',
        'Tanggal_Siar',
        'ID_Format',
        'Nama_Format',
        'Nama_Reporter',
        'Judul_Berita_Topik',
        'Narasumber_Jabatan',
        'Jenis_Laporan',
        'Link_Berita_Audio',
        'Keterangan_Tambahan',
        'Penginput_User'
      ]);
      styleHeaderRow(shBerita, 12);
    }

    // 2. Sheet MASTER_REPORTER
    let shReporter = ss.getSheetByName('MASTER_REPORTER');
    if (!shReporter) {
      shReporter = ss.insertSheet('MASTER_REPORTER');
      shReporter.appendRow(['ID_Reporter', 'Nama_Reporter', 'NIP_ID', 'Jabatan', 'Status_Aktif']);
      styleHeaderRow(shReporter, 5);
      const defaultReporters = [
        ['REP-001', 'Ahmad Fauzi, S.I.Kom', '199003122019021001', 'Reporter Muda', 'Aktif'],
        ['REP-002', 'Siti Rahmawati', '199407212022032004', 'Reporter Pelaksana', 'Aktif'],
        ['REP-003', 'Budi Santoso', '198811052015011002', 'Redaktur / Pengarah Acara', 'Aktif']
      ];
      shReporter.getRange(2, 1, defaultReporters.length, 5).setValues(defaultReporters);
    }

    // 3. Sheet MASTER_FORMAT_ACARA (Dilengkapi Dialog Pagi & Dialog Kentongan)
    let shFormat = ss.getSheetByName('MASTER_FORMAT_ACARA');
    if (!shFormat) {
      shFormat = ss.insertSheet('MASTER_FORMAT_ACARA');
      shFormat.appendRow(['ID_Format', 'Nama_Format', 'Kategori_Form', 'Ikon_FA', 'Status_Aktif']);
      styleHeaderRow(shFormat, 5);
      const defaultFormats = [
        ['FMT-01', 'Buletin', 'STANDAR', 'fa-newspaper', 'Aktif'],
        ['FMT-02', 'Dialog Pagi', 'MULTI_NARA', 'fa-comments', 'Aktif'],
        ['FMT-03', 'Dialog Kentongan', 'MULTI_NARA', 'fa-comments', 'Aktif'],
        ['FMT-04', 'Wawancara', 'STANDAR', 'fa-microphone-lines', 'Aktif'],
        ['FMT-05', 'Siaran Langsung', 'STANDAR', 'fa-tower-broadcast', 'Aktif'],
        ['FMT-06', 'News Feature', 'STANDAR', 'fa-file-audio', 'Aktif'],
        ['FMT-07', 'News Indept', 'MULTI_NARA', 'fa-magnifying-glass-chart', 'Aktif'],
        ['FMT-08', 'ROS (Radio On The Spot)', 'STANDAR', 'fa-street-view', 'Aktif'],
        ['FMT-09', 'Berita Online KBRN', 'KBRN', 'fa-globe', 'Aktif'],
        ['FMT-10', 'Dokumenter', 'MULTI_NARA', 'fa-film', 'Aktif']
      ];
      shFormat.getRange(2, 1, defaultFormats.length, 5).setValues(defaultFormats);
    } else {
      // Auto-Migrasi jika sheet lama masih menggunakan "Dialog" tunggal
      migrateDialogFormatsIfNeeded(shFormat);
    }

    // 4. Sheet MASTER_USER
    let shUser = ss.getSheetByName('MASTER_USER');
    if (!shUser) {
      shUser = ss.insertSheet('MASTER_USER');
      shUser.appendRow(['Username', 'Password_Hash', 'Nama_Lengkap', 'Role', 'Status_Akun']);
      styleHeaderRow(shUser, 5);
      const defaultUsers = [
        ['superadmin', hashSHA256('adminrri'), 'Super Administrator RRI', 'Super Admin', 'Aktif'],
        ['reporter', hashSHA256('rri123'), 'Tim Reporter Pemberitaan', 'User', 'Aktif']
      ];
      shUser.getRange(2, 1, defaultUsers.length, 5).setValues(defaultUsers);
    }

    // 5. Sheet PENGATURAN_SISTEM
    let shConfig = ss.getSheetByName('PENGATURAN_SISTEM');
    if (!shConfig) {
      shConfig = ss.insertSheet('PENGATURAN_SISTEM');
      shConfig.appendRow(['Kunci_Konfigurasi', 'Nilai_Konfigurasi']);
      styleHeaderRow(shConfig, 2);
      const defaultConfigs = [
        ['Nama_Stasiun', 'LPP RRI Jambi'],
        ['Sub_Judul_Aplikasi', 'Sistem Informasi Pelaporan Kinerja dan Administrasi'],
        ['Nama_Pejabat_Tanda_Tangan', 'Nama Ketua Tim Pemberitaan, S.Sos'],
        ['NIP_Pejabat', '19780510 200312 1 002'],
        ['Kota_Stasiun', 'Jambi'],
        ['Status_Maintenance', 'TIDAK']
      ];
      shConfig.getRange(2, 1, defaultConfigs.length, 2).setValues(defaultConfigs);
    }

    // 6. Sheet LOG_AKTIVITAS
    let shLog = ss.getSheetByName('LOG_AKTIVITAS');
    if (!shLog) {
      shLog = ss.insertSheet('LOG_AKTIVITAS');
      shLog.appendRow(['Timestamp', 'Username', 'Role', 'Aksi', 'Detail_Aktivitas']);
      styleHeaderRow(shLog, 5);
    }

    return { status: 'success', message: 'Database siap.' };
  } catch (err) {
    return { status: 'error', message: 'Gagal inisialisasi database: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Auto-migrasi aman untuk mengubah format "Dialog" lama menjadi "Dialog Pagi" & "Dialog Kentongan"
 */
function migrateDialogFormatsIfNeeded(shFormat) {
  if (!shFormat || shFormat.getLastRow() <= 1) return;
  const data = shFormat.getDataRange().getValues();
  let hasDialogPagi = false;
  let hasDialogKentongan = false;
  let genericDialogRowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    const nama = String(data[i][1] || '').trim().toLowerCase();
    if (nama === 'dialog pagi') hasDialogPagi = true;
    if (nama === 'dialog kentongan') hasDialogKentongan = true;
    if (nama === 'dialog') genericDialogRowIndex = i + 1;
  }

  if (genericDialogRowIndex !== -1 && !hasDialogPagi) {
    shFormat.getRange(genericDialogRowIndex, 2, 1, 3).setValues([['Dialog Pagi', 'MULTI_NARA', 'fa-comments']]);
    hasDialogPagi = true;
  }

  if (!hasDialogPagi) {
    shFormat.appendRow(['FMT-DLG1', 'Dialog Pagi', 'MULTI_NARA', 'fa-comments', 'Aktif']);
  }
  if (!hasDialogKentongan) {
    shFormat.appendRow(['FMT-DLG2', 'Dialog Kentongan', 'MULTI_NARA', 'fa-comments', 'Aktif']);
  }
}

/**
 * Styling Baris Header Sheet
 */
function styleHeaderRow(sheet, colCount) {
  const range = sheet.getRange(1, 1, 1, colCount);
  range.setBackground('#002B66')
       .setFontColor('#FFFFFF')
       .setFontWeight('bold')
       .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
}

/**
 * Pencatat Log Aktivitas Sistem
 */
function writeLog(username, role, aksi, detail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let shLog = ss.getSheetByName('LOG_AKTIVITAS');
    if (!shLog) return;
    shLog.appendRow([new Date(), username || 'Sistem', role || '-', aksi, detail]);
    const lastRow = shLog.getLastRow();
    if (lastRow > 1505) {
      shLog.deleteRows(2, lastRow - 1500);
    }
  } catch (e) {
    // Abaikan jika gagal tulis log non-kritis
  }
}

/**
 * ============================================================================
 * SISTEM AUTENTIKASI LOGIN, SESI TOKEN, & PENGATURAN PUBLIK
 * ============================================================================
 */
function getPublicConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shConfig = ss.getSheetByName('PENGATURAN_SISTEM');
  const config = {
    Nama_Stasiun: 'LPP RRI Jambi',
    Sub_Judul_Aplikasi: 'Sistem Informasi Pelaporan Kinerja dan Administrasi',
    Nama_Pejabat_Tanda_Tangan: 'Ketua Tim Pemberitaan',
    NIP_Pejabat: '-',
    Kota_Stasiun: 'Jambi',
    Status_Maintenance: 'TIDAK'
  };
  if (shConfig && shConfig.getLastRow() > 1) {
    const rows = shConfig.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const k = String(rows[i][0]).trim();
      const v = String(rows[i][1]).trim();
      if (k) config[k] = v;
    }
  }
  return config;
}

function createSessionToken(username, role, namaLengkap) {
  const exp = new Date().getTime() + (SESSION_EXPIRY_HOURS * 3600 * 1000);
  const payload = JSON.stringify({ u: username, r: role, n: namaLengkap, exp: exp });
  const encoded = Utilities.base64EncodeWebSafe(payload);
  const secret = ScriptApp.getScriptId() + '_SIPEKA_RRI_2026';
  const signature = hashSHA256(encoded + '.' + secret);
  return encoded + '.' + signature;
}

function verifySessionToken(token) {
  if (!token || typeof token !== 'string' || !token.includes('.')) return null;
  try {
    const parts = token.split('.');
    const encoded = parts[0];
    const sig = parts[1];
    const secret = ScriptApp.getScriptId() + '_SIPEKA_RRI_2026';
    const expectedSig = hashSHA256(encoded + '.' + secret);
    if (sig !== expectedSig) return null;

    const jsonStr = Utilities.newBlob(Utilities.base64DecodeWebSafe(encoded)).getDataAsString();
    const data = JSON.parse(jsonStr);
    if (new Date().getTime() > data.exp) return null;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shUser = ss.getSheetByName('MASTER_USER');
    const rows = shUser.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim().toLowerCase() === String(data.u).toLowerCase()) {
        if (String(rows[i][4]).trim() !== 'Aktif') return null;
        return {
          username: String(rows[i][0]).trim(),
          namaLengkap: String(rows[i][2]).trim(),
          role: String(rows[i][3]).trim()
        };
      }
    }
    return null;
  } catch (e) {
    return null;
  }
}

function loginUser(username, password) {
  try {
    initDatabase();
    const cleanUser = String(username || '').trim();
    const cleanPass = String(password || '');
    if (!cleanUser || !cleanPass) {
      return { status: 'error', message: 'Username dan Password wajib diisi.' };
    }

    const config = getPublicConfig();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shUser = ss.getSheetByName('MASTER_USER');
    const data = shUser.getDataRange().getValues();
    const passHash = hashSHA256(cleanPass);

    for (let i = 1; i < data.length; i++) {
      const dbUser = String(data[i][0]).trim();
      const dbHash = String(data[i][1]).trim();
      const dbNama = String(data[i][2]).trim();
      const dbRole = String(data[i][3]).trim();
      const dbStatus = String(data[i][4]).trim();

      if (dbUser.toLowerCase() === cleanUser.toLowerCase()) {
        const isMatch = (dbHash === passHash) || (dbHash === cleanPass);
        if (!isMatch) {
          return { status: 'error', message: 'Kata sandi yang Anda masukkan salah.' };
        }
        if (dbStatus !== 'Aktif') {
          return { status: 'error', message: 'Akun Anda sedang dinonaktifkan oleh Administrator.' };
        }
        if (config.Status_Maintenance === 'YA' && dbRole !== 'Super Admin') {
          return { status: 'error', message: 'Sistem sedang dalam pemeliharaan (Maintenance). Silakan hubungi Super Admin.' };
        }
        if (dbHash === cleanPass) {
          shUser.getRange(i + 1, 2).setValue(passHash);
        }

        const token = createSessionToken(dbUser, dbRole, dbNama);
        writeLog(dbUser, dbRole, 'LOGIN', 'Berhasil masuk ke aplikasi SiPeKA BERITA RRI');

        return {
          status: 'success',
          message: 'Autentikasi berhasil.',
          session: {
            token: token,
            username: dbUser,
            namaLengkap: dbNama,
            role: dbRole
          },
          bootstrap: getAppBootstrapData(dbUser, dbRole)
        };
      }
    }
    return { status: 'error', message: 'Username tidak ditemukan dalam sistem.' };
  } catch (err) {
    return { status: 'error', message: 'Kesalahan server saat login: ' + err.message };
  }
}

function validateSessionAndBootstrap(token) {
  try {
    initDatabase();
    const user = verifySessionToken(token);
    if (!user) {
      return { status: 'unauthorized', config: getPublicConfig() };
    }
    const config = getPublicConfig();
    if (config.Status_Maintenance === 'YA' && user.role !== 'Super Admin') {
      return { status: 'unauthorized', config: config, message: 'Sistem sedang Maintenance.' };
    }
    return {
      status: 'success',
      session: {
        token: token,
        username: user.username,
        namaLengkap: user.namaLengkap,
        role: user.role
      },
      bootstrap: getAppBootstrapData(user.username, user.role)
    };
  } catch (err) {
    return { status: 'error', message: err.message, config: getPublicConfig() };
  }
}

/**
 * Mengambil seluruh data awal yang dibutuhkan SPA dalam 1 kali panggilan
 */
function getAppBootstrapData(username, role) {
  const config = getPublicConfig();
  const reporters = getMasterReportersList();
  const formats = getMasterFormatsList();
  const berita = getAllBeritaRecords();
  let users = [];
  let logs = [];

  if (role === 'Super Admin') {
    users = getMasterUsersList();
    logs = getRecentActivityLogs(200);
  }

  return {
    config: config,
    reporters: reporters,
    formats: formats,
    berita: berita,
    users: users,
    logs: logs
  };
}

/**
 * ============================================================================
 * HELPER PEMBACAAN DATA MASTER & BERITA
 * ============================================================================
 */
function getMasterReportersList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('MASTER_REPORTER');
  if (!sh || sh.getLastRow() <= 1) return [];
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] && !data[i][1]) continue;
    list.push({
      idReporter: String(data[i][0]).trim(),
      namaReporter: String(data[i][1]).trim(),
      nipId: String(data[i][2]).trim(),
      jabatan: String(data[i][3]).trim(),
      statusAktif: String(data[i][4]).trim() || 'Aktif'
    });
  }
  return list;
}

function getMasterFormatsList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('MASTER_FORMAT_ACARA');
  if (!sh || sh.getLastRow() <= 1) return [];
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0] && !data[i][1]) continue;
    list.push({
      idFormat: String(data[i][0]).trim(),
      namaFormat: String(data[i][1]).trim(),
      kategoriForm: String(data[i][2]).trim() || 'STANDAR',
      ikonFA: String(data[i][3]).trim() || 'fa-tower-broadcast',
      statusAktif: String(data[i][4]).trim() || 'Aktif'
    });
  }
  return list;
}

function getMasterUsersList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('MASTER_USER');
  if (!sh || sh.getLastRow() <= 1) return [];
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    list.push({
      username: String(data[i][0]).trim(),
      namaLengkap: String(data[i][2]).trim(),
      role: String(data[i][3]).trim(),
      statusAkun: String(data[i][4]).trim() || 'Aktif'
    });
  }
  return list;
}

function getRecentActivityLogs(limitCount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('LOG_AKTIVITAS');
  if (!sh || sh.getLastRow() <= 1) return [];
  const data = sh.getDataRange().getValues();
  const logs = [];
  for (let i = data.length - 1; i >= 1; i--) {
    logs.push({
      timestamp: formatTimestamp(data[i][0]),
      username: String(data[i][1]),
      role: String(data[i][2]),
      aksi: String(data[i][3]),
      detail: String(data[i][4])
    });
    if (logs.length >= (limitCount || 200)) break;
  }
  return logs;
}

function getAllBeritaRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('DATABASE_BERITA');
  if (!sh || sh.getLastRow() <= 1) return [];
  const data = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < data.length; i++) {
    const idBerita = String(data[i][0]).trim();
    if (!idBerita) continue;
    list.push({
      idBerita: idBerita,
      timestampInput: formatTimestamp(data[i][1]),
      tanggalSiar: formatDateYMD(data[i][2]),
      idFormat: String(data[i][3]).trim(),
      namaFormat: String(data[i][4]).trim(),
      namaReporter: String(data[i][5]).trim(),
      judulBeritaTopik: String(data[i][6]).trim(),
      narasumberJabatan: String(data[i][7]).trim(),
      jenisLaporan: String(data[i][8]).trim() || 'Umum',
      linkBeritaAudio: String(data[i][9]).trim() || '-',
      keteranganTambahan: String(data[i][10]).trim(),
      penginputUser: String(data[i][11]).trim()
    });
  }
  list.sort(function(a, b) {
    if (b.tanggalSiar !== a.tanggalSiar) {
      return b.tanggalSiar.localeCompare(a.tanggalSiar);
    }
    return b.timestampInput.localeCompare(a.timestampInput);
  });
  return list;
}

/**
 * ============================================================================
 * OPERASI CRUD LAPORAN BERITA (DATABASE_BERITA) DENGAN SCRIPTLOCK
 * ============================================================================
 */
function saveBeritaRecord(token, payload) {
  const user = verifySessionToken(token);
  if (!user) return { status: 'unauthorized', message: 'Sesi Anda telah berakhir. Silakan login kembali.' };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('DATABASE_BERITA');

    const tanggalSiar = formatDateYMD(payload.tanggalSiar);
    const idFormat = String(payload.idFormat || '').trim();
    const namaFormat = String(payload.namaFormat || '').trim();
    const namaReporter = String(payload.namaReporter || '').trim();
    const judulBeritaTopik = String(payload.judulBeritaTopik || '').trim();
    const narasumberJabatan = String(payload.narasumberJabatan || '-').trim() || '-';
    // Menyimpan pilihan Kategori: Umum, RB Tematik, Prioritas Nasional, Astacita
    const jenisLaporan = String(payload.jenisLaporan || 'Umum').trim() || 'Umum';
    const linkBeritaAudio = '-';
    const keteranganTambahan = String(payload.keteranganTambahan || '-').trim() || '-';

    if (!tanggalSiar || !idFormat || !namaReporter || !judulBeritaTopik) {
      return { status: 'error', message: 'Tanggal Siar, Format Acara, Nama Reporter, dan Judul Berita wajib diisi!' };
    }

    // Jika mode Edit (ada idBerita)
    if (payload.idBerita) {
      const targetId = String(payload.idBerita).trim();
      const data = sh.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === targetId) {
          const ownerUser = String(data[i][11]).trim();
          if (user.role !== 'Super Admin' && ownerUser.toLowerCase() !== user.username.toLowerCase()) {
            return { status: 'error', message: 'Akses ditolak: Anda hanya dapat mengedit laporan milik Anda sendiri.' };
          }
          sh.getRange(i + 1, 3, 1, 9).setValues([[
            tanggalSiar,
            idFormat,
            namaFormat,
            namaReporter,
            judulBeritaTopik,
            narasumberJabatan,
            jenisLaporan,
            linkBeritaAudio,
            keteranganTambahan
          ]]);
          writeLog(user.username, user.role, 'EDIT_BERITA', 'Mengubah laporan ID ' + targetId + ' (' + judulBeritaTopik + ')');
          return {
            status: 'success',
            message: 'Laporan berita berhasil diperbarui!',
            berita: getAllBeritaRecords()
          };
        }
      }
      return { status: 'error', message: 'Data laporan berita tidak ditemukan.' };
    } else {
      // Mode Tambah Baru
      const newId = 'BRT-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyyMMddHHmmss') + '-' + Math.floor(100 + Math.random() * 900);
      const now = new Date();
      sh.appendRow([
        newId,
        now,
        tanggalSiar,
        idFormat,
        namaFormat,
        namaReporter,
        judulBeritaTopik,
        narasumberJabatan,
        jenisLaporan,
        linkBeritaAudio,
        keteranganTambahan,
        user.username
      ]);
      writeLog(user.username, user.role, 'INPUT_BERITA', 'Menambahkan laporan ' + namaFormat + ' [' + jenisLaporan + '] oleh ' + namaReporter + ': ' + judulBeritaTopik);
      return {
        status: 'success',
        message: 'Laporan berita berhasil disimpan ke database!',
        berita: getAllBeritaRecords()
      };
    }
  } catch (err) {
    return { status: 'error', message: 'Gagal menyimpan berita: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteBeritaRecord(token, idBerita) {
  const user = verifySessionToken(token);
  if (!user) return { status: 'unauthorized', message: 'Sesi Anda telah berakhir. Silakan login kembali.' };

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('DATABASE_BERITA');
    const data = sh.getDataRange().getValues();
    const targetId = String(idBerita || '').trim();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === targetId) {
        const ownerUser = String(data[i][11]).trim();
        const judul = String(data[i][6]).trim();
        if (user.role !== 'Super Admin' && ownerUser.toLowerCase() !== user.username.toLowerCase()) {
          return { status: 'error', message: 'Akses ditolak: Anda hanya dapat menghapus laporan yang Anda input sendiri.' };
        }
        sh.deleteRow(i + 1);
        writeLog(user.username, user.role, 'HAPUS_BERITA', 'Menghapus laporan ID ' + targetId + ' (' + judul + ')');
        return {
          status: 'success',
          message: 'Laporan berita berhasil dihapus.',
          berita: getAllBeritaRecords()
        };
      }
    }
    return { status: 'error', message: 'ID Berita tidak ditemukan.' };
  } catch (err) {
    return { status: 'error', message: 'Gagal menghapus berita: ' + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * FILTER LAPORAN MULTI-DIMENSI & GENERATOR PDF RESMI BER-KOP LPP RRI
 * ============================================================================
 */
function getFilteredReportServer(token, filterParams) {
  const user = verifySessionToken(token);
  if (!user) return { status: 'unauthorized', message: 'Sesi tidak valid.' };

  const allBerita = getAllBeritaRecords();
  const filtered = filterBeritaHelper(allBerita, filterParams);
  return {
    status: 'success',
    data: filtered
  };
}

function filterBeritaHelper(allBerita, filterParams) {
  const mode = filterParams.mode || 'BULANAN';
  const tanggalHarian = filterParams.tanggalHarian || '';
  const bulan = parseInt(filterParams.bulan, 10);
  const tahun = parseInt(filterParams.tahun, 10);
  const tanggalAwal = filterParams.tanggalAwal || '';
  const tanggalAkhir = filterParams.tanggalAkhir || '';
  const reporter = filterParams.reporter || 'SEMUA';
  const idFormat = filterParams.idFormat || 'SEMUA';
  const keyword = String(filterParams.keyword || '').trim().toLowerCase();

  return allBerita.filter(function(item) {
    const tgl = item.tanggalSiar;
    if (!tgl) return false;

    if (mode === 'HARIAN' && tanggalHarian) {
      if (tgl !== tanggalHarian) return false;
    } else if (mode === 'BULANAN' && !isNaN(bulan) && !isNaN(tahun)) {
      const parts = tgl.split('-');
      const y = parseInt(parts[0], 10);
      const m = parseInt(parts[1], 10);
      if (y !== tahun || m !== bulan) return false;
    } else if (mode === 'INTERVAL') {
      if (tanggalAwal && tgl < tanggalAwal) return false;
      if (tanggalAkhir && tgl > tanggalAkhir) return false;
    }

    if (reporter && reporter !== 'SEMUA') {
      if (item.namaReporter !== reporter) return false;
    }

    if (idFormat && idFormat !== 'SEMUA') {
      if (item.idFormat !== idFormat && item.namaFormat !== idFormat) return false;
    }

    if (keyword) {
      const hay = (item.judulBeritaTopik + ' ' + item.narasumberJabatan + ' ' + item.keteranganTambahan + ' ' + item.jenisLaporan).toLowerCase();
      if (!hay.includes(keyword)) return false;
    }

    return true;
  });
}

function generateOfficialReportPDF(token, filterParams) {
  const user = verifySessionToken(token);
  if (!user) return { status: 'unauthorized', message: 'Sesi Anda telah habis. Silakan login kembali.' };

  try {
    const config = getPublicConfig();
    const formats = getMasterFormatsList();
    const allBerita = getAllBeritaRecords();
    const filtered = filterBeritaHelper(allBerita, filterParams);

    const sorted = filtered.slice().sort(function(a, b) {
      return a.tanggalSiar.localeCompare(b.tanggalSiar);
    });

    const namaBulanArr = [
      '', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
      'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
    ];
    let labelPeriode = '';
    if (filterParams.mode === 'HARIAN') {
      labelPeriode = 'Tanggal: ' + (filterParams.tanggalHarian || '-');
    } else if (filterParams.mode === 'BULANAN') {
      const bIdx = parseInt(filterParams.bulan, 10) || 1;
      labelPeriode = 'Bulan ' + namaBulanArr[bIdx] + ' Tahun ' + (filterParams.tahun || '');
    } else if (filterParams.mode === 'INTERVAL') {
      labelPeriode = 'Periode: ' + (filterParams.tanggalAwal || '-') + ' s/d ' + (filterParams.tanggalAkhir || '-');
    }

    // Hitung Rekap per Format Acara
    const formatCounts = {};
    formats.forEach(function(f) {
      formatCounts[f.namaFormat] = 0;
    });
    sorted.forEach(function(b) {
      if (formatCounts[b.namaFormat] === undefined) {
        formatCounts[b.namaFormat] = 0;
      }
      formatCounts[b.namaFormat]++;
    });

    let rekapRowsHtml = '';
    let noRekap = 1;
    Object.keys(formatCounts).forEach(function(fmtName) {
      rekapRowsHtml += '<tr>' +
        '<td style="text-align:center;">' + (noRekap++) + '</td>' +
        '<td>' + escapeHtmlServer(fmtName) + '</td>' +
        '<td style="text-align:center; font-weight:bold;">' + formatCounts[fmtName] + '</td>' +
      '</tr>';
    });

    // Hitung Rekap Produktivitas per Reporter
    const repCounts = {};
    sorted.forEach(function(b) {
      if (!repCounts[b.namaReporter]) repCounts[b.namaReporter] = 0;
      repCounts[b.namaReporter]++;
    });
    let repRowsHtml = '';
    let noRep = 1;
    Object.keys(repCounts).sort().forEach(function(repName) {
      repRowsHtml += '<tr>' +
        '<td style="text-align:center;">' + (noRep++) + '</td>' +
        '<td>' + escapeHtmlServer(repName) + '</td>' +
        '<td style="text-align:center; font-weight:bold;">' + repCounts[repName] + '</td>' +
      '</tr>';
    });
    if (!repRowsHtml) {
      repRowsHtml = '<tr><td colspan="3" style="text-align:center;">Tidak ada data</td></tr>';
    }

    // Baris Tabel Rincian Berita
    let detailRowsHtml = '';
    if (sorted.length === 0) {
      detailRowsHtml = '<tr><td colspan="7" style="text-align:center; padding:14px;">Tidak ada data laporan berita pada filter yang dipilih.</td></tr>';
    } else {
      sorted.forEach(function(item, idx) {
        detailRowsHtml += '<tr>' +
          '<td style="text-align:center;">' + (idx + 1) + '</td>' +
          '<td style="text-align:center; white-space:nowrap;">' + escapeHtmlServer(item.tanggalSiar) + '</td>' +
          '<td>' + escapeHtmlServer(item.namaFormat) + '</td>' +
          '<td>' + escapeHtmlServer(item.namaReporter) + '</td>' +
          '<td><strong>' + escapeHtmlServer(item.judulBeritaTopik) + '</strong>' +
            (item.keteranganTambahan && item.keteranganTambahan !== '-' ? '<br><span style="font-size:9px;color:#555;">Ket: ' + escapeHtmlServer(item.keteranganTambahan) + '</span>' : '') +
          '</td>' +
          '<td>' + escapeHtmlServer(item.narasumberJabatan) + '</td>' +
          '<td style="text-align:center;">' + escapeHtmlServer(item.jenisLaporan || 'Umum') + '</td>' +
        '</tr>';
      });
    }

    const nowDate = new Date();
    const tglCetak = (config.Kota_Stasiun || 'Jambi') + ', ' + nowDate.getDate() + ' ' + namaBulanArr[nowDate.getMonth() + 1] + ' ' + nowDate.getFullYear();

    const htmlContent = '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>' +
      '@page { size: A4 landscape; margin: 15mm; }' +
      'body { font-family: Arial, Helvetica, sans-serif; color: #111; font-size: 11px; line-height: 1.4; }' +
      '.kop-container { text-align: center; border-bottom: 3px double #002B66; padding-bottom: 10px; margin-bottom: 16px; }' +
      '.kop-instansi { font-size: 18px; font-weight: bold; color: #002B66; letter-spacing: 0.5px; margin: 0; text-transform: uppercase; }' +
      '.kop-judul { font-size: 14px; font-weight: bold; margin: 4px 0 2px 0; text-transform: uppercase; }' +
      '.kop-sub { font-size: 11px; color: #333; margin: 0; }' +
      '.meta-table { width: 100%; margin-bottom: 12px; border: none; }' +
      '.meta-table td { border: none; padding: 2px 4px; font-size: 11px; }' +
      '.grid-summary { width: 100%; border: none; margin-bottom: 16px; }' +
      '.grid-summary > tbody > tr > td { border: none; vertical-align: top; padding: 0; }' +
      'table.data-table { width: 100%; border-collapse: collapse; margin-bottom: 14px; }' +
      'table.data-table th { background-color: #002B66; color: #ffffff; font-weight: bold; text-align: center; padding: 6px; border: 1px solid #001A40; font-size: 10px; text-transform: uppercase; }' +
      'table.data-table td { padding: 5px 6px; border: 1px solid #666; font-size: 10px; vertical-align: top; }' +
      '.section-title { font-size: 12px; font-weight: bold; color: #002B66; margin: 10px 0 6px 0; text-transform: uppercase; }' +
      '.ttd-wrapper { width: 100%; margin-top: 25px; page-break-inside: avoid; }' +
      '.ttd-box { float: right; width: 280px; text-align: center; font-size: 11px; }' +
      '</style></head><body>' +
      '<div class="kop-container">' +
        '<p class="kop-instansi">LEMBAGA PENYIARAN PUBLIK RADIO REPUBLIK INDONESIA (' + escapeHtmlServer(config.Nama_Stasiun) + ')</p>' +
        '<p class="kop-judul">LAPORAN REKAPITULASI & RINCIAN PEMBERITAAN / PROGRAM ACARA (SiPeKA)</p>' +
        '<p class="kop-sub">' + escapeHtmlServer(config.Sub_Judul_Aplikasi) + '</p>' +
      '</div>' +
      '<table class="meta-table">' +
        '<tr>' +
          '<td style="width:15%;"><strong>Periode Laporan</strong></td><td style="width:35%;">: ' + escapeHtmlServer(labelPeriode) + '</td>' +
          '<td style="width:15%;"><strong>Filter Format Acara</strong></td><td style="width:35%;">: ' + escapeHtmlServer(filterParams.namaFormatLabel || 'Semua Format Acara') + '</td>' +
        '</tr>' +
        '<tr>' +
          '<td><strong>Filter Reporter</strong></td><td>: ' + escapeHtmlServer(filterParams.reporter === 'SEMUA' ? 'Semua Reporter' : filterParams.reporter) + '</td>' +
          '<td><strong>Total Berita/Program</strong></td><td>: <strong>' + sorted.length + ' Laporan</strong></td>' +
        '</tr>' +
      '</table>' +
      '<table class="grid-summary"><tbody><tr>' +
        '<td style="width:48%; padding-right:10px;">' +
          '<div class="section-title">A. Rekapitulasi Per Format Acara</div>' +
          '<table class="data-table"><thead><tr><th style="width:40px;">No</th><th>Nama Format Acara</th><th style="width:90px;">Jumlah</th></tr></thead><tbody>' +
            rekapRowsHtml +
            '<tr style="background:#f0f4fa;font-weight:bold;"><td colspan="2" style="text-align:right;">TOTAL KESELURUHAN</td><td style="text-align:center;">' + sorted.length + '</td></tr>' +
          '</tbody></table>' +
        '</td>' +
        '<td style="width:48%; padding-left:10px;">' +
          '<div class="section-title">B. Rekapitulasi Produktivitas Reporter</div>' +
          '<table class="data-table"><thead><tr><th style="width:40px;">No</th><th>Nama Reporter / Broadcaster</th><th style="width:90px;">Jumlah</th></tr></thead><tbody>' +
            repRowsHtml +
            '<tr style="background:#f0f4fa;font-weight:bold;"><td colspan="2" style="text-align:right;">TOTAL PRODUKTIVITAS</td><td style="text-align:center;">' + sorted.length + '</td></tr>' +
          '</tbody></table>' +
        '</td>' +
      '</tr></tbody></table>' +
      '<div class="section-title">C. Tabel Rincian Laporan Berita & Program</div>' +
      '<table class="data-table">' +
        '<thead><tr>' +
          '<th style="width:30px;">No</th>' +
          '<th style="width:75px;">Tanggal Siar</th>' +
          '<th style="width:110px;">Format Acara</th>' +
          '<th style="width:130px;">Nama Reporter</th>' +
          '<th>Judul Berita / Topik Siaran</th>' +
          '<th style="width:160px;">Narasumber & Jabatan</th>' +
          '<th style="width:95px;">Kategori</th>' +
        '</tr></thead>' +
        '<tbody>' + detailRowsHtml + '</tbody>' +
      '</table>' +
      '<div class="ttd-wrapper">' +
        '<div class="ttd-box">' +
          '<p style="margin-bottom:4px;">' + escapeHtmlServer(tglCetak) + '</p>' +
          '<p style="margin-top:0; font-weight:bold;">Mengetahui,<br>Pejabat Berwenang / Ketua Tim Pemberitaan</p>' +
          '<br><br><br><br>' +
          '<p style="margin-bottom:2px; font-weight:bold; text-decoration:underline;">' + escapeHtmlServer(config.Nama_Pejabat_Tanda_Tangan) + '</p>' +
          '<p style="margin-top:0;">NIP. ' + escapeHtmlServer(config.NIP_Pejabat) + '</p>' +
        '</div>' +
      '</div>' +
      '</body></html>';

    const pdfBlob = HtmlService.createHtmlOutput(htmlContent).getAs('application/pdf');
    const base64Pdf = Utilities.base64Encode(pdfBlob.getBytes());
    const fileName = 'Laporan_SiPeKA_Berita_' + String(config.Nama_Stasiun).replace(/\s+/g, '_') + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyyMMdd_HHmm') + '.pdf';

    writeLog(user.username, user.role, 'CETAK_PDF', 'Mengunduh laporan PDF (' + labelPeriode + ', ' + sorted.length + ' baris)');

    return {
      status: 'success',
      fileName: fileName,
      base64: base64Pdf
    };
  } catch (err) {
    return { status: 'error', message: 'Gagal membuat PDF: ' + err.message };
  }
}

function escapeHtmlServer(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * ============================================================================
 * CONTROL PANEL SUPER ADMIN: 1. MANAJEMEN FORMAT ACARA (CRUD)
 * ============================================================================
 */
function saveMasterFormat(token, payload) {
  const user = verifySessionToken(token);
  if (!user || user.role !== 'Super Admin') {
    return { status: 'error', message: 'Akses ditolak: Hanya Super Admin yang diizinkan.' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('MASTER_FORMAT_ACARA');
    const data = sh.getDataRange().getValues();

    const idFormat = String(payload.idFormat || '').trim();
    const namaFormat = String(payload.namaFormat || '').trim();
    const kategoriForm = String(payload.kategoriForm || 'STANDAR').trim();
    const ikonFA = String(payload.ikonFA || 'fa-newspaper').trim();
    const statusAktif = String(payload.statusAktif || 'Aktif').trim();

    if (!namaFormat) return { status: 'error', message: 'Nama Format Acara wajib diisi.' };

    if (idFormat) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === idFormat) {
          sh.getRange(i + 1, 2, 1, 4).setValues([[namaFormat, kategoriForm, ikonFA, statusAktif]]);
          writeLog(user.username, user.role, 'EDIT_FORMAT', 'Memperbarui format acara: ' + namaFormat);
          return {
            status: 'success',
            message: 'Format acara berhasil diperbarui.',
            formats: getMasterFormatsList()
          };
        }
      }
    }

    const newId = 'FMT-' + Math.floor(100 + Math.random() * 900);
    sh.appendRow([newId, namaFormat, kategoriForm, ikonFA, statusAktif]);
    writeLog(user.username, user.role, 'TAMBAH_FORMAT', 'Menambahkan format acara baru: ' + namaFormat);
    return {
      status: 'success',
      message: 'Format acara baru berhasil ditambahkan.',
      formats: getMasterFormatsList()
    };
  } catch (err) {
    return { status: 'error', message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteMasterFormat(token, idFormat) {
  const user = verifySessionToken(token);
  if (!user || user.role !== 'Super Admin') {
    return { status: 'error', message: 'Akses ditolak: Hanya Super Admin yang diizinkan.' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('MASTER_FORMAT_ACARA');
    const data = sh.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(idFormat).trim()) {
        const nama = data[i][1];
        sh.deleteRow(i + 1);
        writeLog(user.username, user.role, 'HAPUS_FORMAT', 'Menghapus format acara: ' + nama);
        return {
          status: 'success',
          message: 'Format acara berhasil dihapus.',
          formats: getMasterFormatsList()
        };
      }
    }
    return { status: 'error', message: 'Format acara tidak ditemukan.' };
  } catch (err) {
    return { status: 'error', message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * CONTROL PANEL SUPER ADMIN: 2. MANAJEMEN DATA REPORTER (CRUD)
 * ============================================================================
 */
function saveMasterReporter(token, payload) {
  const user = verifySessionToken(token);
  if (!user || user.role !== 'Super Admin') {
    return { status: 'error', message: 'Akses ditolak: Hanya Super Admin yang diizinkan.' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('MASTER_REPORTER');
    const data = sh.getDataRange().getValues();

    const idReporter = String(payload.idReporter || '').trim();
    const namaReporter = String(payload.namaReporter || '').trim();
    const nipId = String(payload.nipId || '-').trim() || '-';
    const jabatan = String(payload.jabatan || 'Reporter').trim();
    const statusAktif = String(payload.statusAktif || 'Aktif').trim();

    if (!namaReporter) return { status: 'error', message: 'Nama Reporter wajib diisi.' };

    for (let i = 1; i < data.length; i++) {
      const existingId = String(data[i][0]).trim();
      const existingName = String(data[i][1]).trim().toLowerCase();
      if (existingName === namaReporter.toLowerCase() && existingId !== idReporter) {
        return { status: 'error', message: 'Nama reporter tersebut sudah terdaftar di dalam database.' };
      }
    }

    if (idReporter) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === idReporter) {
          sh.getRange(i + 1, 2, 1, 4).setValues([[namaReporter, nipId, jabatan, statusAktif]]);
          writeLog(user.username, user.role, 'EDIT_REPORTER', 'Memperbarui data reporter: ' + namaReporter);
          return {
            status: 'success',
            message: 'Data reporter berhasil diperbarui.',
            reporters: getMasterReportersList()
          };
        }
      }
    }

    const newId = 'REP-' + Math.floor(1000 + Math.random() * 9000);
    sh.appendRow([newId, namaReporter, nipId, jabatan, statusAktif]);
    writeLog(user.username, user.role, 'TAMBAH_REPORTER', 'Menambahkan reporter baru: ' + namaReporter);
    return {
      status: 'success',
      message: 'Reporter baru berhasil ditambahkan.',
      reporters: getMasterReportersList()
    };
  } catch (err) {
    return { status: 'error', message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteMasterReporter(token, idReporter) {
  const user = verifySessionToken(token);
  if (!user || user.role !== 'Super Admin') {
    return { status: 'error', message: 'Akses ditolak: Hanya Super Admin yang diizinkan.' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('MASTER_REPORTER');
    const data = sh.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(idReporter).trim()) {
        const nama = data[i][1];
        sh.deleteRow(i + 1);
        writeLog(user.username, user.role, 'HAPUS_REPORTER', 'Menghapus reporter: ' + nama);
        return {
          status: 'success',
          message: 'Data reporter berhasil dihapus.',
          reporters: getMasterReportersList()
        };
      }
    }
    return { status: 'error', message: 'Data reporter tidak ditemukan.' };
  } catch (err) {
    return { status: 'error', message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * CONTROL PANEL SUPER ADMIN: 3. MANAJEMEN AKUN PENGGUNA (CRUD & PROTEKSI)
 * ============================================================================
 */
function saveMasterUser(token, payload) {
  const user = verifySessionToken(token);
  if (!user || user.role !== 'Super Admin') {
    return { status: 'error', message: 'Akses ditolak: Hanya Super Admin yang diizinkan.' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('MASTER_USER');
    const data = sh.getDataRange().getValues();

    const mode = payload.mode || 'ADD';
    const username = String(payload.username || '').trim();
    const password = String(payload.password || '');
    const namaLengkap = String(payload.namaLengkap || '').trim();
    let role = String(payload.role || 'User').trim();
    let statusAkun = String(payload.statusAkun || 'Aktif').trim();

    if (!username || !namaLengkap) {
      return { status: 'error', message: 'Username dan Nama Lengkap wajib diisi.' };
    }

    if (username.toLowerCase() === 'superadmin') {
      role = 'Super Admin';
      statusAkun = 'Aktif';
    }

    if (mode === 'EDIT') {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim().toLowerCase() === username.toLowerCase()) {
          sh.getRange(i + 1, 3).setValue(namaLengkap);
          sh.getRange(i + 1, 4).setValue(role);
          sh.getRange(i + 1, 5).setValue(statusAkun);
          if (password) {
            sh.getRange(i + 1, 2).setValue(hashSHA256(password));
          }
          writeLog(user.username, user.role, 'EDIT_USER', 'Memperbarui akun pengguna: ' + username);
          return {
            status: 'success',
            message: 'Akun pengguna berhasil diperbarui.',
            users: getMasterUsersList()
          };
        }
      }
      return { status: 'error', message: 'Username tidak ditemukan.' };
    } else {
      if (!password) {
        return { status: 'error', message: 'Password wajib diisi untuk akun baru.' };
      }
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim().toLowerCase() === username.toLowerCase()) {
          return { status: 'error', message: 'Username sudah terdaftar. Gunakan username lain.' };
        }
      }
      sh.appendRow([username, hashSHA256(password), namaLengkap, role, statusAkun]);
      writeLog(user.username, user.role, 'TAMBAH_USER', 'Menambahkan akun baru: ' + username + ' (' + role + ')');
      return {
        status: 'success',
        message: 'Akun pengguna baru berhasil ditambahkan.',
        users: getMasterUsersList()
      };
    }
  } catch (err) {
    return { status: 'error', message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteMasterUser(token, targetUsername) {
  const user = verifySessionToken(token);
  if (!user || user.role !== 'Super Admin') {
    return { status: 'error', message: 'Akses ditolak: Hanya Super Admin yang diizinkan.' };
  }

  const cleanTarget = String(targetUsername || '').trim();
  if (cleanTarget.toLowerCase() === 'superadmin') {
    return { status: 'error', message: 'Akun utama "superadmin" dilindungi oleh sistem dan tidak dapat dihapus!' };
  }
  if (cleanTarget.toLowerCase() === user.username.toLowerCase()) {
    return { status: 'error', message: 'Anda tidak dapat menghapus akun Anda sendiri yang sedang aktif login.' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('MASTER_USER');
    const data = sh.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === cleanTarget.toLowerCase()) {
        sh.deleteRow(i + 1);
        writeLog(user.username, user.role, 'HAPUS_USER', 'Menghapus akun pengguna: ' + cleanTarget);
        return {
          status: 'success',
          message: 'Akun pengguna berhasil dihapus.',
          users: getMasterUsersList()
        };
      }
    }
    return { status: 'error', message: 'Akun tidak ditemukan.' };
  } catch (err) {
    return { status: 'error', message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function changeOwnPassword(token, oldPass, newPass) {
  const user = verifySessionToken(token);
  if (!user) return { status: 'unauthorized', message: 'Sesi telah berakhir.' };

  if (!oldPass || !newPass || String(newPass).length < 4) {
    return { status: 'error', message: 'Password baru minimal 4 karakter.' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('MASTER_USER');
    const data = sh.getDataRange().getValues();
    const oldHash = hashSHA256(oldPass);

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toLowerCase() === user.username.toLowerCase()) {
        const dbHash = String(data[i][1]).trim();
        if (dbHash !== oldHash && dbHash !== oldPass) {
          return { status: 'error', message: 'Password lama yang Anda masukkan tidak sesuai.' };
        }
        sh.getRange(i + 1, 2).setValue(hashSHA256(newPass));
        writeLog(user.username, user.role, 'UBAH_PASSWORD', 'Mengubah kata sandi pribadi');
        return { status: 'success', message: 'Kata sandi Anda berhasil diperbarui!' };
      }
    }
    return { status: 'error', message: 'Akun tidak ditemukan.' };
  } catch (err) {
    return { status: 'error', message: err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * CONTROL PANEL SUPER ADMIN: 4. PENGATURAN SISTEM & LOG AKTIVITAS
 * ============================================================================
 */
function saveSystemSettings(token, newSettings) {
  const user = verifySessionToken(token);
  if (!user || user.role !== 'Super Admin') {
    return { status: 'error', message: 'Akses ditolak: Hanya Super Admin yang diizinkan.' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('PENGATURAN_SISTEM');
    sh.clearContents();
    sh.appendRow(['Kunci_Konfigurasi', 'Nilai_Konfigurasi']);
    styleHeaderRow(sh, 2);

    const rows = [
      ['Nama_Stasiun', String(newSettings.Nama_Stasiun || 'LPP RRI Jambi').trim()],
      ['Sub_Judul_Aplikasi', String(newSettings.Sub_Judul_Aplikasi || 'Sistem Informasi Pelaporan Kinerja dan Administrasi').trim()],
      ['Nama_Pejabat_Tanda_Tangan', String(newSettings.Nama_Pejabat_Tanda_Tangan || 'Ketua Tim Pemberitaan').trim()],
      ['NIP_Pejabat', String(newSettings.NIP_Pejabat || '-').trim()],
      ['Kota_Stasiun', String(newSettings.Kota_Stasiun || 'Jambi').trim()],
      ['Status_Maintenance', String(newSettings.Status_Maintenance || 'TIDAK').trim()]
    ];
    sh.getRange(2, 1, rows.length, 2).setValues(rows);

    writeLog(user.username, user.role, 'UPDATE_KONFIGURASI', 'Memperbarui Pengaturan Instansi (' + rows[0][1] + ')');

    return {
      status: 'success',
      message: 'Pengaturan sistem dan instansi berhasil disimpan!',
      config: getPublicConfig(),
      logs: getRecentActivityLogs(200)
    };
  } catch (err) {
    return { status: 'error', message: err.message };
  } finally {
    lock.releaseLock();
  }
}

function clearActivityLogs(token) {
  const user = verifySessionToken(token);
  if (!user || user.role !== 'Super Admin') {
    return { status: 'error', message: 'Akses ditolak.' };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName('LOG_AKTIVITAS');
    sh.clearContents();
    sh.appendRow(['Timestamp', 'Username', 'Role', 'Aksi', 'Detail_Aktivitas']);
    styleHeaderRow(sh, 5);
    writeLog(user.username, user.role, 'BERSIHKAN_LOG', 'Membersihkan riwayat log aktivitas');
    return {
      status: 'success',
      message: 'Log aktivitas berhasil dibersihkan.',
      logs: getRecentActivityLogs(200)
    };
  } catch (err) {
    return { status: 'error', message: err.message };
  } finally {
    lock.releaseLock();
  }
}