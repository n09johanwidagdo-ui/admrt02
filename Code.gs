/**
 * ============================================================
 *  APLIKASI ADMINISTRASI RT 02 — BACKEND API (JSON)
 * ============================================================
 *  Mode: API-only. Tidak lagi serve HTML — dipanggil dari
 *  frontend statis (GitHub Pages/Netlify/Vercel) via fetch().
 *  Database tetap 100% di Google Sheets aktif project ini.
 *
 *  CARA KERJA:
 *  Semua request (baca maupun tulis) dikirim sebagai POST
 *  dengan body JSON berformat:
 *    { "action": "namaAction", "token": "...", "payload": {...} }
 *  Header Content-Type WAJIB "text/plain;charset=utf-8" (bukan
 *  application/json) agar browser tidak mengirim CORS preflight
 *  (OPTIONS) — karena Apps Script Web App tidak menangani method
 *  OPTIONS. Dengan text/plain, request dianggap "simple request"
 *  sehingga langsung diproses doPost() dan responsnya otomatis
 *  disertai header Access-Control-Allow-Origin oleh infrastruktur
 *  Apps Script (khusus deployment akses "Anyone").
 *
 *  doGet() disediakan sebagai fallback untuk testing manual via
 *  URL browser: ?action=getWargaList dsb.
 * ============================================================
 */

const IURAN_DEFAULT = 30000;
const DEFAULT_ADMIN_PASSWORD = 'nDalem12345';
const TOKEN_DURATION_SECONDS = 21600; // 6 jam (maksimum CacheService)

const SHEET_NAMES = {
  WARGA: 'Warga',
  KAS: 'Kas',
  ARUSKAS: 'ArusKas',
  PERSURATAN: 'Persuratan'
};

const HEADERS = {
  Warga: ['ID', 'Nama', 'NomorBlok', 'KeteranganHunian'],
  Kas: ['ID', 'Bulan', 'Tahun', 'NomorBlok', 'NamaWarga', 'Status', 'Jumlah'],
  ArusKas: ['ID', 'Tanggal', 'Jenis', 'Keterangan', 'Jumlah'],
  Persuratan: ['ID', 'Tanggal', 'NomorSurat', 'Perihal', 'Kepada', 'LinkArsip']
};

const BULAN_LIST = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];

/* ============================================================
 *  ENTRY POINT WEB APP (API)
 * ============================================================ */
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};
  let payload = {};
  if (params.payload) {
    try { payload = JSON.parse(params.payload); } catch (err) { payload = {}; }
  }
  return handleRequest_(params.action, params.token, payload);
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    body = e.parameter || {};
  }
  return handleRequest_(body.action, body.token, body.payload || {});
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Router utama: satu pintu untuk semua aksi CRUD. */
function handleRequest_(action, token, payload) {
  setupSheets();
  try {
    let result;
    switch (action) {
      // ---- AUTH ----
      case 'loginAdmin': result = loginAdmin(payload.password); break;
      case 'logoutAdmin': result = logoutAdmin(token); break;
      case 'verifySession': result = verifySession(token); break;

      // ---- DASHBOARD ----
      case 'getDashboardData': result = getDashboardData(); break;
      case 'getBulanList': result = getBulanList(); break;

      // ---- WARGA ----
      case 'getWargaList': result = getWargaList(); break;
      case 'addWarga': result = addWarga(token, payload); break;
      case 'updateWarga': result = updateWarga(token, payload); break;
      case 'deleteWarga': result = deleteWarga(token, payload.id); break;

      // ---- KAS ----
      case 'getKasList': result = getKasList(payload); break;
      case 'generateKasBulanan': result = generateKasBulanan(token, payload.bulan, payload.tahun); break;
      case 'updateStatusKas': result = updateStatusKas(token, payload.id, payload.status); break;
      case 'deleteKas': result = deleteKas(token, payload.id); break;
      case 'getKasPerWarga': result = getKasPerWarga(payload.nomorBlok); break;

      // ---- ARUS KAS ----
      case 'getArusKasList': result = getArusKasList(); break;
      case 'addArusKas': result = addArusKas(token, payload); break;
      case 'deleteArusKas': result = deleteArusKas(token, payload.id); break;
      case 'getSaldoKas': result = getSaldoKas(); break;

      // ---- PERSURATAN ----
      case 'getPersuratanList': result = getPersuratanList(); break;
      case 'addPersuratan': result = addPersuratan(token, payload); break;
      case 'deletePersuratan': result = deletePersuratan(token, payload.id); break;

      default:
        result = { success: false, message: 'Action "' + action + '" tidak dikenal.' };
    }
    return jsonOutput_(result);
  } catch (err) {
    return jsonOutput_({ success: false, message: err.message });
  }
}

/* ============================================================
 *  AUTENTIKASI ADMIN (token + CacheService)
 * ============================================================ */
function setAdminPassword(newPassword) {
  PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', newPassword);
  return { success: true, message: 'Password admin berhasil diubah.' };
}

function getAdminPassword_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || DEFAULT_ADMIN_PASSWORD;
}

function loginAdmin(password) {
  if (password === getAdminPassword_()) {
    const token = Utilities.getUuid();
    CacheService.getScriptCache().put('session_' + token, 'admin', TOKEN_DURATION_SECONDS);
    return { success: true, token: token };
  }
  return { success: false, message: 'Password salah. Silakan coba lagi.' };
}

function logoutAdmin(token) {
  if (token) CacheService.getScriptCache().remove('session_' + token);
  return { success: true };
}

function verifySession(token) {
  return { isAdmin: isAdmin_(token) };
}

function isAdmin_(token) {
  if (!token) return false;
  const role = CacheService.getScriptCache().get('session_' + token);
  return role === 'admin';
}

function requireAdmin_(token) {
  if (!isAdmin_(token)) {
    throw new Error('Akses ditolak. Anda harus login sebagai Admin untuk melakukan aksi ini.');
  }
}

/* ============================================================
 *  SETUP SPREADSHEET (auto create sheets & headers)
 * ============================================================ */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(HEADERS).forEach(function (sheetName) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS[sheetName]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADERS[sheetName].length)
        .setFontWeight('bold').setBackground('#0d6efd').setFontColor('#ffffff');
      sheet.autoResizeColumns(1, HEADERS[sheetName].length);
    }
  });

  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);
  return true;
}

/* ============================================================
 *  UTILITAS UMUM
 * ============================================================ */
function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) { setupSheets(); sheet = ss.getSheetByName(name); }
  return sheet;
}

function sheetToObjects_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .map(function (row, idx) {
      const obj = {};
      headers.forEach(function (h, i) { obj[h] = row[i]; });
      obj._row = idx + 2;
      return obj;
    })
    .filter(function (obj) { return obj.ID !== '' && obj.ID !== null && obj.ID !== undefined; });
}

function generateId_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(function (v) { return v !== '' && v !== null; });
  if (ids.length === 0) return 1;
  return Math.max.apply(null, ids.map(Number)) + 1;
}

function findRowById_(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

/* ============================================================
 *  MODUL: WARGA
 * ============================================================ */
function getWargaList() {
  return sheetToObjects_(getSheet_(SHEET_NAMES.WARGA));
}

function addWarga(token, data) {
  requireAdmin_(token);
  const sheet = getSheet_(SHEET_NAMES.WARGA);
  const id = generateId_(sheet);
  sheet.appendRow([id, data.Nama, data.NomorBlok, data.KeteranganHunian]);
  return { success: true, id: id };
}

function updateWarga(token, data) {
  requireAdmin_(token);
  const sheet = getSheet_(SHEET_NAMES.WARGA);
  const row = findRowById_(sheet, data.ID);
  if (row === -1) return { success: false, message: 'Data warga tidak ditemukan' };
  sheet.getRange(row, 2, 1, 3).setValues([[data.Nama, data.NomorBlok, data.KeteranganHunian]]);
  return { success: true };
}

function deleteWarga(token, id) {
  requireAdmin_(token);
  const sheet = getSheet_(SHEET_NAMES.WARGA);
  const row = findRowById_(sheet, id);
  if (row === -1) return { success: false, message: 'Data warga tidak ditemukan' };
  sheet.deleteRow(row);
  return { success: true };
}

/* ============================================================
 *  MODUL: KAS (IURAN BULANAN)
 * ============================================================ */
function getKasList(filter) {
  let data = sheetToObjects_(getSheet_(SHEET_NAMES.KAS));
  if (filter) {
    if (filter.Bulan) data = data.filter(function (d) { return String(d.Bulan) === String(filter.Bulan); });
    if (filter.Tahun) data = data.filter(function (d) { return String(d.Tahun) === String(filter.Tahun); });
  }
  return data;
}

function generateKasBulanan(token, bulan, tahun) {
  requireAdmin_(token);
  const kasSheet = getSheet_(SHEET_NAMES.KAS);
  const wargaList = getWargaList();
  const existingBlok = sheetToObjects_(kasSheet)
    .filter(function (k) { return String(k.Bulan) === String(bulan) && String(k.Tahun) === String(tahun); })
    .map(function (k) { return String(k.NomorBlok); });

  let idCounter = generateId_(kasSheet);
  const rowsToAdd = [];
  wargaList.forEach(function (w) {
    if (existingBlok.indexOf(String(w.NomorBlok)) === -1) {
      rowsToAdd.push([idCounter, bulan, tahun, w.NomorBlok, w.Nama, 'Belum', IURAN_DEFAULT]);
      idCounter++;
    }
  });

  if (rowsToAdd.length > 0) {
    kasSheet.getRange(kasSheet.getLastRow() + 1, 1, rowsToAdd.length, 7).setValues(rowsToAdd);
  }
  return { success: true, generated: rowsToAdd.length, skipped: wargaList.length - rowsToAdd.length };
}

function updateStatusKas(token, id, status) {
  requireAdmin_(token);
  const sheet = getSheet_(SHEET_NAMES.KAS);
  const row = findRowById_(sheet, id);
  if (row === -1) return { success: false, message: 'Data kas tidak ditemukan' };

  const kasData = sheet.getRange(row, 1, 1, 7).getValues()[0];
  sheet.getRange(row, 6).setValue(status);

  if (status === 'Lunas') {
    const keterangan = 'Iuran Kas ' + kasData[1] + ' ' + kasData[2] + ' - Blok ' + kasData[3] + ' (' + kasData[4] + ')';
    addArusKasOtomatis_('Masuk', keterangan, kasData[6] || IURAN_DEFAULT);
  }
  return { success: true };
}

function deleteKas(token, id) {
  requireAdmin_(token);
  const sheet = getSheet_(SHEET_NAMES.KAS);
  const row = findRowById_(sheet, id);
  if (row === -1) return { success: false, message: 'Data kas tidak ditemukan' };
  sheet.deleteRow(row);
  return { success: true };
}

/**
 * Riwayat kas per anggota (semua bulan & tahun untuk satu Nomor Blok),
 * diurutkan dari yang terbaru, lengkap dengan ringkasan lunas/belum/total.
 */
function getKasPerWarga(nomorBlok) {
  const all = sheetToObjects_(getSheet_(SHEET_NAMES.KAS))
    .filter(function (k) { return String(k.NomorBlok) === String(nomorBlok); });

  all.sort(function (a, b) {
    if (Number(b.Tahun) !== Number(a.Tahun)) return Number(b.Tahun) - Number(a.Tahun);
    return BULAN_LIST.indexOf(b.Bulan) - BULAN_LIST.indexOf(a.Bulan);
  });

  const totalLunas = all.filter(function (k) { return k.Status === 'Lunas'; }).length;
  const totalBelum = all.filter(function (k) { return k.Status === 'Belum'; }).length;
  const totalTerbayar = all
    .filter(function (k) { return k.Status === 'Lunas'; })
    .reduce(function (sum, k) { return sum + (Number(k.Jumlah) || 0); }, 0);
  const totalTunggakan = all
    .filter(function (k) { return k.Status === 'Belum'; })
    .reduce(function (sum, k) { return sum + (Number(k.Jumlah) || 0); }, 0);

  return {
    riwayat: all,
    totalLunas: totalLunas,
    totalBelum: totalBelum,
    totalTerbayar: totalTerbayar,
    totalTunggakan: totalTunggakan
  };
}

/* ============================================================
 *  MODUL: ARUS KAS
 * ============================================================ */
function getArusKasList() {
  const data = sheetToObjects_(getSheet_(SHEET_NAMES.ARUSKAS));
  data.sort(function (a, b) { return new Date(b.Tanggal) - new Date(a.Tanggal); });
  return data;
}

function addArusKas(token, data) {
  requireAdmin_(token);
  const sheet = getSheet_(SHEET_NAMES.ARUSKAS);
  const id = generateId_(sheet);
  const tanggal = data.Tanggal || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sheet.appendRow([id, tanggal, data.Jenis, data.Keterangan, Number(data.Jumlah)]);
  return { success: true, id: id };
}

function addArusKasOtomatis_(jenis, keterangan, jumlah) {
  const sheet = getSheet_(SHEET_NAMES.ARUSKAS);
  const id = generateId_(sheet);
  const tanggal = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sheet.appendRow([id, tanggal, jenis, keterangan, Number(jumlah)]);
}

function deleteArusKas(token, id) {
  requireAdmin_(token);
  const sheet = getSheet_(SHEET_NAMES.ARUSKAS);
  const row = findRowById_(sheet, id);
  if (row === -1) return { success: false, message: 'Data tidak ditemukan' };
  sheet.deleteRow(row);
  return { success: true };
}

function getSaldoKas() {
  let saldo = 0;
  getArusKasList().forEach(function (item) {
    const jumlah = Number(item.Jumlah) || 0;
    saldo += item.Jenis === 'Masuk' ? jumlah : -jumlah;
  });
  return saldo;
}

/* ============================================================
 *  MODUL: PERSURATAN
 * ============================================================ */
function getPersuratanList() {
  const data = sheetToObjects_(getSheet_(SHEET_NAMES.PERSURATAN));
  data.sort(function (a, b) { return new Date(b.Tanggal) - new Date(a.Tanggal); });
  return data;
}

function addPersuratan(token, data) {
  requireAdmin_(token);
  const sheet = getSheet_(SHEET_NAMES.PERSURATAN);
  const id = generateId_(sheet);
  const tanggal = data.Tanggal || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sheet.appendRow([id, tanggal, data.NomorSurat, data.Perihal, data.Kepada, data.LinkArsip || '']);
  return { success: true, id: id };
}

function deletePersuratan(token, id) {
  requireAdmin_(token);
  const sheet = getSheet_(SHEET_NAMES.PERSURATAN);
  const row = findRowById_(sheet, id);
  if (row === -1) return { success: false, message: 'Data surat tidak ditemukan' };
  sheet.deleteRow(row);
  return { success: true };
}

/* ============================================================
 *  MODUL: DASHBOARD
 * ============================================================ */
function getDashboardData() {
  const now = new Date();
  const bulanIni = BULAN_LIST[now.getMonth()];
  const tahunIni = now.getFullYear();

  const wargaList = getWargaList();
  const kasList = getKasList({ Bulan: bulanIni, Tahun: tahunIni });
  const saldo = getSaldoKas();

  const totalLunas = kasList.filter(function (k) { return k.Status === 'Lunas'; }).length;
  const totalBelum = kasList.filter(function (k) { return k.Status === 'Belum'; }).length;
  const totalTerkumpulBulanIni = kasList
    .filter(function (k) { return k.Status === 'Lunas'; })
    .reduce(function (sum, k) { return sum + (Number(k.Jumlah) || 0); }, 0);

  return {
    totalWarga: wargaList.length,
    bulanIni: bulanIni,
    tahunIni: tahunIni,
    totalLunas: totalLunas,
    totalBelum: totalBelum,
    totalTerkumpulBulanIni: totalTerkumpulBulanIni,
    saldoKas: saldo
  };
}

function getBulanList() { return BULAN_LIST; }
