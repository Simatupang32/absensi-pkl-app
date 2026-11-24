/**
 * ABSENSI PKL - Google Apps Script (v2)
 * - Login username/password (tidak perlu akun Google)
 * - Selfie kamera wajib
 * - Banyak hotel/lokasi PKL, masing-masing punya radius sendiri
 * - Halaman ADMIN: atur hotel + radius, data siswa, akurasi GPS, rekap hari ini
 *
 * Sheet: Users, Lokasi, Absensi, Pengaturan (dibuat/diupgrade otomatis)
 */

const SH = { USERS: 'Users', LOKASI: 'Lokasi', ABSEN: 'Absensi', CONFIG: 'Pengaturan' };
const TZ = 'Asia/Jakarta';
const SESSION_HOURS = 8;
const SCHEMA_VERSION = '5';
const DEPT_DEFAULT = ['Housekeeping', 'Engineering', 'Accounting', 'Front Office', 'FB Service', 'FB Product', 'Sales'];

const HEADERS = {
  Users: ['Username', 'Password', 'Nama', 'Sekolah', 'Lokasi PKL', 'Role (ADMIN/SISWA)', 'Aktif (YA/TIDAK)',
    'No HP', 'Guru Pembimbing', 'Tanggal Mulai', 'Tanggal Selesai', 'Departemen'],
  Lokasi: ['Nama Lokasi', 'Latitude', 'Longitude', 'Radius (meter)'],
  Absensi: ['Tanggal', 'Jam', 'Username', 'Nama', 'Sekolah', 'Lokasi', 'Jenis',
    'Latitude', 'Longitude', 'Akurasi GPS (m)', 'Jarak (m)', 'Foto',
    'Status', 'Terlambat (menit)', 'Keterangan', 'Dokumen', 'Departemen'],
  Pengaturan: ['Kunci', 'Nilai']
};

/* ===================== WEB APP ===================== */

// Logo aplikasi (disimpan di GitHub)
const LOGO_URL = 'https://raw.githubusercontent.com/Simatupang32/Absensi-PKL/main/icons/favicon-32.png';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Absensi PKL')
    .setFaviconUrl(LOGO_URL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    // Izinkan dibuka di dalam aplikasi HP (GitHub Pages)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ===================== SETUP & UPGRADE ===================== */

function setup() {
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('Script harus dibuat dari Google Sheet (Extensions > Apps Script).');

  // --- Upgrade dari versi lama (jika ada) ---
  const oldUsers = ss.getSheetByName(SH.USERS);
  if (oldUsers && oldUsers.getLastColumn() < 7) {
    oldUsers.insertColumnsAfter(4, 2);
    oldUsers.getRange(1, 5, 1, 2).setValues([['Lokasi PKL', 'Role (ADMIN/SISWA)']]);
    const n = oldUsers.getLastRow() - 1;
    if (n > 0) oldUsers.getRange(2, 6, n, 1).setValue('SISWA');
  }
  const oldAbsen = ss.getSheetByName(SH.ABSEN);
  if (oldAbsen && oldAbsen.getRange(1, 6).getValue() === 'Jenis') {
    oldAbsen.insertColumnBefore(6);
    oldAbsen.getRange(1, 6).setValue('Lokasi');
  }

  // --- Pengaturan ---
  let cfg = ss.getSheetByName(SH.CONFIG);
  let oldCfg = {};
  if (!cfg) {
    cfg = makeSheet_(ss, SH.CONFIG, []);
  } else {
    oldCfg = readConfigRaw_(cfg);
  }
  if (!oldCfg.FOLDER_ID) setConfig_('FOLDER_ID', DriveApp.createFolder('Foto Absensi PKL').getId());
  if (!oldCfg.MAX_AKURASI_METER) setConfig_('MAX_AKURASI_METER', 150);
  if (!oldCfg.DOC_FOLDER_ID) setConfig_('DOC_FOLDER_ID', DriveApp.createFolder('Dokumen Sakit & Izin PKL').getId());
  if (!oldCfg.JAM_MASUK) setConfig_('JAM_MASUK', '08:00');
  if (!oldCfg.JAM_PULANG) setConfig_('JAM_PULANG', '17:00');
  if (oldCfg.TOLERANSI_MENIT === undefined || oldCfg.TOLERANSI_MENIT === '') setConfig_('TOLERANSI_MENIT', 0);
  if (!oldCfg.DEPARTEMEN) setConfig_('DEPARTEMEN', DEPT_DEFAULT.join(', '));

  // --- Lokasi (pindahkan dari pengaturan lama jika ada) ---
  if (!ss.getSheetByName(SH.LOKASI)) {
    const first = oldCfg.LATITUDE
      ? [[oldCfg.NAMA_LOKASI || 'Lokasi PKL', Number(oldCfg.LATITUDE), Number(oldCfg.LONGITUDE), Number(oldCfg.RADIUS_METER) || 100]]
      : [['Swiss-Belhotel (contoh - ganti)', -6.2, 106.816666, 100]];
    makeSheet_(ss, SH.LOKASI, first);
  }
  // hapus kunci lama yang sudah dipindah
  ['NAMA_LOKASI', 'LATITUDE', 'LONGITUDE', 'RADIUS_METER'].forEach(deleteConfig_);

  // --- Users ---
  const users = ss.getSheetByName(SH.USERS) || makeSheet_(ss, SH.USERS,
    [['siswa01', 'rahasia123', 'Contoh Siswa', 'SMK Contoh', '', 'SISWA', 'YA', '', '', '', '', '']]);
  if (users.getLastColumn() < 12) users.getRange(1, 8, 1, 5).setValues([HEADERS.Users.slice(7)]).setFontWeight('bold');
  setUserTextFormat_(users);
  const hasAdmin = users.getDataRange().getDisplayValues().slice(1)
    .some(r => String(r[5]).toUpperCase() === 'ADMIN');
  if (!hasAdmin) users.appendRow(['admin', 'admin123', 'Administrator', '', '', 'ADMIN', 'YA', '', '', '', '', '']);

  // --- Absensi ---
  const absen = ss.getSheetByName(SH.ABSEN) || makeSheet_(ss, SH.ABSEN, []);
  if (absen.getLastColumn() < 17) absen.getRange(1, 13, 1, 5).setValues([HEADERS.Absensi.slice(12)]).setFontWeight('bold');

  PropertiesService.getScriptProperties().setProperty('SCHEMA', SCHEMA_VERSION);
}

function makeSheet_(ss, name, rows) {
  const sh = ss.insertSheet(name);
  const h = HEADERS[name];
  sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, h.length).setValues(rows);
  sh.setFrozenRows(1);
  return sh;
}

function ensureReady_() {
  if (PropertiesService.getScriptProperties().getProperty('SCHEMA') !== SCHEMA_VERSION) setup();
}

function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) { setup(); sh = ss.getSheetByName(name); }
  return sh;
}

/* ===================== PENGATURAN ===================== */

function readConfigRaw_(sh) {
  const c = {};
  sh.getDataRange().getValues().slice(1).forEach(r => { if (r[0]) c[String(r[0]).trim()] = r[1]; });
  return c;
}

function setConfig_(key, value) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.CONFIG);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === key) { sh.getRange(i + 1, 2).setNumberFormat('@').setValue(String(value)); return; }
  }
  sh.appendRow([key, '']);
  sh.getRange(sh.getLastRow(), 2).setNumberFormat('@').setValue(String(value));
}

function deleteConfig_(key) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.CONFIG);
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]).trim() === key) sh.deleteRow(i + 1);
  }
}

function normTime_(v, def) {
  if (v instanceof Date) return Utilities.formatDate(v, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'HH:mm');
  const m = /^(\d{1,2})[:.](\d{2})/.exec(String(v || '').trim());
  return m ? ('0' + m[1]).slice(-2) + ':' + m[2] : def;
}

function toMenit_(hhmm) { const p = hhmm.split(':'); return Number(p[0]) * 60 + Number(p[1]); }

function getConfig_() {
  const c = readConfigRaw_(sheet_(SH.CONFIG));
  return {
    maxAkurasi: Number(c.MAX_AKURASI_METER) || 150,
    folderId: String(c.FOLDER_ID || ''),
    docFolderId: String(c.DOC_FOLDER_ID || ''),
    jamMasuk: normTime_(c.JAM_MASUK, '08:00'),
    jamPulang: normTime_(c.JAM_PULANG, '17:00'),
    toleransi: Math.max(0, Number(c.TOLERANSI_MENIT) || 0),
    departemen: String(c.DEPARTEMEN || DEPT_DEFAULT.join(',')).split(/[,\n]/).map(x => x.trim()).filter(Boolean)
  };
}

/* ===================== DATA ===================== */

function getLokasiList_() {
  const sh = sheet_(SH.LOKASI);
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()
    .map(r => ({ nama: String(r[0]).trim(), lat: Number(r[1]), lng: Number(r[2]), radius: Number(r[3]) || 100 }))
    .filter(l => l.nama && isFinite(l.lat) && isFinite(l.lng));
}

function getUsers_() {
  const sh = sheet_(SH.USERS);
  if (sh.getLastRow() < 2) return [];
  const rg = sh.getRange(2, 1, sh.getLastRow() - 1, 12);
  const val = rg.getValues();
  return rg.getDisplayValues().map((r, i) => ({
    username: r[0].trim(), password: r[1], nama: r[2], sekolah: r[3], lokasi: r[4].trim(),
    role: r[5].trim().toUpperCase() === 'ADMIN' ? 'ADMIN' : 'SISWA',
    aktif: r[6].trim().toUpperCase() === 'TIDAK' ? 'TIDAK' : 'YA',
    hp: r[7], guru: r[8], mulai: val[i][9] ? normDate_(val[i][9]) : '', selesai: val[i][10] ? normDate_(val[i][10]) : '',
    dept: r[11].trim()
  })).filter(u => u.username);
}

function setUserTextFormat_(sh) {
  ['A:B', 'H:H', 'J:K'].forEach(a => sh.getRange(a).setNumberFormat('@'));
}

function findUser_(username) {
  const u = getUsers_().find(x => x.username.toLowerCase() === String(username).toLowerCase());
  if (!u || u.aktif === 'TIDAK') throw new Error('Akun tidak ditemukan atau tidak aktif.');
  return u;
}

function distanceMeter_(lat1, lng1, lat2, lng2) {
  const R = 6371000, rad = x => x * Math.PI / 180;
  const a = Math.sin(rad(lat2 - lat1) / 2) ** 2 +
    Math.cos(rad(lat1)) * Math.cos(rad(lat2)) * Math.sin(rad(lng2 - lng1) / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(a));
}

function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

function normDate_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return String(v).trim();
}

/** Baca semua baris absensi sebagai objek (row = nomor baris di sheet) */
function readAbsen_() {
  const sh = sheet_(SH.ABSEN);
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const rg = sh.getRange(2, 1, n, 17);
  const val = rg.getValues(), disp = rg.getDisplayValues();
  return disp.map((r, i) => ({
    row: i + 2, tgl: normDate_(val[i][0]), jam: r[1], username: r[2], nama: r[3], sekolah: r[4],
    lokasi: r[5], jenis: r[6], lat: r[7], lng: r[8], akurasi: r[9], jarak: r[10], foto: r[11],
    status: r[12] || (r[6] === 'MASUK' || r[6] === 'PULANG' ? 'TEPAT WAKTU' : r[6]),
    telat: Number(r[13]) || 0, ket: r[14], dok: r[15], dept: r[16]
  }));
}

function absenHariIni_() {
  const t = today_();
  return readAbsen_().filter(r => r.tgl === t);
}

function statusHariIni_(username) {
  const st = { MASUK: null, PULANG: null, MASUK_STATUS: '', PULANG_STATUS: '', TELAT: 0, KET: null, KET_JAM: null };
  absenHariIni_().forEach(r => {
    if (r.username !== username) return;
    if (r.jenis === 'MASUK') { st.MASUK = r.jam; st.MASUK_STATUS = r.status; st.TELAT = r.telat; }
    if (r.jenis === 'PULANG') { st.PULANG = r.jam; st.PULANG_STATUS = r.status; }
    if (r.jenis === 'SAKIT' || r.jenis === 'IZIN') { st.KET = r.jenis; st.KET_JAM = r.jam; }
  });
  return st;
}

function filterAbsen_(rows, f) {
  f = f || {};
  const dari = String(f.dari || ''), sampai = String(f.sampai || '');
  const lok = String(f.lokasi || ''), q = String(f.cari || '').trim().toLowerCase();
  const st = String(f.status || ''), dept = String(f.dept || '');
  return rows.filter(r =>
    (!dari || r.tgl >= dari) && (!sampai || r.tgl <= sampai) &&
    (!lok || r.lokasi === lok) &&
    (!st || (st === 'MASUK' || st === 'PULANG' || st === 'SAKIT' || st === 'IZIN' ? r.jenis === st : r.status === st)) &&
    (!dept || r.dept === dept) &&
    (!q || (r.nama + ' ' + r.username + ' ' + r.sekolah + ' ' + r.dept).toLowerCase().indexOf(q) >= 0));
}

/* ===================== SESI ===================== */

function getSession_(token) {
  if (!token) throw new Error('Sesi tidak valid. Silakan login ulang.');
  const raw = CacheService.getScriptCache().get('tok_' + token);
  if (!raw) throw new Error('Sesi habis. Silakan login ulang.');
  return JSON.parse(raw);
}

function requireAdmin_(token) {
  const s = getSession_(token);
  if (findUser_(s.username).role !== 'ADMIN') throw new Error('Hanya admin yang boleh mengakses.');
  return s;
}

function clientData_(token, user) {
  const cfg = getConfig_();
  return {
    token: token,
    role: user.role,
    nama: user.nama,
    sekolah: user.sekolah,
    lokasiTetap: user.lokasi,
    dept: user.dept,
    mulai: user.mulai,
    selesai: user.selesai,
    lokasiList: getLokasiList_(),
    maxAkurasi: cfg.maxAkurasi,
    jamMasuk: cfg.jamMasuk,
    jamPulang: cfg.jamPulang,
    toleransi: cfg.toleransi,
    status: statusHariIni_(user.username)
  };
}

/* ===================== API: LOGIN ===================== */

function login(username, password) {
  ensureReady_();
  username = String(username || '').trim();
  password = String(password || '');
  if (!username || !password) throw new Error('Username dan password wajib diisi.');

  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + username.toLowerCase();
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= 5) throw new Error('Terlalu banyak percobaan. Coba lagi 10 menit lagi.');

  const user = getUsers_().find(u =>
    u.username.toLowerCase() === username.toLowerCase() && u.password === password);
  if (!user) {
    cache.put(failKey, String(fails + 1), 600);
    throw new Error('Username atau password salah.');
  }
  if (user.aktif === 'TIDAK') throw new Error('Akun tidak aktif.');
  cache.remove(failKey);

  const token = Utilities.getUuid();
  cache.put('tok_' + token, JSON.stringify({ username: user.username }), SESSION_HOURS * 3600);
  return clientData_(token, user);
}

/** Dipakai saat halaman dimuat ulang (mis. setelah buka kamera HP) agar tidak perlu login lagi */
function resume(token) {
  ensureReady_();
  const s = getSession_(token);
  return clientData_(token, findUser_(s.username));
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('tok_' + token);
  return true;
}

/* ===================== API: ABSEN ===================== */

function submitAbsen(token, jenis, lokasiNama, lat, lng, akurasi, fotoDataUrl, izin) {
  const user = findUser_(getSession_(token).username);
  jenis = String(jenis || '').toUpperCase();
  if (['MASUK', 'PULANG'].indexOf(jenis) < 0) throw new Error('Jenis absen tidak dikenal.');

  cekMasaPkl_(user);

  // Siswa yang sudah ditetapkan lokasinya oleh admin tidak bisa memilih lokasi lain
  lokasiNama = user.lokasi || String(lokasiNama || '');
  const lokasi = getLokasiList_().find(l => l.nama === lokasiNama);
  if (!lokasi) throw new Error('Pilih hotel/lokasi PKL terlebih dahulu.');

  lat = Number(lat); lng = Number(lng); akurasi = Number(akurasi);
  if (!isFinite(lat) || !isFinite(lng)) throw new Error('Lokasi tidak terbaca. Aktifkan GPS.');

  const cfg = getConfig_();
  if (akurasi > cfg.maxAkurasi) {
    throw new Error('Akurasi GPS terlalu rendah (' + Math.round(akurasi) +
      ' m, maks ' + cfg.maxAkurasi + ' m). Aktifkan GPS/mode akurasi tinggi.');
  }
  const jarak = distanceMeter_(lat, lng, lokasi.lat, lokasi.lng);
  if (jarak > lokasi.radius) {
    throw new Error('Anda berada ' + Math.round(jarak) + ' m dari ' + lokasi.nama +
      '. Maksimal ' + lokasi.radius + ' m.');
  }

  const m = /^data:image\/(jpeg|png);base64,(.+)$/.exec(String(fotoDataUrl || ''));
  if (!m) throw new Error('Foto selfie wajib diambil.');

  // Izin datang terlambat / pulang cepat (opsional)
  let izinKet = '', izinBlob = null;
  if (izin) {
    izinKet = String(izin.ket || '').trim().slice(0, 500);
    if (izinKet.length < 3) throw new Error('Isi alasan izin terlebih dahulu.');
    izinBlob = dokBlob_(user, jenis === 'MASUK' ? 'IZIN-TERLAMBAT' : 'IZIN-PULANG-CEPAT', izin.dok);
  }
  if (m[2].length > 4 * 1024 * 1024) throw new Error('Ukuran foto terlalu besar.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const st = statusHariIni_(user.username);
    if (st.KET) throw new Error('Hari ini Anda sudah tercatat ' + st.KET + '.');
    if (jenis === 'MASUK' && st.MASUK) throw new Error('Anda sudah absen masuk hari ini (' + st.MASUK + ').');
    if (jenis === 'PULANG' && !st.MASUK) throw new Error('Belum absen masuk hari ini.');
    if (jenis === 'PULANG' && st.PULANG) throw new Error('Anda sudah absen pulang hari ini (' + st.PULANG + ').');

    const now = new Date();
    const tgl = Utilities.formatDate(now, TZ, 'yyyy-MM-dd');
    const jam = Utilities.formatDate(now, TZ, 'HH:mm:ss');
    const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), 'image/' + m[1],
      user.username + '_' + jenis + '_' + Utilities.formatDate(now, TZ, 'yyyyMMdd_HHmmss') + '.jpg');
    const file = simpanFile_(getFolder_('FOLDER_ID', 'Foto Absensi PKL'), blob);

    // Hitung status terhadap jam kerja
    const menitSekarang = toMenit_(Utilities.formatDate(now, TZ, 'HH:mm'));
    let status = 'TEPAT WAKTU', telat = '';
    if (jenis === 'MASUK') {
      const selisih = menitSekarang - toMenit_(cfg.jamMasuk);
      if (selisih > cfg.toleransi) { status = izin ? 'IZIN TERLAMBAT' : 'TELAT'; telat = selisih; }
    } else if (menitSekarang < toMenit_(cfg.jamPulang)) {
      status = izin ? 'IZIN PULANG CEPAT' : 'PULANG CEPAT';
    }
    const dokUrl = izinBlob ? simpanFile_(getFolder_('DOC_FOLDER_ID', 'Dokumen Sakit & Izin PKL'), izinBlob).getUrl() : '';

    sheet_(SH.ABSEN).appendRow([tgl, jam, user.username, user.nama, user.sekolah, lokasi.nama, jenis,
      lat, lng, Math.round(akurasi), Math.round(jarak), file.getUrl(), status, telat, izinKet, dokUrl, user.dept]);

    st[jenis] = jam;
    if (jenis === 'MASUK') { st.MASUK_STATUS = status; st.TELAT = telat || 0; }
    if (jenis === 'PULANG') st.PULANG_STATUS = status;
    let pesan = 'Absen ' + jenis.toLowerCase() + ' di ' + lokasi.nama + ' berhasil pukul ' + jam;
    if (status === 'TELAT') pesan += ' — TERLAMBAT ' + telat + ' menit (jam masuk ' + cfg.jamMasuk + ')';
    if (status === 'IZIN TERLAMBAT') pesan += ' — tercatat IZIN datang terlambat (tidak dihitung telat)';
    if (status === 'PULANG CEPAT') pesan += ' — pulang sebelum jam ' + cfg.jamPulang;
    if (status === 'IZIN PULANG CEPAT') pesan += ' — tercatat IZIN pulang cepat';
    return { pesan: pesan, status: st, telat: status === 'TELAT' || status === 'PULANG CEPAT' };
  } finally {
    lock.releaseLock();
  }
}

/** Folder Drive dari Pengaturan; dibuat ulang otomatis jika belum ada / terhapus */
function getFolder_(key, nama) {
  const id = String(readConfigRaw_(sheet_(SH.CONFIG))[key] || '').trim();
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  const baru = DriveApp.createFolder(nama);
  setConfig_(key, baru.getId());
  return baru;
}

function simpanFile_(folder, blob) {
  try { return folder.createFile(blob); }
  catch (e) { throw new Error('Gagal menyimpan file ke Google Drive: ' + e.message); }
}

/** Ubah upload {name, mime, base64} menjadi blob (foto/PDF), atau null jika tidak ada */
function dokBlob_(user, label, dok) {
  if (!dok || !dok.base64) return null;
  const nama = String(dok.name || '').toLowerCase();
  let mime = String(dok.mime || '').toLowerCase();
  if (!mime) mime = /\.pdf$/.test(nama) ? 'application/pdf' : 'image/jpeg';
  if (mime.indexOf('image/') !== 0 && mime !== 'application/pdf') {
    throw new Error('Dokumen harus berupa foto atau PDF.');
  }
  if (String(dok.base64).length > 7 * 1024 * 1024) throw new Error('Ukuran dokumen maksimal 5 MB.');
  const ext = mime === 'application/pdf' ? '.pdf' : '.' + (mime.split('/')[1] || 'jpg').replace('jpeg', 'jpg');
  let bytes;
  try { bytes = Utilities.base64Decode(dok.base64); } catch (e) { throw new Error('File dokumen rusak, coba pilih ulang.'); }
  return Utilities.newBlob(bytes, mime,
    user.username + '_' + label + '_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmmss') + ext);
}

function cekMasaPkl_(user) {
  const t = today_();
  if (user.role === 'SISWA' && user.mulai && t < user.mulai) throw new Error('Masa PKL Anda baru mulai ' + user.mulai + '.');
  if (user.role === 'SISWA' && user.selesai && t > user.selesai) throw new Error('Masa PKL Anda sudah selesai (' + user.selesai + ').');
}

/**
 * Pengajuan SAKIT / IZIN (tidak perlu GPS & selfie).
 * dok: null atau { name, mime, base64 } — surat sakit / surat izin dari guru (gambar atau PDF)
 */
function submitIzin(token, jenis, keterangan, dok) {
  const user = findUser_(getSession_(token).username);
  cekMasaPkl_(user);
  jenis = String(jenis || '').toUpperCase();
  if (['SAKIT', 'IZIN'].indexOf(jenis) < 0) throw new Error('Pilih Sakit atau Izin.');
  keterangan = String(keterangan || '').trim();
  if (keterangan.length < 3) throw new Error('Keterangan wajib diisi.');
  if (keterangan.length > 500) keterangan = keterangan.slice(0, 500);

  const blob = dokBlob_(user, jenis, dok);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const st = statusHariIni_(user.username);
    if (st.KET) throw new Error('Hari ini Anda sudah mengajukan ' + st.KET + '.');
    if (st.MASUK) throw new Error('Anda sudah absen masuk hari ini, tidak bisa mengajukan ' + jenis.toLowerCase() + '.');

    const cfg = getConfig_();
    const url = blob ? simpanFile_(getFolder_('DOC_FOLDER_ID', 'Dokumen Sakit & Izin PKL'), blob).getUrl() : '';
    const now = new Date();
    const jam = Utilities.formatDate(now, TZ, 'HH:mm:ss');
    sheet_(SH.ABSEN).appendRow([today_(), jam, user.username, user.nama, user.sekolah, user.lokasi || '', jenis,
      '', '', '', '', '', jenis, '', keterangan, url, user.dept]);

    st.KET = jenis; st.KET_JAM = jam;
    return { pesan: jenis === 'SAKIT' ? 'Pengajuan sakit terkirim. Semoga lekas sembuh 🙏' : 'Pengajuan izin terkirim.', status: st };
  } finally {
    lock.releaseLock();
  }
}

/* ===================== API: ADMIN ===================== */

function adminGetData(token) {
  requireAdmin_(token);
  return {
    lokasi: getLokasiList_(),
    users: getUsers_(),
    setting: (function (c) { return { maxAkurasi: c.maxAkurasi, jamMasuk: c.jamMasuk, jamPulang: c.jamPulang, toleransi: c.toleransi, departemen: c.departemen }; })(getConfig_())
  };
}

function adminSaveLokasi(token, list) {
  requireAdmin_(token);
  const seen = {};
  const rows = (list || []).map((l, i) => {
    const nama = String(l.nama || '').trim();
    const lat = Number(l.lat), lng = Number(l.lng), radius = Number(l.radius);
    const no = 'Baris ' + (i + 1) + ': ';
    if (!nama) throw new Error(no + 'nama lokasi wajib diisi.');
    if (seen[nama.toLowerCase()]) throw new Error(no + 'nama "' + nama + '" dobel.');
    seen[nama.toLowerCase()] = true;
    if (!isFinite(lat) || lat < -90 || lat > 90 || String(l.lat).trim() === '') throw new Error(no + 'latitude tidak valid.');
    if (!isFinite(lng) || lng < -180 || lng > 180 || String(l.lng).trim() === '') throw new Error(no + 'longitude tidak valid.');
    if (!isFinite(radius) || radius < 10 || radius > 5000) throw new Error(no + 'radius harus 10 - 5000 meter.');
    return [nama, lat, lng, Math.round(radius)];
  });
  if (!rows.length) throw new Error('Minimal harus ada 1 lokasi.');
  writeTable_(SH.LOKASI, rows, 4);
  return getLokasiList_();
}

function adminSaveUsers(token, list) {
  const me = requireAdmin_(token);
  const lokasiNames = getLokasiList_().map(l => l.nama);
  const seen = {};
  const rows = (list || []).map((u, i) => {
    const no = 'Baris ' + (i + 1) + ': ';
    const username = String(u.username || '').trim();
    const password = String(u.password || '');
    if (!username) throw new Error(no + 'username wajib diisi.');
    if (/\s/.test(username)) throw new Error(no + 'username tidak boleh ada spasi.');
    if (seen[username.toLowerCase()]) throw new Error(no + 'username "' + username + '" dobel.');
    seen[username.toLowerCase()] = true;
    if (password.length < 4) throw new Error(no + 'password minimal 4 karakter.');
    const lokasi = String(u.lokasi || '').trim();
    if (lokasi && lokasiNames.indexOf(lokasi) < 0) throw new Error(no + 'lokasi "' + lokasi + '" tidak ada.');
    const role = String(u.role).toUpperCase() === 'ADMIN' ? 'ADMIN' : 'SISWA';
    const aktif = String(u.aktif).toUpperCase() === 'TIDAK' ? 'TIDAK' : 'YA';
    const tgl = v => { v = String(v || '').trim(); if (v && !/^\d{4}-\d{2}-\d{2}$/.test(v)) throw new Error(no + 'format tanggal salah.'); return v; };
    const mulai = tgl(u.mulai), selesai = tgl(u.selesai);
    if (mulai && selesai && selesai < mulai) throw new Error(no + 'tanggal selesai sebelum tanggal mulai.');
    const hp = String(u.hp || '').replace(/[^\d+]/g, '');
    return [username, password, String(u.nama || '').trim(), String(u.sekolah || '').trim(), lokasi, role, aktif,
      hp, String(u.guru || '').trim(), mulai, selesai, String(u.dept || '').trim()];
  });
  if (!rows.some(r => r[5] === 'ADMIN' && r[6] === 'YA')) throw new Error('Minimal harus ada 1 admin yang aktif.');
  if (!rows.some(r => r[0].toLowerCase() === me.username.toLowerCase() && r[5] === 'ADMIN' && r[6] === 'YA')) {
    throw new Error('Anda tidak bisa menghapus/menonaktifkan akun admin Anda sendiri.');
  }
  writeTable_(SH.USERS, rows, 12, true);
  return getUsers_();
}

function adminSaveSetting(token, s) {
  requireAdmin_(token);
  s = s || {};
  const acc = Number(s.maxAkurasi);
  if (!isFinite(acc) || acc < 10 || acc > 5000) throw new Error('Maks akurasi harus 10 - 5000 meter.');
  const jm = normTime_(s.jamMasuk, ''), jp = normTime_(s.jamPulang, '');
  if (!jm || !jp) throw new Error('Jam masuk dan jam pulang wajib diisi (format JJ:MM).');
  if (toMenit_(jp) <= toMenit_(jm)) throw new Error('Jam pulang harus setelah jam masuk.');
  const tol = Number(s.toleransi);
  if (!isFinite(tol) || tol < 0 || tol > 240) throw new Error('Toleransi telat harus 0 - 240 menit.');
  setConfig_('MAX_AKURASI_METER', Math.round(acc));
  setConfig_('JAM_MASUK', jm);
  setConfig_('JAM_PULANG', jp);
  setConfig_('TOLERANSI_MENIT', Math.round(tol));
  const dept = (Array.isArray(s.departemen) ? s.departemen : String(s.departemen || '').split(/[,\n]/))
    .map(x => String(x).trim()).filter(Boolean);
  if (!dept.length) throw new Error('Daftar departemen tidak boleh kosong.');
  setConfig_('DEPARTEMEN', dept.filter((x, i) => dept.indexOf(x) === i).join(', '));
  return true;
}

/* ===================== API: REKAP (ADMIN) ===================== */

function adminGetRekap(token, f) {
  requireAdmin_(token);
  const rows = filterAbsen_(absenDenganDept_(), f)
    .sort((a, b) => (b.tgl + b.jam).localeCompare(a.tgl + a.jam));
  return { total: rows.length, rows: rows.slice(0, 1000), ringkasan: ringkasan_(rows) };
}

/** Absen lama (sebelum ada kolom departemen) diisi dari data siswa saat ini */
function absenDenganDept_() {
  const map = {};
  getUsers_().forEach(u => { map[u.username] = u.dept; });
  return readAbsen_().map(r => { if (!r.dept) r.dept = map[r.username] || ''; return r; });
}

function ringkasan_(rows) {
  const r = { MASUK: 0, TELAT: 0, IZIN_TERLAMBAT: 0, PULANG: 0, PULANG_CEPAT: 0, IZIN_PULANG_CEPAT: 0, SAKIT: 0, IZIN: 0 };
  rows.forEach(x => {
    if (r.hasOwnProperty(x.jenis)) r[x.jenis]++;
    const k = String(x.status).replace(/ /g, '_');
    if (k !== 'SAKIT' && k !== 'IZIN' && r.hasOwnProperty(k)) r[k]++;
  });
  return r;
}

/** items: [{row, tgl, jam, username}] — dicocokkan ulang agar tidak salah hapus */
function adminDeleteAbsen(token, items, hapusFoto) {
  requireAdmin_(token);
  if (!items || !items.length) throw new Error('Tidak ada data yang dipilih.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const byRow = {};
    readAbsen_().forEach(r => { byRow[r.row] = r; });
    const target = [];
    items.forEach(it => {
      const r = byRow[Number(it.row)];
      if (r && r.tgl === it.tgl && r.jam === it.jam && r.username === it.username) target.push(r);
    });
    if (!target.length) throw new Error('Data sudah berubah. Muat ulang rekap lalu coba lagi.');

    if (hapusFoto) {
      target.forEach(r => [r.foto, r.dok].forEach(u => {
        const m = /[-\w]{25,}/.exec(u || '');
        if (m) { try { DriveApp.getFileById(m[0]).setTrashed(true); } catch (e) {} }
      }));
    }
    // hapus dari baris paling bawah, gabungkan baris yang berurutan
    const rowsDesc = target.map(r => r.row).sort((a, b) => b - a);
    const sh = sheet_(SH.ABSEN);
    let i = 0;
    while (i < rowsDesc.length) {
      let start = rowsDesc[i], count = 1;
      while (i + count < rowsDesc.length && rowsDesc[i + count] === start - 1) { start--; count++; }
      sh.deleteRows(start, count);
      i += count;
    }
    return target.length;
  } finally {
    lock.releaseLock();
  }
}

/** format: 'xlsx' atau 'pdf'. Mengembalikan file dalam base64 untuk diunduh di browser */
function adminExport(token, f, format) {
  requireAdmin_(token);
  f = f || {};
  const pdf = format === 'pdf';
  const rows = filterAbsen_(absenDenganDept_(), f)
    .sort((a, b) => (a.tgl + a.jam).localeCompare(b.tgl + b.jam));

  const stTeks = r => r.status === 'TELAT' ? 'TELAT ' + r.telat + ' mnt'
    : r.status === 'IZIN TERLAMBAT' ? 'IZIN TERLAMBAT (' + r.telat + ' mnt)' : r.status;
  const head = pdf
    ? ['No', 'Tanggal', 'Jam', 'Nama', 'Departemen', 'Sekolah', 'Hotel', 'Jenis', 'Status', 'Keterangan', 'Jarak (m)']
    : ['No', 'Tanggal', 'Jam', 'Username', 'Nama', 'Departemen', 'Sekolah', 'Hotel', 'Jenis', 'Status', 'Terlambat (menit)',
       'Keterangan', 'Latitude', 'Longitude', 'Akurasi GPS (m)', 'Jarak (m)', 'Link Foto', 'Link Dokumen'];
  const data = rows.map((r, i) => pdf
    ? [i + 1, r.tgl, r.jam, r.nama, r.dept, r.sekolah, r.lokasi, r.jenis, stTeks(r), r.ket + (r.dok ? ' (ada surat)' : ''), r.jarak]
    : [i + 1, r.tgl, r.jam, r.username, r.nama, r.dept, r.sekolah, r.lokasi, r.jenis, r.status, r.telat || '',
       r.ket, r.lat, r.lng, r.akurasi, r.jarak, r.foto, r.dok]);
  const rk = ringkasan_(rows);

  const periode = (f.dari || 'awal') + ' s/d ' + (f.sampai || 'sekarang');
  const tmp = SpreadsheetApp.create('tmp_export_absensi_' + Date.now());
  try {
    const sh = tmp.getSheets()[0];
    sh.setName('Rekap Absensi');
    sh.getRange(1, 1).setValue('REKAP ABSENSI PKL').setFontWeight('bold').setFontSize(14);
    sh.getRange(2, 1).setValue('Periode: ' + periode + (f.lokasi ? '   |   Hotel: ' + f.lokasi : '') +
      (f.dept ? '   |   Departemen: ' + f.dept : ''));
    sh.getRange(3, 1).setValue('Dicetak: ' + Utilities.formatDate(new Date(), TZ, 'dd-MM-yyyy HH:mm') +
      '   |   Jumlah data: ' + data.length + '   |   Masuk: ' + rk.MASUK + ' (telat ' + rk.TELAT + ', izin terlambat ' + rk.IZIN_TERLAMBAT +
      ')   Pulang: ' + rk.PULANG + ' (pulang cepat ' + rk.PULANG_CEPAT + ', izin pulang cepat ' + rk.IZIN_PULANG_CEPAT +
      ')   Sakit: ' + rk.SAKIT + '   Izin tidak hadir: ' + rk.IZIN);
    sh.getRange(5, 1, 1, head.length).setValues([head])
      .setFontWeight('bold').setBackground('#1d4ed8').setFontColor('#ffffff');
    if (data.length) {
      sh.getRange(6, 2, data.length, 2).setNumberFormat('@'); // tanggal & jam sebagai teks
      sh.getRange(6, 1, data.length, head.length).setValues(data);
    }
    sh.getRange(5, 1, data.length + 1, head.length).setBorder(true, true, true, true, true, true);
    sh.autoResizeColumns(1, head.length);
    if (sh.getMaxColumns() > head.length) sh.deleteColumns(head.length + 1, sh.getMaxColumns() - head.length);
    const lastRow = Math.max(6, data.length + 5);
    if (sh.getMaxRows() > lastRow) sh.deleteRows(lastRow + 1, sh.getMaxRows() - lastRow);
    SpreadsheetApp.flush();

    const base = 'https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?';
    const url = pdf
      ? base + 'format=pdf&size=A4&portrait=false&fitw=true&gridlines=false&sheetnames=false' +
        '&printtitle=false&pagenum=CENTER&top_margin=0.5&bottom_margin=0.5&left_margin=0.4&right_margin=0.4' +
        '&gid=' + sh.getSheetId()
      : base + 'format=xlsx';
    const resp = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) throw new Error('Gagal membuat file (' + resp.getResponseCode() + ').');

    const nama = 'Rekap_Absensi_PKL_' + (f.dari || 'awal') + '_sd_' + (f.sampai || 'sekarang') + (pdf ? '.pdf' : '.xlsx');
    return {
      name: nama,
      mime: pdf ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      base64: Utilities.base64Encode(resp.getBlob().getBytes())
    };
  } finally {
    DriveApp.getFileById(tmp.getId()).setTrashed(true);
  }
}

function writeTable_(name, rows, cols, textAB) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_(name);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, cols).clearContent();
    if (textAB) setUserTextFormat_(sh);
    sh.getRange(2, 1, rows.length, cols).setValues(rows);
  } finally {
    lock.releaseLock();
  }
}
