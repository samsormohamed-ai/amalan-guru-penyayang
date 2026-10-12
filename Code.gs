/**
 * AMALAN GURU PENYAYANG (AGP) — SK JAWI
 * Versi Dalam Talian 2.0: Google Sheet + Apps Script + Google Drive
 *
 * Baharu dalam 2.0: 4 lokasi bertugas, pentadbir turut serta, elemen SUDI,
 * senarai PRS (CSV) dan analisis kekerapan.
 * Sheet versi 1 dinaik taraf secara automatik (lajur baharu ditambah di hujung).
 *
 * Cara pasang (ringkas):
 * 1. Buka Google Sheet > Extensions > Apps Script.
 * 2. Tampal kod ini ke dalam Code.gs. Fail HTML bernama "Index" diisi dengan Index.html.
 * 3. Pilih fungsi "setup" > Run > benarkan akses (Authorize).
 * 4. Deploy > New deployment > Web app  (Execute as: Me | Who has access: Anyone)
 *    Jika sudah pernah deploy: Deploy > Manage deployments > pensel > Version: New version > Deploy.
 * 5. PIN ditetapkan dalam tab "Tetapan" (sel B2). Kosongkan jika tidak mahu kunci PIN.
 */

const SHEET_LAPORAN = 'Laporan AGP';
const SHEET_TETAPAN = 'Tetapan';
const SHEET_PRS     = 'Senarai PRS';
const NAMA_FOLDER   = 'AGP SK Jawi - Gambar & Tandatangan';
const FAVICON_URL   = ''; // pilihan: tampal URL icon-192.png dari GitHub Pages
const MAX_GAMBAR    = 8;

const HARI_MS  = ['Ahad','Isnin','Selasa','Rabu','Khamis','Jumaat','Sabtu'];
const BULAN_MS = ['Januari','Februari','Mac','April','Mei','Jun','Julai','Ogos','September','Oktober','November','Disember'];

const FIELDS = [
  ['id','ID Laporan'],['tarikh','Tarikh'],['hari','Hari'],['bulan','Bulan'],
  ['masaMula','Masa Mula'],['masaTamat','Masa Tamat'],['lokasi','Lokasi'],
  ['guru1','Guru Bertugas 1'],['guru2','Guru Bertugas 2'],['guru3','Guru Bertugas 3'],['guru4','Guru Bertugas 4'],
  ['pentadbir','Pentadbir Turut Serta'],
  ['prs1','PRS 1'],['prs2','PRS 2'],['prs3','PRS 3'],['prs4','PRS 4'],
  ['c_senyum','Guru Menyambut Dengan SENYUMAN'],['c_sapa','Menerima / Memberi SAPAAN'],
  ['c_salam','Murid Mengucap SALAM'],['c_jawab','Murid Menjawab SALAM Guru'],
  ['c_sudi','Guru SUDI Mendengar & Membantu'],
  ['bilMurid','Bil. Murid Terlibat'],['bilGuru','Bil. Guru Terlibat'],
  ['akt1','Aktiviti 1'],['akt2','Aktiviti 2'],['akt3','Aktiviti 3'],['bilGambar','Bil. Gambar'],
  ['dicatatNama','Dicatat Oleh (PRS)'],['disemakNama','Disemak Oleh'],['disemakJawatan','Jawatan Penyemak'],
  ['statusTT','Status Tandatangan'],['dikemaskini','Dikemas Kini'],
  ['gambarJson','Gambar (ID Drive)'],['sig1Id','ID Tandatangan Pencatat'],['sig2Id','ID Tandatangan Penyemak']
];
const KEYS      = FIELDS.map(f => f[0]);
const CHECKS    = ['c_senyum','c_sapa','c_salam','c_jawab','c_sudi'];
const TEXT_COLS = ['tarikh','masaMula','masaTamat'];
const TECH_COLS = ['gambarJson','sig1Id','sig2Id'];

/* ================= WEB APP ================= */
function doGet() {
  const out = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Amalan Guru Penyayang SK Jawi')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  if (FAVICON_URL) out.setFaviconUrl(FAVICON_URL);
  return out;
}

/* Jalankan sekali dari editor untuk sediakan / naik taraf Sheet & folder Drive */
function setup() {
  ensureSheets_();
  const f = getRootFolder_();
  Logger.log('Siap. Sheet: ' + SpreadsheetApp.getActive().getUrl());
  Logger.log('Folder gambar: ' + f.getUrl());
}

/* ================= API ================= */
function apiInit(pin) {
  ensureSheets_();
  return { pinRequired: !!getPin_(), appUrl: ScriptApp.getService().getUrl() };
}

function apiCheckPin(pin) {
  const p = getPin_();
  return !p || String(pin || '').trim() === p;
}

function apiList(pin) {
  checkPin_(pin);
  const sh = ensureSheets_();
  const m = colMap_(sh);
  const last = sh.getLastRow();
  const rows = [];
  if (last > 1) {
    const vals = sh.getRange(2, 1, last - 1, m.width).getValues();
    vals.forEach(v => {
      const o = rowToObj_(v, m);
      if (!o.id) return;
      TECH_COLS.forEach(k => delete o[k]);
      delete o.akt1; delete o.akt2; delete o.akt3;
      o.bilGambar = Number(o.bilGambar) || 0;
      rows.push(o);
    });
  }
  rows.sort((a, b) => String(b.tarikh).localeCompare(String(a.tarikh)) || String(b.masaMula).localeCompare(String(a.masaMula)));
  return { rows: rows, prs: getPrs_(), sheetUrl: SpreadsheetApp.getActive().getUrl() };
}

function apiGet(pin, id) {
  checkPin_(pin);
  const sh = ensureSheets_();
  const m = colMap_(sh);
  const row = findRow_(sh, m, id);
  if (row < 0) throw new Error('Laporan tidak dijumpai. Mungkin telah dipadam.');
  const o = rowToObj_(sh.getRange(row, 1, 1, m.width).getValues()[0], m);
  o.gambar = parseJson_(o.gambarJson).map(p => ({ id: p.id, cap: p.cap || '', src: fileDataUrl_(p.id) })).filter(p => p.src);
  o.dicatatSig = o.sig1Id ? (fileDataUrl_(o.sig1Id) || null) : null;
  o.disemakSig = o.sig2Id ? (fileDataUrl_(o.sig2Id) || null) : null;
  TECH_COLS.forEach(k => delete o[k]);
  o.saved = true;
  return o;
}

function apiSave(pin, r) {
  checkPin_(pin);
  if (!r || !r.tarikh) throw new Error('Sila pilih tarikh laporan.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = ensureSheets_();
    const m = colMap_(sh);
    const id = String(r.id || ('r' + Date.now()));
    let row = findRow_(sh, m, id);
    const oldRow = row > 0 ? sh.getRange(row, 1, 1, m.width).getValues()[0] : null;
    const old = oldRow ? rowToObj_(oldRow, m) : {};
    const folder = getMonthFolder_(r.tarikh);
    const stamp = Date.now();

    const oldPhotos = parseJson_(old.gambarJson);
    const oldIds = {};
    oldPhotos.forEach(p => oldIds[p.id] = true);
    const photos = [];
    (r.gambar || []).slice(0, MAX_GAMBAR).forEach((g, i) => {
      if (g.id && oldIds[g.id]) photos.push({ id: g.id, cap: String(g.cap || '') });
      else if (g.src) {
        const fid = saveDataUrl_(g.src, 'AGP_' + r.tarikh + '_gambar' + (i + 1) + '_' + stamp + '.jpg', folder);
        if (fid) photos.push({ id: fid, cap: String(g.cap || '') });
      }
    });
    const keep = {};
    photos.forEach(p => keep[p.id] = true);
    oldPhotos.forEach(p => { if (!keep[p.id]) trash_(p.id); });

    if (old.sig1Id) trash_(old.sig1Id);
    if (old.sig2Id) trash_(old.sig2Id);
    const sig1Id = r.dicatatSig ? saveDataUrl_(r.dicatatSig, 'AGP_' + r.tarikh + '_TT_pencatat_' + stamp + '.png', folder) : '';
    const sig2Id = r.disemakSig ? saveDataUrl_(r.disemakSig, 'AGP_' + r.tarikh + '_TT_penyemak_' + stamp + '.png', folder) : '';

    const d = parseDate_(r.tarikh);
    const rec = {
      id: id, tarikh: r.tarikh,
      hari: d ? HARI_MS[d.getDay()] : '', bulan: d ? BULAN_MS[d.getMonth()] : '',
      masaMula: r.masaMula || '', masaTamat: r.masaTamat || '', lokasi: r.lokasi || '',
      guru1: r.guru1, guru2: r.guru2, guru3: r.guru3, guru4: r.guru4, pentadbir: r.pentadbir || '',
      prs1: r.prs1, prs2: r.prs2, prs3: r.prs3, prs4: r.prs4,
      bilMurid: num_(r.bilMurid), bilGuru: num_(r.bilGuru),
      akt1: r.akt1, akt2: r.akt2, akt3: r.akt3, bilGambar: photos.length,
      dicatatNama: r.dicatatNama, disemakNama: r.disemakNama, disemakJawatan: r.disemakJawatan,
      statusTT: (sig1Id && sig2Id) ? 'Lengkap' : ((sig1Id || sig2Id) ? 'Separuh' : 'Belum'),
      dikemaskini: Utilities.formatDate(new Date(), tz_(), 'dd/MM/yyyy HH:mm'),
      gambarJson: JSON.stringify(photos), sig1Id: sig1Id, sig2Id: sig2Id
    };
    CHECKS.forEach(k => rec[k] = r[k] ? 'X' : '');

    const out = oldRow ? oldRow.slice() : new Array(m.width).fill('');
    KEYS.forEach(k => { out[m.idx[k] - 1] = safe_(rec[k]); });
    if (row < 0) row = sh.getLastRow() + 1;
    TEXT_COLS.forEach(k => sh.getRange(row, m.idx[k]).setNumberFormat('@'));
    sh.getRange(row, 1, 1, m.width).setValues([out]);
    return { id: id, gambar: photos.map(p => ({ id: p.id })) };
  } finally {
    lock.releaseLock();
  }
}

function apiDelete(pin, id) {
  checkPin_(pin);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = ensureSheets_();
    const m = colMap_(sh);
    const row = findRow_(sh, m, id);
    if (row < 0) return true;
    const o = rowToObj_(sh.getRange(row, 1, 1, m.width).getValues()[0], m);
    parseJson_(o.gambarJson).forEach(p => trash_(p.id));
    if (o.sig1Id) trash_(o.sig1Id);
    if (o.sig2Id) trash_(o.sig2Id);
    sh.deleteRow(row);
    return true;
  } finally {
    lock.releaseLock();
  }
}

/* Senarai PRS: [{nama, kelas}] — menggantikan senarai sedia ada */
function apiSetPrs(pin, list) {
  checkPin_(pin);
  const ss = SpreadsheetApp.getActive();
  const sh = ensurePrsSheet_(ss);
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, 3).clearContent();
  const rows = (list || []).filter(p => p && String(p.nama || '').trim())
    .map((p, i) => [i + 1, safe_(String(p.nama).trim()), safe_(String(p.kelas || '').trim())]);
  if (rows.length) sh.getRange(2, 1, rows.length, 3).setValues(rows);
  return rows.length;
}

/* ================= HELPERS ================= */
function ensureSheets_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_LAPORAN);
  if (!sh) {
    sh = ss.insertSheet(SHEET_LAPORAN, 0);
    ss.getSheets().forEach(s => {
      if ([SHEET_LAPORAN, SHEET_TETAPAN, SHEET_PRS].indexOf(s.getName()) < 0 && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
    });
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, FIELDS.length).setValues([FIELDS.map(f => f[1])]);
    styleHeader_(sh.getRange(1, 1, 1, FIELDS.length));
    sh.setFrozenRows(1);
    sh.setRowHeight(1, 42);
    sh.setColumnWidths(1, FIELDS.length, 140);
    sh.hideColumns(FIELDS.length - 2, 3); // lajur teknikal (ID fail Drive)
  } else {
    // Naik taraf: tambah lajur yang belum wujud di hujung
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    const missing = FIELDS.filter(f => head.indexOf(f[1]) < 0);
    if (missing.length) {
      const start = head.length + 1;
      if (sh.getMaxColumns() < start + missing.length - 1) sh.insertColumnsAfter(sh.getMaxColumns(), start + missing.length - 1 - sh.getMaxColumns());
      sh.getRange(1, start, 1, missing.length).setValues([missing.map(f => f[1])]);
      styleHeader_(sh.getRange(1, start, 1, missing.length));
    }
  }
  let st = ss.getSheetByName(SHEET_TETAPAN);
  if (!st) {
    st = ss.insertSheet(SHEET_TETAPAN);
    st.getRange('A1:B1').setValues([['Tetapan', 'Nilai']]);
    styleHeader_(st.getRange('A1:B1'));
    st.getRange('A2:A3').setValues([['PIN aplikasi'], ['Nota']]).setFontWeight('bold');
    st.getRange('B2').setNumberFormat('@').setValue('');
    st.getRange('B3').setValue('Isi PIN (contoh 1234) dalam sel B2 untuk mengunci aplikasi. Kosongkan B2 jika tidak mahu PIN.');
    st.setColumnWidth(1, 160); st.setColumnWidth(2, 520);
  }
  ensurePrsSheet_(ss);
  return sh;
}

function ensurePrsSheet_(ss) {
  let sh = ss.getSheetByName(SHEET_PRS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_PRS);
    sh.getRange('A1:C1').setValues([['Bil', 'Nama PRS', 'Kelas']]);
    styleHeader_(sh.getRange('A1:C1'));
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 50); sh.setColumnWidth(2, 320); sh.setColumnWidth(3, 160);
  }
  return sh;
}

function getPrs_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_PRS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 2, sh.getLastRow() - 1, 2).getDisplayValues()
    .map(r => ({ nama: String(r[0]).trim(), kelas: String(r[1]).trim() }))
    .filter(p => p.nama);
}

function styleHeader_(rg) {
  rg.setFontWeight('bold').setFontColor('#ffffff').setBackground('#13235b').setWrap(true).setVerticalAlignment('middle');
}

/* Peta lajur ikut tajuk, supaya Sheet lama & baharu sama-sama berfungsi */
function colMap_(sh) {
  const width = sh.getLastColumn();
  const head = sh.getRange(1, 1, 1, width).getValues()[0].map(String);
  const idx = {};
  FIELDS.forEach(f => {
    let i = head.indexOf(f[1]);
    if (i < 0 && f[0] === 'c_jawab') i = head.indexOf('Murid Menjawab SALAM Guru');
    if (i < 0) throw new Error('Lajur "' + f[1] + '" tiada dalam Sheet. Jalankan fungsi setup sekali lagi.');
    idx[f[0]] = i + 1;
  });
  return { idx: idx, width: width };
}

function getPin_() {
  const st = SpreadsheetApp.getActive().getSheetByName(SHEET_TETAPAN);
  if (!st) return '';
  return String(st.getRange('B2').getDisplayValue() || '').trim();
}

function checkPin_(pin) {
  if (!apiCheckPin(pin)) throw new Error('PIN_SALAH');
}

function getRootFolder_() {
  const props = PropertiesService.getScriptProperties();
  const fid = props.getProperty('FOLDER_ID');
  if (fid) { try { const f = DriveApp.getFolderById(fid); if (!f.isTrashed()) return f; } catch (e) {} }
  const f = DriveApp.createFolder(NAMA_FOLDER);
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}

function getMonthFolder_(tarikh) {
  const root = getRootFolder_();
  const name = String(tarikh).slice(0, 7) || 'Lain-lain';
  const it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

function findRow_(sh, m, id) {
  if (!id || sh.getLastRow() < 2) return -1;
  const hit = sh.getRange(2, m.idx.id, sh.getLastRow() - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : -1;
}

function rowToObj_(v, m) {
  const o = {};
  KEYS.forEach(k => o[k] = norm_(v[m.idx[k] - 1], k));
  CHECKS.forEach(k => o[k] = String(o[k]).toUpperCase() === 'X');
  return o;
}

function norm_(v, k) {
  if (v instanceof Date) {
    if (k === 'tarikh') return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
    if (k === 'masaMula' || k === 'masaTamat') return Utilities.formatDate(v, tz_(), 'HH:mm');
    return Utilities.formatDate(v, tz_(), 'dd/MM/yyyy HH:mm');
  }
  if (v === null || v === undefined) return '';
  return v;
}

function tz_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone() || 'Asia/Kuala_Lumpur'; }
function parseJson_(s) { try { const a = JSON.parse(s || '[]'); return Array.isArray(a) ? a : []; } catch (e) { return []; } }
function parseDate_(s) { const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; }
function num_(v) { return (v === '' || v === null || v === undefined || isNaN(Number(v))) ? '' : Number(v); }
function safe_(v) { if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v; return v === undefined || v === null ? '' : v; }

function saveDataUrl_(dataUrl, name, folder) {
  const m = String(dataUrl || '').match(/^data:([^;]+);base64,(.+)$/);
  if (!m) return '';
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name);
  return folder.createFile(blob).getId();
}

function fileDataUrl_(id) {
  try {
    const b = DriveApp.getFileById(id).getBlob();
    return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
  } catch (e) { return ''; }
}

function trash_(id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }
