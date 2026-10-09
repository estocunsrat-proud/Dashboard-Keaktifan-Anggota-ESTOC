/**
 * ==========================================================
 * DASHBOARD KEAKTIFAN ANGGOTA — KSPM ESTOC UNSRAT
 * Backend Google Apps Script (API-only, deploy sebagai Web App)
 * ==========================================================
 *
 * ---------------------------------------------------------
 * ARSITEKTUR MULTI-DIVISI
 * ---------------------------------------------------------
 * 1 periode sekarang terdiri dari BEBERAPA sheet, bukan cuma 1:
 *   - "{periode} | BOD"      -> roster SEMUA anggota (dikelompokkan per
 *                               divisi), PLUS kegiatan yang diurus BOD
 *                               sendiri (Rapat Umum, Rapat Inti, rapat-rapat
 *                               lain) yang berlaku untuk SEMUA anggota.
 *   - "{periode} | {Divisi}" -> 1 sheet per divisi (Training, dst — otomatis
 *                               mengikuti nama grup yang ada di roster BOD,
 *                               KECUALI grup BOD/Board of Director itu sendiri).
 *                               Isinya kegiatan & kehadiran milik divisi itu
 *                               SENDIRI, tapi ROSTER-nya tetap SELURUH
 *                               anggota organisasi (sama seperti BOD) —
 *                               karena peserta 1 kegiatan divisi BISA dari
 *                               divisi lain juga. Status "B" (Bukan Peserta)
 *                               dipakai untuk anggota yang memang bukan
 *                               peserta kegiatan itu.
 * Semua sheet (BOD + tiap divisi) SELALU punya roster yang sama persis
 * (nama & pengelompokan divisi) — cuma beda di kolom kegiatannya masing-
 * masing. Setiap ada tambah/edit/pindah/hapus anggota di BOD, backend ini
 * otomatis menyamakan roster itu ke SEMUA sheet lain juga.
 *
 * HANYA BOD yang bisa: tambah/edit/pindah/hapus anggota, dan bikin periode
 * baru.
 *
 * Password ADMIN beda per divisi, disimpan di Script Properties dengan
 * pola key: ADMIN_PASSWORD_<NAMA_DIVISI_DALAM_HURUF_BESAR_TANPA_SPASI>.
 * Contoh: divisi "BOD" -> ADMIN_PASSWORD_BOD
 *         divisi "Training and Skill Development Division" ->
 *         ADMIN_PASSWORD_TRAINING_AND_SKILL_DEVELOPMENT_DIVISION
 * Kalau BOD bikin divisi baru, WAJIB tambahkan Script Property untuk
 * password divisi itu secara manual sebelum divisi itu bisa dipakai login.
 *
 * Tab "Ikhtisar" beda isinya tergantung siapa yang lihat:
 *   - Admin manapun yang login (BOD atau divisi tertentu): Ikhtisar HANYA
 *     berisi kegiatan divisinya SENDIRI (dihitung dari sheet mereka saja).
 *   - Tampilan publik "Anggota" (tanpa login): Ikhtisar GABUNGAN dari
 *     BOD + semua sheet divisi.
 *
 * Cara pasang: sama seperti sebelumnya — tempel ini sebagai Code.gs,
 * Script Properties diisi ADMIN_PASSWORD_BOD dan ADMIN_PASSWORD_<DIVISI>
 * untuk tiap divisi yang sudah ada, lalu deploy sebagai Web App.
 */

const TIMEZONE = 'Asia/Makassar'; // WITA
const SHEET_SEP = ' | '; // pemisah nama periode & divisi di nama tab, mis. "2026 | Training"
const BOD_LABEL = 'BOD';

const HEADER_ROW_NAMA = 2;
const HEADER_ROW_TANGGAL = 3;
const FIRST_DATA_ROW = 4;
const FIRST_ACTIVITY_COL = 3; // kolom C

const STATUS_VALID = ['H', 'A', 'I', 'B', ''];
const NONAKTIF_LABEL = 'Non-Aktif (Pindah Divisi)'; // sisa desain lama, hanya dipakai file migrasi

function getSS_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------------------------------------------------------
 * NAMA SHEET & AUTH PER DIVISI
 * --------------------------------------------------------- */
function sheetName_(periode, divisi) {
  return `${periode}${SHEET_SEP}${divisi}`;
}

function isBodLabel_(nama) {
  const n = String(nama || '').trim().toLowerCase();
  return n === 'bod' || n === 'board of director';
}

function slugDivisi_(nama) {
  return String(nama).toUpperCase().replace(/[^A-Z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}

function checkDivisiAuth_(divisi, password) {
  const key = 'ADMIN_PASSWORD_' + slugDivisi_(divisi);
  const real = PropertiesService.getScriptProperties().getProperty(key);
  return real && password && password === real;
}

function getDivisiSheet_(ss, periode, divisi) {
  return ss.getSheetByName(sheetName_(periode, divisi)) || null;
}

function getOrCreateDivisiSheet_(ss, periode, divisi) {
  const name = sheetName_(periode, divisi);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 3).setValue('Nama Kegiatan');
    sheet.getRange(HEADER_ROW_NAMA, 1).setValue('No');
    sheet.getRange(HEADER_ROW_NAMA, 2).setValue('Nama');
    sheet.getRange(HEADER_ROW_NAMA, 1, 2, 1).merge();
    sheet.getRange(HEADER_ROW_NAMA, 2, 2, 1).merge();
  }
  return sheet;
}

/* Daftar nama periode (unik, terbaru duluan), dibaca dari semua nama tab
 * yang berpola "{periode} | {divisi}". */
function listPeriode_(ss) {
  const seen = [];
  ss.getSheets().forEach(s => {
    const idx = s.getName().indexOf(SHEET_SEP);
    if (idx === -1) return;
    const periode = s.getName().slice(0, idx);
    if (seen.indexOf(periode) === -1) seen.push(periode);
  });
  return seen.reverse();
}

/* Daftar divisi (termasuk BOD) untuk 1 periode, BOD selalu di depan. */
function listDivisi_(ss, periode) {
  const out = [];
  ss.getSheets().forEach(s => {
    const prefix = periode + SHEET_SEP;
    if (s.getName().indexOf(prefix) === 0) out.push(s.getName().slice(prefix.length));
  });
  out.sort((a, b) => {
    if (isBodLabel_(a)) return -1;
    if (isBodLabel_(b)) return 1;
    return a.localeCompare(b);
  });
  return out;
}

/* ---------------------------------------------------------
 * doGet
 *   (tanpa action)                                   -> halaman dashboard (Admin.html)
 *   action=list_periode_public                        -> daftar periode (semua orang)
 *   action=list_divisi_public&periode=                -> daftar divisi utk periode itu (semua orang, buat dropdown login)
 *   action=public_data&periode=                        -> ikhtisar GABUNGAN (semua orang, read-only)
 *   action=admin_ikhtisar&password=&periode=&divisi=   -> ikhtisar HANYA kegiatan divisi itu (admin)
 *   action=admin_data&password=&periode=&divisi=       -> data 1 sheet (roster+kegiatan BOD, atau kegiatan divisi)
 * --------------------------------------------------------- */
function doGet(e) {
  const action = e.parameter && e.parameter.action;
  if (!action) {
    return serveAdminPage_();
  }

  const ss = getSS_();

  try {
    if (action === 'list_periode_public') {
      return jsonOut_({ ok: true, data: listPeriode_(ss) });
    }

    if (action === 'list_divisi_public') {
      const periode = e.parameter.periode || listPeriode_(ss)[0];
      if (!periode) return jsonOut_({ ok: true, data: [BOD_LABEL] });
      const divisi = listDivisi_(ss, periode);
      return jsonOut_({ ok: true, data: divisi.length ? divisi : [BOD_LABEL] });
    }

    if (action === 'public_data') {
      const data = buildIkhtisarGabungan_(ss, e.parameter.periode);
      if (!data) return jsonOut_({ ok: false, error: 'Periode tidak ditemukan' });
      return jsonOut_({ ok: true, data: data });
    }

    if (action === 'admin_ikhtisar') {
      if (!checkDivisiAuth_(e.parameter.divisi, e.parameter.password)) return jsonOut_({ ok: false, error: 'Password salah' });
      const sheet = getDivisiSheet_(ss, e.parameter.periode, e.parameter.divisi);
      if (!sheet) return jsonOut_({ ok: false, error: 'Sheet tidak ditemukan' });
      return jsonOut_({ ok: true, data: buildIkhtisarDivisi_(sheet) });
    }

    if (action === 'admin_data') {
      if (!checkDivisiAuth_(e.parameter.divisi, e.parameter.password)) return jsonOut_({ ok: false, error: 'Password salah' });
      const sheet = getDivisiSheet_(ss, e.parameter.periode, e.parameter.divisi);
      if (!sheet) return jsonOut_({ ok: false, error: 'Sheet tidak ditemukan' });
      return jsonOut_({ ok: true, data: buildDashboardData_(sheet) });
    }

    return jsonOut_({ ok: false, error: 'Action tidak dikenali' });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

/* Halaman dashboard yang dilayani langsung oleh Apps Script (URL /exec
 * tanpa parameter). File HTML-nya bernama "Admin" (Admin.html) di project
 * ini. API_URL diisi otomatis dengan URL web app ini sendiri. */
function serveAdminPage_() {
  const tpl = HtmlService.createTemplateFromFile('Admin');
  tpl.apiUrl = ScriptApp.getService().getUrl();
  return tpl.evaluate()
    .setTitle('Dashboard Keaktifan — KSPM ESTOC')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ---------------------------------------------------------
 * doPost
 * --------------------------------------------------------- */
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ ok: false, error: 'Body tidak valid JSON' });
  }

  if (body.action === 'login') {
    return jsonOut_({ ok: checkDivisiAuth_(body.divisi, body.password) });
  }
  if (!checkDivisiAuth_(body.divisi, body.password)) {
    return jsonOut_({ ok: false, error: 'Password salah' });
  }

  const ss = getSS_();

  try {
    if (body.action === 'buat_periode') {
      if (!isBodLabel_(body.divisi)) return jsonOut_({ ok: false, error: 'Cuma BOD yang bisa bikin periode baru' });
      return buatPeriodeBaru_(ss, body);
    }

    // action kelola anggota/roster & tambah-divisi cuma boleh dari BOD
    if (['tambah_anggota', 'edit_anggota', 'hapus_anggota'].indexOf(body.action) !== -1 && !isBodLabel_(body.divisi)) {
      return jsonOut_({ ok: false, error: 'Cuma BOD yang bisa mengelola anggota' });
    }

    const sheet = getDivisiSheet_(ss, body.periode, body.divisi);
    if (!sheet) return jsonOut_({ ok: false, error: 'Sheet tidak ditemukan' });

    switch (body.action) {
      case 'tambah_kegiatan': return tambahKegiatan_(sheet, body);
      case 'edit_kegiatan': return editKegiatan_(sheet, body);
      case 'hapus_kegiatan': return hapusKegiatan_(sheet, body);
      case 'update_kehadiran': return updateKehadiran_(sheet, body);
      case 'tambah_anggota': return tambahAnggotaBod_(ss, sheet, body);
      case 'edit_anggota': return editAnggotaBod_(ss, sheet, body);
      case 'hapus_anggota': return hapusAnggotaBod_(ss, sheet, body);
      default: return jsonOut_({ ok: false, error: 'Action tidak dikenali' });
    }
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

/* ===========================================================
 * BACA STRUKTUR & DATA 1 SHEET (dipakai utk BOD maupun sheet divisi —
 * mekanismenya identik, cuma beda isi grupnya)
 * =========================================================== */
function readStrukturAnggota_(values) {
  const divisions = [];
  let current = null;

  for (let r = FIRST_DATA_ROW - 1; r < values.length; r++) {
    const colA = values[r][0];
    const colB = values[r][1];
    const isKosong = (colA === '' || colA === null) && (colB === '' || colB === null);
    if (isKosong) continue;

    const adalahAnggota = (typeof colA === 'number') && colB;
    if (adalahAnggota) {
      if (!current) { current = { nama: '(Tanpa Grup)', anggota: [] }; divisions.push(current); }
      current.anggota.push({ no: colA, nama: colB, row: r + 1 });
    } else if (colA) {
      current = { nama: String(colA), anggota: [] };
      divisions.push(current);
    }
  }
  return divisions;
}

function readKegiatanList_(values, displayValues) {
  const lastCol = values[HEADER_ROW_NAMA - 1].length;
  const list = [];
  for (let c = FIRST_ACTIVITY_COL - 1; c < lastCol; c++) {
    const nama = values[HEADER_ROW_NAMA - 1][c];
    if (!nama) continue;
    const tanggalDisplay = displayValues[HEADER_ROW_TANGGAL - 1][c];
    const tanggalVal = values[HEADER_ROW_TANGGAL - 1][c];
    const tanggalIso = (tanggalVal instanceof Date) ? Utilities.formatDate(tanggalVal, TIMEZONE, 'yyyy-MM-dd') : '';
    list.push({ col: c + 1, nama: String(nama), tanggal: tanggalDisplay, tanggal_iso: tanggalIso });
  }
  return list;
}

function buildDashboardData_(sheet) {
  const values = sheet.getDataRange().getValues();
  const displayValues = sheet.getDataRange().getDisplayValues();

  const kegiatanList = readKegiatanList_(values, displayValues);
  const divisions = readStrukturAnggota_(values);

  divisions.forEach(div => {
    div.anggota.forEach(a => {
      const rowIdx = a.row - 1;
      const kehadiran = {};
      let h = 0, al = 0, i = 0, b = 0;
      kegiatanList.forEach(k => {
        const status = String(values[rowIdx][k.col - 1] || '').trim().toUpperCase();
        kehadiran[k.col] = status;
        if (status === 'H') h++;
        else if (status === 'A') al++;
        else if (status === 'I') i++;
        else if (status === 'B') b++;
      });
      const totalTertandai = h + al + i;
      a.kehadiran = kehadiran;
      a.stats = {
        H: h, A: al, I: i, B: b,
        total_kegiatan: kegiatanList.length,
        total_tertandai: totalTertandai,
        persen_hadir: totalTertandai > 0 ? Math.round((h / totalTertandai) * 1000) / 10 : null
      };
    });
  });

  return { kegiatan: kegiatanList, divisi: divisions };
}

/* Versi ringan (tanpa stats) khusus operasi tambah/edit/hapus/arsip anggota. */
function getFullStructure_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return [];
  const values = sheet.getRange(FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, 2).getValues();

  const divisions = [];
  let current = null;
  values.forEach((rowVals, idx) => {
    const row = FIRST_DATA_ROW + idx;
    const colA = rowVals[0], colB = rowVals[1];
    const kosong = (colA === '' || colA === null) && (colB === '' || colB === null);
    if (kosong) return;

    const adalahAnggota = (typeof colA === 'number') && colB;
    if (adalahAnggota) {
      if (!current) { current = { nama: '(Tanpa Grup)', headerRow: null, anggota: [] }; divisions.push(current); }
      current.anggota.push({ no: colA, nama: colB, row });
    } else if (colA) {
      current = { nama: String(colA), headerRow: row, anggota: [] };
      divisions.push(current);
    }
  });
  return divisions;
}

function ensureRowCount_(sheet, rowNum) {
  if (rowNum > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), rowNum - sheet.getMaxRows());
  }
}

function insertAnggotaDiDivisi_(sheet, divisiNama, existingRowValues, namaBaru) {
  const structure = getFullStructure_(sheet);
  const div = structure.find(d => d.nama === divisiNama);
  let insertRow;

  if (div) {
    insertRow = div.anggota.length ? div.anggota[div.anggota.length - 1].row + 1 : div.headerRow + 1;
    sheet.insertRowBefore(insertRow);
  } else {
    const headerRow = sheet.getLastRow() + 1;
    insertRow = headerRow + 1;
    ensureRowCount_(sheet, insertRow);
    sheet.getRange(headerRow, 1).setValue(divisiNama);
  }

  if (existingRowValues) {
    existingRowValues[1] = namaBaru;
    sheet.getRange(insertRow, 1, 1, existingRowValues.length).setValues([existingRowValues]);
  } else {
    sheet.getRange(insertRow, 2).setValue(namaBaru);
  }
  renumberSemua_(sheet);
}

function renumberSemua_(sheet) {
  const structure = getFullStructure_(sheet);
  let n = 1;
  structure.forEach(div => {
    div.anggota.forEach(a => {
      sheet.getRange(a.row, 1).setValue(n);
      n++;
    });
  });
}

/* Hapus permanen 1 anggota (dicari berdasarkan nama) dari grup tertentu di sheet. */
function hapusAnggotaByNama_(sheet, groupNama, nama) {
  const structure = getFullStructure_(sheet);
  const group = structure.find(d => d.nama === groupNama);
  if (!group) return;
  const target = group.anggota.find(a => a.nama === nama);
  if (!target) return;
  sheet.deleteRow(target.row);
  renumberSemua_(sheet);
}

/* ===========================================================
 * PERIODE (BOD only)
 * =========================================================== */
function buatPeriodeBaru_(ss, body) {
  const namaPeriode = String(body.nama_periode || '').trim();
  const sumberPeriode = String(body.sumber_periode || '').trim();
  if (!namaPeriode) return jsonOut_({ ok: false, error: 'Nama periode wajib diisi' });
  if (ss.getSheetByName(sheetName_(namaPeriode, BOD_LABEL))) return jsonOut_({ ok: false, error: 'Periode dengan nama itu sudah ada' });

  const bodSumber = getDivisiSheet_(ss, sumberPeriode, BOD_LABEL);
  if (!bodSumber) return jsonOut_({ ok: false, error: 'Periode sumber tidak ditemukan' });

  const bodBaru = ss.insertSheet(sheetName_(namaPeriode, BOD_LABEL));
  bodBaru.getRange(1, 3).setValue('Nama Kegiatan');
  bodBaru.getRange(HEADER_ROW_NAMA, 1).setValue('No');
  bodBaru.getRange(HEADER_ROW_NAMA, 2).setValue('Nama');
  bodBaru.getRange(HEADER_ROW_NAMA, 1, 2, 1).merge();
  bodBaru.getRange(HEADER_ROW_NAMA, 2, 2, 1).merge();

  const lastRowSumber = bodSumber.getLastRow();
  if (lastRowSumber >= FIRST_DATA_ROW) {
    const dataAB = bodSumber.getRange(FIRST_DATA_ROW, 1, lastRowSumber - FIRST_DATA_ROW + 1, 2).getValues();
    ensureRowCount_(bodBaru, FIRST_DATA_ROW + dataAB.length - 1);
    bodBaru.getRange(FIRST_DATA_ROW, 1, dataAB.length, 2).setValues(dataAB);
  }

  // bikinkan juga sheet kosong utk tiap divisi (selain grup BOD), isi SELURUH roster
  // (semua grup) — tanpa kegiatan & kehadiran
  const struktur = getFullStructure_(bodBaru);
  struktur.forEach(d => {
    if (isBodLabel_(d.nama)) return;
    pastikanSheetDivisiSiap_(ss, namaPeriode, d.nama);
  });

  return jsonOut_({ ok: true, periode: namaPeriode });
}

/* ===========================================================
 * KEGIATAN & KEHADIRAN (1 sheet — BOD atau divisi, mekanismenya sama)
 * =========================================================== */
function tambahKegiatan_(sheet, body) {
  const namaKegiatan = String(body.nama_kegiatan || '').trim();
  if (!namaKegiatan) return jsonOut_({ ok: false, error: 'Nama kegiatan wajib diisi' });
  if (!body.tanggal) return jsonOut_({ ok: false, error: 'Tanggal wajib diisi' });

  const values = sheet.getDataRange().getValues();
  const headerRow = values[HEADER_ROW_NAMA - 1];
  let lastCol = FIRST_ACTIVITY_COL - 1;
  for (let c = FIRST_ACTIVITY_COL - 1; c < headerRow.length; c++) {
    if (headerRow[c]) lastCol = c + 1;
  }
  const newCol = Math.max(lastCol + 1, FIRST_ACTIVITY_COL);

  sheet.getRange(HEADER_ROW_NAMA, newCol).setValue(namaKegiatan);
  const tglDate = new Date(body.tanggal + 'T00:00:00');
  sheet.getRange(HEADER_ROW_TANGGAL, newCol).setValue(tglDate).setNumberFormat('dd/MM/yyyy');

  return jsonOut_({ ok: true, col: newCol });
}

function editKegiatan_(sheet, body) {
  const col = Number(body.col);
  const namaKegiatan = String(body.nama_kegiatan || '').trim();
  if (!col || col < FIRST_ACTIVITY_COL) return jsonOut_({ ok: false, error: 'Kolom kegiatan tidak valid' });
  if (!namaKegiatan) return jsonOut_({ ok: false, error: 'Nama kegiatan wajib diisi' });
  if (!body.tanggal) return jsonOut_({ ok: false, error: 'Tanggal wajib diisi' });

  sheet.getRange(HEADER_ROW_NAMA, col).setValue(namaKegiatan);
  const tglDate = new Date(body.tanggal + 'T00:00:00');
  sheet.getRange(HEADER_ROW_TANGGAL, col).setValue(tglDate).setNumberFormat('dd/MM/yyyy');

  return jsonOut_({ ok: true });
}

function hapusKegiatan_(sheet, body) {
  const col = Number(body.col);
  if (!col || col < FIRST_ACTIVITY_COL) return jsonOut_({ ok: false, error: 'Kolom kegiatan tidak valid' });
  sheet.deleteColumn(col);
  return jsonOut_({ ok: true });
}

function updateKehadiran_(sheet, body) {
  const row = Number(body.row);
  const col = Number(body.col);
  const status = String(body.status || '').trim().toUpperCase();

  if (!row || !col || row < FIRST_DATA_ROW || col < FIRST_ACTIVITY_COL) {
    return jsonOut_({ ok: false, error: 'Posisi sel tidak valid' });
  }
  if (STATUS_VALID.indexOf(status) === -1) {
    return jsonOut_({ ok: false, error: 'Status harus H, A, I, B, atau kosong' });
  }

  sheet.getRange(row, col).setValue(status);
  return jsonOut_({ ok: true });
}

/* ===========================================================
 * KELOLA ANGGOTA — BOD ONLY. `sheet` di sini SELALU sheet BOD (roster),
 * dan tiap perubahan di-propagate otomatis ke SEMUA sheet divisi lain
 * (karena roster-nya memang harus sama di semua sheet).
 * =========================================================== */

/* Cari 1 anggota berdasarkan NAMA di sembarang grup dalam 1 sheet.
 * Dipakai buat operasi lintas-sheet karena nomor baris beda-beda tiap sheet. */
function cariAnggotaDiSheetByNama_(sheet, nama) {
  const struktur = getFullStructure_(sheet);
  for (const g of struktur) {
    const found = g.anggota.find(a => a.nama === nama);
    if (found) return { groupNama: g.nama, row: found.row };
  }
  return null;
}

/* Pindahkan posisi 1 anggota (dicari by nama) ke grup lain DI SHEET YANG
 * SAMA — data kehadiran di sheet itu tetap dipertahankan apa adanya,
 * cuma posisinya (grup/divisinya) yang berubah. */
function pindahkanPosisiAnggota_(sheet, namaLama, divisiBaru, namaBaru) {
  const info = cariAnggotaDiSheetByNama_(sheet, namaLama);
  if (!info) return;
  const lastCol = Math.max(sheet.getLastColumn(), FIRST_ACTIVITY_COL - 1);
  const rowValues = sheet.getRange(info.row, 1, 1, lastCol).getValues()[0];
  sheet.deleteRow(info.row);
  insertAnggotaDiDivisi_(sheet, divisiBaru, rowValues, namaBaru);
}

/* Isi sheet divisi yang BARU DIBUAT dengan salinan roster lengkap dari
 * BOD (semua grup, semua anggota) — karena tiap sheet wajib punya roster
 * yang sama, biar siapa saja bisa jadi peserta kegiatan divisi manapun. */
function syncSheetBaruDenganRoster_(ss, periode, sheetBaru) {
  const bodSheet = getDivisiSheet_(ss, periode, BOD_LABEL);
  if (!bodSheet) return;
  const struktur = getFullStructure_(bodSheet);
  struktur.forEach(d => d.anggota.forEach(a => insertAnggotaDiDivisi_(sheetBaru, d.nama, null, a.nama)));
}

/* Pastikan sheet divisi `divisi` ada; kalau baru dibuat, langsung isi
 * roster lengkap (bukan cuma nama satu anggota) supaya konsisten. */
function pastikanSheetDivisiSiap_(ss, periode, divisi) {
  if (isBodLabel_(divisi)) return null;
  const sudahAda = !!ss.getSheetByName(sheetName_(periode, divisi));
  const sheet = getOrCreateDivisiSheet_(ss, periode, divisi);
  if (!sudahAda) syncSheetBaruDenganRoster_(ss, periode, sheet);
  return sheet;
}

function tambahAnggotaBod_(ss, bodSheet, body) {
  const nama = String(body.nama || '').trim();
  const divisi = String(body.divisi_anggota || body.divisi_baru || '').trim() || '(Tanpa Grup)';
  if (!nama) return jsonOut_({ ok: false, error: 'Nama anggota wajib diisi' });

  insertAnggotaDiDivisi_(bodSheet, divisi, null, nama);

  // tambahkan juga ke SEMUA sheet divisi lain yang sudah ada (roster harus sama di mana-mana)
  listDivisi_(ss, body.periode).filter(d => !isBodLabel_(d)).forEach(divName => {
    const sheet = getDivisiSheet_(ss, body.periode, divName);
    if (sheet) insertAnggotaDiDivisi_(sheet, divName, null, nama);
  });

  // kalau divisi yang dipilih ternyata baru (belum ada sheet-nya), buat + sync roster
  // (otomatis ikut memasukkan anggota baru ini juga, karena sudah ditambahkan ke BOD di atas)
  pastikanSheetDivisiSiap_(ss, body.periode, divisi);

  return jsonOut_({ ok: true });
}

function editAnggotaBod_(ss, bodSheet, body) {
  const row = Number(body.row);
  const namaBaru = String(body.nama || '').trim();
  const divisiBaru = body.divisi_anggota ? String(body.divisi_anggota).trim() : null;
  if (!row) return jsonOut_({ ok: false, error: 'Baris anggota tidak valid' });
  if (!namaBaru) return jsonOut_({ ok: false, error: 'Nama wajib diisi' });

  const structure = getFullStructure_(bodSheet);
  let divisiSekarang = null, namaLama = null;
  structure.forEach(d => {
    const found = d.anggota.find(a => a.row === row);
    if (found) { divisiSekarang = d.nama; namaLama = found.nama; }
  });
  if (divisiSekarang === null) return jsonOut_({ ok: false, error: 'Anggota tidak ditemukan' });

  const pindahDivisi = divisiBaru && divisiBaru !== divisiSekarang;
  const daftarSheetLain = listDivisi_(ss, body.periode).filter(d => !isBodLabel_(d));

  if (!pindahDivisi) {
    bodSheet.getRange(row, 2).setValue(namaBaru);
    if (namaBaru !== namaLama) {
      // nama berubah tapi divisi tetap -> ikut ubah nama di SEMUA sheet lain juga
      daftarSheetLain.forEach(divName => {
        const sheet = getDivisiSheet_(ss, body.periode, divName);
        if (!sheet) return;
        const info = cariAnggotaDiSheetByNama_(sheet, namaLama);
        if (info) sheet.getRange(info.row, 2).setValue(namaBaru);
      });
    }
    return jsonOut_({ ok: true });
  }

  // pindah divisi: reposisikan di BOD (bawa data lengkap barisnya)
  const lastCol = Math.max(bodSheet.getLastColumn(), FIRST_ACTIVITY_COL - 1);
  const rowValues = bodSheet.getRange(row, 1, 1, lastCol).getValues()[0];
  bodSheet.deleteRow(row);
  insertAnggotaDiDivisi_(bodSheet, divisiBaru, rowValues, namaBaru);

  // reposisikan juga di SEMUA sheet divisi lain — data kehadiran di tiap sheet
  // itu TIDAK ikut berubah, cuma grup/label divisinya yang disamakan ulang
  daftarSheetLain.forEach(divName => {
    const sheet = getDivisiSheet_(ss, body.periode, divName);
    if (sheet) pindahkanPosisiAnggota_(sheet, namaLama, divisiBaru, namaBaru);
  });

  // kalau divisi tujuan baru (belum ada sheet-nya), buat + sync roster
  pastikanSheetDivisiSiap_(ss, body.periode, divisiBaru);

  return jsonOut_({ ok: true });
}

function hapusAnggotaBod_(ss, bodSheet, body) {
  const row = Number(body.row);
  if (!row || row < FIRST_DATA_ROW) return jsonOut_({ ok: false, error: 'Baris anggota tidak valid' });

  const structure = getFullStructure_(bodSheet);
  let nama = null;
  structure.forEach(d => {
    const found = d.anggota.find(a => a.row === row);
    if (found) nama = found.nama;
  });
  if (nama === null) return jsonOut_({ ok: false, error: 'Anggota tidak ditemukan' });

  bodSheet.deleteRow(row);
  renumberSemua_(bodSheet);

  // hapus juga dari SEMUA sheet divisi lain
  listDivisi_(ss, body.periode).filter(d => !isBodLabel_(d)).forEach(divName => {
    const sheet = getDivisiSheet_(ss, body.periode, divName);
    if (!sheet) return;
    const info = cariAnggotaDiSheetByNama_(sheet, nama);
    if (info) { sheet.deleteRow(info.row); renumberSemua_(sheet); }
  });

  return jsonOut_({ ok: true });
}

/* ===========================================================
 * IKHTISAR GABUNGAN (BOD + semua divisi) — dipakai tampilan publik
 * "Anggota" (tanpa login) MAUPUN tab "Ikhtisar" di halaman admin siapa pun.
 * Anggota dicocokkan lintas-sheet berdasarkan NAMA (asumsi nama unik).
 * =========================================================== */
function buildIkhtisarGabungan_(ss, periode) {
  if (!periode) return null;
  const bodSheet = getDivisiSheet_(ss, periode, BOD_LABEL);
  if (!bodSheet) return null;

  const bodFull = buildDashboardData_(bodSheet);
  const gabungan = {}; // nama -> { H,A,I,B, divisiSekarang }

  bodFull.divisi.forEach(d => {
    d.anggota.forEach(a => {
      gabungan[a.nama] = { H: a.stats.H, A: a.stats.A, I: a.stats.I, B: a.stats.B, divisiSekarang: d.nama };
    });
  });

  const perKegiatan = bodFull.kegiatan.map(k => {
    let h = 0, a = 0, i = 0, b = 0;
    bodFull.divisi.forEach(d => d.anggota.forEach(ang => {
      const st = ang.kehadiran[k.col] || '';
      if (st === 'H') h++; else if (st === 'A') a++; else if (st === 'I') i++; else if (st === 'B') b++;
    }));
    const total = h + a + i;
    return { nama: `BOD: ${k.nama}`, tanggal: k.tanggal, H: h, A: a, I: i, B: b,
      persen_hadir: total > 0 ? Math.round((h / total) * 1000) / 10 : null };
  });

  const namaDivisiDinamis = bodFull.divisi.map(d => d.nama).filter(n => !isBodLabel_(n));
  const perDivisi = [];
  let totalKegiatan = bodFull.kegiatan.length;

  namaDivisiDinamis.forEach(divName => {
    const dSheet = getDivisiSheet_(ss, periode, divName);
    if (!dSheet) { perDivisi.push({ nama: divName, jumlah_anggota: 0, persen_hadir_rata2: null }); return; }
    const full = buildDashboardData_(dSheet);
    totalKegiatan += full.kegiatan.length;

    full.kegiatan.forEach(k => {
      let h = 0, a = 0, i = 0, b = 0;
      full.divisi.forEach(g => g.anggota.forEach(ang => {
        const st = ang.kehadiran[k.col] || '';
        if (st === 'H') h++; else if (st === 'A') a++; else if (st === 'I') i++; else if (st === 'B') b++;

        if (!gabungan[ang.nama]) gabungan[ang.nama] = { H: 0, A: 0, I: 0, B: 0, divisiSekarang: divName };
        if (st === 'H') gabungan[ang.nama].H++;
        else if (st === 'A') gabungan[ang.nama].A++;
        else if (st === 'I') gabungan[ang.nama].I++;
        else if (st === 'B') gabungan[ang.nama].B++;
      }));
      const total = h + a + i;
      perKegiatan.push({ nama: `${divName}: ${k.nama}`, tanggal: k.tanggal, H: h, A: a, I: i, B: b,
        persen_hadir: total > 0 ? Math.round((h / total) * 1000) / 10 : null });
    });

    const grupAktif = full.divisi.find(g => g.nama === divName);
    const anggotaAktif = grupAktif ? grupAktif.anggota : [];
    const punyaData = anggotaAktif.filter(a => a.stats.persen_hadir !== null);
    const rata2 = punyaData.length
      ? Math.round((punyaData.reduce((s, a) => s + a.stats.persen_hadir, 0) / punyaData.length) * 10) / 10
      : null;
    perDivisi.push({ nama: divName, jumlah_anggota: anggotaAktif.length, persen_hadir_rata2: rata2 });
  });

  let totH = 0, totA = 0, totI = 0, totB = 0;
  Object.keys(gabungan).forEach(nama => {
    const g = gabungan[nama];
    const total = g.H + g.A + g.I;
    g.persen_hadir = total > 0 ? Math.round((g.H / total) * 1000) / 10 : null;
    totH += g.H; totA += g.A; totI += g.I; totB += g.B;
  });
  const totalTertandaiSemua = totH + totA + totI;

  const divisiMap = {};
  Object.keys(gabungan).forEach(nama => {
    const g = gabungan[nama];
    if (!divisiMap[g.divisiSekarang]) divisiMap[g.divisiSekarang] = [];
    divisiMap[g.divisiSekarang].push({ nama, stats: { H: g.H, A: g.A, I: g.I, B: g.B, persen_hadir: g.persen_hadir } });
  });
  const divisiOut = Object.keys(divisiMap).map(nama => ({ nama, anggota: divisiMap[nama] }));

  return {
    total_kegiatan: totalKegiatan,
    total_anggota: Object.keys(gabungan).length,
    keseluruhan: { H: totH, A: totA, I: totI, B: totB,
      persen_hadir: totalTertandaiSemua > 0 ? Math.round((totH / totalTertandaiSemua) * 1000) / 10 : null },
    per_kegiatan: perKegiatan,
    per_divisi: perDivisi,
    divisi: divisiOut
  };
}

/* ===========================================================
 * IKHTISAR 1 DIVISI — dipakai tab "Ikhtisar" di halaman admin.
 * Hanya menghitung kegiatan di sheet divisi itu sendiri. Bentuk
 * outputnya sama dengan buildIkhtisarGabungan_ supaya frontend
 * memakai fungsi render yang sama.
 * =========================================================== */
function buildIkhtisarDivisi_(sheet) {
  const full = buildDashboardData_(sheet);
  const perKegiatan = full.kegiatan.map(k => {
    let h = 0, a = 0, i = 0, b = 0;
    full.divisi.forEach(d => d.anggota.forEach(ang => {
      const st = ang.kehadiran[k.col] || '';
      if (st === 'H') h++; else if (st === 'A') a++; else if (st === 'I') i++; else if (st === 'B') b++;
    }));
    const total = h + a + i;
    return { nama: k.nama, tanggal: k.tanggal, H: h, A: a, I: i, B: b,
      persen_hadir: total > 0 ? Math.round((h / total) * 1000) / 10 : null };
  });

  let totH = 0, totA = 0, totI = 0, totB = 0, totalAnggota = 0;
  const perDivisi = [];
  const divisiOut = full.divisi.map(d => {
    const punyaData = d.anggota.filter(a => a.stats.persen_hadir !== null);
    const rata2 = punyaData.length
      ? Math.round((punyaData.reduce((s, a) => s + a.stats.persen_hadir, 0) / punyaData.length) * 10) / 10
      : null;
    perDivisi.push({ nama: d.nama, jumlah_anggota: d.anggota.length, persen_hadir_rata2: rata2 });
    totalAnggota += d.anggota.length;
    d.anggota.forEach(a => { totH += a.stats.H; totA += a.stats.A; totI += a.stats.I; totB += a.stats.B; });
    return { nama: d.nama, anggota: d.anggota.map(a => ({ nama: a.nama,
      stats: { H: a.stats.H, A: a.stats.A, I: a.stats.I, B: a.stats.B, persen_hadir: a.stats.persen_hadir } })) };
  });
  const tot = totH + totA + totI;

  return {
    total_kegiatan: full.kegiatan.length,
    total_anggota: totalAnggota,
    keseluruhan: { H: totH, A: totA, I: totI, B: totB,
      persen_hadir: tot > 0 ? Math.round((totH / tot) * 1000) / 10 : null },
    per_kegiatan: perKegiatan,
    per_divisi: perDivisi,
    divisi: divisiOut
  };
}
