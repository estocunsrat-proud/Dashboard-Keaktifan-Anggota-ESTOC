/**
 * ==========================================================
 * DASHBOARD KEAKTIFAN ANGGOTA — KSPM ESTOC UNSRAT
 * Backend Google Apps Script (API-only, deploy sebagai Web App)
 * ==========================================================
 * Project ini BERDIRI SENDIRI, terpisah dari project Reservasi OA,
 * tapi memakai pola yang sama: Apps Script cuma jadi API/backend,
 * sedangkan HALAMAN ADMIN-nya (index.html) di-hosting terpisah di
 * Netlify lewat repo GitHub — persis seperti mockup nasabah Reservasi OA.
 * Bedanya, project ini TIDAK ADA bagian nasabah — cuma satu halaman admin.
 *
 * Cara pasang:
 * 1. Buka spreadsheet "Keaktifan_Anggota" > Extensions > Apps Script.
 * 2. Tempel file ini sebagai Code.gs (Admin.html TIDAK dipakai lagi di sini).
 * 3. Project Settings > Script Properties > tambah key ADMIN_PASSWORD.
 * 4. PENTING: ganti nama tab sheet yang sekarang ("Sheet1") jadi "2026" —
 *    klik kanan tab di bawah > Rename. Ini jadi sheet periode pertama.
 * 5. Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 *    Salin URL Web App-nya — itu yang dipakai sebagai API_URL di index.html.
 * 6. index.html (di repo GitHub terpisah) di-deploy ke Netlify. Lihat README.md.
 *
 * ---------------------------------------------------------
 * KONSEP PERIODE
 * ---------------------------------------------------------
 * 1 periode (mis. tahun kepengurusan) = 1 sheet/tab tersendiri di
 * spreadsheet yang sama. Nama tab = nama periode (mis. "2026").
 * Saat admin bikin "Periode baru", script akan:
 *  - membuat sheet baru dengan nama periode itu
 *  - menyalin daftar anggota (No, Nama, Divisi) dari periode sumber
 *  - TIDAK menyalin kegiatan/kehadiran — periode baru selalu mulai bersih
 * Anggota yang dihapus/diedit di satu periode TIDAK memengaruhi periode lain,
 * karena tiap periode memang sheet yang benar-benar terpisah.
 *
 * ---------------------------------------------------------
 * STRUKTUR SETIAP SHEET PERIODE
 * ---------------------------------------------------------
 * Kolom A = No, Kolom B = Nama (berlaku utk baris anggota)
 * Baris DIVISI: kolom A berisi teks nama divisi, kolom B kosong.
 * Baris ANGGOTA: kolom A berisi nomor urut, kolom B berisi nama.
 * Mulai kolom C ke kanan = 1 kolom per kegiatan:
 *   - baris HEADER_ROW_NAMA (2)    -> nama kegiatan
 *   - baris HEADER_ROW_TANGGAL (3) -> tanggal kegiatan
 *   - baris FIRST_DATA_ROW (4) dst -> status kehadiran per anggota: "H" / "A" / "I" / kosong
 *     (H = Hadir, A = Alpa, I = Izin)
 */

const SPREADSHEET_ID = ''; // isi kalau standalone, kosongkan kalau bound ke spreadsheet
const TIMEZONE = 'Asia/Makassar'; // WITA

const HEADER_ROW_NAMA = 2;
const HEADER_ROW_TANGGAL = 3;
const FIRST_DATA_ROW = 4;
const FIRST_ACTIVITY_COL = 3; // kolom C

const STATUS_VALID = ['H', 'A', 'I', ''];

function getSS_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

/* ---------------------------------------------------------
 * AUTH ADMIN
 * --------------------------------------------------------- */
function checkAdminAuth_(password) {
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  return real && password && password === real;
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getPeriodeSheet_(ss, namaPeriode) {
  const sheet = ss.getSheetByName(String(namaPeriode || ''));
  return sheet || null;
}

/* ---------------------------------------------------------
 * doGet
 *   (tanpa action)                          -> pesan status API saja (bukan halaman)
 *   action=list_periode&password=           -> daftar nama periode (tab)
 *   action=admin_data&password=&periode=    -> data dashboard periode itu
 * --------------------------------------------------------- */
function doGet(e) {
  const action = e.parameter && e.parameter.action;
  if (!action) {
    return jsonOut_({ ok: true, message: 'API Dashboard Keaktifan Anggota — KSPM ESTOC. Halaman admin ada di Netlify, bukan di sini.' });
  }

  try {
    if (action === 'list_periode') {
      if (!checkAdminAuth_(e.parameter.password)) return jsonOut_({ ok: false, error: 'Password salah' });
      return jsonOut_({ ok: true, data: listPeriode_(getSS_()) });
    }

    if (action === 'admin_data') {
      if (!checkAdminAuth_(e.parameter.password)) return jsonOut_({ ok: false, error: 'Password salah' });
      const sheet = getPeriodeSheet_(getSS_(), e.parameter.periode);
      if (!sheet) return jsonOut_({ ok: false, error: 'Periode tidak ditemukan' });
      return jsonOut_({ ok: true, data: buildDashboardData_(sheet) });
    }

    return jsonOut_({ ok: false, error: 'Action tidak dikenali' });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

/* Nama-nama tab periode, terbaru duluan (tab yang dibuat belakangan = lebih baru) */
function listPeriode_(ss) {
  return ss.getSheets().map(s => s.getName()).reverse();
}

/* ---------------------------------------------------------
 * doPost — semua operasi TULIS data (semuanya wajib admin, kecuali login)
 * --------------------------------------------------------- */
function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut_({ ok: false, error: 'Body tidak valid JSON' });
  }

  if (body.action === 'login') {
    return jsonOut_({ ok: checkAdminAuth_(body.password) });
  }
  if (!checkAdminAuth_(body.password)) {
    return jsonOut_({ ok: false, error: 'Password salah' });
  }

  const ss = getSS_();

  try {
    if (body.action === 'buat_periode') {
      return buatPeriodeBaru_(ss, body);
    }

    // sisa action butuh sheet periode yang valid
    const sheet = getPeriodeSheet_(ss, body.periode);
    if (!sheet) return jsonOut_({ ok: false, error: 'Periode tidak ditemukan' });

    switch (body.action) {
      case 'tambah_kegiatan': return tambahKegiatan_(sheet, body);
      case 'edit_kegiatan': return editKegiatan_(sheet, body);
      case 'hapus_kegiatan': return hapusKegiatan_(sheet, body);
      case 'update_kehadiran': return updateKehadiran_(sheet, body);
      case 'tambah_anggota': return tambahAnggota_(sheet, body);
      case 'edit_anggota': return editAnggota_(sheet, body);
      case 'hapus_anggota': return hapusAnggota_(sheet, body);
      default: return jsonOut_({ ok: false, error: 'Action tidak dikenali' });
    }
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

/* ===========================================================
 * BACA STRUKTUR & DATA DASHBOARD
 * =========================================================== */

// Versi lengkap (dengan stats kehadiran) untuk ditampilkan di dashboard
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
      if (!current) { current = { nama: '(Tanpa Divisi)', anggota: [] }; divisions.push(current); }
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
      let h = 0, al = 0, i = 0;
      kegiatanList.forEach(k => {
        const status = String(values[rowIdx][k.col - 1] || '').trim().toUpperCase();
        kehadiran[k.col] = status;
        if (status === 'H') h++;
        else if (status === 'A') al++;
        else if (status === 'I') i++;
      });
      const totalTertandai = h + al + i;
      a.kehadiran = kehadiran;
      a.stats = {
        H: h, A: al, I: i,
        total_kegiatan: kegiatanList.length,
        total_tertandai: totalTertandai,
        persen_hadir: totalTertandai > 0 ? Math.round((h / totalTertandai) * 1000) / 10 : null
      };
    });
  });

  return { kegiatan: kegiatanList, divisi: divisions };
}

/* Versi ringan (tanpa stats) khusus untuk operasi tambah/edit/hapus anggota —
 * cukup baca kolom A:B saja, dan simpan row header tiap divisi supaya tahu
 * di mana harus sisip/hapus baris. */
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
      if (!current) { current = { nama: '(Tanpa Divisi)', headerRow: null, anggota: [] }; divisions.push(current); }
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

/* Sisip 1 baris anggota ke dalam divisi tertentu (bikin divisi baru kalau belum ada).
 * existingRowValues: kalau diisi (array 1 baris penuh, hasil capture sebelum dihapus),
 * dipakai untuk MEMINDAHKAN anggota (data kehadiran ikut terbawa). Kalau null,
 * berarti anggota baru (kehadiran kosong semua). */
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

/* Nomor ulang kolom "No" secara berurutan (1..N) untuk semua baris anggota,
 * baris divisi tidak disentuh. Dipanggil tiap kali ada tambah/pindah/hapus anggota. */
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

/* ===========================================================
 * PERIODE
 * =========================================================== */
function buatPeriodeBaru_(ss, body) {
  const namaPeriode = String(body.nama_periode || '').trim();
  const sumberPeriode = String(body.sumber_periode || '').trim();
  if (!namaPeriode) return jsonOut_({ ok: false, error: 'Nama periode wajib diisi' });
  if (ss.getSheetByName(namaPeriode)) return jsonOut_({ ok: false, error: 'Periode dengan nama itu sudah ada' });

  const sumberSheet = ss.getSheetByName(sumberPeriode);
  if (!sumberSheet) return jsonOut_({ ok: false, error: 'Periode sumber tidak ditemukan' });

  const newSheet = ss.insertSheet(namaPeriode);
  newSheet.getRange(1, 3).setValue('Nama Kegiatan');
  newSheet.getRange(HEADER_ROW_NAMA, 1).setValue('No');
  newSheet.getRange(HEADER_ROW_NAMA, 2).setValue('Nama');
  newSheet.getRange(HEADER_ROW_NAMA, 1, 2, 1).merge();
  newSheet.getRange(HEADER_ROW_NAMA, 2, 2, 1).merge();

  // salin daftar anggota (kolom A:B saja) dari periode sumber — TANPA kegiatan/kehadiran
  const lastRowSumber = sumberSheet.getLastRow();
  if (lastRowSumber >= FIRST_DATA_ROW) {
    const dataAB = sumberSheet.getRange(FIRST_DATA_ROW, 1, lastRowSumber - FIRST_DATA_ROW + 1, 2).getValues();
    ensureRowCount_(newSheet, FIRST_DATA_ROW + dataAB.length - 1);
    newSheet.getRange(FIRST_DATA_ROW, 1, dataAB.length, 2).setValues(dataAB);
  }

  return jsonOut_({ ok: true, periode: namaPeriode });
}

/* ===========================================================
 * KEGIATAN & KEHADIRAN
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

/* Menghapus 1 kolom kegiatan beserta seluruh data kehadiran di kolom itu.
 * Kolom-kolom kegiatan lain otomatis bergeser ke kiri (deleteColumn bawaan Sheets),
 * itu aman karena kolom dibaca ulang dari header tiap kali data dimuat. */
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
    return jsonOut_({ ok: false, error: 'Status harus H, A, I, atau kosong' });
  }

  sheet.getRange(row, col).setValue(status);
  return jsonOut_({ ok: true });
}

/* ===========================================================
 * KELOLA ANGGOTA (hanya memengaruhi sheet periode yang dikirim)
 * =========================================================== */
function tambahAnggota_(sheet, body) {
  const nama = String(body.nama || '').trim();
  const divisi = String(body.divisi || '').trim() || '(Tanpa Divisi)';
  if (!nama) return jsonOut_({ ok: false, error: 'Nama anggota wajib diisi' });

  insertAnggotaDiDivisi_(sheet, divisi, null, nama);
  return jsonOut_({ ok: true });
}

function editAnggota_(sheet, body) {
  const row = Number(body.row);
  const namaBaru = String(body.nama || '').trim();
  const divisiBaru = body.divisi ? String(body.divisi).trim() : null;
  if (!row) return jsonOut_({ ok: false, error: 'Baris anggota tidak valid' });
  if (!namaBaru) return jsonOut_({ ok: false, error: 'Nama wajib diisi' });

  const structure = getFullStructure_(sheet);
  let divisiSekarang = null;
  structure.forEach(d => { if (d.anggota.some(a => a.row === row)) divisiSekarang = d.nama; });
  if (divisiSekarang === null) return jsonOut_({ ok: false, error: 'Anggota tidak ditemukan' });

  if (!divisiBaru || divisiBaru === divisiSekarang) {
    sheet.getRange(row, 2).setValue(namaBaru);
    return jsonOut_({ ok: true });
  }

  // pindah divisi: bawa serta data kehadiran yang sudah ada di baris itu
  const lastCol = Math.max(sheet.getLastColumn(), FIRST_ACTIVITY_COL - 1);
  const rowValues = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  sheet.deleteRow(row);
  insertAnggotaDiDivisi_(sheet, divisiBaru, rowValues, namaBaru);
  return jsonOut_({ ok: true });
}

function hapusAnggota_(sheet, body) {
  const row = Number(body.row);
  if (!row || row < FIRST_DATA_ROW) return jsonOut_({ ok: false, error: 'Baris anggota tidak valid' });

  sheet.deleteRow(row);
  renumberSemua_(sheet);
  return jsonOut_({ ok: true });
}
