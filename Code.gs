// Website link afiliasi Shopee: Google Apps Script + Google Sheets
// Sheet "Produk": baris 1 = header. Kolom dicari dari NAMA header (urutan boleh berubah):
// nama, link, gambar, kategori, klik, badge (opsional: NEW/HOT/RECOMMENDED), subjudul (opsional)

const SHEET = 'Produk';
const NAMA_SITUS = 'Rekomendasi Belanja'; // ganti sesuai keinginan
const BADGE_VALID = ['NEW', 'HOT', 'RECOMMENDED'];

function doGet(e) {
  // Mode API untuk halaman di GitHub Pages: ?aksi=produk (data) dan ?aksi=klik&id=NOMOR_BARIS
  const aksi = e && e.parameter && e.parameter.aksi;
  if (aksi === 'produk') return json({ nama: NAMA_SITUS, produk: getProduk() });
  if (aksi === 'klik') { catatKlik(e.parameter.id); return json({ ok: true }); }

  const t = HtmlService.createTemplateFromFile('index');
  t.namaSitus = NAMA_SITUS;
  return t.evaluate()
    .setTitle(NAMA_SITUS)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // izinkan ditampilkan di iframe situs lain
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Jalankan SEKALI dari editor untuk membuat sheet + contoh data
function siapkan() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) sh = ss.insertSheet(SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['nama', 'link', 'gambar', 'kategori', 'klik', 'badge', 'subjudul']);
    sh.appendRow(['Contoh produk', 'https://s.shopee.co.id/contoh', '', 'Umum', 0, 'NEW', 'Sub judul contoh']);
    sh.setFrozenRows(1);
  }
}

// Nomor kolom (mulai dari 1) berdasarkan nama header; pakai nilai cadangan kalau header tidak ditemukan
function kolom(sh, nama, cadangan) {
  const header = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  const i = header.map(h => String(h).trim().toLowerCase()).indexOf(nama);
  return i >= 0 ? i + 1 : cadangan;
}

// Dipanggil dari halaman: daftar produk (terbaru di atas)
function getProduk() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  const rows = sh.getDataRange().getValues();
  const header = rows.shift().map(h => String(h).trim().toLowerCase());
  const idx = (nama, cadangan) => { const i = header.indexOf(nama); return i >= 0 ? i : cadangan; };
  const cNama = idx('nama', 0), cLink = idx('link', 1), cGambar = idx('gambar', 2),
        cKategori = idx('kategori', 3), cBadge = idx('badge', 5), cSub = idx('subjudul', 6);
  return rows
    .map((r, i) => {
      const badge = String(r[cBadge] || '').trim().toUpperCase();
      return {
        id: i + 2, // nomor baris di sheet
        nama: String(r[cNama] || ''),
        link: String(r[cLink] || ''),
        gambar: String(r[cGambar] || ''),
        kategori: String(r[cKategori] || ''),
        subjudul: String(r[cSub] || '').trim(),
        badge: BADGE_VALID.indexOf(badge) >= 0 ? badge : ''
      };
    })
    .filter(p => p.nama && /^https?:\/\//i.test(p.link))
    .reverse();
}

// Dipanggil saat tombol beli diklik. Kolom "klik" dicari dari header (cadangan: kolom F)
function catatKlik(id) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  id = Number(id);
  if (!Number.isInteger(id) || id < 2 || id > sh.getLastRow()) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sel = sh.getRange(id, kolom(sh, 'klik', 6));
    sel.setValue((Number(sel.getValue()) || 0) + 1);
  } finally {
    lock.releaseLock();
  }
}

// Pasang dropdown NEW / HOT / RECOMMENDED di kolom "badge". Kosong tetap boleh.
// Kalau header "badge" belum ada, dibuat di kolom kosong paling kanan.
function pasangDropdownBadge() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  let kb = kolom(sh, 'badge', 0);
  if (!kb) { kb = sh.getLastColumn() + 1; sh.getRange(1, kb).setValue('badge'); }
  const aturan = SpreadsheetApp.newDataValidation()
    .requireValueInList(BADGE_VALID, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(2, kb, Math.max(sh.getMaxRows() - 1, 1), 1).setDataValidation(aturan);
}
