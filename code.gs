// Google Apps Script – penerima data Porsiku (seperti Google Form)
// Tempel di Extensions > Apps Script pada Google Sheet, lalu Deploy > Manage deployments > Edit > New version.
// Deploy sebagai Web app: Execute as "Me", Who has access "Anyone".
// Saat pertama dijalankan, izinkan akses Google Drive (dipakai untuk menyimpan foto bukti bayar).
const TOKEN = 'PORSIKU1234'; // samakan dengan CONFIG.TOKEN di script.js
const FOLDER = 'Bukti Bayar Porsiku';

const SHEETS = {
  pengguna: {
    nama: 'Data Pengguna',
    kolom: ['nama','jk','usia','bb','tb','aktivitas','alergi','imt','kategori','bmr','tee'],
    judul: ['Nama','Jenis kelamin','Usia','BB (kg)','TB (cm)','Aktivitas','Alergi','IMT','Kategori IMT','BMR (kkal)','Total energi (kkal)']
  },
  konsultasi: {
    nama: 'Konsultasi',
    kolom: ['kode','nama','jk','usia','bb','tb','aktivitas','alergi','imt','kategori','bmr','tee','menu','keluhan','_bukti'],
    judul: ['Kode','Nama','Jenis kelamin','Usia','BB (kg)','TB (cm)','Aktivitas','Alergi','IMT','Kategori IMT','BMR (kkal)','Total energi (kkal)','Menu dipilih','Keluhan','Bukti bayar']
  }
};

function balas(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const cfg = SHEETS[d.tipe];
    if (d.token !== TOKEN || !cfg) return balas({ ok: false, pesan: 'ditolak' });

    // simpan foto bukti bayar ke Drive, ambil link-nya
    let link = '';
    if (d.tipe === 'konsultasi' && d.bukti) {
      const b64 = String(d.bukti).replace(/^data:image\/\w+;base64,/, '');
      const it = DriveApp.getFoldersByName(FOLDER);
      const folder = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER);
      const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', (d.kode || 'bukti') + ' - ' + d.nama + '.jpg');
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      link = file.getUrl();
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(cfg.nama) || ss.insertSheet(cfg.nama);
      if (sh.getLastRow() === 0) {
        sh.appendRow(['Waktu', ...cfg.judul]);
        sh.getRange(1, 1, 1, cfg.judul.length + 1).setFontWeight('bold').setBackground('#dcebe8');
        sh.setFrozenRows(1);
      }
      sh.appendRow([new Date(), ...cfg.kolom.map(k => k === '_bukti' ? link : (d[k] === undefined ? '' : d[k]))]);
    } finally {
      lock.releaseLock();
    }
    return balas({ ok: true, url: link });
  } catch (err) {
    return balas({ ok: false, pesan: String(err) });
  }
}