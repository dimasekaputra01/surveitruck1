// Penerima data Survei Truk Gate Pelabuhan
// SECRET harus sama persis dengan SECRET di index.html
const SECRET = 'truk-gate-9f3k2m7x';
const SHEET_NAME = 'Data';
const FOLDER_NAME = 'Foto Survei Truk';

function doGet() {
  return ContentService.createTextOutput('Skrip survei aktif');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.secret !== SECRET) return out({ ok: false, error: 'Kode rahasia salah' });
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (d.action === 'row') return out(saveRow(sh, d));
    if (d.action === 'foto') return out(saveFoto(sh, d));
    return out({ ok: false, error: 'Aksi tidak dikenal' });
  } catch (err) {
    return out({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

// Kolom terakhir di Sheet selalu "ID" (penanda unik tiap truk)
function findRow(sh, id) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const col = sh.getLastColumn();
  const ids = sh.getRange(2, col, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return 0;
}

function saveRow(sh, d) {
  if (sh.getLastRow() === 0) sh.appendRow(d.header);
  const id = d.row[d.row.length - 1];
  const r = findRow(sh, id);
  if (r) {
    // truk sudah ada: perbarui baris, pertahankan link foto yang sudah tersimpan
    const fc = d.header.indexOf('Foto kendaraan') + 1;
    if (fc) d.row[fc - 1] = sh.getRange(r, fc).getValue();
    sh.getRange(r, 1, 1, d.row.length).setValues([d.row]);
  } else {
    sh.appendRow(d.row);
  }
  return { ok: true };
}

function saveFoto(sh, d) {
  const r = findRow(sh, d.id);
  if (!r) return { ok: false, error: 'Baris truk belum ada di Sheet' };
  const m = String(d.foto).match(/^data:(.*?);base64,(.*)$/);
  if (!m) return { ok: false, error: 'Format foto tidak valid' };
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'truk_' + d.id + '.jpg');
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
  const url = folder.createFile(blob).getUrl();
  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const fc = header.indexOf('Foto kendaraan') + 1;
  if (fc) sh.getRange(r, fc).setValue(url);
  return { ok: true };
}
