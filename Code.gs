/**
 * Andala Brand Check → Google Sheets
 * Tempel seluruh kode ini di Extensions > Apps Script pada spreadsheet Anda,
 * lalu Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 */

const TOKEN = "andala-brandcheck";   // samakan dengan SHEET_TOKEN di index.html
const SHEET_NAME = "Hasil";

const HEADERS = [
  ["id", "ID"], ["waktu", "Waktu"], ["status", "Status"],
  ["nama_bisnis", "Nama bisnis"], ["bidang", "Bidang"], ["usia_bisnis", "Usia bisnis"],
  ["skor_total", "Skor total"], ["tingkat", "Tingkat"], ["diagnosis", "Diagnosis"], ["pola_kedua", "Pola kedua"],
  ["akar_masalah", "Akar masalah"], ["tahap_brand", "Tahap brand"],
  ["skor_identitas", "Identitas"], ["skor_positioning", "Positioning"], ["skor_konsistensi", "Konsistensi"],
  ["skor_digital", "Digital"], ["skor_konten", "Konten"], ["skor_kepercayaan", "Kepercayaan"], ["skor_konversi", "Konversi"],
  ["kepribadian_sekarang", "Kepribadian sekarang"], ["kepribadian_diinginkan", "Kepribadian diinginkan"],
  ["titik_buta", "Jumlah titik buta"], ["layanan_relevan", "Layanan relevan"],
  ["kesiapan_waktu", "Waktu mulai"], ["pengambil_keputusan", "Pengambil keputusan"], ["pengalaman_agensi", "Pengalaman agensi"],
  ["nama_kontak", "Nama kontak"], ["jabatan", "Jabatan"], ["whatsapp", "WhatsApp"], ["persetujuan", "Persetujuan"],
  ["waktu_kontak", "Waktu kontak"], ["jawaban_lengkap", "Jawaban lengkap"], ["halaman", "Halaman"]
];
const CONTACT_KEYS = ["nama_kontak", "jabatan", "whatsapp", "persetujuan"];

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    if (data.token !== TOKEN) return json_({ ok: false, error: "token" });
    if (!data.id || !/^[a-z0-9-]{6,40}$/i.test(String(data.id))) return json_({ ok: false, error: "id" });

    const sheet = getSheet_();
    const keys = HEADERS.map(h => h[0]);
    const now = new Date();
    const row = findRow_(sheet, String(data.id));

    if (row) {
      // Perbarui baris yang sama (misalnya saat kontak dikirim setelah hasil)
      const current = sheet.getRange(row, 1, 1, keys.length).getValues()[0];
      keys.forEach((k, i) => {
        if (k === "id" || k === "waktu") return;
        if (data[k] !== undefined && data[k] !== "") current[i] = clean_(data[k]);
      });
      if (data.type === "lead") {
        current[keys.indexOf("status")] = "Lead (menghubungi)";
        current[keys.indexOf("waktu_kontak")] = now;
      }
      sheet.getRange(row, 1, 1, keys.length).setValues([current]);
    } else {
      const values = keys.map(k => {
        if (k === "waktu") return now;
        if (k === "status") return data.type === "lead" ? "Lead (menghubungi)" : "Hasil saja";
        if (k === "waktu_kontak") return data.type === "lead" ? now : "";
        if (data.type !== "lead" && CONTACT_KEYS.indexOf(k) > -1) return "";
        return data[k] === undefined ? "" : clean_(data[k]);
      });
      sheet.appendRow(values);
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function doGet() {
  return json_({ ok: true, service: "Andala Brand Check" });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS.map(h => h[1]));
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#0A0A0A").setFontColor("#FFBF00");
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(HEADERS.length - 1, 420);
  }
  return sheet;
}

function findRow_(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const found = sheet.getRange(2, 1, last - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  return found ? found.getRow() : 0;
}

// Cegah teks yang diawali = + - @ dieksekusi sebagai rumus
function clean_(v) {
  if (typeof v === "number") return v;
  const s = String(v).slice(0, 5000);
  return (/^[=+\-@]/.test(s) || /^0\d+$/.test(s)) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
