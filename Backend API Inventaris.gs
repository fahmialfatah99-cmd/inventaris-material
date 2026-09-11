/**
 * ============================================================
 * BACKEND API INVENTARIS MATERIAL + UPLOAD FOTO
 * ============================================================
 *
 * Struktur Google Sheets:
 * 1. Sheet master material: "Sheet1"
 *    - Baris 1 sampai 3 dianggap header/judul bebas.
 *    - Data material dibaca mulai baris 4.
 *    - Kolom A: Tipe Material
 *    - Kolom B: Kode Material
 *    - Kolom C: Deskripsi Material
 *
 * 2. Sheet log transaksi: "Log_Transaksi"
 *    - Dibuat otomatis jika belum ada.
 *    - Jika sudah ada tetapi belum punya kolom foto, kolom akan ditambahkan.
 *
 * 3. Foto transaksi:
 *    - Disimpan ke Google Drive folder "Foto_Transaksi".
 *    - Link foto dan ID file dicatat ke Log_Transaksi.
 *
 * Endpoint:
 * GET  -> Mengambil daftar material.
 * GET  ?action=ping -> Cek API aktif.
 * POST -> Menyimpan transaksi + foto.
 */

const MASTER_SHEET_NAME = "Sheet1";
const LOG_SHEET_NAME = "Log_Transaksi";
const PHOTO_FOLDER_NAME = "Foto_Transaksi";
const TIMEZONE = "Asia/Jakarta";

const MATERIAL_CACHE_KEY = "materials_v1";
const MATERIAL_CACHE_SECONDS = 60; // Cache 1 menit agar tetap cepat tapi data tidak terlalu lama.

/**
 * ============================================================
 * ENDPOINT GET
 * ============================================================
 */
function doGet(e) {
  try {
    // Endpoint ping untuk cek koneksi API
    if (e && e.parameter && e.parameter.action === "ping") {
      return jsonOutput({
        status: "success",
        message: "PONG",
        timestamp: Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm:ss")
      });
    }

    const cache = CacheService.getScriptCache();
    const cached = cache.get(MATERIAL_CACHE_KEY);

    // Jika cache masih ada, langsung kembalikan cache agar cepat
    if (cached) {
      return ContentService
        .createTextOutput(cached)
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);

    if (!masterSheet) {
      throw new Error(`Sheet '${MASTER_SHEET_NAME}' tidak ditemukan.`);
    }

    const dataRange = masterSheet.getDataRange().getDisplayValues();
    const materials = [];
    let currentType = "";

    /**
     * Data mulai baris 4.
     * Index array baris 4 adalah 3.
     */
    for (let i = 3; i < dataRange.length; i++) {
      const row = dataRange[i] || [];

      const typeVal = String(row[0] || "").trim();
      const codeVal = String(row[1] || "").trim();
      const descVal = String(row[2] || "").trim();

      // Handle cell tipe yang di-merge.
      // Jika kolom tipe kosong, gunakan tipe dari baris sebelumnya.
      if (typeVal !== "") {
        currentType = typeVal;
      }

      if (descVal !== "") {
        materials.push({
          type: currentType || "-",
          code: codeVal || "-",
          description: descVal
        });
      }
    }

    const responseObject = {
      status: "success",
      total: materials.length,
      data: materials
    };

    const responseText = JSON.stringify(responseObject);

    /**
     * CacheService memiliki batas ukuran value.
     * Jika data terlalu besar, jangan simpan cache.
     */
    if (responseText.length < 90000) {
      cache.put(MATERIAL_CACHE_KEY, responseText, MATERIAL_CACHE_SECONDS);
    }

    return ContentService
      .createTextOutput(responseText)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return jsonOutput({
      status: "error",
      message: error.toString()
    });
  }
}

/**
 * ============================================================
 * ENDPOINT POST
 * ============================================================
 */
function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    // Tunggu maksimal 30 detik jika ada request bersamaan
    lock.tryLock(30000);

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Payload kosong. Tidak ada data yang diterima.");
    }

    const data = JSON.parse(e.postData.contents);

    const logSheet = setupLogSheet();

    const now = new Date();

    const transactionId =
      "TRX-" + Utilities.formatDate(now, TIMEZONE, "yyyyMMdd-HHmmss");

    const timestamp =
      Utilities.formatDate(now, TIMEZONE, "yyyy-MM-dd HH:mm:ss");

    let fotoUrl = "";
    let fotoFileId = "";

    /**
     * Jika frontend mengirim foto base64, simpan ke Google Drive.
     */
    if (data.fotoBase64 && String(data.fotoBase64).trim() !== "") {
      const savedFoto = saveFotoToDrive(data, transactionId);
      fotoUrl = savedFoto.url;
      fotoFileId = savedFoto.fileId;
    }

    const rowData = [
      transactionId,
      timestamp,
      data.tanggal || Utilities.formatDate(now, TIMEZONE, "yyyy-MM-dd"),
      data.jenisTransaksi || "Pengambilan",
      data.namaMandor || "-",
      data.lokasiProyek || "-",
      data.kodeMaterial || "-",
      data.deskripsiMaterial || "-",
      Number(data.qty) || 0,
      data.satuan || "Pcs",
      data.keterangan || "-",
      fotoUrl,
      fotoFileId
    ];

    logSheet.appendRow(rowData);

    return jsonOutput({
      status: "success",
      transactionId: transactionId,
      fotoUrl: fotoUrl,
      fotoFileId: fotoFileId,
      message: "Data transaksi berhasil dicatat ke Google Sheets."
    });

  } catch (error) {
    return jsonOutput({
      status: "error",
      message: error.toString()
    });

  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================
 * SETUP SHEET LOG TRANSAKSI
 * ============================================================
 */
function setupLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName(LOG_SHEET_NAME);

  const headers = [
    "ID Transaksi",
    "Timestamp",
    "Tanggal",
    "Jenis Transaksi",
    "Nama Mandor",
    "Lokasi / Nama Proyek",
    "Kode Material",
    "Deskripsi Material",
    "Jumlah (Qty)",
    "Satuan",
    "Kondisi / Keterangan",
    "Foto URL",
    "Foto File ID"
  ];

  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET_NAME);
    logSheet.appendRow(headers);
  } else {
    const lastCol = logSheet.getLastColumn();

    /**
     * Jika sheet sudah ada tetapi jumlah kolom kurang,
     * tambahkan kolom yang kurang, misalnya Foto URL dan Foto File ID.
     */
    if (lastCol < headers.length) {
      logSheet
        .getRange(1, lastCol + 1, 1, headers.length - lastCol)
        .setValues([headers.slice(lastCol)]);
    }
  }

  // Rapikan header
  logSheet
    .getRange(1, 1, 1, headers.length)
    .setFontWeight("bold")
    .setBackground("#e2e8f0")
    .setWrap(true);

  logSheet.setFrozenRows(1);

  return logSheet;
}

/**
 * ============================================================
 * SIMPAN FOTO KE GOOGLE DRIVE
 * ============================================================
 */
function saveFotoToDrive(data, transactionId) {
  const base64Full = String(data.fotoBase64 || "");

  if (!base64Full.trim()) {
    return {
      url: "",
      fileId: ""
    };
  }

  /**
   * Batasi ukuran base64 agar tidak membebani Apps Script.
   * 6.000.000 karakter base64 kira-kira sekitar 4.5MB file biner.
   */
  if (base64Full.length > 6000000) {
    throw new Error("Ukuran foto terlalu besar untuk diproses oleh Apps Script.");
  }

  let base64 = base64Full;

  // Jika format data URL: data:image/jpeg;base64,xxxx
  if (base64.includes(",")) {
    base64 = base64.split(",")[1];
  }

  // Hapus spasi / line break jika ada
  base64 = base64.replace(/\s/g, "");

  if (!base64) {
    throw new Error("Format foto base64 tidak valid.");
  }

  const contentType = String(data.fotoMime || "image/jpeg").toLowerCase();

  let ext = "jpg";

  if (contentType.includes("png")) {
    ext = "png";
  } else if (contentType.includes("webp")) {
    ext = "webp";
  } else if (contentType.includes("jpeg")) {
    ext = "jpg";
  }

  const decoded = Utilities.base64Decode(base64);

  const folder = getOrCreateFolder(PHOTO_FOLDER_NAME);

  const fileName =
    transactionId +
    "-" +
    Utilities.formatDate(new Date(), TIMEZONE, "HHmmss") +
    "." +
    ext;

  const blob = Utilities.newBlob(decoded, contentType, fileName);

  const file = folder.createFile(blob);

  file.setDescription(
    "Transaksi: " + transactionId +
    " | Material: " + (data.deskripsiMaterial || "-") +
    " | Mandor: " + (data.namaMandor || "-")
  );

  /**
   * Coba buat file bisa diakses dengan link.
   * Jika kebijakan Google Workspace melarang, abaikan error.
   */
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    // Biarkan tetap tersimpan di Drive tanpa sharing publik.
  }

  return {
    fileId: file.getId(),
    url: file.getUrl()
  };
}

/**
 * ============================================================
 * AMBIL FOLDER DRIVE, BUAT JIKA BELUM ADA
 * ============================================================
 */
function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(folderName);
}

/**
 * ============================================================
 * HELPER: OUTPUT JSON
 * ============================================================
 */
function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ============================================================
 * OPSIONAL: HAPUS CACHE MATERIAL SECARA MANUAL
 * ============================================================
 * Jalankan fungsi ini dari Apps Script jika Anda baru mengubah
 * daftar material di Sheet1 dan ingin cache langsung diperbarui.
 */
function clearMaterialCache() {
  const cache = CacheService.getScriptCache();
  cache.remove(MATERIAL_CACHE_KEY);
}

/**
 * ============================================================
 * OPSIONAL: TEST SETUP SHEET
 * ============================================================
 * Jalankan sekali untuk memastikan Log_Transaksi terbentuk.
 */
function testSetupLogSheet() {
  setupLogSheet();
}
