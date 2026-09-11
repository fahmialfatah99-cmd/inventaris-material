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
 *    - Mendukung multi-cabang/kota (Lampung, Bengkulu, Palembang, Semarang, dll).
 *    - Jika sudah ada tetapi belum punya kolom Kota / Cabang atau Foto, kolom akan disinkronkan otomatis.
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

    const forceFresh = e && e.parameter && (e.parameter.nocache === "1" || e.parameter.action === "refresh");
    const cache = CacheService.getScriptCache();
    const cached = forceFresh ? null : cache.get(MATERIAL_CACHE_KEY);

    // Jika cache masih ada dan bukan force refresh, kembalikan cache
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

      let finalDesc = descVal;
      let finalCode = codeVal;

      // Toleransi jika user mengisi deskripsi di kolom B (kolom C kosong)
      if (!finalDesc && finalCode) {
        finalDesc = finalCode;
        finalCode = "-";
      }

      if (finalDesc !== "") {
        materials.push({
          id: `mat_${i + 1}`,
          type: currentType || "-",
          code: finalCode || "-",
          description: finalDesc
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

    const photoUrls = [];
    const photoFileIds = [];

    /**
     * Handle multi-foto (array) atau single foto (legacy)
     */
    const photosToProcess = [];
    if (Array.isArray(data.photos) && data.photos.length > 0) {
      photosToProcess.push(...data.photos);
    } else if (data.fotoBase64 && String(data.fotoBase64).trim() !== "") {
      photosToProcess.push({
        fotoBase64: data.fotoBase64,
        fotoMime: data.fotoMime || "image/jpeg",
        fotoNama: data.fotoNama || "foto.jpg"
      });
    }

    for (let p = 0; p < photosToProcess.length; p++) {
      const pData = photosToProcess[p];
      if (pData && pData.fotoBase64 && String(pData.fotoBase64).trim() !== "") {
        const savedFoto = saveFotoToDrive(pData, transactionId + (photosToProcess.length > 1 ? "-" + (p + 1) : ""), data);
        if (savedFoto.url) photoUrls.push(savedFoto.url);
        if (savedFoto.fileId) photoFileIds.push(savedFoto.fileId);
      }
    }

    const fotoUrlString = photoUrls.join("\n");
    const fotoFileIdString = photoFileIds.join(", ");

    /**
     * Handle multi-material (array) atau single material (legacy)
     */
    let itemsToProcess = [];
    if (Array.isArray(data.items) && data.items.length > 0) {
      itemsToProcess = data.items;
    } else if (Array.isArray(data.materials) && data.materials.length > 0) {
      itemsToProcess = data.materials;
    } else {
      itemsToProcess = [{
        kodeMaterial: data.kodeMaterial || "-",
        deskripsiMaterial: data.deskripsiMaterial || "-",
        qty: Number(data.qty) || 0,
        satuan: data.satuan || "Pcs",
        keterangan: data.keterangan || "-"
      }];
    }

    const rowsToAdd = [];
    const kota = String(data.kota || data.cabang || "Lampung").trim();

    // Dapatkan pemetaan kolom dinamis berdasarkan header di Google Sheet
    const headerRow = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getDisplayValues()[0];
    const headerMap = {};
    headerRow.forEach((h, idx) => {
      headerMap[String(h || "").trim().toLowerCase()] = idx;
    });

    const findCol = (keys, fallback) => {
      for (const k of keys) {
        for (const h in headerMap) {
          if (h.includes(k)) return headerMap[h];
        }
      }
      return fallback;
    };

    const numCols = logSheet.getLastColumn();

    for (let i = 0; i < itemsToProcess.length; i++) {
      const item = itemsToProcess[i];
      const itemKet = item.keterangan || data.keterangan || "-";
      const row = new Array(numCols).fill("-");

      const assign = (keys, val, fallback) => {
        const c = findCol(keys, fallback);
        if (c >= 0 && c < numCols) row[c] = val;
      };

      assign(["id transaksi"], transactionId, 0);
      assign(["timestamp"], timestamp, 1);
      assign(["tanggal"], data.tanggal || Utilities.formatDate(now, TIMEZONE, "yyyy-MM-dd"), 2);
      assign(["kota", "cabang"], kota, 3);
      assign(["jenis transaksi"], data.jenisTransaksi || "Pengambilan", 4);
      assign(["mandor"], data.namaMandor || "-", 5);
      assign(["proyek", "lokasi"], data.lokasiProyek || "-", 6);
      assign(["kode material"], item.kodeMaterial || item.code || "-", 7);
      assign(["deskripsi material"], item.deskripsiMaterial || item.description || "-", 8);
      assign(["jumlah", "qty"], Number(item.qty) || 0, 9);
      assign(["satuan"], item.satuan || "Pcs", 10);
      assign(["kondisi", "keterangan"], itemKet, 11);
      assign(["foto url"], fotoUrlString, 12);
      assign(["foto file id"], fotoFileIdString, 13);

      rowsToAdd.push(row);
    }

    if (rowsToAdd.length > 0) {
      const lastRow = logSheet.getLastRow();
      logSheet.getRange(lastRow + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
    }

    return jsonOutput({
      status: "success",
      transactionId: transactionId,
      kota: kota,
      totalItems: rowsToAdd.length,
      totalPhotos: photoUrls.length,
      fotoUrls: photoUrls,
      fotoFileIds: photoFileIds,
      message: `${rowsToAdd.length} material transaksi cabang ${kota} berhasil dicatat ke Google Sheets.`
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

  const defaultHeaders = [
    "ID Transaksi",
    "Timestamp",
    "Tanggal",
    "Kota / Cabang",
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
    logSheet.appendRow(defaultHeaders);
  } else {
    const lastCol = logSheet.getLastColumn();

    if (lastCol === 0) {
      logSheet.appendRow(defaultHeaders);
    } else {
      const headerRow = logSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
      const headerLower = headerRow.map(h => String(h || "").trim().toLowerCase());

      // 1. Pastikan kolom Kota / Cabang ada
      const hasKota = headerLower.some(h => h.includes("kota") || h.includes("cabang"));
      if (!hasKota) {
        let tanggalCol = -1;
        for (let c = 0; c < headerLower.length; c++) {
          if (headerLower[c].includes("tanggal")) {
            tanggalCol = c + 1; // 1-indexed
            break;
          }
        }

        if (tanggalCol > 0) {
          logSheet.insertColumnAfter(tanggalCol);
          logSheet.getRange(1, tanggalCol + 1).setValue("Kota / Cabang");
        } else {
          logSheet.getRange(1, logSheet.getLastColumn() + 1).setValue("Kota / Cabang");
        }
      }

      // 2. Pastikan kolom Foto URL dan Foto File ID ada
      const curLastCol = logSheet.getLastColumn();
      const curHeaders = logSheet.getRange(1, 1, 1, curLastCol).getDisplayValues()[0].map(h => String(h || "").trim().toLowerCase());
      
      if (!curHeaders.some(h => h.includes("foto url"))) {
        logSheet.getRange(1, logSheet.getLastColumn() + 1).setValue("Foto URL");
      }
      if (!curHeaders.some(h => h.includes("foto file id"))) {
        logSheet.getRange(1, logSheet.getLastColumn() + 1).setValue("Foto File ID");
      }
    }
  }

  // Rapikan header
  const finalLastCol = logSheet.getLastColumn();
  logSheet
    .getRange(1, 1, 1, finalLastCol)
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
function saveFotoToDrive(photoData, fileSuffix, parentData) {
  const base64Full = String(photoData.fotoBase64 || "");

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

  const contentType = String(photoData.fotoMime || "image/jpeg").toLowerCase();

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
    fileSuffix +
    "-" +
    Utilities.formatDate(new Date(), TIMEZONE, "HHmmss") +
    "." +
    ext;

  const blob = Utilities.newBlob(decoded, contentType, fileName);

  const file = folder.createFile(blob);

  const kota = (parentData && (parentData.kota || parentData.cabang)) || "-";
  const mandor = (parentData && parentData.namaMandor) || "-";
  const proyek = (parentData && parentData.lokasiProyek) || "-";

  file.setDescription(
    "Transaksi: " + fileSuffix +
    " | Kota: " + kota +
    " | Mandor: " + mandor +
    " | Proyek: " + proyek
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
