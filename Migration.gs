/**
 * Migration.gs
 * ============
 * Fungsi migrasi data dari spreadsheet lama ke database baru.
 * Menghandle berbagai format tab bulanan dari spreadsheet lama.
 *
 * Old Spreadsheet ID: 1_s8PykfM8gK5eDnQtE6dxmeDj-_TOk1vqc3qHkDbf-w
 */

/**
 * Mapping nama bulan Indonesia ke nomor bulan (0-indexed untuk Date constructor).
 */
const INDONESIAN_MONTHS = {
  'januari': 0,
  'februari': 1,
  'maret': 2,
  'april': 3,
  'mei': 4,
  'juni': 5,
  'juli': 6,
  'agustus': 7,
  'september': 8,
  'oktober': 9,
  'november': 10,
  'desember': 11
};

/**
 * Parse tanggal format Indonesia (contoh: '1 Mei 2026', '13 Juni 2026') menjadi objek Date.
 * Menangani berbagai variasi format tanggal yang mungkin ada di spreadsheet lama.
 *
 * @param {string} dateStr - String tanggal dalam format Indonesia
 * @returns {Date|null} Objek Date atau null jika parsing gagal
 */
function parseIndonesianDate(dateStr) {
  if (!dateStr) return null;

  // Jika sudah berupa Date object
  if (dateStr instanceof Date && !isNaN(dateStr.getTime())) {
    return dateStr;
  }

  // Konversi ke string dan bersihkan
  const str = String(dateStr).trim();
  if (!str) return null;

  // Coba parse format: "DD NamaBulan YYYY" (contoh: "1 Mei 2026", "13 Juni 2026")
  const parts = str.split(/\s+/);
  if (parts.length >= 3) {
    const day = parseInt(parts[0], 10);
    const monthName = parts[1].toLowerCase();
    const year = parseInt(parts[2], 10);

    if (!isNaN(day) && !isNaN(year) && INDONESIAN_MONTHS.hasOwnProperty(monthName)) {
      const month = INDONESIAN_MONTHS[monthName];
      return new Date(year, month, day);
    }
  }

  // Coba parse format: "DD-MM-YYYY" atau "DD/MM/YYYY"
  const dateRegex = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/;
  const match = str.match(dateRegex);
  if (match) {
    const day = parseInt(match[1], 10);
    const month = parseInt(match[2], 10) - 1; // 0-indexed
    const year = parseInt(match[3], 10);
    return new Date(year, month, day);
  }

  // Fallback: coba native Date parsing
  const nativeDate = new Date(str);
  if (!isNaN(nativeDate.getTime())) {
    return nativeDate;
  }

  Logger.log(`Gagal parse tanggal: "${str}"`);
  return null;
}

/**
 * Migrasi data dari spreadsheet lama ke database baru.
 * Membaca semua tab pada spreadsheet lama, mencari header 'DATE TO BE POSTED',
 * kemudian membaca dan memindahkan data ke Master_Content pada spreadsheet baru.
 *
 * Struktur spreadsheet lama:
 * - Multiple tabs (satu per bulan)
 * - Header info di baris 1-16 (skip)
 * - Header data di baris yang mengandung 'DATE TO BE POSTED'
 * - Data mulai dari baris setelah header, kolom B-H (index 1-7)
 *
 * Mapping kolom lama → baru:
 * - DATE TO BE POSTED → Date_To_Be_Posted
 * - PIC → PIC_Design
 * - CONTENT → Content_Title
 * - PLATFORM → Platform
 * - TEMPLATE → Template
 * - RATIO → Ratio
 * - STATUS → Status
 *
 * @returns {string} JSON string berisi status migrasi dan jumlah data yang dimigrasikan
 */
function migrateFromOldSpreadsheet() {
  try {
    const OLD_SPREADSHEET_ID = '1_s8PykfM8gK5eDnQtE6dxmeDj-_TOk1vqc3qHkDbf-w';

    // Buka spreadsheet lama
    let oldSs;
    try {
      oldSs = SpreadsheetApp.openById(OLD_SPREADSHEET_ID);
    } catch (e) {
      return sendResponse(false, null, `Gagal membuka spreadsheet lama: ${e.message}. Pastikan Anda memiliki akses.`);
    }

    // Buka spreadsheet baru (tujuan)
    const newSheet = getSheet('Master_Content');

    // Ambil data yang sudah ada di sheet baru untuk menentukan nomor ID berikutnya
    const lastNum = getLastContentNumber();
    let idCounter = lastNum + 1;

    // Kumpulkan semua data migrasi
    const allMigratedRows = [];
    const allSheets = oldSs.getSheets();

    Logger.log(`Ditemukan ${allSheets.length} tab di spreadsheet lama.`);

    for (const oldSheet of allSheets) {
      const sheetName = oldSheet.getName();
      Logger.log(`Memproses tab: "${sheetName}"...`);

      try {
        // Ambil semua data dari tab
        const data = oldSheet.getDataRange().getValues();

        if (data.length === 0) {
          Logger.log(`Tab "${sheetName}" kosong, lewati.`);
          continue;
        }

        // Cari baris header yang mengandung 'DATE TO BE POSTED'
        let headerRowIndex = -1;
        let colMapping = {};

        for (let i = 0; i < data.length; i++) {
          const row = data[i];
          for (let j = 0; j < row.length; j++) {
            const cellValue = String(row[j]).trim().toUpperCase();
            if (cellValue === 'DATE TO BE POSTED' || cellValue.includes('DATE TO BE POSTED')) {
              headerRowIndex = i;

              // Map posisi kolom berdasarkan header yang ditemukan
              for (let k = 0; k < row.length; k++) {
                const header = String(row[k]).trim().toUpperCase();
                if (header.includes('DATE')) colMapping.date = k;
                else if (header === 'PIC' || header.includes('PIC')) colMapping.pic = k;
                else if (header === 'CONTENT' || header.includes('CONTENT')) colMapping.content = k;
                else if (header === 'PLATFORM' || header.includes('PLATFORM')) colMapping.platform = k;
                else if (header === 'TEMPLATE' || header.includes('TEMPLATE')) colMapping.template = k;
                else if (header === 'RATIO' || header.includes('RATIO')) colMapping.ratio = k;
                else if (header === 'STATUS' || header.includes('STATUS')) colMapping.status = k;
              }

              break;
            }
          }
          if (headerRowIndex !== -1) break;
        }

        if (headerRowIndex === -1) {
          Logger.log(`Tab "${sheetName}": Header 'DATE TO BE POSTED' tidak ditemukan, lewati.`);
          continue;
        }

        // Fallback ke posisi default (kolom B-H = index 1-7) jika mapping tidak lengkap
        if (colMapping.date === undefined) colMapping.date = 1;
        if (colMapping.pic === undefined) colMapping.pic = 2;
        if (colMapping.content === undefined) colMapping.content = 3;
        if (colMapping.platform === undefined) colMapping.platform = 4;
        if (colMapping.template === undefined) colMapping.template = 5;
        if (colMapping.ratio === undefined) colMapping.ratio = 6;
        if (colMapping.status === undefined) colMapping.status = 7;

        Logger.log(`Tab "${sheetName}": Header ditemukan di baris ${headerRowIndex + 1}. Mapping: ${JSON.stringify(colMapping)}`);

        // Baca data mulai dari baris setelah header
        let rowsFromTab = 0;
        for (let i = headerRowIndex + 1; i < data.length; i++) {
          const row = data[i];

          // Ambil nilai tanggal
          const rawDate = row[colMapping.date];

          // Skip baris tanpa tanggal valid
          const parsedDate = parseIndonesianDate(rawDate);
          if (!parsedDate) {
            continue;
          }

          // Ambil data dari kolom yang dimapping
          const pic = row[colMapping.pic] ? String(row[colMapping.pic]).trim() : '';
          const content = row[colMapping.content] ? String(row[colMapping.content]).trim() : '';
          const platform = row[colMapping.platform] ? String(row[colMapping.platform]).trim() : '';
          const template = row[colMapping.template] ? String(row[colMapping.template]).trim() : '';
          const ratio = row[colMapping.ratio] ? String(row[colMapping.ratio]).trim() : '';
          const status = row[colMapping.status] ? String(row[colMapping.status]).trim() : 'Draft';

          // Skip baris yang tidak memiliki data substantif (semua kolom kosong selain tanggal)
          if (!pic && !content && !platform) {
            continue;
          }

          // Generate ID baru
          const year = new Date().getFullYear();
          const contentId = `CNT-${year}-${String(idCounter).padStart(3, '0')}`;
          idCounter++;

          // Hitung Bulan_Tahun dari tanggal
          const bulanTahun = formatDateToYearMonth(parsedDate);

          // Normalisasi status
          const normalizedStatus = normalizeStatus_(status);

          // Siapkan baris untuk sheet baru
          allMigratedRows.push([
            contentId,
            parsedDate,
            bulanTahun,
            pic,
            content,
            platform,
            template,
            ratio,
            normalizedStatus
          ]);

          rowsFromTab++;
        }

        Logger.log(`Tab "${sheetName}": ${rowsFromTab} baris berhasil dimigrasikan.`);

      } catch (tabError) {
        Logger.log(`Error memproses tab "${sheetName}": ${tabError.message}`);
        // Lanjutkan ke tab berikutnya
        continue;
      }
    }

    // Batch write semua data ke Master_Content
    if (allMigratedRows.length > 0) {
      const startRow = newSheet.getLastRow() + 1;
      newSheet.getRange(startRow, 1, allMigratedRows.length, 9).setValues(allMigratedRows);

      // Bersihkan cache setelah migrasi
      clearContentCache();

      Logger.log(`=== MIGRASI SELESAI ===`);
      Logger.log(`Total ${allMigratedRows.length} baris berhasil dimigrasikan.`);

      return {
        success: true,
        migratedCount: allMigratedRows.length
      };

    } else {
      Logger.log('Tidak ada data yang bisa dimigrasikan.');
      return {
        success: true,
        migratedCount: 0
      };
    }

  } catch (error) {
    Logger.log(`Error migrateFromOldSpreadsheet: ${error.message}`);
    throw new Error(`Gagal melakukan migrasi: ${error.message}`);
  }
}

/**
 * Normalisasi nilai status agar konsisten dengan opsi yang valid.
 * Menangani variasi penulisan dari spreadsheet lama.
 *
 * @param {string} status - Nilai status dari spreadsheet lama
 * @returns {string} Status yang sudah dinormalisasi
 * @private
 */
function normalizeStatus_(status) {
  if (!status) return 'Draft';

  const lowerStatus = status.toLowerCase().trim();

  // Mapping variasi penulisan ke status standar
  if (lowerStatus === 'draft' || lowerStatus === 'draf') {
    return 'Draft';
  }
  if (lowerStatus === 'in progress' || lowerStatus === 'inprogress' ||
      lowerStatus === 'in-progress' || lowerStatus === 'progress' ||
      lowerStatus === 'proses' || lowerStatus === 'dalam proses') {
    return 'In Progress';
  }
  if (lowerStatus === 'review' || lowerStatus === 'reviewing' ||
      lowerStatus === 'revisi' || lowerStatus === 'acc') {
    return 'Review';
  }
  if (lowerStatus === 'completed' || lowerStatus === 'complete' ||
      lowerStatus === 'done' || lowerStatus === 'selesai' ||
      lowerStatus === 'posted' || lowerStatus === 'post') {
    return 'Completed';
  }

  // Default jika tidak cocok dengan mapping manapun
  return 'Draft';
}
