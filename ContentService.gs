/**
 * ContentService.gs
 * =================
 * Service layer untuk operasi CRUD pada sheet Master_Content.
 * Semua fungsi bersifat global agar dapat dipanggil dari frontend via google.script.run.
 *
 * Kolom Master_Content:
 * ID_Content | Date_To_Be_Posted | Bulan_Tahun | PIC_Design | Content_Title | Platform | Template | Ratio | Status
 */

// Konstanta index kolom (0-based) untuk Master_Content
const MC_COL = {
  ID: 0,
  DATE: 1,
  BULAN_TAHUN: 2,
  PIC: 3,
  TITLE: 4,
  PLATFORM: 5,
  TEMPLATE: 6,
  RATIO: 7,
  STATUS: 8,
  FILE_URL: 9,
  FORMAT: 10,
  DESCRIPTION: 11,
  EVENT_ID: 12
};

/**
 * Mengambil data konten berdasarkan bulan-tahun tertentu.
 * Menggunakan CacheService untuk mempercepat akses berulang (cache 300 detik).
 *
 * @param {string} yearMonth - Format 'YYYY-MM' (contoh: '2026-06')
 * @returns {string} JSON string berisi array objek konten
 */
function getContentByMonth(yearMonth) {
  try {
    const sheet = getSheet('Master_Content');
    
    // Auto-migration: check and create 12th column 'Description' and 13th column 'Event_ID' if missing
    const lastCol = sheet.getLastColumn();
    if (lastCol < 13) {
      const maxCols = sheet.getMaxColumns();
      if (maxCols < 13) {
        sheet.insertColumnsAfter(maxCols, 13 - maxCols);
      }
      
      const header12 = sheet.getRange(1, 12);
      if (!header12.getValue()) {
        header12.setValue('Description');
        header12.setFontWeight('bold');
        header12.setBackground('#4285F4');
        header12.setFontColor('#FFFFFF');
        header12.setHorizontalAlignment('center');
        sheet.setColumnWidth(12, 250);
      }

      const header13 = sheet.getRange(1, 13);
      if (!header13.getValue()) {
        header13.setValue('Event_ID');
        header13.setFontWeight('bold');
        header13.setBackground('#4285F4');
        header13.setFontColor('#FFFFFF');
        header13.setHorizontalAlignment('center');
        sheet.setColumnWidth(13, 150);
      }
    }

    // Ambil semua data dari sheet
    const data = sheet.getDataRange().getValues();

    // Skip header (baris pertama)
    const contents = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const bulanTahun = normalizeYearMonth(row[MC_COL.BULAN_TAHUN]);

      // Filter berdasarkan Bulan_Tahun
      if (bulanTahun === yearMonth) {
        contents.push({
          id: String(row[MC_COL.ID] || ''),
          date: row[MC_COL.DATE] instanceof Date ? Utilities.formatDate(row[MC_COL.DATE], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(row[MC_COL.DATE] || ''),
          bulanTahun: String(bulanTahun || ''),
          pic: String(row[MC_COL.PIC] || ''),
          title: String(row[MC_COL.TITLE] || ''),
          platform: String(row[MC_COL.PLATFORM] || ''),
          template: String(row[MC_COL.TEMPLATE] || ''),
          ratio: String(row[MC_COL.RATIO] || ''),
          status: String(row[MC_COL.STATUS] || ''),
          fileUrl: String(row[MC_COL.FILE_URL] || ''),
          format: String(row[MC_COL.FORMAT] || 'desain'),
          description: String(row[MC_COL.DESCRIPTION] || ''),
          eventId: String(row[MC_COL.EVENT_ID] || '')
        });
      }
    }

    // Sort berdasarkan tanggal ascending
    contents.sort((a, b) => new Date(a.date) - new Date(b.date));

    return contents;

  } catch (error) {
    Logger.log(`Error getContentByMonth: ${error.message}`);
    throw new Error(`Gagal mengambil data konten: ${error.message}`);
  }
}

/**
 * Menambahkan konten baru ke Master_Content.
 *
 * @param {Object} formData - Data form: {date, pic, title, platform, template, ratio, status}
 * @returns {string} JSON string dengan ID yang di-generate
 */
function addContent(formData, userName) {
  try {
    const sheet = getSheet('Master_Content');

    // Generate ID otomatis
    const contentId = generateContentId();

    // Parse tanggal dan hitung Bulan_Tahun
    const dateObj = new Date(formData.date);
    if (isNaN(dateObj.getTime())) {
      throw new Error('Format tanggal tidak valid.');
    }
    const bulanTahun = formatDateToYearMonth(dateObj);

    // Sinkronisasi ke Google Calendar
    let eventId = '';
    try {
      eventId = syncCalendarEvent({
        id: contentId,
        title: formData.title,
        platform: formData.platform,
        ratio: formData.ratio,
        pic: formData.pic,
        date: dateObj,
        status: formData.status || 'Draft',
        template: formData.template,
        fileUrl: '',
        description: formData.description,
        eventId: ''
      });
    } catch (e) {
      Logger.log(`Gagal sinkronisasi Google Calendar saat tambah konten: ${e.message}`);
    }

    // Siapkan baris data baru sesuai urutan kolom (13 kolom)
    const newRow = [
      contentId,
      dateObj,
      bulanTahun,
      formData.pic || '',
      formData.title || '',
      formData.platform || '',
      formData.template || '',
      formData.ratio || '',
      formData.status || 'Draft',
      '', // fileUrl
      formData.format || 'desain',
      formData.description || '', // description
      eventId // Event_ID
    ];

    // Append baris ke sheet
    sheet.appendRow(newRow);
    SpreadsheetApp.flush();

    // Bersihkan cache
    clearContentCache(bulanTahun);

    // Catat Log Aktivitas
    logActivity('Tambah Konten', `Menambahkan konten baru: "${formData.title}" (ID: ${contentId})`, userName);

    // Kirim Notifikasi WhatsApp ke Grup
    try {
      notifyNewContentWhatsApp({
        title: formData.title,
        platform: formData.platform,
        ratio: formData.ratio,
        pic: formData.pic,
        date: dateObj
      }, userName);
    } catch (e) {
      Logger.log(`Gagal mengirim notifikasi WhatsApp saat tambah konten: ${e.message}`);
    }

    Logger.log(`Konten baru ditambahkan: ${contentId}`);
    return { id: contentId };

  } catch (error) {
    Logger.log(`Error addContent: ${error.message}`);
    throw new Error(`Gagal menambahkan konten: ${error.message}`);
  }
}

/**
 * Menambahkan beberapa konten baru secara massal ke Master_Content.
 *
 * @param {Array<Object>} formDataArray - Array data form: [{date, pic, title, platform, template, ratio, status}]
 * @returns {Object} Hasil operasi: {success: true, count: number}
 */
function addMultipleContents(formDataArray, userName) {
  try {
    if (!Array.isArray(formDataArray) || formDataArray.length === 0) {
      throw new Error('Data konten kosong atau tidak valid.');
    }

    const sheet = getSheet('Master_Content');
    const lastRow = sheet.getLastRow();
    const lastNum = getLastContentNumber();
    let nextNumber = lastNum + 1;
    const year = new Date().getFullYear();

    const newRows = [];
    const bulanTahunsToClear = new Set();

    for (let i = 0; i < formDataArray.length; i++) {
      const formData = formDataArray[i];
      
      const paddedNumber = String(nextNumber).padStart(3, '0');
      const contentId = `CNT-${year}-${paddedNumber}`;
      nextNumber++;

      // Parse tanggal dan hitung Bulan_Tahun
      const dateObj = new Date(formData.date);
      if (isNaN(dateObj.getTime())) {
        throw new Error(`Format tanggal tidak valid pada baris ke-${i + 1}.`);
      }
      const bulanTahun = formatDateToYearMonth(dateObj);
      bulanTahunsToClear.add(bulanTahun);

      // Sinkronisasi ke Google Calendar
      let eventId = '';
      try {
        eventId = syncCalendarEvent({
          id: contentId,
          title: formData.title,
          platform: formData.platform,
          ratio: formData.ratio,
          pic: formData.pic,
          date: dateObj,
          status: formData.status || 'Draft',
          template: formData.template,
          fileUrl: '',
          description: formData.description,
          eventId: ''
        });
      } catch (e) {
        Logger.log(`Gagal sync kalender pada bulk item ke-${i+1}: ${e.message}`);
      }

      // Siapkan baris data baru sesuai urutan kolom (13 kolom)
      const newRow = [
        contentId,
        dateObj,
        bulanTahun,
        formData.pic || '',
        formData.title || '',
        formData.platform || '',
        formData.template || '',
        formData.ratio || '',
        formData.status || 'Draft',
        '', // fileUrl
        formData.format || 'desain',
        formData.description || '', // description
        eventId // Event_ID
      ];
      newRows.push(newRow);
    }

    // Pastikan sheet memiliki baris & kolom yang cukup untuk menampung data baru
    const maxRows = sheet.getMaxRows();
    const maxCols = sheet.getMaxColumns();
    const requiredRows = lastRow + newRows.length;

    if (requiredRows > maxRows) {
      sheet.insertRowsAfter(maxRows, requiredRows - maxRows);
    }
    if (maxCols < 13) {
      sheet.insertColumnsAfter(maxCols, 13 - maxCols);
    }

    // Tulis semua baris sekaligus (13 kolom)
    sheet.getRange(lastRow + 1, 1, newRows.length, 13).setValues(newRows);
    SpreadsheetApp.flush();

    // Bersihkan cache untuk semua Bulan_Tahun yang terpengaruh
    bulanTahunsToClear.forEach(bulanTahun => {
      clearContentCache(bulanTahun);
    });

    // Catat Log Aktivitas
    logActivity('Tambah Banyak Konten', `Menambahkan ${newRows.length} konten baru secara massal`, userName);

    // Kirim notifikasi WhatsApp ke Grup (1 pesan digest untuk massal)
    try {
      notifyBulkContentsWhatsApp(newRows.length, userName);
    } catch (e) {
      Logger.log(`Gagal mengirim notifikasi WhatsApp bulk: ${e.message}`);
    }

    Logger.log(`${newRows.length} konten baru ditambahkan secara bulk.`);
    return { success: true, count: newRows.length };

  } catch (error) {
    Logger.log(`Error addMultipleContents: ${error.message}`);
    throw new Error(`Gagal menambahkan konten massal: ${error.message}`);
  }
}

/**
 * Mengupdate konten yang sudah ada berdasarkan ID_Content.
 *
 * @param {string} id - ID_Content yang akan diupdate
 * @param {Object} formData - Data form: {date, pic, title, platform, template, ratio, status}
 * @returns {string} JSON string status update
 */
function updateContent(id, formData, userName) {
  try {
    const sheet = getSheet('Master_Content');
    const data = sheet.getDataRange().getValues();

    // Cari baris berdasarkan ID_Content
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][MC_COL.ID].toString().trim() === id.toString().trim()) {
        rowIndex = i + 1; // +1 karena sheet 1-indexed
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error(`Konten dengan ID "${id}" tidak ditemukan.`);
    }

    const oldBulanTahun = normalizeYearMonth(data[rowIndex - 1][MC_COL.BULAN_TAHUN]);

    // Parse tanggal dan hitung Bulan_Tahun
    const dateObj = new Date(formData.date);
    if (isNaN(dateObj.getTime())) {
      throw new Error('Format tanggal tidak valid.');
    }
    const bulanTahun = formatDateToYearMonth(dateObj);

    // Preserve existing File_URL & Event_ID
    const existingFileUrl = data[rowIndex - 1][MC_COL.FILE_URL] || '';
    const existingEventId = data[rowIndex - 1][MC_COL.EVENT_ID] || '';

    // Sinkronisasi/Update Google Calendar
    let newEventId = existingEventId;
    try {
      newEventId = syncCalendarEvent({
        id: id,
        title: formData.title,
        platform: formData.platform,
        ratio: formData.ratio,
        pic: formData.pic,
        date: dateObj,
        status: formData.status || 'Draft',
        template: formData.template,
        fileUrl: existingFileUrl,
        description: formData.description,
        eventId: existingEventId
      });
    } catch (e) {
      Logger.log(`Gagal sync Google Calendar saat updateContent: ${e.message}`);
    }

    // Update seluruh baris (13 kolom, dimulai dari kolom 1)
    const updatedRow = [
      id,
      dateObj,
      bulanTahun,
      formData.pic || '',
      formData.title || '',
      formData.platform || '',
      formData.template || '',
      formData.ratio || '',
      formData.status || 'Draft',
      existingFileUrl,
      formData.format || 'desain',
      formData.description || '', // description
      newEventId // Event_ID
    ];

    const maxCols = sheet.getMaxColumns();
    if (maxCols < 13) {
      sheet.insertColumnsAfter(maxCols, 13 - maxCols);
    }

    sheet.getRange(rowIndex, 1, 1, 13).setValues([updatedRow]);
    SpreadsheetApp.flush();

    // Bersihkan cache
    clearContentCache(oldBulanTahun);
    if (oldBulanTahun !== bulanTahun) {
      clearContentCache(bulanTahun);
    }

    // Catat Log Aktivitas
    logActivity('Edit Konten', `Mengubah konten: "${formData.title}" (ID: ${id})`, userName);

    Logger.log(`Konten ${id} berhasil diupdate.`);
    return true;

  } catch (error) {
    Logger.log(`Error updateContent: ${error.message}`);
    throw new Error(`Gagal mengupdate konten: ${error.message}`);
  }
}

/**
 * Menghapus konten berdasarkan ID_Content.
 *
 * @param {string} id - ID_Content yang akan dihapus
 * @returns {string} JSON string status penghapusan
 */
function deleteContent(id, userName) {
  try {
    const sheet = getSheet('Master_Content');
    const data = sheet.getDataRange().getValues();

    // Cari baris berdasarkan ID_Content
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][MC_COL.ID].toString().trim() === id.toString().trim()) {
        rowIndex = i + 1; // +1 karena sheet 1-indexed
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error(`Konten dengan ID "${id}" tidak ditemukan.`);
    }

    const title = data[rowIndex - 1][MC_COL.TITLE] || '';
    const oldBulanTahun = normalizeYearMonth(data[rowIndex - 1][MC_COL.BULAN_TAHUN]);
    const eventId = data[rowIndex - 1][MC_COL.EVENT_ID] || '';

    // Hapus event kalender Google
    if (eventId) {
      try {
        deleteCalendarEvent(eventId);
      } catch (e) {
        Logger.log(`Gagal menghapus event kalender saat deleteContent: ${e.message}`);
      }
    }

    // Hapus baris
    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();

    // Bersihkan cache
    clearContentCache(oldBulanTahun);

    // Catat Log Aktivitas
    logActivity('Hapus Konten', `Menghapus konten: "${title}" (ID: ${id})`, userName);

    Logger.log(`Konten ${id} berhasil dihapus.`);
    return true;

  } catch (error) {
    Logger.log(`Error deleteContent: ${error.message}`);
    throw new Error(`Gagal menghapus konten: ${error.message}`);
  }
}

/**
 * Mengupdate status konten saja (tanpa mengubah field lain).
 * Berguna untuk fitur drag-and-drop pada Kanban board.
 *
 * @param {string} id - ID_Content yang akan diupdate statusnya
 * @param {string} newStatus - Status baru (Draft, In Progress, Review, Completed)
 * @returns {string} JSON string status update
 */
function updateContentStatus(id, newStatus, userName) {
  try {
    const sheet = getSheet('Master_Content');
    const data = sheet.getDataRange().getValues();

    // Cari baris berdasarkan ID_Content
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][MC_COL.ID].toString().trim() === id.toString().trim()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error(`Konten dengan ID "${id}" tidak ditemukan.`);
    }

    const title = data[rowIndex - 1][MC_COL.TITLE] || '';
    const oldStatus = data[rowIndex - 1][MC_COL.STATUS] || '';
    const oldBulanTahun = normalizeYearMonth(data[rowIndex - 1][MC_COL.BULAN_TAHUN]);

    // Validasi status yang diperbolehkan
    const validStatuses = ['Draft', 'In Progress', 'Review', 'Completed'];
    if (!validStatuses.includes(newStatus)) {
      throw new Error(`Status "${newStatus}" tidak valid. Gunakan: ${validStatuses.join(', ')}`);
    }

    const row = data[rowIndex - 1];

    // Update kolom Status saja (kolom ke-9, index 9 pada 1-based)
    sheet.getRange(rowIndex, MC_COL.STATUS + 1).setValue(newStatus);
    SpreadsheetApp.flush();

    // Update Google Calendar status
    const eventId = row[MC_COL.EVENT_ID] || '';
    if (eventId) {
      try {
        syncCalendarEvent({
          id: id,
          title: row[MC_COL.TITLE],
          platform: row[MC_COL.PLATFORM],
          ratio: row[MC_COL.RATIO],
          pic: row[MC_COL.PIC],
          date: row[MC_COL.DATE],
          status: newStatus,
          template: row[MC_COL.TEMPLATE],
          fileUrl: row[MC_COL.FILE_URL],
          description: row[MC_COL.DESCRIPTION],
          eventId: eventId
        });
      } catch (e) {
        Logger.log(`Gagal sync kalender saat updateContentStatus: ${e.message}`);
      }
    }

    // Bersihkan cache
    clearContentCache(oldBulanTahun);

    // Catat Log Aktivitas
    logActivity('Ubah Status', `Mengubah status "${title}" (${id}): "${oldStatus}" ➔ "${newStatus}"`, userName);

    Logger.log(`Status konten ${id} diubah menjadi "${newStatus}".`);
    return true;

  } catch (error) {
    Logger.log(`Error updateContentStatus: ${error.message}`);
    throw new Error(`Gagal mengupdate status: ${error.message}`);
  }
}

/**
 * Mengambil statistik konten untuk bulan tertentu.
 * Menghitung total, jumlah per status, dan jumlah per PIC.
 *
 * @param {string} yearMonth - Format 'YYYY-MM'
 * @returns {string} JSON string berisi objek statistik
 */
function getContentStats(yearMonth) {
  try {
    const sheet = getSheet('Master_Content');
    const data = sheet.getDataRange().getValues();

    // Filter data berdasarkan bulan
    const monthData = [];
    for (let i = 1; i < data.length; i++) {
      const rowBulanTahun = normalizeYearMonth(data[i][MC_COL.BULAN_TAHUN]);
      if (rowBulanTahun === yearMonth) {
        monthData.push(data[i]);
      }
    }

    // Hitung total
    const total = monthData.length;

    // Hitung per status
    const countByStatus = {
      draft: 0,
      inProgress: 0,
      review: 0,
      completed: 0
    };

    // Hitung per PIC menggunakan Map
    const picMap = new Map();

    for (const row of monthData) {
      const status = row[MC_COL.STATUS];
      const pic = row[MC_COL.PIC] || 'Unassigned';

      // Count by status
      switch (status) {
        case 'Draft':
          countByStatus.draft++;
          break;
        case 'In Progress':
          countByStatus.inProgress++;
          break;
        case 'Review':
          countByStatus.review++;
          break;
        case 'Completed':
          countByStatus.completed++;
          break;
      }

      // Count per PIC
      if (!picMap.has(pic)) {
        picMap.set(pic, { name: pic, total: 0, draft: 0, inProgress: 0, review: 0, completed: 0 });
      }
      const picStats = picMap.get(pic);
      picStats.total++;

      switch (status) {
        case 'Draft':
          picStats.draft++;
          break;
        case 'In Progress':
          picStats.inProgress++;
          break;
        case 'Review':
          picStats.review++;
          break;
        case 'Completed':
          picStats.completed++;
          break;
      }
    }

    // Konversi Map ke array
    const perPic = Array.from(picMap.values());

    const stats = {
      total: total,
      countByStatus: countByStatus,
      perPic: perPic
    };

    return stats;

  } catch (error) {
    Logger.log(`Error getContentStats: ${error.message}`);
    throw new Error(`Gagal mengambil statistik: ${error.message}`);
  }
}

/**
 * Mengambil semua data konten dari Master_Content tanpa filter.
 *
 * @returns {string} JSON string berisi array seluruh objek konten
 */
function getAllContent() {
  try {
    const sheet = getSheet('Master_Content');
    const data = sheet.getDataRange().getValues();

    const contents = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // Lewati baris kosong
      if (!row[MC_COL.ID] && !row[MC_COL.TITLE]) continue;

      contents.push({
        id: row[MC_COL.ID],
        date: row[MC_COL.DATE] instanceof Date ? Utilities.formatDate(row[MC_COL.DATE], Session.getScriptTimeZone(), "yyyy-MM-dd") : row[MC_COL.DATE],
        bulanTahun: normalizeYearMonth(row[MC_COL.BULAN_TAHUN]),
        pic: row[MC_COL.PIC],
        title: row[MC_COL.TITLE],
        platform: row[MC_COL.PLATFORM],
        template: row[MC_COL.TEMPLATE],
        ratio: row[MC_COL.RATIO],
        status: row[MC_COL.STATUS],
        fileUrl: row[MC_COL.FILE_URL] || '',
        format: row[MC_COL.FORMAT] || 'desain',
        description: row[MC_COL.DESCRIPTION] || '',
        eventId: row[MC_COL.EVENT_ID] || ''
      });
    }

    // Sort berdasarkan tanggal ascending
    contents.sort((a, b) => new Date(a.date) - new Date(b.date));

    return contents;

  } catch (error) {
    Logger.log(`Error getAllContent: ${error.message}`);
    throw new Error(`Gagal mengambil semua konten: ${error.message}`);
  }
}

/**
 * Diagnostic function to print database contents and properties.
 */
function getDebugInfo(yearMonth) {
  try {
    const sheet = getSheet('Master_Content');
    const data = sheet.getDataRange().getValues();
    const lastRow = sheet.getLastRow();
    const headers = data[0];
    
    let matchCount = 0;
    const rows = data.slice(1).map(row => {
      const bT = normalizeYearMonth(row[MC_COL.BULAN_TAHUN]);
      if (bT === yearMonth) {
        matchCount++;
      }
      return {
        id: row[MC_COL.ID],
        date: row[MC_COL.DATE] instanceof Date ? Utilities.formatDate(row[MC_COL.DATE], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(row[MC_COL.DATE]),
        bulanTahun: String(row[MC_COL.BULAN_TAHUN]),
        normalizedBulanTahun: bT,
        title: row[MC_COL.TITLE]
      };
    });
    
    return {
      success: true,
      lastRow: lastRow,
      totalRows: data.length,
      yearMonthPassed: yearMonth,
      matchCount: matchCount,
      spreadsheetTz: getSpreadsheet().getSpreadsheetTimeZone(),
      scriptTz: Session.getScriptTimeZone(),
      headers: headers,
      rows: rows.slice(0, 15) // limit to first 15 rows for display
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Menghapus beberapa konten sekaligus berdasarkan ID_Content.
 *
 * @param {Array<string>} ids - Array berisi ID_Content yang akan dihapus
 * @returns {Object} Status keberhasilan dan jumlah data yang terhapus
 */
function deleteMultipleContents(ids, userName) {
  try {
    if (!Array.isArray(ids) || ids.length === 0) {
      throw new Error('Daftar ID kosong atau tidak valid.');
    }

    const sheet = getSheet('Master_Content');
    const data = sheet.getDataRange().getValues();

    // Gunakan set untuk pencarian cepat dan trim spasi
    const idSet = new Set(ids.map(id => id.toString().trim()));
    const deletedBulanTahuns = new Set();

    let deletedCount = 0;
    // Cari baris dari bawah ke atas agar indeks baris tidak bergeser saat dihapus
    for (let i = data.length - 1; i >= 1; i--) {
      const rowId = data[i][MC_COL.ID].toString().trim();
      if (idSet.has(rowId)) {
        const bulanTahun = normalizeYearMonth(data[i][MC_COL.BULAN_TAHUN]);
        if (bulanTahun) {
          deletedBulanTahuns.add(bulanTahun);
        }
        
        const eventId = data[i][MC_COL.EVENT_ID] || '';
        if (eventId) {
          try {
            deleteCalendarEvent(eventId);
          } catch (e) {
            Logger.log(`Gagal menghapus event kalender saat deleteMultipleContents: ${e.message}`);
          }
        }

        sheet.deleteRow(i + 1); // +1 karena sheet 1-indexed
        deletedCount++;
      }
    }

    SpreadsheetApp.flush();

    // Bersihkan cache untuk semua Bulan_Tahun yang terpengaruh
    deletedBulanTahuns.forEach(bt => {
      clearContentCache(bt);
    });

    // Catat Log Aktivitas
    logActivity('Hapus Banyak Konten', `Menghapus ${deletedCount} konten secara massal`, userName);

    Logger.log(`${deletedCount} konten berhasil dihapus secara massal.`);
    return { success: true, count: deletedCount };

  } catch (error) {
    Logger.log(`Error deleteMultipleContents: ${error.message}`);
    throw new Error(`Gagal menghapus konten terpilih: ${error.message}`);
  }
}

