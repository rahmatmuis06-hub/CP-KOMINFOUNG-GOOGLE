/**
 * FileService.gs
 * ==============
 * Service layer untuk operasi upload file ke Google Drive.
 * Menangani listing folder, upload file, dan penghapusan file.
 *
 * Root Folder ID sudah ditentukan.
 */

// ID folder Google Drive untuk masing-masing jenis karya
const FOLDER_DESAIN_ID = '1loi9fCX5xK_djBjQsJPwTGDHGfZELlOZ';
const FOLDER_VIDEO_ID = '1U86cbUTleuKzP8sEBBNfqlic4UJ28i6O';

/**
 * Mengambil daftar subfolder langsung di dalam folder tertentu (non-recursive / lazy load).
 *
 * @param {string} parentId - ID folder induk (bisa 'desain', 'video', atau ID subfolder riil)
 * @returns {Array<Object>} Array objek {id, name, parentId}
 */
function getSubFolders(parentId) {
  try {
    let parentFolderId = parentId;
    if (parentId === 'desain') {
      parentFolderId = FOLDER_DESAIN_ID;
    } else if (parentId === 'video') {
      parentFolderId = FOLDER_VIDEO_ID;
    } else if (!parentId) {
      throw new Error('Parent ID tidak boleh kosong.');
    }

    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const subFolders = parentFolder.getFolders();
    const folders = [];

    while (subFolders.hasNext()) {
      const folder = subFolders.next();
      folders.push({
        id: folder.getId(),
        name: folder.getName(),
        parentId: parentId
      });
    }

    // Urutkan berdasarkan nama A-Z
    folders.sort((a, b) => a.name.localeCompare(b.name));
    return folders;

  } catch (error) {
    Logger.log(`Error getSubFolders: ${error.message}`);
    throw new Error(`Gagal mengambil subfolder: ${error.message}`);
  }
}

/**
 * Membuat folder baru di Google Drive di dalam folder induk tertentu.
 *
 * @param {string} name - Nama folder baru yang akan dibuat
 * @param {string} parentId - ID folder induk (bisa 'desain', 'video', atau ID subfolder riil)
 * @returns {Object} Info folder yang berhasil dibuat
 */
function createDriveFolder(name, parentId) {
  try {
    if (!name || !name.trim()) {
      throw new Error('Nama folder tidak boleh kosong.');
    }

    let parentFolderId = parentId;
    if (parentId === 'desain') {
      parentFolderId = FOLDER_DESAIN_ID;
    } else if (parentId === 'video') {
      parentFolderId = FOLDER_VIDEO_ID;
    }

    const parentFolder = DriveApp.getFolderById(parentFolderId);
    const newFolder = parentFolder.createFolder(name.trim());

    Logger.log(`Folder baru "${name}" berhasil dibuat di bawah folder ID ${parentFolderId}.`);

    return {
      success: true,
      id: newFolder.getId(),
      name: newFolder.getName(),
      parentId: parentId
    };

  } catch (error) {
    Logger.log(`Error createDriveFolder: ${error.message}`);
    throw new Error(`Gagal membuat folder baru: ${error.message}`);
  }
}

/**
 * Upload file untuk konten tertentu ke Google Drive.
 * Otomatis mengubah status konten menjadi 'Completed' setelah berhasil upload.
 *
 * @param {string} contentId - ID_Content terkait
 * @param {string} folderId - ID subfolder tujuan di Google Drive (atau 'desain' / 'video')
 * @param {string} fileBase64 - Data file dalam format base64
 * @param {string} fileName - Nama file asli
 * @param {string} mimeType - MIME type file
 * @returns {Object} {success, fileUrl, folderName}
 */
function uploadFileForContent(contentId, folderId, fileBase64, fileName, mimeType) {
  try {
    // Validasi input
    if (!contentId || !folderId || !fileBase64 || !fileName) {
      throw new Error('Parameter upload tidak lengkap.');
    }

    // Tentukan folder ID tujuan
    let targetFolderId = folderId;
    let folderName = '';
    
    if (folderId === 'desain') {
      targetFolderId = FOLDER_DESAIN_ID;
      folderName = 'Desain (Utama)';
    } else if (folderId === 'video') {
      targetFolderId = FOLDER_VIDEO_ID;
      folderName = 'Video (Utama)';
    } else {
      const targetFolder = DriveApp.getFolderById(folderId);
      folderName = targetFolder.getName();
    }

    // Decode base64 ke blob
    const decodedBytes = Utilities.base64Decode(fileBase64);
    const blob = Utilities.newBlob(decodedBytes, mimeType, fileName);

    // Ambil folder tujuan
    const targetFolder = DriveApp.getFolderById(targetFolderId);

    // Upload file ke folder tujuan
    const uploadedFile = targetFolder.createFile(blob);
    
    // Set file agar bisa diakses oleh siapa saja yang memiliki link
    uploadedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileUrl = uploadedFile.getUrl();
    const fileId = uploadedFile.getId();

    // Update kolom File_URL dan Status di Master_Content
    const sheet = getSheet('Master_Content');
    const data = sheet.getDataRange().getValues();

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][MC_COL.ID].toString().trim() === contentId.toString().trim()) {
        rowIndex = i + 1; // 1-indexed
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error(`Konten dengan ID "${contentId}" tidak ditemukan.`);
    }

    // Update File_URL (kolom 10) dan Status (kolom 9) menjadi 'Completed'
    const existingFileUrlValue = data[rowIndex - 1][MC_COL.FILE_URL] || '';
    let updatedFileUrlValue = '';
    if (existingFileUrlValue.toString().trim()) {
      const existingUrls = existingFileUrlValue.toString().split('\n').map(u => u.trim()).filter(Boolean);
      existingUrls.push(fileUrl);
      updatedFileUrlValue = existingUrls.join('\n');
    } else {
      updatedFileUrlValue = fileUrl;
    }

    sheet.getRange(rowIndex, MC_COL.FILE_URL + 1).setValue(updatedFileUrlValue);
    sheet.getRange(rowIndex, MC_COL.STATUS + 1).setValue('Completed');
    SpreadsheetApp.flush();

    // Sinkronisasi ke Google Calendar (Update fileUrl dan status)
    const eventId = data[rowIndex - 1][MC_COL.EVENT_ID] || '';
    if (eventId) {
      try {
        syncCalendarEvent({
          id: contentId,
          title: data[rowIndex - 1][MC_COL.TITLE],
          platform: data[rowIndex - 1][MC_COL.PLATFORM],
          ratio: data[rowIndex - 1][MC_COL.RATIO],
          pic: data[rowIndex - 1][MC_COL.PIC],
          date: data[rowIndex - 1][MC_COL.DATE],
          status: 'Completed',
          template: data[rowIndex - 1][MC_COL.TEMPLATE],
          fileUrl: updatedFileUrlValue,
          description: data[rowIndex - 1][MC_COL.DESCRIPTION],
          eventId: eventId
        });
      } catch (e) {
        Logger.log(`Gagal sinkronisasi Google Calendar saat upload file: ${e.message}`);
      }
    }

    // Bersihkan cache
    const bulanTahun = normalizeYearMonth(data[rowIndex - 1][MC_COL.BULAN_TAHUN]);
    clearContentCache(bulanTahun);

    // Kirim notifikasi WhatsApp ke Grup
    try {
      let uploader = 'Anggota Tim';
      try {
        const activeEmail = Session.getActiveUser().getEmail();
        uploader = activeEmail ? activeEmail.split('@')[0] : 'Anggota Tim';
      } catch(e) {}
      notifyFileUploadWhatsApp(contentId, fileName, fileUrl, uploader);
    } catch (e) {
      Logger.log(`Gagal mengirim notifikasi WhatsApp upload file: ${e.message}`);
    }

    Logger.log(`File "${fileName}" berhasil diupload ke folder "${folderName}" untuk konten ${contentId}.`);

    return {
      success: true,
      fileUrl: fileUrl,
      fileId: fileId,
      folderName: folderName
    };

  } catch (error) {
    Logger.log(`Error uploadFileForContent: ${error.message}`);
    throw new Error(`Gagal mengupload file: ${error.message}`);
  }
}

/**
 * Menghapus file yang sudah diupload untuk konten tertentu berdasarkan URL file spesifik.
 *
 * @param {string} contentId - ID_Content terkait
 * @param {string} fileUrl - URL file spesifik yang ingin dihapus
 * @returns {Object} {success: true}
 */
function deleteUploadedFile(contentId, fileUrl, userName) {
  try {
    const sheet = getSheet('Master_Content');
    const data = sheet.getDataRange().getValues();

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][MC_COL.ID].toString().trim() === contentId.toString().trim()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error(`Konten dengan ID "${contentId}" tidak ditemukan.`);
    }

    const currentFileUrls = data[rowIndex - 1][MC_COL.FILE_URL] || '';
    const urlsArray = currentFileUrls.toString().split('\n').map(url => url.trim()).filter(Boolean);

    let fileName = '';
    // Hapus file dari Drive jika ada
    if (fileUrl) {
      try {
        // Extract file ID dari URL Google Drive
        const fileIdMatch = fileUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (fileIdMatch && fileIdMatch[1]) {
          const file = DriveApp.getFileById(fileIdMatch[1]);
          fileName = file.getName();
          file.setTrashed(true);
          Logger.log(`File "${fileName}" (${fileIdMatch[1]}) berhasil dihapus dari Drive.`);
        }
      } catch (driveErr) {
        Logger.log(`Warning: Gagal menghapus file dari Drive: ${driveErr.message}`);
        // Lanjutkan meskipun gagal hapus dari Drive
      }
    }

    const logDetails = fileName ? `Menghapus file: "${fileName}" dari konten ID: ${contentId}` : `Menghapus file lampiran dari konten ID: ${contentId}`;
    logActivity('Hapus File Konten', logDetails, userName);

    // Saring untuk membuang fileUrl yang dicocokkan
    const updatedUrlsArray = urlsArray.filter(url => url !== fileUrl);
    const updatedValue = updatedUrlsArray.join('\n');

    // Simpan kembali sisa URL
    sheet.getRange(rowIndex, MC_COL.FILE_URL + 1).setValue(updatedValue);
    
    // Jika semua file habis terhapus, status dikembalikan ke 'In Progress' jika saat ini 'Completed'
    let updatedStatus = data[rowIndex - 1][MC_COL.STATUS];
    if (updatedUrlsArray.length === 0) {
      const currentStatus = data[rowIndex - 1][MC_COL.STATUS];
      if (currentStatus === 'Completed') {
        sheet.getRange(rowIndex, MC_COL.STATUS + 1).setValue('In Progress');
        updatedStatus = 'In Progress';
      }
    }
    
    SpreadsheetApp.flush();

    // Update Google Calendar event
    const eventId = data[rowIndex - 1][MC_COL.EVENT_ID] || '';
    if (eventId) {
      try {
        syncCalendarEvent({
          id: contentId,
          title: data[rowIndex - 1][MC_COL.TITLE],
          platform: data[rowIndex - 1][MC_COL.PLATFORM],
          ratio: data[rowIndex - 1][MC_COL.RATIO],
          pic: data[rowIndex - 1][MC_COL.PIC],
          date: data[rowIndex - 1][MC_COL.DATE],
          status: updatedStatus,
          template: data[rowIndex - 1][MC_COL.TEMPLATE],
          fileUrl: updatedValue,
          description: data[rowIndex - 1][MC_COL.DESCRIPTION],
          eventId: eventId
        });
      } catch (e) {
        Logger.log(`Gagal sync kalender saat deleteUploadedFile: ${e.message}`);
      }
    }

    // Bersihkan cache
    const bulanTahun = normalizeYearMonth(data[rowIndex - 1][MC_COL.BULAN_TAHUN]);
    clearContentCache(bulanTahun);

    Logger.log(`File ${fileUrl} untuk konten ${contentId} berhasil dihapus.`);
    return { success: true };

  } catch (error) {
    Logger.log(`Error deleteUploadedFile: ${error.message}`);
    throw new Error(`Gagal menghapus file: ${error.message}`);
  }
}

/**
 * Menghapus seluruh file yang sudah diupload untuk konten tertentu sekaligus.
 *
 * @param {string} contentId - ID_Content terkait
 * @param {string} userName - Nama admin/user yang melakukan aksi
 * @returns {Object} {success: true, count: number}
 */
function deleteAllUploadedFiles(contentId, userName) {
  try {
    const sheet = getSheet('Master_Content');
    const data = sheet.getDataRange().getValues();

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][MC_COL.ID].toString().trim() === contentId.toString().trim()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error(`Konten dengan ID "${contentId}" tidak ditemukan.`);
    }

    const currentFileUrls = data[rowIndex - 1][MC_COL.FILE_URL] || '';
    if (!currentFileUrls.toString().trim()) {
      return { success: true, count: 0 };
    }
    
    const urlsArray = currentFileUrls.toString().split('\n').map(url => url.trim()).filter(Boolean);

    // Hapus semua file dari Drive
    let deletedCount = 0;
    for (const fileUrl of urlsArray) {
      try {
        const fileIdMatch = fileUrl.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (fileIdMatch && fileIdMatch[1]) {
          const file = DriveApp.getFileById(fileIdMatch[1]);
          file.setTrashed(true);
          deletedCount++;
        }
      } catch (driveErr) {
        Logger.log(`Warning: Gagal menghapus file dari Drive: ${driveErr.message}`);
      }
    }

    // Bersihkan kolom File_URL (kolom 10)
    sheet.getRange(rowIndex, MC_COL.FILE_URL + 1).setValue('');
    
    // Kembalikan status ke 'In Progress' jika saat ini 'Completed'
    let updatedStatus = data[rowIndex - 1][MC_COL.STATUS];
    const currentStatus = data[rowIndex - 1][MC_COL.STATUS];
    if (currentStatus === 'Completed') {
      sheet.getRange(rowIndex, MC_COL.STATUS + 1).setValue('In Progress');
      updatedStatus = 'In Progress';
    }
    
    SpreadsheetApp.flush();

    // Update Google Calendar event
    const eventId = data[rowIndex - 1][MC_COL.EVENT_ID] || '';
    if (eventId) {
      try {
        syncCalendarEvent({
          id: contentId,
          title: data[rowIndex - 1][MC_COL.TITLE],
          platform: data[rowIndex - 1][MC_COL.PLATFORM],
          ratio: data[rowIndex - 1][MC_COL.RATIO],
          pic: data[rowIndex - 1][MC_COL.PIC],
          date: data[rowIndex - 1][MC_COL.DATE],
          status: updatedStatus,
          template: data[rowIndex - 1][MC_COL.TEMPLATE],
          fileUrl: '',
          description: data[rowIndex - 1][MC_COL.DESCRIPTION],
          eventId: eventId
        });
      } catch (e) {
        Logger.log(`Gagal sync kalender saat deleteAllUploadedFiles: ${e.message}`);
      }
    }

    // Bersihkan cache
    const bulanTahun = normalizeYearMonth(data[rowIndex - 1][MC_COL.BULAN_TAHUN]);
    clearContentCache(bulanTahun);

    // Catat Log Aktivitas
    logActivity('Hapus Semua File Konten', `Menghapus seluruh berkas lampiran (${deletedCount} file) dari konten ID: ${contentId}`, userName);

    Logger.log(`Seluruh file (${deletedCount} file) untuk konten ${contentId} berhasil dihapus.`);
    return { success: true, count: deletedCount };

  } catch (error) {
    Logger.log(`Error deleteAllUploadedFiles: ${error.message}`);
    throw new Error(`Gagal menghapus semua file: ${error.message}`);
  }
}

/**
 * Mendapatkan detail (nama dan url) dari daftar URL file yang dipisahkan baris baru.
 *
 * @param {string} fileUrlsString - Daftar URL dipisahkan baris baru
 * @returns {Array<Object>} Array dari {name, url}
 */
function getFileDetails(fileUrlsString) {
  try {
    if (!fileUrlsString || !fileUrlsString.trim()) {
      return [];
    }
    const urls = fileUrlsString.toString().split('\n').map(u => u.trim()).filter(Boolean);
    const details = [];
    
    for (let i = 0; i < urls.length; i++) {
      const url = urls[i];
      let name = 'File ' + (i + 1);
      try {
        const fileIdMatch = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (fileIdMatch && fileIdMatch[1]) {
          const file = DriveApp.getFileById(fileIdMatch[1]);
          name = file.getName();
        }
      } catch (err) {
        Logger.log(`Gagal mengambil nama file untuk ${url}: ${err.message}`);
      }
      details.push({ name: name, url: url });
    }
    return details;
  } catch (error) {
    Logger.log(`Error getFileDetails: ${error.message}`);
    return [];
  }
}

