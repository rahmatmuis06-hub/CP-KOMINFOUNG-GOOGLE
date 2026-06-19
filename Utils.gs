/**
 * Utils.gs
 * ========
 * Fungsi-fungsi helper/utilitas untuk Dashboard Content Planning Kominfo.
 * Menyediakan akses spreadsheet, generate ID, format tanggal, cache, dan response.
 */

// Cache global untuk menyimpan instance spreadsheet selama satu kali eksekusi request.
let _cachedSpreadsheet = null;

/**
 * Mendapatkan spreadsheet database berdasarkan ID yang tersimpan di PropertiesService.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} Objek Spreadsheet
 */
function getSpreadsheet() {
  if (_cachedSpreadsheet) {
    return _cachedSpreadsheet;
  }
  const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!spreadsheetId) {
    throw new Error('SPREADSHEET_ID belum dikonfigurasi. Jalankan initialSetup() terlebih dahulu.');
  }
  _cachedSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
  return _cachedSpreadsheet;
}

/**
 * Mendapatkan sheet tertentu berdasarkan nama.
 * @param {string} sheetName - Nama sheet yang dicari
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} Objek Sheet
 */
function getSheet(sheetName) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" tidak ditemukan.`);
  }
  return sheet;
}

/**
 * Auto-generate ID konten dengan format CNT-YYYY-NNN.
 * YYYY = tahun berjalan, NNN = nomor urut zero-padded berdasarkan jumlah baris di Master_Content.
 * @returns {string} ID konten yang di-generate (contoh: CNT-2026-001)
 */
function getLastContentNumber() {
  const sheet = getSheet('Master_Content');
  const lastRow = sheet.getLastRow();
  let maxNum = 0;
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < data.length; i++) {
      const id = String(data[i][0]).trim();
      const match = id.match(/CNT-\d{4}-(\d+)/);
      if (match) {
        const num = parseInt(match[1], 10);
        if (num > maxNum) {
          maxNum = num;
        }
      }
    }
  }
  return maxNum;
}

/**
 * Auto-generate ID konten dengan format CNT-YYYY-NNN.
 * YYYY = tahun berjalan, NNN = nomor urut zero-padded berdasarkan ID tertinggi.
 * @returns {string} ID konten yang di-generate (contoh: CNT-2026-001)
 */
function generateContentId() {
  const lastNum = getLastContentNumber();
  const year = new Date().getFullYear();
  const nextNumber = lastNum + 1;
  const paddedNumber = String(nextNumber).padStart(3, '0');
  return `CNT-${year}-${paddedNumber}`;
}

/**
 * Mendapatkan tahun-bulan saat ini dalam format 'YYYY-MM'.
 * @returns {string} Format 'YYYY-MM' (contoh: '2026-06')
 */
function getCurrentYearMonth() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, '0');
  return `${year}-${month}`;
}

/**
 * Mengkonversi objek Date menjadi string format 'YYYY-MM'.
 * @param {Date} date - Objek Date yang akan diformat
 * @returns {string} Format 'YYYY-MM'
 */
function formatDateToYearMonth(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    // Coba parse jika string
    date = new Date(date);
    if (isNaN(date.getTime())) {
      return '';
    }
  }
  try {
    const tz = getSpreadsheet().getSpreadsheetTimeZone();
    return Utilities.formatDate(date, tz, 'yyyy-MM');
  } catch (e) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    return `${year}-${month}`;
  }
}

/**
 * Mengkonversi Date ke format tampilan 'DD MMM YYYY' dengan nama bulan Indonesia.
 * Contoh: 1 Jan 2026, 15 Mei 2026, 30 Agu 2026
 * @param {Date} date - Objek Date yang akan diformat
 * @returns {string} Format 'DD MMM YYYY'
 */
function formatDateDisplay(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    date = new Date(date);
    if (isNaN(date.getTime())) {
      return '';
    }
  }

  // Nama bulan singkat dalam Bahasa Indonesia
  const monthNames = [
    'Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun',
    'Jul', 'Agu', 'Sep', 'Okt', 'Nov', 'Des'
  ];

  const day = date.getDate();
  const month = monthNames[date.getMonth()];
  const year = date.getFullYear();

  return `${day} ${month} ${year}`;
}

/**
 * Membersihkan cache data konten dari CacheService.
 * Dipanggil setiap kali ada perubahan data (add/update/delete).
 * @param {string} [yearMonth] - Format 'YYYY-MM' (opsional). Jika diisi, hanya membersihkan bulan tersebut.
 */
function clearContentCache(yearMonth) {
  const cache = CacheService.getScriptCache();
  if (yearMonth) {
    cache.remove(`content_${yearMonth}`);
    Logger.log(`Cache untuk ${yearMonth} dibersihkan.`);
  } else {
    // Hapus cache untuk semua kemungkinan bulan (12 bulan tahun ini dan tahun lalu)
    const currentYear = new Date().getFullYear();
    const keysToRemove = [];

    for (let year = currentYear - 1; year <= currentYear + 1; year++) {
      for (let month = 1; month <= 12; month++) {
        const monthStr = String(month).padStart(2, '0');
        keysToRemove.push(`content_${year}-${monthStr}`);
      }
    }

    // Hapus semua cache key yang mungkin ada
    cache.removeAll(keysToRemove);
    Logger.log('Semua cache konten dibersihkan.');
  }
}

/**
 * Membuat response standar dalam format JSON string.
 * @param {boolean} success - Status keberhasilan operasi
 * @param {*} data - Data yang dikembalikan (opsional)
 * @param {string} message - Pesan informasi (opsional)
 * @returns {string} JSON string dari objek response
 */
function sendResponse(success, data, message) {
  return JSON.stringify({
    success: success,
    data: data || null,
    message: message || ''
  });
}

/**
 * Menormalisasi nilai Bulan_Tahun dari spreadsheet (bisa berupa Date atau String)
 * menjadi format string 'YYYY-MM'.
 * @param {*} val - Nilai dari kolom Bulan_Tahun
 * @returns {string} Format 'YYYY-MM'
 */
function normalizeYearMonth(val) {
  if (val instanceof Date) {
    return formatDateToYearMonth(val);
  }
  if (val) {
    const str = String(val).trim();
    // Jika format 'YYYY-MM'
    if (str.match(/^\d{4}-\d{2}/)) {
      return str.substring(0, 7);
    }
    
    // Coba parse sebagai tanggal/date string
    const parsedDate = new Date(str);
    if (!isNaN(parsedDate.getTime())) {
      return formatDateToYearMonth(parsedDate);
    }

    // Coba deteksi jika nama bulan Indonesia (misal: "Juni 2026", "Jun 2026")
    const monthsIndo = {
      'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04', 'mei': '05', 'jun': '06',
      'jul': '07', 'agu': '08', 'sep': '09', 'okt': '10', 'nov': '11', 'des': '12',
      'agt': '08'
    };
    const lowerStr = str.toLowerCase();
    for (const key in monthsIndo) {
      if (lowerStr.includes(key)) {
        const yearMatch = lowerStr.match(/\d{4}/);
        if (yearMatch) {
          return `${yearMatch[0]}-${monthsIndo[key]}`;
        }
      }
    }
    return str;
  }
  return '';
}

/**
 * Mencatat aktivitas pengguna ke sheet Activity_Log.
 * Jika sheet Activity_Log belum ada, fungsi akan membuatnya secara otomatis.
 *
 * @param {string} activity - Jenis aktivitas (misal: 'Tambah Konten', 'Ubah Status')
 * @param {string} details - Detail aktivitas (misal: 'Menambahkan konten dengan ID CNT-2026-001')
 * @param {string} [userName] - Nama pengguna yang melakukan aksi (opsional, jika kosong diambil dari profil aktif)
 */
function logActivity(activity, details, userName) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('Activity_Log');
    
    // Jika sheet log belum ada, buat secara otomatis
    if (!sheet) {
      sheet = ss.insertSheet('Activity_Log');
      
      const headers = ['Waktu', 'Pengguna', 'Aktivitas', 'Detail'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format header
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#EA4335'); // Merah untuk log aktivitas
      headerRange.setFontColor('#FFFFFF');
      headerRange.setHorizontalAlignment('center');
      
      sheet.setColumnWidth(1, 180);  // Waktu
      sheet.setColumnWidth(2, 150);  // Pengguna
      sheet.setColumnWidth(3, 150);  // Aktivitas
      sheet.setColumnWidth(4, 350);  // Detail
      sheet.setFrozenRows(1);
    }
    
    // Gunakan user dari parameter, jika kosong coba ambil dari Google Session
    let actor = userName || '';
    if (!actor) {
      const email = Session.getActiveUser().getEmail();
      actor = email ? email : 'Anonymous';
    }
    
    const timestamp = new Date();
    
    // Append baris ke sheet log
    sheet.appendRow([timestamp, actor, activity, details]);
    SpreadsheetApp.flush();
    
    Logger.log(`Log dicatat: ${actor} - ${activity} - ${details}`);
    
  } catch (error) {
    Logger.log(`Gagal mencatat log aktivitas: ${error.message}`);
  }
}

/**
 * Mengambil log aktivitas untuk ditampilkan di Admin Area.
 * Mengembalikan array log terurut dari yang terbaru (limit 200 log).
 *
 * @returns {Array<Object>} Daftar log aktivitas
 */
function getActivityLogs() {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Activity_Log');
    if (!sheet) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return []; // Hanya ada header
    }
    
    const logs = [];
    const tz = ss.getSpreadsheetTimeZone();
    
    // Baca dari baris terakhir (terbaru) ke atas (lewatkan baris pertama/header)
    const limit = Math.max(1, data.length - 200);
    for (let i = data.length - 1; i >= limit; i--) {
      const row = data[i];
      let formattedTime = '';
      if (row[0] instanceof Date) {
        try {
          formattedTime = Utilities.formatDate(row[0], tz, 'dd MMM yyyy, HH:mm');
        } catch(e) {
          formattedTime = row[0].toLocaleString();
        }
      } else {
        formattedTime = String(row[0] || '');
      }
      
      logs.push({
        time: formattedTime,
        user: String(row[1] || 'Anonymous'),
        activity: String(row[2] || ''),
        details: String(row[3] || '')
      });
    }
    
    return logs;
  } catch (error) {
    Logger.log(`Error getActivityLogs: ${error.message}`);
    throw new Error(`Gagal memuat log aktivitas: ${error.message}`);
  }
}

/**
 * Mendapatkan peran pengguna aktif saat ini.
 * Memeriksa apakah pengguna memiliki hak akses Admin/Kadiv.
 *
 * @returns {Object} Peran pengguna: {email, role, isAdmin, spreadsheetUrl}
 */
function getUserRole() {
  try {
    const activeEmail = Session.getActiveUser().getEmail();
    const effectiveEmail = Session.getEffectiveUser().getEmail();
    
    let spreadsheetId = '';
    let spreadsheetUrl = null;
    let spreadsheetError = null;
    try {
      spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
      if (spreadsheetId) {
        const ss = SpreadsheetApp.openById(spreadsheetId);
        spreadsheetUrl = ss.getUrl();
      } else {
        spreadsheetError = 'SPREADSHEET_ID belum dikonfigurasi di Script Properties.';
      }
    } catch(e) {
      spreadsheetError = e.message;
      Logger.log("Gagal membuka spreadsheet dalam getUserRole: " + e.message);
    }
    
    // Jika pengguna adalah pemilik script (effective user), otomatis Admin
    if (activeEmail && activeEmail === effectiveEmail) {
      return {
        email: activeEmail,
        role: 'Admin (Owner)',
        isAdmin: true,
        spreadsheetUrl: spreadsheetUrl,
        spreadsheetError: spreadsheetError
      };
    }
    
    // Jika email kosong (akses publik tanpa login/otorisasi), default bukan admin
    if (!activeEmail) {
      return {
        email: '',
        role: 'User (Tamu)',
        isAdmin: false,
        spreadsheetUrl: spreadsheetUrl,
        spreadsheetError: spreadsheetError
      };
    }
    
    // Cari di Team_Members sheet
    if (spreadsheetId && !spreadsheetError) {
      try {
        const sheet = getSheet('Team_Members');
        const data = sheet.getDataRange().getValues();
        for (let i = 1; i < data.length; i++) {
          const name = String(data[i][0]);
          const role = String(data[i][1]);
          const email = String(data[i][2]).toLowerCase().trim();
          
          if (email === activeEmail.toLowerCase().trim()) {
            const isAdmin = role.toLowerCase().includes('admin') || role.toLowerCase().includes('kadiv');
            return {
              email: activeEmail,
              role: `${name} (${role})`,
              isAdmin: isAdmin,
              spreadsheetUrl: spreadsheetUrl,
              spreadsheetError: spreadsheetError
            };
          }
        }
      } catch(e) {
        spreadsheetError = "Gagal memuat Team_Members: " + e.message;
      }
    }
    
    return {
      email: activeEmail,
      role: 'User (Luar Tim)',
      isAdmin: false,
      spreadsheetUrl: spreadsheetUrl,
      spreadsheetError: spreadsheetError
    };
    
  } catch (error) {
    Logger.log(`Error getUserRole: ${error.message}`);
    // Fallback aman
    return {
      email: '',
      role: 'User (Error)',
      isAdmin: false,
      spreadsheetUrl: null,
      spreadsheetError: error.message
    };
  }
}

/**
 * Mengubah koneksi database ke spreadsheet baru berdasarkan ID atau URL.
 *
 * @param {string} urlOrId - URL lengkap atau ID Spreadsheet tujuan
 * @param {string} userName - Nama admin yang melakukan perubahan
 * @returns {Object} Status koneksi: {success, title, needsSetup}
 */
function updateSpreadsheetLink(urlOrId, userName) {
  try {
    if (!urlOrId || !urlOrId.trim()) {
      throw new Error('URL atau ID Spreadsheet tidak boleh kosong.');
    }
    
    let targetId = urlOrId.trim();
    // Ekstrak ID jika diinput berupa URL lengkap
    if (urlOrId.includes('/d/')) {
      const parts = urlOrId.split('/d/');
      if (parts.length > 1) {
        targetId = parts[1].split('/')[0];
      }
    }
    
    // Coba buka spreadsheet untuk memastikan ID valid dan dapat diakses
    let ss;
    try {
      ss = SpreadsheetApp.openById(targetId);
    } catch(e) {
      throw new Error('Spreadsheet tidak ditemukan atau Anda tidak memiliki izin akses (Share) ke Spreadsheet tersebut.');
    }
    
    const spreadsheetTitle = ss.getName();
    
    // Periksa apakah spreadsheet memiliki struktur database yang benar
    const hasMaster = ss.getSheetByName('Master_Content') !== null;
    const hasTeam = ss.getSheetByName('Team_Members') !== null;
    const needsSetup = !hasMaster || !hasTeam;
    
    // Simpan ID spreadsheet baru ke Script Properties
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', targetId);
    
    // Bersihkan semua cache konten lama agar data baru langsung tampil
    clearContentCache();
    
    // Catat log aktivitas perubahan database
    logActivity('Ubah Link Database', `Menghubungkan ke database baru: "${spreadsheetTitle}" (${targetId})`, userName);
    
    return {
      success: true,
      title: spreadsheetTitle,
      url: ss.getUrl(),
      needsSetup: needsSetup
    };
    
  } catch (error) {
    Logger.log(`Error updateSpreadsheetLink: ${error.message}`);
    throw new Error(error.message);
  }
}


