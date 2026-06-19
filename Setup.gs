/**
 * Setup.gs
 * ========
 * Fungsi setup satu kali (one-time) untuk inisialisasi database spreadsheet.
 * Jalankan initialSetup() dari Apps Script Editor untuk pertama kali.
 */

/**
 * Inisialisasi lengkap: membuat spreadsheet, sheet, header, validasi data, dan seed data tim.
 *
 * Langkah-langkah:
 * 1. Buat Spreadsheet baru dengan nama 'DB_Dashboard_Kominfo_[timestamp]'
 * 2. Buat sheet Master_Content dengan header kolom
 * 3. Buat sheet Team_Members dengan header kolom
 * 4. Hapus sheet default 'Sheet1'
 * 5. Simpan ID spreadsheet ke PropertiesService
 * 6. Setup data validation (dropdown) untuk kolom Platform, Ratio, Status
 * 7. Seed data anggota tim
 * 8. Log URL spreadsheet
 *
 * @returns {string} JSON string berisi status dan URL spreadsheet
 */
function initialSetup() {
  try {
    // 1. Buat spreadsheet baru dengan timestamp
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');
    const ssName = `DB_Dashboard_Kominfo_${timestamp}`;
    const ss = SpreadsheetApp.create(ssName);
    ss.setSpreadsheetTimeZone('Asia/Jakarta');
    const spreadsheetId = ss.getId();
    const spreadsheetUrl = ss.getUrl();

    Logger.log(`Spreadsheet dibuat: ${ssName}`);
    Logger.log(`ID: ${spreadsheetId}`);

    // 2. Buat sheet Master_Content dengan header
    const masterContentSheet = ss.getSheets()[0]; // Sheet pertama yang sudah ada
    masterContentSheet.setName('Master_Content');

    const contentHeaders = [
      'ID_Content',
      'Date_To_Be_Posted',
      'Bulan_Tahun',
      'PIC_Design',
      'Content_Title',
      'Platform',
      'Template',
      'Ratio',
      'Status',
      'File_URL',
      'Format',
      'Description',
      'Event_ID'
    ];

    masterContentSheet.getRange(1, 1, 1, contentHeaders.length).setValues([contentHeaders]);

    // Format header Master_Content
    const contentHeaderRange = masterContentSheet.getRange(1, 1, 1, contentHeaders.length);
    contentHeaderRange.setFontWeight('bold');
    contentHeaderRange.setBackground('#4285F4');
    contentHeaderRange.setFontColor('#FFFFFF');
    contentHeaderRange.setHorizontalAlignment('center');

    // Set lebar kolom agar rapi
    masterContentSheet.setColumnWidth(1, 130);  // ID_Content
    masterContentSheet.setColumnWidth(2, 150);  // Date_To_Be_Posted
    masterContentSheet.setColumnWidth(3, 100);  // Bulan_Tahun
    masterContentSheet.setColumnWidth(4, 130);  // PIC_Design
    masterContentSheet.setColumnWidth(5, 250);  // Content_Title
    masterContentSheet.setColumnWidth(6, 100);  // Platform
    masterContentSheet.setColumnWidth(7, 120);  // Template
    masterContentSheet.setColumnWidth(8, 80);   // Ratio
    masterContentSheet.setColumnWidth(9, 100);  // Status
    masterContentSheet.setColumnWidth(10, 200); // File_URL
    masterContentSheet.setColumnWidth(11, 120); // Format

    // Freeze header row
    masterContentSheet.setFrozenRows(1);

    Logger.log('Sheet Master_Content berhasil dibuat.');

    // 3. Buat sheet Team_Members dengan header
    const teamSheet = ss.insertSheet('Team_Members');

    const teamHeaders = ['Nama_Lengkap', 'Role', 'Email'];
    teamSheet.getRange(1, 1, 1, teamHeaders.length).setValues([teamHeaders]);

    // Format header Team_Members
    const teamHeaderRange = teamSheet.getRange(1, 1, 1, teamHeaders.length);
    teamHeaderRange.setFontWeight('bold');
    teamHeaderRange.setBackground('#34A853');
    teamHeaderRange.setFontColor('#FFFFFF');
    teamHeaderRange.setHorizontalAlignment('center');

    // Set lebar kolom
    teamSheet.setColumnWidth(1, 180);  // Nama_Lengkap
    teamSheet.setColumnWidth(2, 150);  // Role
    teamSheet.setColumnWidth(3, 250);  // Email

    // Freeze header row
    teamSheet.setFrozenRows(1);

    Logger.log('Sheet Team_Members berhasil dibuat.');

    // 3b. Buat sheet Activity_Log dengan header
    const logSheet = ss.insertSheet('Activity_Log');
    const logHeaders = ['Waktu', 'Pengguna', 'Aktivitas', 'Detail'];
    logSheet.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]);
    
    const logHeaderRange = logSheet.getRange(1, 1, 1, logHeaders.length);
    logHeaderRange.setFontWeight('bold');
    logHeaderRange.setBackground('#EA4335'); // Merah untuk log aktivitas
    logHeaderRange.setFontColor('#FFFFFF');
    logHeaderRange.setHorizontalAlignment('center');
    
    logSheet.setColumnWidth(1, 180);  // Waktu
    logSheet.setColumnWidth(2, 150);  // Pengguna
    logSheet.setColumnWidth(3, 150);  // Aktivitas
    logSheet.setColumnWidth(4, 350);  // Detail
    logSheet.setFrozenRows(1);
    
    Logger.log('Sheet Activity_Log berhasil dibuat.');

    // 4. Hapus sheet default 'Sheet1' jika ada
    const defaultSheet = ss.getSheetByName('Sheet1');
    if (defaultSheet) {
      ss.deleteSheet(defaultSheet);
      Logger.log('Sheet1 default berhasil dihapus.');
    }

    // 5. Simpan ID spreadsheet ke PropertiesService
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheetId);
    Logger.log('SPREADSHEET_ID tersimpan di PropertiesService.');

    // 6. Setup data validation (dropdown) untuk Master_Content
    setupDataValidation_(masterContentSheet);

    // 7. Seed data anggota tim
    seedTeamMembers(teamSheet);

    // 8. Log URL spreadsheet
    Logger.log(`=== SETUP BERHASIL ===`);
    Logger.log(`Spreadsheet URL: ${spreadsheetUrl}`);
    Logger.log(`Buka spreadsheet di: ${spreadsheetUrl}`);

    // 9. Return informasi setup
    return {
      success: true,
      spreadsheetUrl: spreadsheetUrl
    };

  } catch (error) {
    Logger.log(`Error initialSetup: ${error.message}`);
    throw new Error(`Gagal melakukan setup: ${error.message}`);
  }
}

/**
 * Setup data validation (dropdown) untuk kolom Platform, Ratio, dan Status di Master_Content.
 * Fungsi internal, hanya dipanggil oleh initialSetup().
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - Sheet Master_Content
 * @private
 */
function setupDataValidation_(sheet) {
  // Tentukan jumlah baris untuk validasi (1000 baris cukup untuk awal)
  const maxRows = 1000;

  // Platform validation (kolom 6)
  const platformRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Flyer', 'Reels', 'TikTok', 'Feeds', 'Story'], true)
    .setAllowInvalid(false)
    .setHelpText('Pilih platform: Flyer, Reels, TikTok, Feeds, atau Story')
    .build();
  sheet.getRange(2, 6, maxRows, 1).setDataValidation(platformRule);

  // Ratio validation (kolom 8)
  const ratioRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['16:9', '1:1', '4:5', '9:16'], true)
    .setAllowInvalid(false)
    .setHelpText('Pilih ratio: 16:9, 1:1, 4:5, atau 9:16')
    .build();
  sheet.getRange(2, 8, maxRows, 1).setDataValidation(ratioRule);

  // Status validation (kolom 9)
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Draft', 'In Progress', 'Review', 'Completed'], true)
    .setAllowInvalid(false)
    .setHelpText('Pilih status: Draft, In Progress, Review, atau Completed')
    .build();
  sheet.getRange(2, 9, maxRows, 1).setDataValidation(statusRule);

  Logger.log('Data validation berhasil di-setup untuk Platform, Ratio, dan Status.');
}

/**
 * Memasukkan data awal anggota tim Kominfo ke sheet Team_Members.
 * Menggunakan batch write untuk efisiensi.
 */
function seedTeamMembers(optSheet) {
  try {
    const sheet = optSheet || getSheet('Team_Members');

    // Data anggota tim awal
    const teamData = [
      ['Fatha',    'Kadiv',             ''],
      ['Bela',     'Admin',             ''],
      ['Anti',     'Admin',             ''],
      ['Riska',    'Admin',             ''],
      ['Beby',     'Admin',             ''],
      ['Adibah',   'Admin',             ''],
      ['Surya',    'Graphic Designer',  ''],
      ['Sulfajri', 'Graphic Designer',  ''],
      ['Nadine',   'Graphic Designer',  ''],
      ['Gibran',   'Graphic Designer',  ''],
      ['Bintang',  'Graphic Designer',  ''],
      ['Ijep',     'Videographer',      ''],
      ['Edo',      'Videographer',      ''],
      ['Hendro',   'Videographer',      '']
    ];

    // Batch write mulai dari baris 2 (setelah header)
    sheet.getRange(2, 1, teamData.length, 3).setValues(teamData);

    Logger.log(`${teamData.length} anggota tim berhasil di-seed ke Team_Members.`);

  } catch (error) {
    Logger.log(`Error seedTeamMembers: ${error.message}`);
    throw error;
  }
}

/**
 * Mengecek status setup apakah spreadsheet sudah dikonfigurasi.
 *
 * @returns {string} JSON string berisi status setup dan URL spreadsheet
 */
function getSetupStatus() {
  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');

    if (spreadsheetId) {
      // Coba buka spreadsheet untuk verifikasi masih valid
      try {
        const ss = SpreadsheetApp.openById(spreadsheetId);
        const url = ss.getUrl();
        return {
          isSetup: true,
          spreadsheetUrl: url
        };
      } catch (e) {
        // Spreadsheet ID ada tapi tidak bisa dibuka (mungkin sudah dihapus)
        return {
          isSetup: false,
          spreadsheetUrl: null
        };
      }
    }

    return {
      isSetup: false,
      spreadsheetUrl: null
    };

  } catch (error) {
    Logger.log(`Error getSetupStatus: ${error.message}`);
    throw new Error(`Gagal mengecek status setup: ${error.message}`);
  }
}
