/**
 * TeamService.gs
 * ==============
 * Service layer untuk operasi CRUD pada sheet Team_Members.
 * Mengelola data anggota tim Kominfo.
 *
 * Kolom Team_Members:
 * Nama_Lengkap | Role | Email
 */

// Konstanta index kolom (0-based) untuk Team_Members
const TM_COL = {
  NAMA: 0,
  ROLE: 1,
  EMAIL: 2
};

/**
 * Mengambil semua anggota tim dari sheet Team_Members.
 *
 * @returns {string} JSON string berisi array objek anggota tim
 */
function getTeamMembers() {
  try {
    const sheet = getSheet('Team_Members');
    const data = sheet.getDataRange().getValues();

    const members = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // Lewati baris kosong
      if (!row[TM_COL.NAMA]) continue;

      members.push({
        name: row[TM_COL.NAMA],
        role: row[TM_COL.ROLE],
        email: row[TM_COL.EMAIL] || ''
      });
    }

    return members;

  } catch (error) {
    Logger.log(`Error getTeamMembers: ${error.message}`);
    throw new Error(`Gagal mengambil data tim: ${error.message}`);
  }
}

/**
 * Menambahkan anggota tim baru ke sheet Team_Members.
 *
 * @param {Object} data - Data anggota: {name, role, email}
 * @returns {string} JSON string status penambahan
 */
function addTeamMember(data, userName) {
  try {
    const sheet = getSheet('Team_Members');

    // Validasi input
    if (!data.name || !data.role) {
      throw new Error('Nama dan Role wajib diisi.');
    }

    // Cek duplikasi berdasarkan nama (opsional, untuk menghindari data ganda)
    const existingData = sheet.getDataRange().getValues();
    for (let i = 1; i < existingData.length; i++) {
      if (existingData[i][TM_COL.NAMA] === data.name) {
        throw new Error(`Anggota dengan nama "${data.name}" sudah ada.`);
      }
    }

    // Append baris baru
    const newRow = [
      data.name,
      data.role,
      data.email || ''
    ];
    sheet.appendRow(newRow);
    SpreadsheetApp.flush();

    // Catat Log Aktivitas
    logActivity('Tambah Anggota Tim', `Menambahkan anggota tim baru: "${data.name}" (${data.role})`, userName);

    Logger.log(`Anggota tim baru ditambahkan: ${data.name}`);
    return true;

  } catch (error) {
    Logger.log(`Error addTeamMember: ${error.message}`);
    throw new Error(`Gagal menambahkan anggota tim: ${error.message}`);
  }
}

/**
 * Mengupdate data anggota tim berdasarkan email.
 *
 * @param {string} oldEmail - Email lama untuk identifikasi baris
 * @param {Object} data - Data baru: {name, role, email}
 * @returns {string} JSON string status update
 */
function updateTeamMember(oldName, data, userName) {
  try {
    const sheet = getSheet('Team_Members');
    const sheetData = sheet.getDataRange().getValues();

    // Cari baris berdasarkan nama
    let rowIndex = -1;
    for (let i = 1; i < sheetData.length; i++) {
      if (sheetData[i][TM_COL.NAMA] === oldName) {
        rowIndex = i + 1; // +1 karena sheet 1-indexed
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error(`Anggota dengan nama "${oldName}" tidak ditemukan.`);
    }

    // Update baris
    const updatedRow = [
      data.name || sheetData[rowIndex - 1][TM_COL.NAMA],
      data.role || sheetData[rowIndex - 1][TM_COL.ROLE],
      data.email || ''
    ];

    sheet.getRange(rowIndex, 1, 1, 3).setValues([updatedRow]);
    SpreadsheetApp.flush();

    // Catat Log Aktivitas
    logActivity('Edit Anggota Tim', `Mengubah data anggota tim: "${updatedRow[0]}" (Role: ${updatedRow[1]}, Email: ${updatedRow[2]})`, userName);

    Logger.log(`Anggota tim diupdate: ${oldName} -> ${data.name}`);
    return true;

  } catch (error) {
    Logger.log(`Error updateTeamMember: ${error.message}`);
    throw new Error(`Gagal mengupdate anggota tim: ${error.message}`);
  }
}

/**
 * Menghapus anggota tim berdasarkan email.
 *
 * @param {string} email - Email anggota yang akan dihapus
 * @returns {string} JSON string status penghapusan
 */
function deleteTeamMember(name, userName) {
  try {
    const sheet = getSheet('Team_Members');
    const data = sheet.getDataRange().getValues();

    // Cari baris berdasarkan nama
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][TM_COL.NAMA] === name) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      throw new Error(`Anggota dengan nama "${name}" tidak ditemukan.`);
    }

    const memberName = data[rowIndex - 1][TM_COL.NAMA] || '';

    // Hapus baris
    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush();

    // Catat Log Aktivitas
    logActivity('Hapus Anggota Tim', `Menghapus anggota tim: "${memberName}"`, userName);

    Logger.log(`Anggota tim dihapus: ${name}`);
    return true;

  } catch (error) {
    Logger.log(`Error deleteTeamMember: ${error.message}`);
    throw new Error(`Gagal menghapus anggota tim: ${error.message}`);
  }
}

/**
 * Mendapatkan daftar PIC untuk dropdown di form konten.
 * Mengembalikan array {name, role} dari semua anggota tim.
 *
 * @returns {string} JSON string berisi array objek PIC
 */
function getPICList() {
  try {
    const sheet = getSheet('Team_Members');
    const data = sheet.getDataRange().getValues();

    const picList = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // Lewati baris kosong
      if (!row[TM_COL.NAMA]) continue;

      picList.push({
        name: row[TM_COL.NAMA],
        role: row[TM_COL.ROLE]
      });
    }

    return picList;

  } catch (error) {
    Logger.log(`Error getPICList: ${error.message}`);
    throw new Error(`Gagal mengambil daftar PIC: ${error.message}`);
  }
}
