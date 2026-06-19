/**
 * ExternalDataService.gs
 * =======================
 * Service layer untuk menarik data dari spreadsheet eksternal:
 * 1. Daftar Program Kerja (Proker) Divisi Kominfo.
 * 2. Daftar Kepanitiaan Anggota Kominfo.
 */

// Konstanta ID Spreadsheet Eksternal diambil dari Script Properties
const SPREADSHEET_PROKER_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_PROKER_ID') || '';
const SPREADSHEET_KEPANITIAAN_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_KEPANITIAAN_ID') || '';

/**
 * Mengambil data Program Kerja (Proker) dari spreadsheet eksternal.
 * 
 * @returns {Array<Object>} Daftar Program Kerja
 */
function getProkerData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_PROKER_ID);
    const sheet = ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();

    if (data.length <= 5) {
      return [];
    }

    // Array nama bulan sesuai urutan kolom timeline proker (dimulai dari indeks kolom ke-4 / B)
    const monthNames = [
      'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 
      'Agustus', 'September', 'Oktober', 'November', 'Desember'
    ];

    // Temukan baris header utama (No., Program Kerja, Pelaksanaan)
    let headerRowIdx = -1;
    for (let i = 0; i < data.length; i++) {
      const rowStr = data[i].map(c => String(c).toLowerCase().trim());
      if (rowStr.includes('program kerja') || rowStr.includes('no.')) {
        headerRowIdx = i;
        break;
      }
    }

    if (headerRowIdx === -1) {
      headerRowIdx = 5; // Fallback jika tidak ditemukan
    }

    const prokerList = [];
    
    // Data dimulai dari baris setelah sub-header nomor minggu (headerRowIdx + 2)
    for (let i = headerRowIdx + 2; i < data.length; i++) {
      const row = data[i];
      const no = String(row[1] || '').trim();
      const name = String(row[2] || '').trim();
      const execution = String(row[3] || '').trim();

      // Jika kolom Program Kerja kosong, lewati baris
      if (!name || name.toLowerCase() === 'program kerja') continue;

      // Cari bulan-bulan aktif berdasarkan ceklis pada kolom-kolom timeline (indeks 4 s.d selesai)
      const activeMonths = [];
      for (let col = 4; col < row.length; col++) {
        const cellVal = String(row[col] || '').trim();
        if (cellVal !== '') {
          // Setiap bulan memakan 5 sub-kolom (minggu 1 - 5)
          const monthIdx = Math.floor((col - 4) / 5);
          const monthName = monthNames[monthIdx];
          if (monthName && !activeMonths.includes(monthName)) {
            activeMonths.push(monthName);
          }
        }
      }

      prokerList.push({
        rowIdx: i + 1,
        no: no,
        name: name,
        execution: execution,
        activeMonths: activeMonths
      });
    }

    Logger.log(`Berhasil menarik ${prokerList.length} data Proker dari Spreadsheet.`);
    return prokerList;

  } catch (error) {
    Logger.log(`Error getProkerData: ${error.message}`);
    throw new Error(`Gagal menarik data Proker: ${error.message}`);
  }
}

/**
 * Mengambil data Kepanitiaan dari spreadsheet eksternal.
 * Mengelompokkan aktivitas/event per bulan, serta menghitung keaktifan anggota secara dinamis.
 * 
 * @returns {Object} Data kepanitiaan per bulan dan leaderboard keaktifan
 */
function getKepanitiaanData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_KEPANITIAAN_ID);
    const sheet = ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();

    if (data.length <= 3) {
      return { kepanitiaan: [], leaderboard: [] };
    }

    // Ambil daftar nama anggota tim aktif untuk referensi pencocokan nama panitia
    let teamNames = [];
    try {
      const members = getTeamMembers();
      teamNames = members.map(m => m.name.toLowerCase().trim());
    } catch(e) {
      Logger.log(`Warning: Gagal memuat daftar anggota tim internal, menggunakan fallback nama. ${e.message}`);
      // Fallback default nama dari seed data
      teamNames = [
        'fatha', 'bela', 'anti', 'riska', 'beby', 'adibah', 
        'surya', 'sulfajri', 'nadine', 'gibran', 'bintang', 
        'ijep', 'edo', 'hendro', 'sul'
      ];
    }

    // Cari baris header bulan
    let headerRowIdx = -1;
    for (let i = 0; i < data.length; i++) {
      const rowLower = data[i].map(c => String(c).toUpperCase().trim());
      if (rowLower.includes('FEBRUARI') || rowLower.includes('MARET')) {
        headerRowIdx = i;
        break;
      }
    }

    if (headerRowIdx === -1) {
      headerRowIdx = 3; // Fallback jika tidak ditemukan
    }

    const months = [
      'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI', 'JULI', 
      'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER'
    ];

    const kepanitiaanList = [];
    const workloadCounts = {}; // Menyimpan hitungan keaktifan per anggota

    // Iterasi kolom B s.d L (indeks 1 s.d 11) untuk masing-masing bulan
    for (let col = 1; col <= 11; col++) {
      const monthName = String(data[headerRowIdx][col] || '').trim().toUpperCase();
      if (!months.includes(monthName)) continue;

      const eventsInMonth = [];
      let currentEvent = null;

      // Iterasi ke bawah untuk mengambil data kegiatan dan panitia di bawah bulan aktif
      for (let rowIdx = headerRowIdx + 1; rowIdx < data.length; rowIdx++) {
        const cellVal = String(data[rowIdx][col] || '').trim();
        if (cellVal === '') continue;

        const cellValLower = cellVal.toLowerCase();
        
        // Cek apakah cell tersebut berisi nama anggota tim (panitia)
        const isMember = teamNames.some(name => cellValLower === name || cellValLower.includes(name));

        if (!isMember) {
          // Jika BUKAN anggota tim, berarti ini adalah Judul Event/Kegiatan baru
          currentEvent = {
            title: cellVal,
            members: []
          };
          eventsInMonth.push(currentEvent);
        } else {
          // Jika ADALAH anggota tim, tambahkan ke anggota kepanitiaan event aktif saat ini
          if (currentEvent) {
            // Cari nama asli anggota tim dari referensi (agar format huruf/kapital seragam)
            let formattedName = cellVal;
            try {
              const matchedMember = getTeamMembers().find(m => m.name.toLowerCase().trim() === cellValLower || cellValLower.includes(m.name.toLowerCase().trim()));
              if (matchedMember) {
                formattedName = matchedMember.name;
              }
            } catch(e) {}

            currentEvent.members.push(formattedName);

            // Tambahkan hitungan beban kerja (workload) untuk leaderboard
            workloadCounts[formattedName] = (workloadCounts[formattedName] || 0) + 1;
          }
        }
      }

      // Hanya masukkan bulan yang memiliki kegiatan/event
      if (eventsInMonth.length > 0) {
        kepanitiaanList.push({
          month: monthName,
          events: eventsInMonth
        });
      }
    }

    // Konversi objek workload ke array leaderboard terurut
    const leaderboard = Object.keys(workloadCounts).map(name => {
      return {
        name: name,
        count: workloadCounts[name]
      };
    });

    // Urutkan leaderboard dari keaktifan tertinggi ke terendah
    leaderboard.sort((a, b) => b.count - a.count);

    Logger.log(`Berhasil memuat data Kepanitiaan: ${kepanitiaanList.length} bulan aktif, ${leaderboard.length} panitia terdaftar.`);

    return {
      kepanitiaan: kepanitiaanList,
      leaderboard: leaderboard
    };

  } catch (error) {
    Logger.log(`Error getKepanitiaanData: ${error.message}`);
    throw new Error(`Gagal menarik data Kepanitiaan: ${error.message}`);
  }
}

/**
 * Menyimpan data Program Kerja (Proker) ke spreadsheet eksternal.
 * Bisa berupa penambahan baru atau update data lama.
 * 
 * @param {Object} proker - Data proker { rowIdx, no, name, execution, activeMonths }
 * @returns {Object} Hasil sukses
 */
function saveProker(proker) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_PROKER_ID);
    const sheet = ss.getSheets()[0];
    
    const monthNames = [
      'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 
      'Agustus', 'September', 'Oktober', 'November', 'Desember'
    ];
    
    let rowIdx = proker.rowIdx;
    
    if (rowIdx === null || rowIdx === undefined || rowIdx === '' || rowIdx < 0) {
      // Tambah Baru: Append ke baris terakhir
      const lastRow = sheet.getLastRow();
      rowIdx = lastRow + 1;
    }
    
    // Tulis data utama
    sheet.getRange(rowIdx, 2).setValue(proker.no); // Kolom B (No)
    sheet.getRange(rowIdx, 3).setValue(proker.name); // Kolom C (Program Kerja)
    sheet.getRange(rowIdx, 4).setValue(proker.execution); // Kolom D (Pelaksanaan)
    
    // Tulis timeline bulan (Kolom E s.d Kolom BB)
    for (let monthIdx = 0; monthIdx < monthNames.length; monthIdx++) {
      const monthName = monthNames[monthIdx];
      const startCol = 5 + (monthIdx * 5); // Kolom pertama di masing-masing bulan
      
      const isMonthActive = proker.activeMonths.includes(monthName);
      
      if (isMonthActive) {
        sheet.getRange(rowIdx, startCol).setValue("v"); // Tandai checklist 'v' di minggu pertama
        // Kosongkan 4 minggu lainnya di bulan tersebut
        for (let w = 1; w < 5; w++) {
          sheet.getRange(rowIdx, startCol + w).setValue("");
        }
      } else {
        // Kosongkan semua minggu di bulan tersebut
        for (let w = 0; w < 5; w++) {
          sheet.getRange(rowIdx, startCol + w).setValue("");
        }
      }
    }
    
    Logger.log(`Berhasil menyimpan proker: ${proker.name} di baris ${rowIdx}`);
    return { success: true, rowIdx: rowIdx };
  } catch (error) {
    Logger.log(`Error saveProker: ${error.message}`);
    throw new Error(`Gagal menyimpan data Proker: ${error.message}`);
  }
}

/**
 * Menambahkan kegiatan kepanitiaan baru di bawah bulan tertentu pada spreadsheet eksternal.
 * 
 * @param {Object} eventData - { month, title, members }
 * @returns {Object} Hasil sukses
 */
function addKepanitiaanEvent(eventData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_KEPANITIAAN_ID);
    const sheet = ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    
    const months = [
      'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI', 'JULI', 
      'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER'
    ];
    
    const targetMonth = String(eventData.month).toUpperCase().trim();
    const monthIdx = months.indexOf(targetMonth);
    if (monthIdx === -1) {
      throw new Error(`Bulan "${eventData.month}" tidak valid.`);
    }
    
    // Kolom bulan di spreadsheet (Kolom B = Februari = index 1 = col 2)
    const col = monthIdx + 2; 
    
    // Cari baris terakhir yang tidak kosong pada kolom ini (di bawah header)
    let lastRow = 3; // header biasanya di baris ke-3
    for (let r = data.length - 1; r >= 3; r--) {
      if (data[r] && String(data[r][col - 1] || '').trim() !== '') {
        lastRow = r + 1; // 1-based row number
        break;
      }
    }
    
    // Tulis Event Title di baris berikutnya (beri 1 baris kosong untuk kerapihan jika baris terakhir bukan header)
    let eventRow = lastRow + 1;
    if (lastRow > 4) {
      eventRow = lastRow + 2; // Berikan 1 baris kosong pembatas antar event
    }
    
    sheet.getRange(eventRow, col).setValue(eventData.title);
    
    // Tulis anggota panitia di baris-baris berikutnya
    for (let i = 0; i < eventData.members.length; i++) {
      sheet.getRange(eventRow + 1 + i, col).setValue(eventData.members[i]);
    }
    
    Logger.log(`Berhasil menambahkan event kepanitiaan: ${eventData.title} pada bulan ${targetMonth}`);
    return { success: true };
  } catch (error) {
    Logger.log(`Error addKepanitiaanEvent: ${error.message}`);
    throw new Error(`Gagal menambah kegiatan kepanitiaan: ${error.message}`);
  }
}

/**
 * Mengupdate daftar seluruh kegiatan kepanitiaan untuk bulan tertentu.
 * Digunakan untuk aksi edit dan hapus kegiatan panitia.
 * 
 * @param {string} monthName - Nama bulan (misal: "FEBRUARI")
 * @param {Array<Object>} events - Array objek event [{ title, members: [] }]
 * @returns {Object} Hasil sukses
 */
function updateKepanitiaanEvents(monthName, events) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_KEPANITIAAN_ID);
    const sheet = ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    
    // Cari baris header bulan
    let headerRowIdx = -1;
    for (let i = 0; i < data.length; i++) {
      const rowLower = data[i].map(c => String(c).toUpperCase().trim());
      if (rowLower.includes('FEBRUARI') || rowLower.includes('MARET')) {
        headerRowIdx = i;
        break;
      }
    }
    if (headerRowIdx === -1) headerRowIdx = 3; // Fallback
    
    const months = [
      'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI', 'JULI', 
      'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER'
    ];
    
    const targetMonth = String(monthName).toUpperCase().trim();
    const monthIdx = months.indexOf(targetMonth);
    if (monthIdx === -1) {
      throw new Error(`Bulan "${monthName}" tidak valid.`);
    }
    
    // Kolom bulan di spreadsheet (Kolom B = Februari = index 1 = col 2)
    const col = monthIdx + 2; 
    
    // 1. Bersihkan seluruh isi kolom di bawah header
    const lastRow = sheet.getLastRow();
    if (lastRow > headerRowIdx + 1) {
      sheet.getRange(headerRowIdx + 2, col, lastRow - headerRowIdx, 1).clearContent();
    }
    
    // 2. Tulis ulang seluruh event dan panitia
    let currentRow = headerRowIdx + 2;
    for (let i = 0; i < events.length; i++) {
      const ev = events[i];
      if (!ev.title || String(ev.title).trim() === '') continue;
      
      // Tulis judul event
      sheet.getRange(currentRow, col).setValue(ev.title.trim());
      currentRow++;
      
      // Tulis anggota
      if (ev.members && ev.members.length > 0) {
        for (let j = 0; j < ev.members.length; j++) {
          sheet.getRange(currentRow, col).setValue(ev.members[j].trim());
          currentRow++;
        }
      }
      
      // Beri baris kosong pembatas antar event
      currentRow++;
    }
    
    Logger.log(`Berhasil mengupdate kepanitiaan bulan ${targetMonth}.`);
    return { success: true };
  } catch (error) {
    Logger.log(`Error updateKepanitiaanEvents: ${error.message}`);
    throw new Error(`Gagal menyimpan kepanitiaan: ${error.message}`);
  }
}

/**
 * Menyimpan banyak Program Kerja sekaligus.
 * 
 * @param {Array<Object>} prokers - Array data proker [{ no, name, execution, activeMonths }]
 * @returns {Object} Hasil sukses
 */
function saveMultipleProkers(prokers) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_PROKER_ID);
    const sheet = ss.getSheets()[0];
    const monthNames = [
      'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 
      'Agustus', 'September', 'Oktober', 'November', 'Desember'
    ];
    
    let lastRow = sheet.getLastRow();
    
    for (let i = 0; i < prokers.length; i++) {
      const p = prokers[i];
      const rowIdx = lastRow + 1 + i;
      
      sheet.getRange(rowIdx, 2).setValue(p.no); // Kolom B
      sheet.getRange(rowIdx, 3).setValue(p.name); // Kolom C
      sheet.getRange(rowIdx, 4).setValue(p.execution); // Kolom D
      
      // Tulis timeline
      for (let monthIdx = 0; monthIdx < monthNames.length; monthIdx++) {
        const monthName = monthNames[monthIdx];
        const startCol = 5 + (monthIdx * 5);
        const isMonthActive = p.activeMonths.includes(monthName);
        
        if (isMonthActive) {
          sheet.getRange(rowIdx, startCol).setValue("v");
          for (let w = 1; w < 5; w++) {
            sheet.getRange(rowIdx, startCol + w).setValue("");
          }
        } else {
          for (let w = 0; w < 5; w++) {
            sheet.getRange(rowIdx, startCol + w).setValue("");
          }
        }
      }
    }
    
    Logger.log(`Berhasil menyimpan ${prokers.length} proker sekaligus.`);
    return { success: true };
  } catch (error) {
    Logger.log(`Error saveMultipleProkers: ${error.message}`);
    throw new Error(`Gagal menyimpan banyak proker: ${error.message}`);
  }
}

