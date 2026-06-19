/**
 * NotificationService.gs
 * =======================
 * Modul terintegrasi untuk mengelola:
 * 1. Pengiriman Notifikasi WhatsApp Grup menggunakan Fonnte API.
 * 2. Sinkronisasi jadwal konten dengan Google Calendar.
 * 3. Trigger terjadwal harian untuk pengingat deadline konten.
 */

/**
 * Mengirim pesan WhatsApp ke grup terdaftar menggunakan Fonnte API.
 * 
 * @param {string} message - Isi pesan yang akan dikirim
 * @returns {boolean} Status keberhasilan pengiriman
 */
function sendWhatsAppGroupMessage(message) {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const token = scriptProps.getProperty('WHATSAPP_API_TOKEN');
    const target = scriptProps.getProperty('WHATSAPP_TARGET');

    if (!token || !token.trim() || !target || !target.trim()) {
      Logger.log('WhatsApp notification skipped: API Token atau Target Grup belum disetting.');
      return false;
    }

    const url = 'https://api.fonnte.com/send';
    const payload = {
      target: target.trim(),
      message: message
    };

    const options = {
      method: 'post',
      headers: {
        'Authorization': token.trim()
      },
      payload: payload,
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    Logger.log(`Fonnte API Response Code: ${responseCode}, Response: ${responseText}`);

    if (responseCode === 200) {
      const resJson = JSON.parse(responseText);
      if (resJson.status === true) {
        return true;
      }
    }
    return false;

  } catch (error) {
    Logger.log(`Error sendWhatsAppGroupMessage: ${error.message}`);
    return false;
  }
}

/**
 * Mendapatkan kalender khusus "Content Planning Kominfo".
 * Jika belum ada, otomatis membuat kalender baru dan menyimpan ID-nya.
 * 
 * @returns {GoogleAppsScript.Calendar.Calendar} Objek Google Calendar
 */
function getOrCreateCalendar() {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    let calendarId = scriptProps.getProperty('CALENDAR_ID');
    let calendar = null;

    if (calendarId) {
      try {
        calendar = CalendarApp.getCalendarById(calendarId);
      } catch (e) {
        Logger.log(`Kalender lama ID ${calendarId} tidak dapat diakses atau telah dihapus.`);
      }
    }

    // Jika kalender belum ada atau tidak dapat diakses, buat yang baru
    if (!calendar) {
      calendar = CalendarApp.createCalendar('Content Planning Kominfo', {
        summary: 'Kalender Perencanaan Konten Divisi Kominfo. Sinkronisasi otomatis dari Dashboard.',
        timeZone: 'Asia/Jakarta'
      });
      calendarId = calendar.getId();
      scriptProps.setProperty('CALENDAR_ID', calendarId);
      Logger.log(`Kalender baru berhasil dibuat dengan ID: ${calendarId}`);
    }

    return calendar;

  } catch (error) {
    Logger.log(`Error getOrCreateCalendar: ${error.message}`);
    throw new Error(`Gagal mengkoneksikan Google Calendar: ${error.message}`);
  }
}

/**
 * Mensinkronkan baris konten dengan Google Calendar event.
 * Membuat event baru jika belum ada, atau memperbarui yang sudah ada.
 * 
 * @param {Object} contentData - Data konten lengkap
 * @returns {string} ID Event Google Calendar yang berhasil disinkronkan
 */
function syncCalendarEvent(contentData) {
  try {
    const calendar = getOrCreateCalendar();
    const eventDate = new Date(contentData.date);

    if (isNaN(eventDate.getTime())) {
      Logger.log(`Gagal sinkronisasi kalender: Tanggal "${contentData.date}" tidak valid.`);
      return '';
    }

    // Format judul event
    const eventTitle = `[${contentData.platform}] ${contentData.title} - ${contentData.pic || 'Unassigned'}`;

    // Format deskripsi event
    let eventDescription = `📅 INFORMASI PERENCANAAN KONTEN\n`;
    eventDescription += `===============================\n`;
    eventDescription += `ID Konten: ${contentData.id}\n`;
    eventDescription += `Platform: ${contentData.platform}\n`;
    eventDescription += `Ratio: ${contentData.ratio || '-'}\n`;
    eventDescription += `PIC: ${contentData.pic || 'Belum Ditentukan'}\n`;
    eventDescription += `Status: ${contentData.status || 'Draft'}\n\n`;
    
    if (contentData.description) {
      eventDescription += `📝 Keterangan:\n"${contentData.description}"\n\n`;
    }
    
    if (contentData.template) {
      eventDescription += `🎨 Link Template Desain:\n${contentData.template}\n\n`;
    }

    if (contentData.fileUrl) {
      eventDescription += `📁 File Hasil Akhir:\n${contentData.fileUrl}\n`;
    }

    let event = null;

    // Jika sudah ada Event ID, coba cari event-nya
    if (contentData.eventId) {
      try {
        event = calendar.getEventById(contentData.eventId);
      } catch (e) {
        Logger.log(`Event ID ${contentData.eventId} tidak ditemukan, akan dibuat event baru.`);
      }
    }

    if (event) {
      // Update event yang sudah ada
      event.setTitle(eventTitle);
      event.setDescription(eventDescription);
      
      // Karena event ini all-day event
      event.setAllDayDate(eventDate);
      Logger.log(`Event kalender diupdate: ${contentData.eventId}`);
      return contentData.eventId;
    } else {
      // Buat all-day event baru
      const newEvent = calendar.createAllDayEvent(eventTitle, eventDate, {
        description: eventDescription
      });
      const newEventId = newEvent.getId();
      Logger.log(`Event kalender baru dibuat: ${newEventId}`);
      return newEventId;
    }

  } catch (error) {
    Logger.log(`Error syncCalendarEvent: ${error.message}`);
    return '';
  }
}

/**
 * Menghapus event di Google Calendar berdasarkan Event ID.
 * 
 * @param {string} eventId - ID event kalender yang akan dihapus
 */
function deleteCalendarEvent(eventId) {
  try {
    if (!eventId) return;
    const calendar = getOrCreateCalendar();
    const event = calendar.getEventById(eventId);
    if (event) {
      event.deleteEvent();
      Logger.log(`Event kalender ${eventId} berhasil dihapus.`);
    }
  } catch (error) {
    Logger.log(`Error deleteCalendarEvent: ${error.message}`);
  }
}

/**
 * Mengirim notifikasi WA grup saat ada konten baru didaftarkan.
 * 
 * @param {Object} contentData - Data konten yang didaftarkan
 * @param {string} userName - Pengguna yang melakukan penambahan
 */
function notifyNewContentWhatsApp(contentData, userName) {
  const tglDisplay = formatDateDisplay(new Date(contentData.date));
  const operator = userName || 'Sistem';

  let message = `📅 *[KONTEN BARU DIDAFTARKAN]*\n`;
  message += `Halo Tim Kominfo! Konten baru telah didaftarkan oleh *${operator}*:\n\n`;
  message += `📌 *Judul:* "${contentData.title}"\n`;
  message += `📱 *Platform:* ${contentData.platform}\n`;
  message += `📐 *Rasio:* ${contentData.ratio || '-'}\n`;
  message += `👥 *PIC:* ${contentData.pic || 'Belum Ditentukan'}\n`;
  message += `🗓️ *Jadwal Posting:* ${tglDisplay}\n\n`;
  message += `Silakan buka dashboard perencanaan untuk melihat detail selengkapnya!`;

  sendWhatsAppGroupMessage(message);
}

/**
 * Mengirim notifikasi WA grup saat ada beberapa konten sekaligus ditambahkan (Bulk).
 * 
 * @param {number} count - Jumlah konten yang ditambahkan
 * @param {string} userName - Pengguna yang melakukan aksi
 */
function notifyBulkContentsWhatsApp(count, userName) {
  const operator = userName || 'Sistem';

  let message = `📅 *[PENAMBAHAN KONTEN MASSAL]*\n`;
  message += `Halo Tim Kominfo! Sebanyak *${count} konten baru* telah ditambahkan secara massal oleh *${operator}*.\n\n`;
  message += `Silakan cek tabel Dashboard atau Kalender untuk melihat pembaruan jadwal posting terbaru. Semangat!`;

  sendWhatsAppGroupMessage(message);
}

/**
 * Mengirim notifikasi WA grup saat file selesai diunggah.
 * 
 * @param {string} contentId - ID Konten terkait
 * @param {string} fileName - Nama berkas
 * @param {string} fileUrl - Link berkas di Drive
 * @param {string} userName - Pengunggah berkas
 */
function notifyFileUploadWhatsApp(contentId, fileName, fileUrl, userName) {
  try {
    const sheet = getSheet('Master_Content');
    const data = sheet.getDataRange().getValues();
    
    let contentTitle = 'Konten';
    let pic = 'Belum Ditentukan';
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][MC_COL.ID].toString().trim() === contentId.toString().trim()) {
        contentTitle = data[i][MC_COL.TITLE] || 'Konten';
        pic = data[i][MC_COL.PIC] || 'Belum Ditentukan';
        break;
      }
    }

    const uploader = userName || 'Anggota Tim';

    let message = `🚀 *[FILE SELESAI DIUNGGAH]*\n`;
    message += `Halo Tim Kominfo! Lampiran berkas baru telah berhasil diunggah:\n\n`;
    message += `📌 *Konten:* "${contentTitle}" (ID: ${contentId})\n`;
    message += `👥 *PIC:* ${pic}\n`;
    message += `📁 *Nama File:* ${fileName}\n`;
    message += `🔗 *Link File:* ${fileUrl}\n`;
    message += `👤 *Uploader:* ${uploader}\n\n`;
    message += `Status konten ini otomatis diperbarui menjadi *Completed* ✅. Terima kasih!`;

    sendWhatsAppGroupMessage(message);

  } catch (error) {
    Logger.log(`Error notifyFileUploadWhatsApp: ${error.message}`);
  }
}

/**
 * Memindai semua konten dan mengirim ringkasan (digest) notifikasi
 * untuk konten yang mendekati deadline (H-2 s.d Hari H) tetapi masih Draft / In Progress.
 */
function checkAndSendDailyDeadlineAlerts() {
  try {
    const sheet = getSheet('Master_Content');
    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return; // Hanya ada header

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const alertList = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = row[MC_COL.STATUS];
      
      // Hanya scan jika status Draft atau In Progress
      if (status === 'Draft' || status === 'In Progress') {
        const postDate = new Date(row[MC_COL.DATE]);
        if (isNaN(postDate.getTime())) continue;

        postDate.setHours(0, 0, 0, 0);
        
        const diffTime = postDate.getTime() - today.getTime();
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));

        // Peringatan untuk deadline Hari H (0), Besok (1), atau Lusa (2)
        if (diffDays >= 0 && diffDays <= 2) {
          let deadlineText = '';
          if (diffDays === 0) {
            deadlineText = '*HARI INI*';
          } else if (diffDays === 1) {
            deadlineText = '*BESOK*';
          } else {
            deadlineText = '*LUSA* (H-2)';
          }

          alertList.push({
            id: row[MC_COL.ID],
            title: row[MC_COL.TITLE],
            pic: row[MC_COL.PIC] || 'Belum Ditentukan',
            platform: row[MC_COL.PLATFORM],
            dateStr: formatDateDisplay(postDate),
            deadlineLabel: deadlineText,
            status: status
          });
        }
      }
    }

    // Jika tidak ada konten yang mendekati deadline, jangan kirim notifikasi
    if (alertList.length === 0) {
      Logger.log('Daily Deadline Alert: Tidak ada konten mendekati deadline.');
      return;
    }

    // Susun pesan ringkasan (digest)
    let message = `⚠️ *[PERINGATAN DEADLINE KONTEN]*\n`;
    message += `Halo Tim Kominfo! Berikut adalah daftar konten yang mendekati/sudah masuk jadwal posting, namun *masih Draft atau In Progress*:\n\n`;

    alertList.forEach((item, index) => {
      message += `${index + 1}. 📌 *"${item.title}"* (${item.platform})\n`;
      message += `   👥 *PIC:* ${item.pic}\n`;
      message += `   🗓️ *Target Posting:* ${item.dateStr} (${item.deadlineLabel})\n`;
      message += `   🔄 *Status:* ${item.status}\n\n`;
    });

    message += `Mohon kepada PIC terkait untuk segera menyelesaikan konten tersebut dan mengunggah berkas finalnya ke dashboard. Terima kasih! 🙏`;

    // Kirim pesan ke grup WA
    sendWhatsAppGroupMessage(message);
    Logger.log(`Daily Deadline Alert berhasil dikirim untuk ${alertList.length} konten.`);

  } catch (error) {
    Logger.log(`Error checkAndSendDailyDeadlineAlerts: ${error.message}`);
  }
}

/**
 * Membuat trigger otomatis harian (Time-driven trigger) untuk mendeteksi deadline.
 * Pemicu akan berjalan setiap hari pada pukul 08:00 - 09:00 pagi.
 * 
 * @returns {boolean} Status keberhasilan pembuatan trigger
 */
function setupNotificationTriggers() {
  try {
    // Hapus trigger lama terlebih dahulu untuk mencegah duplikasi
    removeNotificationTriggers();

    // Buat trigger harian baru
    ScriptApp.newTrigger('checkAndSendDailyDeadlineAlerts')
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .create();

    Logger.log('Trigger harian checkAndSendDailyDeadlineAlerts berhasil dibuat.');
    return true;

  } catch (error) {
    Logger.log(`Error setupNotificationTriggers: ${error.message}`);
    throw new Error(`Gagal membuat trigger: ${error.message}`);
  }
}

/**
 * Menghapus semua trigger yang memanggil fungsi checkAndSendDailyDeadlineAlerts.
 * 
 * @returns {boolean} Status keberhasilan penghapusan trigger
 */
function removeNotificationTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deleted = false;

    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'checkAndSendDailyDeadlineAlerts') {
        ScriptApp.deleteTrigger(triggers[i]);
        deleted = true;
      }
    }
    
    if (deleted) {
      Logger.log('Trigger harian checkAndSendDailyDeadlineAlerts dibersihkan.');
    }
    return true;

  } catch (error) {
    Logger.log(`Error removeNotificationTriggers: ${error.message}`);
    throw new Error(`Gagal menghapus trigger: ${error.message}`);
  }
}

/**
 * Mengecek apakah trigger deadline harian aktif di project Apps Script.
 * 
 * @returns {boolean} True jika trigger aktif, False jika tidak
 */
function getNotificationTriggerStatus() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'checkAndSendDailyDeadlineAlerts') {
        return true;
      }
    }
    return false;
  } catch (e) {
    Logger.log(`Error getNotificationTriggerStatus: ${e.message}`);
    return false;
  }
}

/**
 * Mengetes seluruh sistem integrasi (WhatsApp dan Google Calendar).
 * Mengirim pesan tes ke grup WA dan memverifikasi izin Google Calendar.
 * 
 * @returns {Object} JSON hasil uji coba
 */
function testSystemIntegrations() {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const token = scriptProps.getProperty('WHATSAPP_API_TOKEN');
    const target = scriptProps.getProperty('WHATSAPP_TARGET');

    const result = {
      whatsapp: { success: false, message: 'Belum diuji coba.' },
      calendar: { success: false, message: 'Belum diuji coba.', name: '', url: '' }
    };

    // 1. Tes Google Calendar
    try {
      const calendar = getOrCreateCalendar();
      result.calendar.name = calendar.getName();
      
      // ID Kalender untuk url publik
      result.calendar.url = `https://calendar.google.com/calendar/r?cid=${encodeURIComponent(calendar.getId())}`;
      
      // Buat event dummy dan langsung hapus
      const testDate = new Date();
      const testEvent = calendar.createAllDayEvent('🧪 TEST INTEGRASI DASHBOARD KOMINFO (DAPAT DIHAPUS)', testDate);
      testEvent.deleteEvent();
      
      result.calendar.success = true;
      result.calendar.message = 'Koneksi & izin Google Calendar BERHASIL. Event uji coba berhasil dibuat dan dihapus.';
    } catch (e) {
      result.calendar.success = false;
      result.calendar.message = `Gagal uji coba Google Calendar: ${e.message}`;
    }

    // 2. Tes WhatsApp Fonnte
    if (!token || !target) {
      result.whatsapp.success = false;
      result.whatsapp.message = 'Token API atau Target WA kosong. Silakan isi terlebih dahulu.';
    } else {
      const msg = `🤖 *[TEST INTEGRASI DASHBOARD KOMINFO]*\n\nHalo Tim Kominfo! Koneksi bot notifikasi WhatsApp berhasil terhubung ke sistem Dashboard Perencanaan Konten Anda.\n\n_Pesan dikirim pada: ${new Date().toLocaleString()}_`;
      const waSuccess = sendWhatsAppGroupMessage(msg);
      
      if (waSuccess) {
        result.whatsapp.success = true;
        result.whatsapp.message = 'Pesan uji coba BERHASIL dikirim ke grup WhatsApp. Periksa HP Anda!';
      } else {
        result.whatsapp.success = false;
        result.whatsapp.message = 'Gagal mengirim pesan uji coba. Periksa apakah API Token valid dan nomor/ID grup terdaftar dengan benar di Fonnte.';
      }
    }

    return result;

  } catch (error) {
    Logger.log(`Error testSystemIntegrations: ${error.message}`);
    throw new Error(`Uji coba gagal: ${error.message}`);
  }
}

/**
 * Menyimpan pengaturan integrasi notifikasi (API Token & Target WA) ke Script Properties.
 * 
 * @param {string} token - Fonnte API Token
 * @param {string} target - Fonnte Target WhatsApp (No HP / ID Grup)
 * @param {string} userName - Pengguna yang melakukan perubahan
 * @returns {boolean} Status keberhasilan penyimpanan
 */
function saveNotificationSettings(token, target, userName) {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const oldToken = scriptProps.getProperty('WHATSAPP_API_TOKEN') || '';
    
    let tokenToSave = token || '';
    if (tokenToSave.includes('...') || tokenToSave.includes('***')) {
      tokenToSave = oldToken;
    }
    
    scriptProps.setProperty('WHATSAPP_API_TOKEN', tokenToSave);
    scriptProps.setProperty('WHATSAPP_TARGET', target || '');

    // Catat log
    logActivity('Ubah Pengaturan Integrasi', `Memperbarui API Token dan Target WhatsApp (${target})`, userName);
    return true;

  } catch (error) {
    Logger.log(`Error saveNotificationSettings: ${error.message}`);
    throw new Error(`Gagal menyimpan pengaturan: ${error.message}`);
  }
}

/**
 * Mendapatkan pengaturan integrasi notifikasi dari Script Properties.
 * 
 * @returns {Object} Konfigurasi tersimpan beserta status trigger
 */
function getNotificationSettings() {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const token = scriptProps.getProperty('WHATSAPP_API_TOKEN') || '';
    const target = scriptProps.getProperty('WHATSAPP_TARGET') || '';
    
    // Masking token untuk keamanan
    let maskedToken = '';
    if (token.length > 12) {
      maskedToken = `${token.substring(0, 6)}...${token.substring(token.length - 6)}`;
    } else if (token) {
      maskedToken = '*** Terkonfigurasi ***';
    }

    let calendarUrl = '';
    let calendarName = '';
    const calendarId = scriptProps.getProperty('CALENDAR_ID');
    if (calendarId) {
      calendarUrl = `https://calendar.google.com/calendar/r?cid=${encodeURIComponent(calendarId)}`;
      calendarName = 'Content Planning Kominfo';
    }

    const triggerActive = getNotificationTriggerStatus();

    return {
      maskedToken: maskedToken,
      hasToken: token ? true : false,
      target: target,
      calendarName: calendarName,
      calendarUrl: calendarUrl,
      triggerActive: triggerActive
    };

  } catch (error) {
    Logger.log(`Error getNotificationSettings: ${error.message}`);
    return {
      token: '',
      maskedToken: '',
      target: '',
      calendarName: '',
      calendarUrl: '',
      triggerActive: false
    };
  }
}
