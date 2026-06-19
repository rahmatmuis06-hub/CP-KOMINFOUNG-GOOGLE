# Dashboard Content Planning Kominfo

Dashboard Web App interaktif untuk manajemen perencanaan konten Divisi Kominfo, dibangun dengan Google Apps Script + Google Spreadsheet.

## 🚀 Cara Setup

### Langkah 1: Buat Project Google Apps Script
1. Buka [Google Apps Script](https://script.google.com)
2. Klik **"New Project"**
3. Beri nama: `Dashboard Kominfo`

### Langkah 2: Upload File
Copy-paste semua file dari folder ini ke Apps Script Editor:

**File Backend (.gs):**
- `Code.gs` — Entry point utama
- `ContentService.gs` — CRUD konten
- `TeamService.gs` — CRUD anggota tim
- `Utils.gs` — Fungsi helper
- `Setup.gs` — Setup database otomatis
- `Migration.gs` — Migrasi data dari spreadsheet lama

**File Frontend (.html):**
- `index.html` — Template utama
- `StyleSheet.html` — CSS Design System
- `JavaScript.html` — Logika client-side
- `Sidebar.html` — Navigasi sidebar
- `Dashboard.html` — Tampilan tabel data
- `KanbanBoard.html` — Papan Kanban
- `CalendarView.html` — Kalender posting
- `Statistics.html` — Statistik beban kerja
- `ContentForm.html` — Form input konten
- `TeamView.html` — Manajemen tim
- `OnboardingGuide.html` — Panduan penggunaan

**Config:**
- `appsscript.json` — Manifest (perlu enable di Settings > Show "appsscript.json")

### Langkah 3: Deploy sebagai Web App
1. Klik **Deploy** > **New deployment**
2. Pilih type: **Web app**
3. Deskripsi: `Dashboard Kominfo v1.0`
4. Execute as: **Me** (akun Anda)
5. Who has access: **Anyone within organization** (atau sesuai kebutuhan)
6. Klik **Deploy**
7. Copy URL Web App yang diberikan

### Langkah 4: Setup Database
1. Buka URL Web App
2. Klik tombol **"⚙️ Setup Database Baru"**
3. Sistem akan otomatis membuat Google Spreadsheet baru dengan:
   - Sheet `Master_Content` (header + data validation)
   - Sheet `Team_Members` (data anggota tim pre-filled)
4. Spreadsheet URL akan ditampilkan setelah setup selesai

### Langkah 5: Migrasi Data (Opsional)
Jika ingin mengimpor data dari spreadsheet lama:
1. Pastikan database baru sudah di-setup (Langkah 4)
2. Klik tombol **"📥 Migrasi Data Lama"** di halaman Setup
3. Data dari spreadsheet lama akan otomatis dipindahkan

## 📋 Fitur

| Fitur | Deskripsi |
|-------|-----------|
| 📊 Dashboard | Tabel data konten dengan filter bulan/tahun |
| 📋 Papan Kanban | Drag & drop status: Draft → In Progress → Review → Completed |
| 📅 Kalender | Tampilan kalender bulanan jadwal posting |
| 📈 Statistik | Chart beban kerja per anggota + leaderboard |
| 👥 Tim Saya | Manajemen anggota tim (tambah/edit/hapus) |
| 📖 Panduan | Onboarding guide untuk anggota baru |

## 🎨 Teknologi
- **Backend**: Google Apps Script (V8 Runtime)
- **Database**: Google Spreadsheet
- **Frontend**: HTML5, CSS3 (Glassmorphism Dark Mode), Vanilla JavaScript
- **Font**: Inter (Google Fonts)
- **Auth**: Google OAuth bawaan

## 📂 Struktur Database

### Sheet: Master_Content
| Kolom | Tipe | Keterangan |
|-------|------|------------|
| ID_Content | String | Auto-generated (CNT-YYYY-NNN) |
| Date_To_Be_Posted | Date | Tanggal posting |
| Bulan_Tahun | String | Format YYYY-MM (filter) |
| PIC_Design | String | Penanggung jawab |
| Content_Title | String | Judul konten |
| Platform | String | Flyer/Reels/TikTok/Feeds/Story |
| Template | String | Link template desain |
| Ratio | String | 16:9, 1:1, 4:5, 9:16 |
| Status | String | Draft/In Progress/Review/Completed |

### Sheet: Team_Members
| Kolom | Tipe | Keterangan |
|-------|------|------------|
| Nama_Lengkap | String | Nama anggota |
| Role | String | Kadiv/Admin/Graphic Designer/Videographer |
| Email | String | Email Google |

## 🔧 Alternatif: Deploy dengan Clasp
```bash
npm install -g @google/clasp
clasp login
clasp create --type webapp --title "Dashboard Kominfo"
clasp push
clasp deploy
```

## 📝 Versi
- **v1.0.0** — Initial release
