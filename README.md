# Web Print App

<p align="center">
  <strong>Local printer web console + file manager for Windows</strong>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/version-4.0.3--advanced-0969DA?style=flat-square" alt="Version">
  <img src="https://img.shields.io/badge/Python-3.8%2B-3776AB?style=flat-square&logo=python&logoColor=white" alt="Python 3.8+">
  <img src="https://img.shields.io/badge/Windows-7%20to%2011-0078D4?style=flat-square&logo=windows11&logoColor=white" alt="Windows">
  <img src="https://img.shields.io/badge/FastAPI-WebApp-009688?style=flat-square&logo=fastapi&logoColor=white" alt="FastAPI">
</p>

Web Print App adalah aplikasi web lokal untuk mengelola file dan mencetak langsung ke printer yang terpasang pada PC/host. Print diproses oleh host Windows itu sendiri—tanpa Google Apps Script, cloud print relay, atau Print Bridge tambahan.

> **Target utama:** PC Windows yang terhubung ke printer lokal/network printer. Browser dari perangkat lain di LAN/VPN dapat membuka WebApp dan mengirim job print ke PC tersebut.

## Fitur utama

### Local printing

- Deteksi printer Windows dan default printer.
- Persistent sequential print queue berbasis SQLite.
- Status Windows Spooler dan Windows Job ID.
- Cancel job sebelum maupun setelah masuk spooler.
- Status printer: offline, paused, paper out, paper jam, error, dan kondisi driver.
- Recovery antrean setelah restart dengan proteksi terhadap duplicate print.
- A4, F4, Legal, Letter.
- Portrait / landscape.
- Fit / actual size.
- Duplex dan color/mono bila didukung driver.
- Copies, page range, sheet Excel, priority, dan scheduled print.
- PDF dan gambar dicetak sebagai satu spool document per job.
- Word, Excel, dan PowerPoint melalui Microsoft Office COM pada Windows.

### File manager

- Multiple root folders.
- Root folder dapat ditambahkan dari WebApp.
- Browse, search, pagination, sorting.
- Upload dan download.
- Rename, delete, dan create folder untuk admin.
- Preview PDF, image, Excel, Word, dan format Office lain dengan fallback engine.
- Upload dari perangkat lalu preview/print tanpa harus menyimpan ke root utama.
- Path traversal protection dan Windows filename validation.

### WebApp & administration

- Responsive mobile/desktop UI.
- PWA/installable WebApp.
- Realtime status melalui Server-Sent Events.
- Browser notification untuk perubahan printer/job.
- Admin dan user roles.
- PBKDF2 password hashing untuk akun tambahan.
- CSRF protection, login throttling, session expiry, API scopes, dan print quotas.
- Print history dan statistik.
- Export history ke CSV/XLSX.
- Audit log dan error log.
- Backup/restore dan storage cleanup.
- Maintenance mode.
- Optional Telegram notifications.
- Optional self-update + rollback.
- Windows Task Scheduler autostart + single instance.

## Persyaratan

### Minimum

- Windows 7/8/10/11 untuk fungsi print Windows.
- Python 3.8 atau lebih baru.
- Printer dan driver printer sudah terpasang di Windows.

### Untuk dokumen Microsoft Office

Untuk hasil Word/Excel/PowerPoint yang mengikuti layout Office secara native, Microsoft Office perlu terpasang pada PC host. LibreOffice dapat digunakan sebagai fallback untuk beberapa operasi preview.

## Instalasi cepat

### 1. Clone repository

```powershell
git clone https://github.com/baska-pro/web-print-app.git
cd web-print-app
```

Atau download ZIP dari GitHub lalu extract.

### 2. Buat konfigurasi

PowerShell:

```powershell
Copy-Item .env.example .env
notepad .env
```

Minimal ubah:

```env
WEBAPP_USERNAME=admin
WEBAPP_PASSWORD=GANTI_DENGAN_PASSWORD_KUAT
```

`WEBAPP_PASSWORD` wajib diisi. Jika belum ada root folder, aplikasi tetap dapat dimulai dan root dapat ditambahkan melalui tombol **+ Folder** di sidebar.

### 3. Jalankan

```powershell
python webapp-PrintBot.py
```

Script akan memeriksa dan meng-install dependency Python yang belum tersedia.

Buka:

```text
http://127.0.0.1:8000
```

Dari perangkat lain dalam jaringan:

```text
http://IP-PC-HOST:8000
```

Pastikan Windows Firewall mengizinkan port yang digunakan jika WebApp hendak diakses dari perangkat lain.

## Konfigurasi `.env`

| Variable | Default / contoh | Fungsi |
|---|---|---|
| `FILE_MANAGER_ROOTS` | kosong | Root awal, format `Label:C:\\Folder;Label2:D:\\Arsip` |
| `ROOTS_BASE_DIR` | `./data/roots` | Base folder untuk root yang dibuat dari WebApp |
| `DATA_DIR` | `./data` | Database, logs, preview, backup, archive, token runtime |
| `DEFAULT_PRINTER` | kosong | Kosong = default printer Windows |
| `WEBAPP_HOST` | `0.0.0.0` | Interface listen |
| `WEBAPP_PORT` | `8000` | Port WebApp |
| `WEBAPP_USERNAME` | `admin` | Admin utama |
| `WEBAPP_PASSWORD` | **wajib** | Password admin utama |
| `AUTO_START_TASK` | `1` | Buat/perbaiki Task Scheduler Windows |
| `TASK_NAME` | `PrintBot-WebApp` | Nama scheduled task |
| `PRINT_RENDER_DPI` | `144` | DPI render print/preview |
| `MAX_UPLOAD_MB` | `50` | Batas upload |
| `MAX_DOWNLOAD_MB` | `50` | Batas download |
| `MAX_JOBS_PER_HOUR` | `30` | Quota job/user |
| `MAX_COPIES_PER_DAY` | `200` | Quota copies harian/user |
| `WEBAPP_API_TOKEN` | kosong | Dibuat otomatis jika kosong |
| `WEBAPP_API_SCOPES` | `read,print` | Scope API |
| `TELEGRAM_BOT_TOKEN` | kosong | Optional notifikasi admin |
| `TELEGRAM_ADMIN_CHAT_ID` | kosong | Tujuan notifikasi Telegram |
| `UPDATE_URL` | raw GitHub URL | Sumber update script |
| `AUTO_UPDATE_APPLY` | `0` | Jangan auto-apply secara default |
| `LOG_LEVEL` | `INFO` | Level logging |

Lihat [`.env.example`](.env.example) untuk seluruh opsi.

## Windows Task Scheduler

Saat `AUTO_START_TASK=1`, aplikasi mencoba memastikan scheduled task tersedia. Task menggunakan sesi Windows user yang login sehingga printer per-user dan Microsoft Office COM tetap dapat digunakan.

Jika Task Scheduler bermasalah, admin WebApp juga memiliki fungsi repair/restart.

## Self update

Template `.env.example` sudah menunjuk ke source utama repository ini:

```env
UPDATE_URL=https://raw.githubusercontent.com/baska-pro/web-print-app/main/webapp-PrintBot.py
AUTO_UPDATE_APPLY=0
```

Dengan `AUTO_UPDATE_APPLY=0`, aplikasi dapat mengecek update tanpa otomatis mengganti file. Update manual membuat backup source lama untuk rollback.

## Data runtime

Secara default seluruh data runtime berada di `./data`:

```text
data/
├── printbot.db
├── users.json
├── roots.json
├── webapp_secret.key
├── api_token.key
├── logs/
├── previews/
├── uploads/
├── print_archive/
├── backup/
├── updates/
└── exports/
```

Folder tersebut di-ignore Git dan **tidak boleh dipush ke repository**.

## Security

Aplikasi ini memiliki akses file manager dan printer host. Jangan expose port `8000` langsung ke internet tanpa perlindungan.

Untuk akses jarak jauh, prioritaskan:

- VPN/Tailscale/WireGuard, atau
- reverse proxy HTTPS + autentikasi tambahan.

Jika dijalankan melalui HTTPS, aktifkan:

```env
WEBAPP_SECURE_COOKIE=1
```

Lihat [SECURITY.md](SECURITY.md).

## Dependency manual

Walaupun script dapat auto-install dependency, instalasi manual tersedia:

```powershell
python -m pip install -r requirements.txt
```

Tool sistem seperti LibreOffice/Poppler/Ghostscript bersifat fallback/peningkatan kualitas preview. Di Windows, engine utama memanfaatkan PyMuPDF, Pillow/openpyxl/python-docx dan Microsoft Office COM bila Office tersedia.

## Update source langsung dengan Git

Jika Anda mengelola instalasi menggunakan Git:

```powershell
git pull
```

Simpan `.env` dan `data/` hanya di PC host; keduanya sudah di-ignore.

## Struktur repository

```text
web-print-app/
├── webapp-PrintBot.py
├── .env.example
├── .gitignore
├── requirements.txt
├── VERSION
├── CHANGELOG.md
├── SECURITY.md
└── .github/
    └── workflows/
        └── ci.yml
```

## Catatan kompatibilitas

Source mempertahankan jalur kompatibilitas Python 3.8 untuk Windows lama dan menggunakan versi dependency yang lebih konservatif ketika berjalan di Python < 3.9.

Fungsi printer lokal bergantung pada Windows/pywin32. Pada non-Windows, sebagian fungsi WebApp/file manager/preview dapat berjalan, tetapi engine print Windows tidak tersedia.

## Project status

Current version: **4.0.3-advanced**

Lihat [CHANGELOG.md](CHANGELOG.md) untuk ringkasan versi.
