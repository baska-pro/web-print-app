# Web Print App v4.0.3-advanced

Initial public release of Web Print App / PrintBot WebApp Local Printer Edition.

## Highlights

- Local Windows printer engine; no Google Apps Script or Print Bridge required.
- Persistent SQLite print queue with sequential processing.
- Windows Spooler monitoring, Windows Job ID tracking, and print cancellation.
- PDF/image printing plus Microsoft Word, Excel, and PowerPoint via Office COM.
- A4, F4, Legal, Letter; portrait/landscape; Fit/Actual; duplex and color modes.
- Responsive file manager with search, pagination, upload/download, rename, delete, and folders.
- Multi-page document preview and device-upload print flow.
- Admin/user roles, PBKDF2 password hashing, CSRF protection, login throttling, sessions, quotas, and API scopes.
- Realtime updates through Server-Sent Events and installable PWA support.
- Backup/restore, audit/error logs, storage cleanup, history export, maintenance mode.
- Optional Telegram administrator notifications.
- Optional source update/rollback mechanism.
- Windows Task Scheduler autostart and single-instance protection.
- Python 3.8 compatibility path for older Windows environments.

## Installation

1. Extract `Web-Print-App-v4.0.3-advanced.zip`.
2. Copy `.env.example` to `.env`.
3. Set `WEBAPP_PASSWORD` to a strong password.
4. Run `python webapp-PrintBot.py`.
5. Open `http://127.0.0.1:8000`.

The application automatically checks and installs missing Python dependencies. Manual installation is also available with `python -m pip install -r requirements.txt`.

## Security

Do not commit or distribute `.env`, `data/`, `printbot.db`, token files, backups, or logs.

For remote access, prefer VPN/Tailscale/WireGuard or an HTTPS reverse proxy with additional authentication.

## Compatibility

- Primary target: Windows 7/8/10/11.
- Python: 3.8+.
- Microsoft Office is recommended for native Word/Excel/PowerPoint rendering and printing.
- Non-Windows systems can use parts of the WebApp/file manager, but Windows local-print functionality is unavailable.
