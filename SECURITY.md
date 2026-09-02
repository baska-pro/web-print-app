# Security Policy

## Important deployment notes

Web Print App exposes file-management and local-printer controls through a browser. Treat it as an administrative application.

- Do **not** commit `.env`, `data/`, API tokens, Telegram tokens, session keys, SQLite databases, or backup archives.
- Set a strong `WEBAPP_PASSWORD` before first launch.
- Keep `AUTO_UPDATE_APPLY=0` unless you explicitly want unattended source replacement.
- If the app is reachable outside a trusted LAN/VPN, place it behind HTTPS and set `WEBAPP_SECURE_COOKIE=1`.
- Prefer a VPN, private network, or authenticated reverse proxy instead of exposing port `8000` directly to the public internet.
- Review `FILE_MANAGER_ROOTS` carefully: users can browse files below configured roots according to their role.
- Rotate `WEBAPP_API_TOKEN` and Telegram credentials if they are ever exposed.

## Reporting a vulnerability

Please avoid publishing credentials, private paths, database contents, or exploitable details in a public issue. Open a minimal issue in this repository requesting a private security contact path, or contact the repository owner through their published GitHub profile contact information.

## Supported version

Security fixes target the latest version on the `main` branch.
