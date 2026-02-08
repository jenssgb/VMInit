# VMInit

Automated setup script for fresh Windows 11 Lab VMs.

## What it does

1. **Removes bloatware** – Uninstalls ~40 pre-installed Windows 11 apps (Clipchamp, Xbox, Solitaire, News, Cortana, Spotify, TikTok, etc.)
2. **Cleans up Edge** – Blank start/new-tab page, disables news feed, shopping assistant, sidebar, recommendations
3. **Cleans desktop** – Removes shortcuts, solid gray wallpaper, taskbar cleanup (Widgets/Chat/Search/Task View weg), Windows-Tipps deaktiviert
4. **Installs Microsoft 365 Apps** – Full Office suite (Word, Excel, PowerPoint, Outlook, OneNote, Access, Publisher) via ODT, German + English
5. **Installs Microsoft Teams** – New Teams v2.0, always latest version via official bootstrapper
6. **Desktop shortcuts** – Outlook, Word, Excel, Teams, OneDrive auf dem Desktop
7. **M365 Account onboarding** – "M365 Account anmelden" Shortcut auf dem Desktop → öffnet direkt die Account-Verbindung (SSO für alle Apps, kein Enrollment nötig)
8. **Installs dev tools** – VS Code, Git, PowerShell 7, Windows Terminal (via winget)
9. **Refreshes Explorer & Restart** – Applies all changes, then prompts to restart

## Usage

Open **PowerShell** on the fresh VM and run:

```powershell
irm https://raw.githubusercontent.com/jenssgb/VMInit/master/Setup.ps1 | iex
```

The script auto-elevates to Administrator (UAC prompt) – no need to manually "Run as Admin".

## Requirements

- Windows 11
- Internet connection
