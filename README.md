# VMInit

Automated setup script for fresh Windows 11 Lab VMs.

## What it does

1. **Removes bloatware** – Uninstalls ~40 pre-installed Windows 11 apps (Clipchamp, Xbox, Solitaire, News, Cortana, Spotify, TikTok, etc.)
2. **Cleans up Edge** – Blank start/new-tab page, disables news feed, shopping assistant, sidebar, recommendations
3. **Cleans desktop** – Removes shortcuts, solid gray wallpaper, taskbar cleanup (Widgets/Chat/Search/Task View weg), Windows-Tipps deaktiviert
4. **Installs Microsoft 365 Apps** – Full Office suite (Word, Excel, PowerPoint, Outlook, OneNote, Access, Publisher) via ODT, German + English
5. **Installs Microsoft Teams** – New Teams v2.0, always latest version via official bootstrapper
6. **Desktop shortcuts** – New Outlook, Word, Excel, Teams, OneDrive, M365 Copilot auf dem Desktop (mit korrekten Icons)
7. **Installs dev tools** – VS Code, Git, PowerShell 7, Windows Terminal (via winget)
8. **Refreshes Explorer & Restart** – Applies all changes, then prompts to restart

## Demo-Account einrichten

Nach dem Neustart einfach **Outlook** auf dem Desktop öffnen und sich mit dem M365 Demo-Tenant anmelden. SSO propagiert automatisch zu Word, Excel, Teams und OneDrive — kein Enrollment nötig.

## Usage

Open **PowerShell** on the fresh VM and run:

```powershell
irm https://raw.githubusercontent.com/jenssgb/VMInit/master/Setup.ps1 | iex
```

The script auto-elevates to Administrator (UAC prompt) – no need to manually "Run as Admin".

## Requirements

- Windows 11
- Internet connection
