# VMInit

Automated setup script for fresh Windows 11 Lab VMs.

## What it does

1. **Removes bloatware** – Uninstalls ~40 pre-installed Windows 11 apps (Clipchamp, Xbox, Solitaire, News, Cortana, Spotify, TikTok, etc.)
2. **Cleans up Edge** – Blank start/new-tab page, disables news feed, shopping assistant, sidebar, recommendations
3. **Installs Microsoft 365 Apps** – Full Office suite (Word, Excel, PowerPoint, Outlook, OneNote, Access, Publisher) via ODT, German + English
4. **Installs Microsoft Teams** – New Teams v2.0, always latest version via official bootstrapper

## Usage

Open **PowerShell** on the fresh VM and run:

```powershell
irm https://raw.githubusercontent.com/jenssgb/VMInit/master/Setup.ps1 | iex
```

The script auto-elevates to Administrator (UAC prompt) – no need to manually "Run as Admin".

## Requirements

- Windows 11
- Internet connection
