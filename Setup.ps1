<#
.SYNOPSIS
    VMInit - Automated Windows 11 Lab VM Setup
.DESCRIPTION
    Installs M365 Apps, Microsoft Teams, removes bloatware, and cleans up Edge.
    Run on a fresh Windows 11 VM with a single command:

    irm https://raw.githubusercontent.com/jenssgb/VMInit/master/Setup.ps1 | iex

.NOTES
    Author: jenssgb
    Requires: Internet connection (auto-elevates to Admin)
#>

# ── Auto-elevate to Administrator ──
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Restarting as Administrator..." -ForegroundColor Yellow
    $script = "irm https://raw.githubusercontent.com/jenssgb/VMInit/master/Setup.ps1 | iex"
    Start-Process powershell -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -Command $script"
    return
}

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'   # speeds up Invoke-WebRequest dramatically
$tempDir = "$env:TEMP\VMInit"
New-Item -Path $tempDir -ItemType Directory -Force | Out-Null

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  VMInit - Windows 11 Lab VM Setup" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

$steps = 7

# ──────────────────────────────────────────────
# 1. REMOVE BLOATWARE
# ──────────────────────────────────────────────
Write-Host "[1/$steps] Removing Windows 11 bloatware..." -ForegroundColor Yellow

$bloatApps = @(
    "Clipchamp.Clipchamp"
    "Microsoft.BingFinance"
    "Microsoft.BingNews"
    "Microsoft.BingSports"
    "Microsoft.BingTranslator"
    "Microsoft.BingWeather"
    "Microsoft.GamingApp"
    "Microsoft.GetHelp"
    "Microsoft.Getstarted"
    "Microsoft.MicrosoftOfficeHub"
    "Microsoft.MicrosoftSolitaireCollection"
    "Microsoft.MicrosoftStickyNotes"
    "Microsoft.People"
    "Microsoft.PowerAutomateDesktop"
    "Microsoft.Todos"
    "Microsoft.WindowsAlarms"
    "Microsoft.WindowsCommunicationsApps"
    "Microsoft.WindowsFeedbackHub"
    "Microsoft.WindowsMaps"
    "Microsoft.WindowsSoundRecorder"
    "Microsoft.Xbox.TCUI"
    "Microsoft.XboxGameOverlay"
    "Microsoft.XboxGamingOverlay"
    "Microsoft.XboxIdentityProvider"
    "Microsoft.XboxSpeechToTextOverlay"
    "Microsoft.YourPhone"
    "Microsoft.ZuneMusic"
    "Microsoft.ZuneVideo"
    "MicrosoftCorporationII.QuickAssist"
    "Microsoft.549981C3F5F10"  # Cortana
    "Microsoft.MixedReality.Portal"
    "Microsoft.SkypeApp"
    "Microsoft.WindowsCamera"
    "Disney.37853FC22B2CE"
    "SpotifyAB.SpotifyMusic"
    "BytedancePte.Ltd.TikTok"
    "king.com.CandyCrushSaga"
    "king.com.CandyCrushSodaSaga"
    "Facebook.Facebook"
    "9E2F88E3.Twitter"
    "AmazonVideo.PrimeVideo"
    "Microsoft.OutlookForWindows"
)

foreach ($app in $bloatApps) {
    $pkg = Get-AppxPackage -Name $app -AllUsers -ErrorAction SilentlyContinue
    if ($pkg) {
        $pkg | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
        Write-Host "  Removed: $app" -ForegroundColor DarkGray
    }
    # Also remove provisioned packages so they don't come back for new users
    $prov = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -eq $app }
    if ($prov) {
        $prov | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Out-Null
    }
}

Write-Host "  Bloatware removed." -ForegroundColor Green

# ──────────────────────────────────────────────
# 2. CLEAN UP MICROSOFT EDGE
# ──────────────────────────────────────────────
Write-Host "[2/$steps] Cleaning up Microsoft Edge..." -ForegroundColor Yellow

$edgePolicies = "HKLM:\SOFTWARE\Policies\Microsoft\Edge"
New-Item -Path $edgePolicies -Force | Out-Null

# Disable first run experience / welcome page
Set-ItemProperty -Path $edgePolicies -Name "HideFirstRunExperience" -Value 1 -Type DWord
Set-ItemProperty -Path $edgePolicies -Name "StartupBoostEnabled" -Value 0 -Type DWord

# Set homepage & new tab to blank
Set-ItemProperty -Path $edgePolicies -Name "HomepageLocation" -Value "about:blank" -Type String
Set-ItemProperty -Path $edgePolicies -Name "HomepageIsNewTabPage" -Value 0 -Type DWord
Set-ItemProperty -Path $edgePolicies -Name "NewTabPageLocation" -Value "about:blank" -Type String
Set-ItemProperty -Path $edgePolicies -Name "RestoreOnStartup" -Value 4 -Type DWord  # Open a specific page

# Startup pages
$startupPages = "HKLM:\SOFTWARE\Policies\Microsoft\Edge\RestoreOnStartupURLs"
New-Item -Path $startupPages -Force | Out-Null
Set-ItemProperty -Path $startupPages -Name "1" -Value "about:blank" -Type String

# Disable news feed / Discover / Shopping / Copilot sidebar
Set-ItemProperty -Path $edgePolicies -Name "NewTabPageContentEnabled" -Value 0 -Type DWord
Set-ItemProperty -Path $edgePolicies -Name "NewTabPageQuickLinksEnabled" -Value 0 -Type DWord
Set-ItemProperty -Path $edgePolicies -Name "HubsSidebarEnabled" -Value 0 -Type DWord
Set-ItemProperty -Path $edgePolicies -Name "EdgeShoppingAssistantEnabled" -Value 0 -Type DWord
Set-ItemProperty -Path $edgePolicies -Name "PersonalizationReportingEnabled" -Value 0 -Type DWord
Set-ItemProperty -Path $edgePolicies -Name "SpotlightExperiencesAndRecommendationsEnabled" -Value 0 -Type DWord
Set-ItemProperty -Path $edgePolicies -Name "ShowRecommendationsEnabled" -Value 0 -Type DWord
Set-ItemProperty -Path $edgePolicies -Name "EdgeCollectionsEnabled" -Value 0 -Type DWord
Set-ItemProperty -Path $edgePolicies -Name "ShowMicrosoftRewards" -Value 0 -Type DWord

# Disable msn/bing redirect on new tab
Set-ItemProperty -Path $edgePolicies -Name "NewTabPageAllowedBackgroundTypes" -Value 3 -Type DWord
Set-ItemProperty -Path $edgePolicies -Name "NewTabPageHideDefaultTopSites" -Value 1 -Type DWord

Write-Host "  Edge cleaned up." -ForegroundColor Green

# ──────────────────────────────────────────────
# 3. CLEAN DESKTOP
# ──────────────────────────────────────────────
Write-Host "[3/$steps] Cleaning up desktop..." -ForegroundColor Yellow

# Remove desktop shortcuts
$desktopPaths = @(
    [Environment]::GetFolderPath('Desktop'),
    [Environment]::GetFolderPath('CommonDesktopDirectory')
)
foreach ($dp in $desktopPaths) {
    Get-ChildItem -Path $dp -Filter "*.lnk" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
}
Write-Host "  Desktop shortcuts removed." -ForegroundColor DarkGray

# Set solid gray wallpaper
Add-Type -TypeDefinition @"
using System.Runtime.InteropServices;
public class Wallpaper {
    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@
# Set wallpaper to none (solid color)
[Wallpaper]::SystemParametersInfo(0x0014, 0, "", 0x0001 -bor 0x0002) | Out-Null
# Set desktop background color to neutral gray (RGB 88,88,88) via registry
Set-ItemProperty -Path "HKCU:\Control Panel\Colors" -Name "Background" -Value "88 88 88"
Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "WallPaper" -Value ""
Write-Host "  Wallpaper set to solid gray." -ForegroundColor DarkGray

# Taskbar cleanup: remove Widgets, Chat, Search, Task View
$taskbarKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
Set-ItemProperty -Path $taskbarKey -Name "ShowTaskViewButton" -Value 0 -Type DWord -ErrorAction SilentlyContinue
Set-ItemProperty -Path $taskbarKey -Name "TaskbarDa" -Value 0 -Type DWord -ErrorAction SilentlyContinue       # Widgets
Set-ItemProperty -Path $taskbarKey -Name "TaskbarMn" -Value 0 -Type DWord -ErrorAction SilentlyContinue       # Chat
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "SearchboxTaskbarMode" -Value 0 -Type DWord -ErrorAction SilentlyContinue
Write-Host "  Taskbar cleaned up." -ForegroundColor DarkGray

# Disable Windows tips & suggestions notifications
$contentDelivery = "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
Set-ItemProperty -Path $contentDelivery -Name "SubscribedContent-310093Enabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue
Set-ItemProperty -Path $contentDelivery -Name "SubscribedContent-338389Enabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue
Set-ItemProperty -Path $contentDelivery -Name "SubscribedContent-338393Enabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue
Set-ItemProperty -Path $contentDelivery -Name "SubscribedContent-353694Enabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue
Set-ItemProperty -Path $contentDelivery -Name "SubscribedContent-353696Enabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue
Set-ItemProperty -Path $contentDelivery -Name "SoftLandingEnabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue
Set-ItemProperty -Path $contentDelivery -Name "SystemPaneSuggestionsEnabled" -Value 0 -Type DWord -ErrorAction SilentlyContinue
Write-Host "  Windows tips & suggestions disabled." -ForegroundColor DarkGray

Write-Host "  Desktop cleaned up." -ForegroundColor Green

# ──────────────────────────────────────────────
# 4. INSTALL MICROSOFT 365 APPS (Office)
# ──────────────────────────────────────────────
Write-Host "[4/$steps] Microsoft 365 Apps..." -ForegroundColor Yellow

$officeInstalled = Test-Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
if ($officeInstalled) {
    Write-Host "  Already installed - skipping." -ForegroundColor Green
} else {
    # Download ODT
    $odtExe = "$tempDir\ODTSetup.exe"
    $odtUrl = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
    Write-Host "  Downloading Office Deployment Tool..." -ForegroundColor DarkGray
    Invoke-WebRequest -Uri $odtUrl -OutFile $odtExe -UseBasicParsing

    # Download config XML from the repo
    $configXml = "$tempDir\Office365.xml"
    $configUrl = "https://raw.githubusercontent.com/jenssgb/VMInit/master/Office365.xml"
    Invoke-WebRequest -Uri $configUrl -OutFile $configXml -UseBasicParsing

    # Run ODT - download and install
    Write-Host "  Downloading and installing Office (this takes a few minutes)..." -ForegroundColor DarkGray
    $odtProcess = Start-Process -FilePath $odtExe -ArgumentList "/configure `"$configXml`"" -Wait -PassThru -NoNewWindow
    if ($odtProcess.ExitCode -eq 0) {
        Write-Host "  Microsoft 365 Apps installed." -ForegroundColor Green
    } else {
        Write-Host "  Office installation finished with exit code: $($odtProcess.ExitCode)" -ForegroundColor Red
    }
}

# ──────────────────────────────────────────────
# 5. INSTALL MICROSOFT TEAMS (new, v2.0)
# ──────────────────────────────────────────────
Write-Host "[5/$steps] Microsoft Teams..." -ForegroundColor Yellow

$teamsInstalled = Get-AppxPackage -Name "MSTeams" -AllUsers -ErrorAction SilentlyContinue
if ($teamsInstalled) {
    Write-Host "  Already installed - skipping." -ForegroundColor Green
} else {
    $teamsBootstrapper = "$tempDir\teamsbootstrapper.exe"
    $teamsUrl = "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409"
    Write-Host "  Downloading Teams bootstrapper..." -ForegroundColor DarkGray
    Invoke-WebRequest -Uri $teamsUrl -OutFile $teamsBootstrapper -UseBasicParsing

    $teamsProcess = Start-Process -FilePath $teamsBootstrapper -ArgumentList "-p" -Wait -PassThru -NoNewWindow
    if ($teamsProcess.ExitCode -eq 0) {
        Write-Host "  Microsoft Teams installed." -ForegroundColor Green
    } else {
        Write-Host "  Teams installation finished with exit code: $($teamsProcess.ExitCode)" -ForegroundColor Red
    }
}

# ──────────────────────────────────────────────
# 6. INSTALL DEV TOOLS (winget)
# ──────────────────────────────────────────────
Write-Host "[6/$steps] Installing dev tools..." -ForegroundColor Yellow

$wingetTools = @(
    @{ Id = "Microsoft.VisualStudioCode"; Name = "VS Code" },
    @{ Id = "Git.Git";                    Name = "Git" },
    @{ Id = "Microsoft.PowerShell";       Name = "PowerShell 7" },
    @{ Id = "Microsoft.WindowsTerminal";  Name = "Windows Terminal" }
)

foreach ($tool in $wingetTools) {
    $installed = winget list --id $tool.Id --accept-source-agreements 2>&1 | Select-String $tool.Id
    if ($installed) {
        Write-Host "  $($tool.Name) already installed - skipping." -ForegroundColor DarkGray
    } else {
        Write-Host "  Installing $($tool.Name)..." -ForegroundColor DarkGray
        winget install --id $tool.Id -e --silent --accept-package-agreements --accept-source-agreements | Out-Null
        Write-Host "  $($tool.Name) installed." -ForegroundColor DarkGray
    }
}

Write-Host "  Dev tools ready." -ForegroundColor Green

# ──────────────────────────────────────────────
# 7. REFRESH EXPLORER & CLEANUP
# ──────────────────────────────────────────────
Write-Host "[7/$steps] Finishing up..." -ForegroundColor Yellow

# Restart Explorer to apply taskbar & desktop changes
Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  VMInit completed!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Please restart the VM to apply all changes." -ForegroundColor Yellow
Write-Host ""
