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

# ──────────────────────────────────────────────
# 1. REMOVE BLOATWARE
# ──────────────────────────────────────────────
Write-Host "[1/4] Removing Windows 11 bloatware..." -ForegroundColor Yellow

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
Write-Host "[2/4] Cleaning up Microsoft Edge..." -ForegroundColor Yellow

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
# 3. INSTALL MICROSOFT 365 APPS (Office)
# ──────────────────────────────────────────────
Write-Host "[3/4] Installing Microsoft 365 Apps..." -ForegroundColor Yellow

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

# ──────────────────────────────────────────────
# 4. INSTALL MICROSOFT TEAMS (new, v2.0)
# ──────────────────────────────────────────────
Write-Host "[4/4] Installing Microsoft Teams..." -ForegroundColor Yellow

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

# ──────────────────────────────────────────────
# CLEANUP
# ──────────────────────────────────────────────
Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  VMInit completed!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Please restart the VM to apply all changes." -ForegroundColor Yellow
Write-Host ""
