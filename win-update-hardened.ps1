#Requires -RunAsAdministrator
#Requires -Version 5.1
<#
.SYNOPSIS
  Windows Update via PSWindowsUpdate (encoding-safe v3 - FIX 2026-01-07).
.DESCRIPTION
  - Uses ExecutionPolicy Bypass for this process only.
  - Trusts PSGallery, installs NuGet provider and PSWindowsUpdate.
  - Unblocks module files.
  - Adds Microsoft Update service.
  - Runs 3 update cycles with AutoReboot.
  - EXPLICITLY SKIPS DRIVERS during installation.
#>

$ErrorActionPreference = 'Stop'

# --- Logging ------------------------------------------------------------------
$logDir = Join-Path $env:ProgramData 'InstallLogs'
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$logFile = Join-Path $logDir ("windowsupdate-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
Start-Transcript -Path $logFile -Force | Out-Null

# --- Execution Policy (process only) ------------------------------------------
try { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force } catch {}

# --- Network check ------------------------------------------------------------
try { 
    Resolve-DnsName download.windowsupdate.com -ErrorAction Stop | Out-Null 
} catch { 
    Write-Error "No network or Windows Update DNS resolution - aborting."
    Stop-Transcript | Out-Null
    exit 1
}

# --- PSGallery & NuGet --------------------------------------------------------
try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted } catch {}
# Added -ErrorAction SilentlyContinue to avoid stop on minor version mismatches if already installed
try { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue } catch {}

# --- PSWindowsUpdate ----------------------------------------------------------
$moduleName = 'PSWindowsUpdate'
if (-not (Get-Module -ListAvailable -Name $moduleName)) {
    Install-Module -Name $moduleName -Force -Scope AllUsers -AllowClobber
}

# Unblock module files (RemoteSigned friendly)
try {
    $mod = Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1
    if ($mod) {
        Get-ChildItem -Path $mod.ModuleBase -Recurse -Include *.ps1,*.psm1,*.psd1 | ForEach-Object {
            try { Unblock-File -Path $_.FullName } catch {}
        }
    }
} catch {}

Import-Module $moduleName -Force

# --- Windows Update settings ---------------------------------------------------
# [FIX 2026-01-07]: Removed Set-WUSettings command as it caused parameter errors in newer builds.
# Driver exclusion is now handled strictly by the -NotCategory 'Drivers' flag in the install command.
Write-Host "Configuration: Driver exclusion will be applied during installation phase."

# --- Add Microsoft Update ------------------------------------------------------
$muId = '7971f918-a847-4430-9279-4a52d1efe18d'
if (-not (Get-WUServiceManager | Where-Object ServiceID -eq $muId)) {
    try {
        Add-WUServiceManager -ServiceID $muId -Confirm:$false
    } catch {
        Write-Warning "Could not register Microsoft Update Service. Continuing with standard Windows Update."
    }
}

# --- Preview updates (for log) -------------------------------------------------
Write-Host "Checking for updates..."
Get-WindowsUpdate -MicrosoftUpdate -IsInstalled:$false | Format-Table -AutoSize | Out-String | Write-Host

# --- Install updates -----------------------------------------------------------
# SAFETY NOTE: -NotCategory 'Drivers' ensures no drivers are installed, preventing crashes.
Write-Host "Starting installation cycle..."
Install-WindowsUpdate -AcceptAll -MicrosoftUpdate -AutoReboot -RecurseCycle 3 -NotCategory 'Drivers'

Stop-Transcript | Out-Null