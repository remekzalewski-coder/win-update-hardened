#Requires -RunAsAdministrator
#Requires -Version 5.1

<#
.SYNOPSIS
  win-update-hardened.ps1 — Production-grade Windows Update via PSWindowsUpdate.
  Features: Driver exclusion, Environment sanitization, Full Audit Logging.

.DESCRIPTION
  - Transcript logging to C:\ProgramData\InstallLogs
  - Network preflight (DNS check)
  - Forces TLS 1.2 for PSGallery / Microsoft endpoints
  - Temporarily sets PSGallery InstallationPolicy=Trusted, then ALWAYS restores original (finally block)
  - Installs/Imports PSWindowsUpdate module
  - Registers Microsoft Update service (if missing)
  - EXPLICITLY SKIPS DRIVERS via -NotCategory 'Drivers'
  - Recursive update cycles (default 3)

.PARAMETER PreviewOnly
  If set, only lists available updates (excluding drivers) without installing.

.PARAMETER RecurseCycle
  Number of install cycles (default 3) to catch cumulative updates.

.PARAMETER AutoReboot
  If $true (default), the system will reboot automatically if required.

.EXAMPLE
  .\win-update-hardened.ps1 -PreviewOnly
.EXAMPLE
  .\win-update-hardened.ps1 -AutoReboot $false
#>

[CmdletBinding()]
param(
    [switch]$PreviewOnly,
    
    [ValidateRange(1,10)]
    [int]$RecurseCycle = 3,
    
    [bool]$AutoReboot = $true
)

$ErrorActionPreference = 'Stop'

# --- Logging ------------------------------------------------------------------
$logDir  = Join-Path $env:ProgramData 'InstallLogs'
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$logFile = Join-Path $logDir ("windowsupdate-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
Start-Transcript -Path $logFile -Force | Out-Null

Write-Host "Log file: $logFile"

# Track state for restoration in 'finally' block
$originalPSGalleryPolicy = $null
$psGalleryExists = $false

try {
    # --- Execution Policy (process only) ----------------------------------------
    try { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force } catch {}

    # --- TLS 1.2 enforcement -----------------------------------------------------
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    } catch {
        Write-Warning "Could not enforce TLS 1.2 via ServicePointManager. Continuing (may fail on older systems)."
    }

    # --- Network preflight -------------------------------------------------------
    try {
        Resolve-DnsName download.windowsupdate.com -ErrorAction Stop | Out-Null
    } catch {
        throw "No network / DNS resolution for Windows Update endpoints. Aborting."
    }

    # --- PSGallery policy (save -> temp trust) ----------------------------------
    try {
        $repo = Get-PSRepository -Name 'PSGallery' -ErrorAction Stop
        $psGalleryExists = $true
        $originalPSGalleryPolicy = $repo.InstallationPolicy
        
        if ($repo.InstallationPolicy -ne 'Trusted') {
            Write-Host "Temporarily trusting PSGallery..."
            Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction Stop
        }
    } catch {
        Write-Warning "PSGallery repository not available or cannot be queried. Module install may fail. Details: $($_.Exception.Message)"
    }

    # --- NuGet provider ----------------------------------------------------------
    try {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue | Out-Null
    } catch {
        Write-Warning "NuGet provider install had issues (may still be OK if already present). Details: $($_.Exception.Message)"
    }

    # --- PSWindowsUpdate module --------------------------------------------------
    $moduleName = 'PSWindowsUpdate'
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Write-Host "Installing PSWindowsUpdate module..."
        Install-Module -Name $moduleName -Repository PSGallery -Scope AllUsers -Force -AllowClobber -ErrorAction Stop
    }

    # Unblock module files (RemoteSigned-friendly)
    $mod = Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1
    if ($mod -and (Test-Path $mod.ModuleBase)) {
        Get-ChildItem -Path $mod.ModuleBase -Recurse -Include *.ps1,*.psm1,*.psd1 -ErrorAction SilentlyContinue |
            ForEach-Object { try { Unblock-File -Path $_.FullName -ErrorAction SilentlyContinue } catch {} }
    }

    Import-Module $moduleName -Force -ErrorAction Stop

    # --- Add Microsoft Update (optional) ----------------------------------------
    $muId = '7971f918-a847-4430-9279-4a52d1efe18d'
    try {
        if (-not (Get-WUServiceManager | Where-Object ServiceID -eq $muId)) {
            Add-WUServiceManager -ServiceID $muId -Confirm:$false | Out-Null
        }
    } catch {
        Write-Warning "Could not register Microsoft Update Service. Continuing with standard Windows Update."
    }

    # --- Preview (must match install filters) ------------------------------------
    Write-Host "Checking for updates (excluding Drivers)..."
    Get-WindowsUpdate -MicrosoftUpdate -IsInstalled:$false -NotCategory 'Drivers' |
        Format-Table -AutoSize | Out-String | Write-Host

    if ($PreviewOnly) {
        Write-Host "PreviewOnly set. Exiting without installation."
        return
    }

    # --- Install -----------------------------------------------------------------
    Write-Host "Starting installation cycle (RecurseCycle=$RecurseCycle, AutoReboot=$AutoReboot, excluding Drivers)..."

    if ($AutoReboot) {
        Install-WindowsUpdate -AcceptAll -MicrosoftUpdate -AutoReboot -RecurseCycle $RecurseCycle -NotCategory 'Drivers'
    } else {
        Install-WindowsUpdate -AcceptAll -MicrosoftUpdate -IgnoreReboot -RecurseCycle $RecurseCycle -NotCategory 'Drivers'
        Write-Host "AutoReboot disabled. If reboot is required, restart manually and re-run the script."
    }

} catch {
    Write-Error "CRITICAL FAILURE: $($_.Exception.Message)"
    throw
} finally {
    # --- Cleanup / Restoration ---------------------------------------------------
    if ($psGalleryExists -and $null -ne $originalPSGalleryPolicy) {
        try {
            # Only restore if it was different
            $current = (Get-PSRepository -Name 'PSGallery').InstallationPolicy
            if ($current -ne $originalPSGalleryPolicy) {
                Write-Host "Restoring PSGallery policy to '$originalPSGalleryPolicy'..."
                Set-PSRepository -Name 'PSGallery' -InstallationPolicy $originalPSGalleryPolicy -ErrorAction SilentlyContinue | Out-Null
            }
        } catch {}
    }

    try { Stop-Transcript | Out-Null } catch {}
}
