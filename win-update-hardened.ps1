#Requires -RunAsAdministrator
#Requires -Version 5.1
<#
.SYNOPSIS
  win-update-hardened (v3.2 - Sanitized, TLS12, PSGallery restore)

.DESCRIPTION
  - Transcript logging to C:\ProgramData\InstallLogs
  - Forces TLS 1.2 for older images
  - Temporarily sets PSGallery to Trusted, then restores original policy in finally
  - Installs/Imports PSWindowsUpdate
  - Registers Microsoft Update service
  - Preview + Install updates (Drivers excluded)
  - AutoReboot + RecurseCycle 3

.WARNING
  This script can reboot the machine automatically.
#>

$ErrorActionPreference = 'Stop'

$logDir = Join-Path $env:ProgramData 'InstallLogs'
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$logFile = Join-Path $logDir ("win-update-hardened-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
Start-Transcript -Path $logFile -Force | Out-Null
Write-Host "Log file: $logFile"

function Set-Tls12 {
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
}

function Test-WUConnectivity {
    try {
        Resolve-DnsName download.windowsupdate.com -ErrorAction Stop | Out-Null
        return $true
    } catch { return $false }
}

$moduleName = 'PSWindowsUpdate'
$psGalleryOriginalPolicy = $null

try {
    try { Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force } catch {}
    Set-Tls12

    if (-not (Test-WUConnectivity)) {
        throw "No network or Windows Update DNS resolution (download.windowsupdate.com)."
    }

    # Save PSGallery policy and set temporary Trusted
    try {
        $repo = Get-PSRepository -Name PSGallery -ErrorAction Stop
        $psGalleryOriginalPolicy = $repo.InstallationPolicy
        if ($repo.InstallationPolicy -ne 'Trusted') {
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
        }
    } catch {
        Write-Warning "Could not read/set PSGallery policy. Continuing (module install may fail)."
    }

    # NuGet provider
    try { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue | Out-Null } catch {}

    # Install module if missing
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Install-Module -Name $moduleName -Force -Scope AllUsers -AllowClobber
    }

    Import-Module $moduleName -Force

    # Unblock module files (RemoteSigned friendly)
    try {
        $mod = Get-Module -ListAvailable -Name $moduleName | Sort-Object Version -Descending | Select-Object -First 1
        if ($mod) {
            Get-ChildItem -Path $mod.ModuleBase -Recurse -Include *.ps1,*.psm1,*.psd1 |
                ForEach-Object { try { Unblock-File -Path $_.FullName } catch {} }
        }
    } catch {}

    # Register Microsoft Update service (optional)
    $muId = '7971f918-a847-4430-9279-4a52d1efe18d'
    try {
        if (-not (Get-WUServiceManager | Where-Object ServiceID -eq $muId)) {
            Add-WUServiceManager -ServiceID $muId -Confirm:$false | Out-Null
        }
    } catch {
        Write-Warning "Could not register Microsoft Update Service. Continuing with standard Windows Update."
    }

    # Preview (match install filters!)
    Write-Host "Checking for updates (excluding Drivers)..."
    Get-WindowsUpdate -MicrosoftUpdate -IsInstalled:$false -NotCategory 'Drivers' |
        Format-Table -AutoSize | Out-String | Write-Host

    # Install
    Write-Host "Starting installation cycle (AutoReboot, RecurseCycle 3, Drivers excluded)..."
    Install-WindowsUpdate -AcceptAll -MicrosoftUpdate -AutoReboot -RecurseCycle 3 -NotCategory 'Drivers'

    Write-Host "Update run completed." -ForegroundColor Green

} catch {
    Write-Error "Critical script failure: $($_.Exception.Message)"
    exit 2
} finally {
    # Restore PSGallery policy
    if ($psGalleryOriginalPolicy) {
        try { Set-PSRepository -Name PSGallery -InstallationPolicy $psGalleryOriginalPolicy } catch {}
    }
    try { Stop-Transcript | Out-Null } catch {}
}

