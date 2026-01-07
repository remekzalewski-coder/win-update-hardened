# win-update-hardened
Production-grade PowerShell script for automated Windows Updates. Features environment sanitization, explicit driver exclusion, auto-reboot logic, and full audit logging. Designed for stability and unattended execution.
⚠️ WARNING: This script runs with -AutoReboot. The system will restart automatically if updates require it. Do not run this on a machine with unsaved work.

⚙️ How It Works (Under the Hood)
Initialization: Sets up a transcript log file in C:\ProgramData\InstallLogs.

Safety Checks: Verifies network connectivity and enforces TLS 1.2.

Dependency Management:

Saves current PSGallery policy.

Temporarily sets policy to Trusted.

Installs/Imports PSWindowsUpdate module.

Unblocks downloaded module files to bypass execution restrictions.

Execution:

Registers Microsoft Update Service (if missing).

Runs Install-WindowsUpdate with AcceptAll, AutoReboot, and RecurseCycle 3.

Crucially: Filters out updates tagged as Drivers.

Cleanup: Restores the original PSGallery policy and stops logging.

📝 Logging
Every execution creates a timestamped log file for audit purposes: Location: C:\ProgramData\InstallLogs\windowsupdate-YYYYMMDD-HHMMSS.log

⚖️ License
Distributed under the MIT License. See LICENSE for more information.
