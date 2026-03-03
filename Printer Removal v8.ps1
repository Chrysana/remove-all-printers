#Requires -RunAsAdministrator
# ==============================================
# Ultimate Nuclear Printer Cleanup Script v8
# ==============================================
[CmdletBinding(SupportsShouldProcess)]
param(
    [int]$InitialWaitSeconds = 30,
    [int]$GpupdateWaitSeconds = 15,
    [string]$LogDir = "C:\logon_script",
    [switch]$OpenLog
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ------------------------------------
# Logging helper
# ------------------------------------
$LogFile = Join-Path $LogDir "ultimate_printer_cleanup_log.txt"

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $Line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -Path $LogFile -Value $Line
    Write-Verbose $Line
}

function Rotate-Log {
    # Keep the log under ~2 MB by trimming oldest lines
    $MaxBytes = 2MB
    if ((Test-Path $LogFile) -and (Get-Item $LogFile).Length -gt $MaxBytes) {
        $Lines = Get-Content $LogFile
        $Trimmed = $Lines | Select-Object -Last 500
        Set-Content -Path $LogFile -Value $Trimmed
    }
}

# ------------------------------------
# Initialise log directory
# ------------------------------------
if (-not (Test-Path $LogDir)) {
    New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
}
Rotate-Log
Write-Log "=== Ultimate Printer Cleanup Run ==="

$KeepPrinters = @(
    "Microsoft Print to PDF",
    "Microsoft XPS Document Writer",
    "Fax",
    "OneNote (Desktop)",
    "Send To OneNote"
)

$CannotRemove = [System.Collections.Generic.List[string]]::new()

# ------------------------------------
# Wait for initial GPO pass
# ------------------------------------
Write-Log "Waiting ${InitialWaitSeconds}s for initial GPO run..."
Start-Sleep -Seconds $InitialWaitSeconds

# ------------------------------------
# Ensure spooler is restarted even if the script fails
# ------------------------------------
$SpoolerStopped = $false
try {
    # -------------------------
    # 1. Stop spooler
    # -------------------------
    Stop-Service -Name Spooler -Force -ErrorAction SilentlyContinue
    $SpoolerStopped = $true
    Write-Log "Print spooler stopped."
    Start-Sleep -Seconds 2

    # -------------------------
    # 2. Clear spooler folder
    # -------------------------
    $SpoolDir = "$env:SystemRoot\System32\spool\PRINTERS"
    if (Test-Path $SpoolDir) {
        if ($PSCmdlet.ShouldProcess($SpoolDir, 'Clear spooler cache')) {
            Remove-Item -Path "$SpoolDir\*" -Force -Recurse -ErrorAction SilentlyContinue
            Write-Log "Cleared spooler cache."
        }
    }

    # -------------------------
    # 3. Remove printers via Remove-Printer
    # -------------------------
    try {
        $Printers = Get-Printer -ErrorAction Stop | Where-Object { $KeepPrinters -notcontains $_.Name }
        foreach ($printer in $Printers) {
            if ($PSCmdlet.ShouldProcess($printer.Name, 'Remove-Printer')) {
                try {
                    Remove-Printer -Name $printer.Name -ErrorAction Stop
                    Write-Log "Removed printer: $($printer.Name)"
                }
                catch {
                    Write-Log "Failed Remove-Printer '$($printer.Name)': $($_.Exception.Message)" 'WARN'
                    $CannotRemove.Add($printer.Name)
                }
            }
        }
    }
    catch {
        Write-Log "Get-Printer error: $($_.Exception.Message)" 'WARN'
    }

    # -------------------------
    # 4. Remove printers from registry (system + current user)
    # -------------------------
    # These paths store printer entries as direct children — no recursion needed.
    $DirectChildPaths = @(
        "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers",
        "HKCU:\Printers\Connections",
        "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Devices",
        "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts"
    )

    foreach ($path in $DirectChildPaths) {
        if (-not (Test-Path $path)) { continue }
        $Entries = Get-ChildItem $path -ErrorAction SilentlyContinue
        foreach ($entry in $Entries) {
            if ($KeepPrinters -notcontains $entry.PSChildName) {
                if ($PSCmdlet.ShouldProcess($entry.PSPath, 'Remove registry key')) {
                    try {
                        Remove-Item -Path $entry.PSPath -Recurse -Force -ErrorAction Stop
                        Write-Log "Removed registry key: $($entry.PSPath)"
                    }
                    catch {
                        Write-Log "Failed registry removal '$($entry.PSPath)': $($_.Exception.Message)" 'WARN'
                        $CannotRemove.Add($entry.PSChildName)
                    }
                }
            }
        }
    }

    # Remove from all loaded user hives (skip system accounts and _Classes hives)
    if (-not (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue)) {
        New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS | Out-Null
    }
    $UserHives = Get-ChildItem 'HKU:' -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch '(DEFAULT|S-1-5-18|S-1-5-19|S-1-5-20|_Classes$)' }

    foreach ($hive in $UserHives) {
        foreach ($subPath in @(
            'Printers\Connections',
            'Software\Microsoft\Windows NT\CurrentVersion\Devices',
            'Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts'
        )) {
            $FullPath = Join-Path $hive.PSPath $subPath
            if (-not (Test-Path $FullPath)) { continue }
            $Entries = Get-ChildItem $FullPath -ErrorAction SilentlyContinue
            foreach ($entry in $Entries) {
                if ($PSCmdlet.ShouldProcess($entry.PSPath, 'Remove registry key from user hive')) {
                    try {
                        Remove-Item -Path $entry.PSPath -Recurse -Force -ErrorAction Stop
                        Write-Log "Removed from user hive: $($entry.PSPath)"
                    }
                    catch {
                        Write-Log "Failed user hive removal '$($entry.PSPath)': $($_.Exception.Message)" 'WARN'
                        $CannotRemove.Add($entry.PSChildName)
                    }
                }
            }
        }
    }

    # -------------------------
    # 5. Remove legacy LanMan network printers
    # -------------------------
    $LanManPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\LanMan Print Services\Servers"
    if (Test-Path $LanManPath) {
        foreach ($server in (Get-ChildItem $LanManPath -ErrorAction SilentlyContinue)) {
            $PrintersPath = "$($server.PSPath)\Printers"
            if (-not (Test-Path $PrintersPath)) { continue }
            foreach ($printer in (Get-ChildItem $PrintersPath -ErrorAction SilentlyContinue)) {
                if ($PSCmdlet.ShouldProcess($printer.PSChildName, 'Remove LanMan printer')) {
                    try {
                        Remove-Item -Path $printer.PSPath -Recurse -Force -ErrorAction Stop
                        Write-Log "Removed legacy network printer: $($printer.PSChildName)"
                    }
                    catch {
                        Write-Log "Failed legacy removal '$($printer.PSChildName)': $($_.Exception.Message)" 'WARN'
                        $CannotRemove.Add($printer.PSChildName)
                    }
                }
            }
        }
    }

    # -------------------------
    # 6. Purge Explorer ghost printers from shell namespace
    # -------------------------
    # Only target the Desktop\NameSpace entries whose default value matches a
    # printer name — avoids blanket deletion of unrelated shell extensions.
    # HKCU:\Software\Classes\CLSID is intentionally excluded: it contains ALL
    # user COM registrations (GUIDs), not printer names, so filtering by printer
    # name there has no effect and would delete unrelated system objects.
    $NamespacePath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace"
    if (Test-Path $NamespacePath) {
        $GhostEntries = Get-ChildItem $NamespacePath -ErrorAction SilentlyContinue | Where-Object {
            $defaultVal = (Get-ItemProperty -Path $_.PSPath -Name '(default)' -ErrorAction SilentlyContinue).'(default)'
            # Keep the entry only if its display name is a printer we want to remove
            $defaultVal -and ($KeepPrinters -notcontains $defaultVal)
        }

        if ($GhostEntries) {
            # Stop Explorer only when there is actually something to clean
            Write-Log "Stopping Explorer to clean ghost namespace entries..."
            Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2

            foreach ($entry in $GhostEntries) {
                if ($PSCmdlet.ShouldProcess($entry.PSPath, 'Remove ghost namespace entry')) {
                    try {
                        Remove-Item -Path $entry.PSPath -Recurse -Force -ErrorAction Stop
                        Write-Log "Removed ghost namespace entry: $($entry.PSPath)"
                    }
                    catch {
                        Write-Log "Failed namespace removal '$($entry.PSPath)': $($_.Exception.Message)" 'WARN'
                        $CannotRemove.Add($entry.PSChildName)
                    }
                }
            }

            # Delete icon cache so stale printer icons don't linger
            Get-ChildItem "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\iconcache*.db" -ErrorAction SilentlyContinue |
                ForEach-Object {
                    try {
                        Remove-Item $_.FullName -Force -ErrorAction Stop
                        Write-Log "Deleted icon cache: $($_.FullName)"
                    }
                    catch {
                        Write-Log "Failed to delete icon cache '$($_.FullName)': $($_.Exception.Message)" 'WARN'
                    }
                }

            Start-Process explorer.exe
            Write-Log "Explorer restarted."
        }
        else {
            Write-Log "No ghost namespace entries found; Explorer not disturbed."
        }
    }
}
finally {
    # Always restart the spooler, even if something above threw
    if ($SpoolerStopped) {
        Start-Service -Name Spooler -ErrorAction SilentlyContinue
        Write-Log "Print spooler restarted."
    }
}

# -------------------------
# 7. Report printers that could not be removed
# -------------------------
$Unique = $CannotRemove | Select-Object -Unique
if ($Unique.Count -gt 0) {
    Write-Log "=== Printers/keys that could NOT be removed ===" 'WARN'
    $Unique | ForEach-Object { Write-Log "  - $_" 'WARN' }
    Write-Log "These are likely GPO/server-deployed or require manual intervention." 'WARN'
}
else {
    Write-Log "All printers removed successfully."
}

# -------------------------
# 8. Run gpupdate silently
# -------------------------
Write-Log "Waiting ${GpupdateWaitSeconds}s before gpupdate..."
Start-Sleep -Seconds $GpupdateWaitSeconds
Start-Process -FilePath "gpupdate.exe" -ArgumentList "/force" -WindowStyle Hidden
Write-Log "gpupdate /force started."

# Open log only when explicitly requested (not suitable for silent logon scripts)
if ($OpenLog) {
    Start-Process notepad.exe $LogFile
}

Write-Log "=== Cleanup complete ==="
