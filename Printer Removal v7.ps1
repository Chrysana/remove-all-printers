# ==============================================
# Ultimate Nuclear Printer Cleanup Script
# ==============================================

# Wait 30 seconds for intial GPO to run
Start-Sleep -Seconds 30

$LogDir = "C:\logon_script"
$LogFile = Join-Path $LogDir "ultimate_printer_cleanup_log.txt"

if (-not (Test-Path $LogDir)) {
    New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
}

$TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Add-Content -Path $LogFile -Value "`n=== Ultimate Printer Cleanup Run: $TimeStamp ==="

$KeepPrinters = @(
    "Microsoft Print to PDF",
    "Microsoft XPS Document Writer",
    "Fax",
    "OneNote (Desktop)",
    "Send To OneNote"
)

$CannotRemove = @()

# -------------------------
# 1. Stop spooler
# -------------------------
Stop-Service -Name Spooler -Force -ErrorAction SilentlyContinue
Add-Content -Path $LogFile -Value "Print spooler stopped."
Start-Sleep -Seconds 2

# -------------------------
# 2. Clear spooler folder
# -------------------------
$SpoolDir = "C:\Windows\System32\spool\PRINTERS"
if (Test-Path $SpoolDir) {
    Remove-Item -Path "$SpoolDir\*" -Force -Recurse -ErrorAction SilentlyContinue
    Add-Content -Path $LogFile -Value "Cleared spooler cache."
}

# -------------------------
# 3. Remove printers via Remove-Printer (WMI)
# -------------------------
try {
    $Printers = Get-Printer | Where-Object { $KeepPrinters -notcontains $_.Name }
    foreach ($printer in $Printers) {
        try {
            Remove-Printer -Name $printer.Name -ErrorAction Stop
            Add-Content -Path $LogFile -Value "Removed printer via Remove-Printer: $($printer.Name)"
        }
        catch {
            Add-Content -Path $LogFile -Value "Failed Remove-Printer: $($printer.Name)"
            $CannotRemove += $printer.Name
        }
    }
}
catch {
    Add-Content -Path $LogFile -Value "No WMI printers found or error: $($_.Exception.Message)"
}

# -------------------------
# 4. Remove printers from registry paths (system + user)
# -------------------------
$RegistryPaths = @(
    "HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers",
    "HKCU:\Printers\Connections",
    "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Devices",
    "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts"
)

# Ensure HKU drive exists
if (-not (Get-PSDrive -Name HKU -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS | Out-Null
}
$UserHives = Get-ChildItem 'HKU:' | Where-Object { $_.Name -notmatch 'DEFAULT|S-1-5-18|S-1-5-19|S-1-5-20' }

foreach ($path in $RegistryPaths) {
    if (Test-Path $path) {
        Get-ChildItem $path -Recurse | ForEach-Object {
            if ($KeepPrinters -notcontains $_.PSChildName) {
                try {
                    Remove-Item -Path $_.PSPath -Recurse -Force -ErrorAction Stop
                    Add-Content -Path $LogFile -Value "Removed registry printer: $($_.PSPath)"
                }
                catch {
                    Add-Content -Path $LogFile -Value "Failed registry removal: $($_.PSPath)"
                    $CannotRemove += $_.PSChildName
                }
            }
        }
    }
}

# Remove from loaded user hives
foreach ($hive in $UserHives) {
    foreach ($subPath in @("Printers\Connections","Software\Microsoft\Windows NT\CurrentVersion\Devices","Software\Microsoft\Windows NT\CurrentVersion\PrinterPorts")) {
        $FullPath = "$($hive.PSPath)\$subPath"
        if (Test-Path $FullPath) {
            Get-ChildItem $FullPath -Recurse | ForEach-Object {
                try {
                    Remove-Item -Path $_.PSPath -Recurse -Force -ErrorAction Stop
                    Add-Content -Path $LogFile -Value "Removed from loaded hive: $($_.PSPath)"
                }
                catch {
                    Add-Content -Path $LogFile -Value "Failed removal from loaded hive: $($_.PSPath)"
                    $CannotRemove += $_.PSChildName
                }
            }
        }
    }
}

# -------------------------
# 5. Remove legacy network printers (LanMan)
# -------------------------
$LanManPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Print\Providers\LanMan Print Services\Servers"
if (Test-Path $LanManPath) {
    $Servers = Get-ChildItem $LanManPath
    foreach ($server in $Servers) {
        $Printers = Get-ChildItem "$($server.PSPath)\Printers" -ErrorAction SilentlyContinue
        foreach ($printer in $Printers) {
            try {
                Remove-Item -Path $printer.PSPath -Recurse -Force -ErrorAction Stop
                Add-Content -Path $LogFile -Value "Removed legacy network printer: $($printer.PSChildName)"
            }
            catch {
                Add-Content -Path $LogFile -Value "Failed legacy removal: $($printer.PSChildName)"
                $CannotRemove += $printer.PSChildName
            }
        }
    }
}

# -------------------------
# 6. Purge Explorer shell ghost printers
# -------------------------
# Stop Explorer
Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
Add-Content -Path $LogFile -Value "Explorer stopped."

$ShellPaths = @(
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace",
    "HKCU:\Software\Classes\CLSID"
)

foreach ($path in $ShellPaths) {
    if (Test-Path $path) {
        Get-ChildItem $path -Recurse | ForEach-Object {
            if ($KeepPrinters -notcontains $_.PSChildName) {
                try {
                    Remove-Item -Path $_.PSPath -Recurse -Force -ErrorAction Stop
                    Add-Content -Path $LogFile -Value "Removed shell object: $($_.PSPath)"
                }
                catch {
                    Add-Content -Path $LogFile -Value "Failed shell removal: $($_.PSPath)"
                    $CannotRemove += $_.PSChildName
                }
            }
        }
    }
}

# Delete icon cache
$IconCacheFiles = Get-ChildItem "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\iconcache*.db" -ErrorAction SilentlyContinue
foreach ($file in $IconCacheFiles) {
    try {
        Remove-Item $file.FullName -Force -ErrorAction Stop
        Add-Content -Path $LogFile -Value "Deleted icon cache: $($file.FullName)"
    }
    catch {
        Add-Content -Path $LogFile -Value "Failed to delete icon cache: $($file.FullName)"
    }
}

# Restart Explorer
Start-Process explorer.exe
Add-Content -Path $LogFile -Value "Explorer restarted."

# Restart spooler
Start-Service -Name Spooler -ErrorAction SilentlyContinue
Add-Content -Path $LogFile -Value "Print spooler restarted."

# -------------------------
# 7. Report printers that could not be removed
# -------------------------
$CannotRemove = $CannotRemove | Select-Object -Unique
if ($CannotRemove.Count -gt 0) {
    Add-Content -Path $LogFile -Value "`n=== Printers that could NOT be removed programmatically ==="
    foreach ($printer in $CannotRemove) {
        Add-Content -Path $LogFile -Value $printer
    }
    Add-Content -Path $LogFile -Value "`nThese are likely COM/Explorer ghost printers or GPO/server-deployed."
} else {
    Add-Content -Path $LogFile -Value "`nAll printers removed successfully."
}

# -------------------------
# 8. Run gpupdate silently after 15s
# -------------------------
Start-Sleep -Seconds 15
Start-Process -FilePath "gpupdate.exe" -ArgumentList "/force" -WindowStyle Hidden
Add-Content -Path $LogFile -Value "gpupdate started successfully."

# Open log automatically
Start-Process notepad.exe $LogFile
