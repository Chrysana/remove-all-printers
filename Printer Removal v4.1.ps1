# Silent printer cleanup and gpupdate script with WMI + PrintUIEntry, spooler restart, and verification

# Wait 30 seconds for intial GPO to run
Start-Sleep -Seconds 30

# Define log directory and file
$LogDir = "C:\logon_script"
$LogFile = Join-Path $LogDir "logon_script_log.txt"

# Ensure log directory exists
if (-not (Test-Path $LogDir)) {
    New-Item -Path $LogDir -ItemType Directory -Force | Out-Null
}

# Start log
$TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Add-Content -Path $LogFile -Value "`n=== Logon Script Run: $TimeStamp ==="

# Built-in printers to keep
$KeepPrinters = @(
    "Microsoft Print to PDF",
    "Microsoft XPS Document Writer",
    "Fax",
    "OneNote (Desktop)",
    "Send To OneNote"
)

try {
    # Restart the Print Spooler to unlock stuck printers
    Restart-Service -Name "Spooler" -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3
    Add-Content -Path $LogFile -Value "Print Spooler restarted."

    # Remove all non-built-in printers via WMI
    $PrintersToRemove = Get-WmiObject -Class Win32_Printer | Where-Object { $KeepPrinters -notcontains $_.Name }
    foreach ($printer in $PrintersToRemove) {
        Add-Content -Path $LogFile -Value "Removing printer via WMI: $($printer.Name)"
        $printer.Delete() | Out-Null
    }

    # Check for remaining non-built-in printers
    $RemainingPrinters = Get-WmiObject -Class Win32_Printer | Where-Object { $KeepPrinters -notcontains $_.Name }

    # Remove any remaining printers via PrintUIEntry
    foreach ($printer in $RemainingPrinters) {
        Add-Content -Path $LogFile -Value "Removing printer via PrintUIEntry: $($printer.Name)"
        Start-Process -FilePath "rundll32.exe" -ArgumentList "printui.dll,PrintUIEntry /dl /n `"$($printer.Name)`"" -Wait -NoNewWindow
    }

    # Wait 15 seconds before verification
    Start-Sleep -Seconds 15

    # Verification: check for any remaining non-built-in printers
    $FinalRemaining = Get-WmiObject -Class Win32_Printer | Where-Object { $KeepPrinters -notcontains $_.Name }
    if ($FinalRemaining) {
        Add-Content -Path $LogFile -Value "Warning: Some non-built-in printers remain after cleanup:"
        $FinalRemaining | ForEach-Object { Add-Content -Path $LogFile -Value " - $($_.Name)" }
    } else {
        Add-Content -Path $LogFile -Value "Verification complete: no non-built-in printers remain."
    }

    # Wait 5 seconds before gpupdate
    Start-Sleep -Seconds 5

    # Run gpupdate silently
    Start-Process -FilePath "gpupdate.exe" -ArgumentList "/force" -WindowStyle Hidden
    Add-Content -Path $LogFile -Value "gpupdate started successfully."
}
catch {
    Add-Content -Path $LogFile -Value "Error: $($_.Exception.Message)"
}

Add-Content -Path $LogFile -Value "=== Script complete ===`n"
