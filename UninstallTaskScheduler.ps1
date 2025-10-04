# Task Scheduler Uninstall Script
# Windows Bluetooth Volume Fix (Win10): Remove all scheduled tasks and scripts
# Created: October 2025

param(
    [string]$TaskName = "Win10-BluetoothPolicy-TaskScheduler",
    [string]$ScriptPath = "C:\Windows\Scripts\WindowsPolicyScript.ps1"
)

# Function to check admin privileges
function Test-AdminPrivileges {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to remove all related scheduled tasks
function Remove-PolicyTasks {
    param([string]$BaseName)

    Write-Host "Removing scheduled tasks..." -ForegroundColor Yellow

    $tasksToRemove = @(
        $BaseName,
        "$BaseName-ManualTest",
        "$BaseName-WindowsUpdateTrigger"
    )

    $removedCount = 0
    $totalTasks = $tasksToRemove.Count

    foreach ($taskName in $tasksToRemove) {
        try {
            $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
            if ($existingTask) {
                Write-Host "Removing task: $taskName" -ForegroundColor Cyan
                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction Stop
                Write-Host "Successfully removed: $taskName" -ForegroundColor Green
                $removedCount++
            } else {
                Write-Host "Task not found: $taskName (already removed or never existed)" -ForegroundColor Gray
            }
        } catch {
            Write-Host "Error removing task '$taskName': $($_.Exception.Message)" -ForegroundColor Red
            Write-Host "You may need to remove this task manually using Task Scheduler" -ForegroundColor Yellow
        }
    }

    Write-Host "`nTask removal summary:" -ForegroundColor White
    Write-Host "  Tasks removed: $removedCount out of $totalTasks" -ForegroundColor $(if ($removedCount -eq $totalTasks) { "Green" } else { "Yellow" })

    return $removedCount
}

# Function to remove installed scripts and directories
function Remove-ScriptFiles {
    param([string]$MainScriptPath)

    Write-Host "`nRemoving script files..." -ForegroundColor Yellow

    $removedFiles = 0
    $scriptDir = Split-Path $MainScriptPath -Parent

    # Remove the main script file
    if (Test-Path $MainScriptPath) {
        try {
            Write-Host "Removing script file: $MainScriptPath" -ForegroundColor Cyan
            Remove-Item $MainScriptPath -Force -ErrorAction Stop
            Write-Host "Successfully removed: $MainScriptPath" -ForegroundColor Green
            $removedFiles++
        } catch {
            Write-Host "Error removing script file: $($_.Exception.Message)" -ForegroundColor Red
        }
    } else {
        Write-Host "Script file not found: $MainScriptPath (already removed or never existed)" -ForegroundColor Gray
    }

    # Check if script directory is empty and remove if so
    if (Test-Path $scriptDir) {
        try {
            $remainingFiles = Get-ChildItem $scriptDir -ErrorAction SilentlyContinue
            if ($remainingFiles.Count -eq 0) {
                Write-Host "Removing empty script directory: $scriptDir" -ForegroundColor Cyan
                Remove-Item $scriptDir -Force -ErrorAction Stop
                Write-Host "Successfully removed directory: $scriptDir" -ForegroundColor Green
            } else {
                Write-Host "Script directory contains other files, keeping: $scriptDir" -ForegroundColor Yellow
                Write-Host "Remaining files: $($remainingFiles.Count)" -ForegroundColor Yellow
            }
        } catch {
            Write-Host "Error checking/removing script directory: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    return $removedFiles
}

# Function to clean up log files (optional)
function Remove-LogFiles {
    Write-Host "`nCleaning up log files..." -ForegroundColor Yellow

    $logFiles = @(
        "C:\Windows\Logs\BluetoothPolicyAutomation.log",
        "C:\Windows\Logs\BluetoothPolicyAutomation_*.log"
    )

    $removedLogs = 0

    foreach ($logPattern in $logFiles) {
        try {
            $matchingLogs = Get-ChildItem $logPattern -ErrorAction SilentlyContinue
            foreach ($logFile in $matchingLogs) {
                Write-Host "Removing log file: $($logFile.FullName)" -ForegroundColor Cyan
                Remove-Item $logFile.FullName -Force -ErrorAction Stop
                Write-Host "Successfully removed: $($logFile.FullName)" -ForegroundColor Green
                $removedLogs++
            }
        } catch {
            Write-Host "Error removing log files: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    if ($removedLogs -eq 0) {
        Write-Host "No log files found to remove" -ForegroundColor Gray
    } else {
        Write-Host "Removed $removedLogs log file(s)" -ForegroundColor Green
    }

    return $removedLogs
}

# Function to display current system status
function Show-SystemStatus {
    param([string]$BaseName)

    Write-Host "`n=== Current System Status ===`n" -ForegroundColor Magenta

    # Check for remaining tasks
    $remainingTasks = @()
    $tasksToCheck = @(
        $BaseName,
        "$BaseName-ManualTest", 
        "$BaseName-WindowsUpdateTrigger"
    )

    foreach ($taskName in $tasksToCheck) {
        $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($task) {
            $remainingTasks += $taskName
        }
    }

    if ($remainingTasks.Count -eq 0) {
        Write-Host "✓ All scheduled tasks have been removed" -ForegroundColor Green
    } else {
        Write-Host "⚠ Some tasks may still exist:" -ForegroundColor Yellow
        foreach ($task in $remainingTasks) {
            Write-Host "  - $task" -ForegroundColor Red
        }
    }

    # Check for remaining script files
    if (Test-Path $ScriptPath) {
        Write-Host "⚠ Script file still exists: $ScriptPath" -ForegroundColor Yellow
    } else {
        Write-Host "✓ Script files have been removed" -ForegroundColor Green
    }

    Write-Host "`nNote: Any registry changes made by the script remain active." -ForegroundColor Cyan
    Write-Host "If you want to revert registry changes, you'll need to do so manually." -ForegroundColor Cyan
}

# Function to show manual cleanup instructions
function Show-ManualCleanupInstructions {
    Write-Host "`n=== Manual Cleanup Instructions ===`n" -ForegroundColor Yellow

    Write-Host "If automatic removal failed, you can manually remove remaining items:`n" -ForegroundColor White

    Write-Host "Remove Tasks via Task Scheduler:" -ForegroundColor Cyan
    Write-Host "1. Press Win+R, type 'taskschd.msc', press Enter" -ForegroundColor White
    Write-Host "2. Look for tasks containing 'Win10-BluetoothPolicy'" -ForegroundColor White
    Write-Host "3. Right-click each task and select 'Delete'" -ForegroundColor White

    Write-Host "`nRemove Script Files:" -ForegroundColor Cyan
    Write-Host "1. Open File Explorer as Administrator" -ForegroundColor White
    Write-Host "2. Navigate to C:\Windows\Scripts\" -ForegroundColor White
    Write-Host "3. Delete 'WindowsPolicyScript.ps1' if it exists" -ForegroundColor White
    Write-Host "4. Remove the Scripts folder if it's empty" -ForegroundColor White

    Write-Host "`nRemove Log Files:" -ForegroundColor Cyan
    Write-Host "1. Navigate to C:\Windows\Logs\" -ForegroundColor White
    Write-Host "2. Delete any files starting with 'BluetoothPolicyAutomation'" -ForegroundColor White

    Write-Host "`nRevert Registry Changes (Optional):" -ForegroundColor Cyan
    Write-Host "The script modified Bluetooth AVRCP settings. If you want to revert:" -ForegroundColor White
    Write-Host "1. Open Registry Editor (regedit) as Administrator" -ForegroundColor White
    Write-Host "2. Navigate to: HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\BthAvrcpTg" -ForegroundColor White
    Write-Host "3. Look for values that were modified by the script" -ForegroundColor White
    Write-Host "4. Consult Windows documentation for default values" -ForegroundColor White
}

# Main execution
Write-Host "`n=== Windows 10 Bluetooth Policy – Task Scheduler Uninstaller ===`n" -ForegroundColor Magenta
Write-Host "Removing Bluetooth AVRCP Registry Management System" -ForegroundColor Magenta
Write-Host "Task Name: $TaskName" -ForegroundColor White
Write-Host "Script Path: $ScriptPath`n" -ForegroundColor White

# Display PowerShell version information
Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor Cyan
Write-Host "PowerShell Edition: $($PSVersionTable.PSEdition)`n" -ForegroundColor Cyan

# Check admin privileges
if (-not (Test-AdminPrivileges)) {
    Write-Host "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "Please right-click on PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "Administrator privileges confirmed" -ForegroundColor Green

# Show what will be removed
Write-Host "`n=== Uninstallation Plan ===`n" -ForegroundColor White
Write-Host "The following items will be removed:" -ForegroundColor Yellow
Write-Host "• Scheduled Task: $TaskName" -ForegroundColor White
Write-Host "• Scheduled Task: $TaskName-ManualTest" -ForegroundColor White  
Write-Host "• Scheduled Task: $TaskName-WindowsUpdateTrigger" -ForegroundColor White
Write-Host "• Script File: $ScriptPath" -ForegroundColor White
Write-Host "• Log Files: C:\Windows\Logs\BluetoothPolicyAutomation*.log" -ForegroundColor White
Write-Host "• Empty script directory (if applicable)" -ForegroundColor White

# Final confirmation
Write-Host "`nWARNING: This action cannot be undone!" -ForegroundColor Red
$confirmation = Read-Host "Do you want to continue? Type 'YES' to proceed"

if ($confirmation -ne "YES") {
    Write-Host "`nUninstallation cancelled by user." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 0
}

# Perform uninstallation
Write-Host "`n=== Starting Uninstallation ===`n" -ForegroundColor Green

# Remove scheduled tasks
$tasksRemoved = Remove-PolicyTasks -BaseName $TaskName

# Remove script files
$filesRemoved = Remove-ScriptFiles -MainScriptPath $ScriptPath

# Remove log files
$logsRemoved = Remove-LogFiles

# Show final status
Show-SystemStatus -BaseName $TaskName

# Summary
Write-Host "`n=== Uninstallation Summary ===`n" -ForegroundColor Green
Write-Host "Tasks removed: $tasksRemoved" -ForegroundColor White
Write-Host "Files removed: $filesRemoved" -ForegroundColor White
Write-Host "Logs removed: $logsRemoved" -ForegroundColor White

if ($tasksRemoved -eq 3 -and $filesRemoved -gt 0) {
    Write-Host "`n✓ Uninstallation completed successfully!" -ForegroundColor Green
    Write-Host "The Windows 10 Bluetooth Policy automation has been completely removed." -ForegroundColor White
} else {
    Write-Host "`n⚠ Uninstallation completed with some issues" -ForegroundColor Yellow
    Write-Host "Some items may require manual removal. See instructions below." -ForegroundColor Yellow
    Show-ManualCleanupInstructions
}

Write-Host "`nPress Enter to exit..." -ForegroundColor Gray
Read-Host
