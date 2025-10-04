# Windows Bluetooth Volume Fix

## Overview
Automatically monitors Windows Updates and manages Bluetooth AVRCP registry settings to prevent Bluetooth audio volume control issues. Runs via Task Scheduler with multiple triggers and provides system notifications when registry changes require a restart.

**Example**: "This script can fix the max volume issue when the Bluetooth ear phone connected."

**What it does:**
- Monitors for Windows Updates and registry changes
- Automatically fixes `DisableAbsoluteVolume` registry value when Windows updates reset it
- Shows restart notifications when changes are made
- Runs automatically via 5 different triggers (startup, daily, weekly, logon, Windows Update completion)

## Key Features

### Automatic Detection & Fixing
- **Windows Update Monitoring**: Detects Event ID 19 (update completion) and triggers within 2 minutes
- **Registry Management**: Automatically sets `DisableAbsoluteVolume` to 1 when Windows updates reset it to 0
- **Registry Path**: `HKLM\SYSTEM\ControlSet001\Control\Bluetooth\Audio\AVRCP\CT`

### System Notifications
- **Toast Notifications**: Modern Windows 10/11 notifications with interactive buttons
- **Balloon Tips**: Traditional system tray notifications with click-to-restart
- **Message Box**: Fallback for maximum compatibility
- **Context-aware**: Different messages for Windows Update vs scheduled triggers

### Multi-Trigger Execution
1. **System Startup** (2-minute delay)
2. **Daily at 3:00 AM** (maintenance window)
3. **Weekly on Sundays at 1:00 AM** (comprehensive check)
4. **User Logon** (1-minute delay)
5. **Windows Update Completion** (2-minute delay after Event ID 19)

## System Requirements

### Operating System
- **Windows 10** (all editions)

### PowerShell Versions
- **PowerShell 5.1** (built into Windows 10) ✅
- **PowerShell 7.0+** (including 7.4.7) ✅
- **PowerShell Core 6.x** ✅

### Other Requirements
- **Administrator privileges** (required for Task Scheduler and registry access)
- **About 10 minutes** for setup

## Installation

### Files Included
- `WindowsPolicyScript.ps1` - Main automation script
- `SetupTaskScheduler.ps1` - Setup script  
- `Setup.bat` - Easy installer
- `UninstallTaskScheduler.ps1` - Uninstall script
- `Uninstall.bat` - Easy uninstaller

### Quick Installation
1. Download all files to same folder
2. **Right-click** `Setup.bat` → **"Run as administrator"**
3. Wait for setup completion
4. Verify 3 tasks created in Task Scheduler

### Manual Installation
```powershell
# Open PowerShell as Administrator
cd "C:\Path\To\Files"
.\SetupTaskScheduler.ps1
```

### Verify Installation
Check Task Scheduler for these tasks:
- `Win10-BluetoothPolicy-TaskScheduler` (main task)
- `Win10-BluetoothPolicy-TaskScheduler-WindowsUpdateTrigger` (event trigger)
- `Win10-BluetoothPolicy-TaskScheduler-ManualTest` (testing)

## Uninstallation

### Quick Uninstall
1. **Right-click** `Uninstall.bat` → **"Run as administrator"**
2. Choose PowerShell version (5.1 or 7+)
3. Type **"YES"** to confirm removal

### Manual Uninstall
```powershell
# Open PowerShell as Administrator
.\UninstallTaskScheduler.ps1
```

### What Gets Removed
- **All scheduled tasks** (3 tasks)
- **Script files** (`C:\Windows\Scripts\WindowsPolicyScript.ps1`)
- **Log files** (`C:\Windows\Logs\BluetoothPolicyAutomation*.log`)
- **Empty directories** (automatic cleanup)

**Note**: Registry changes are preserved (beneficial for Bluetooth audio).

## Logging System

### Log Location
`C:\Windows\Logs\BluetoothPolicyAutomation.log`

### Log Format
```
[2025-10-05 15:30:45] [LEVEL] [CONTEXT] Message
```

### Execution Contexts
- **[SCHEDULED]**: Regular maintenance runs (startup, daily, weekly, logon)
- **[WU-TRIGGER]**: Triggered by Windows Update completion

### Log Levels
- **[INFO]**: Normal operations and status updates
- **[SUCCESS]**: Successful registry changes and notifications
- **[WARNING]**: Non-critical issues (notification fallbacks)
- **[ERROR]**: Problems requiring attention

### Example Log Entries
```
[2025-10-05 15:30:45] [INFO] [SCHEDULED] Policy execution started
[2025-10-05 16:45:22] [SUCCESS] [WU-TRIGGER] Registry value updated to 1
[2025-10-05 16:45:23] [SUCCESS] [WU-TRIGGER] User notification displayed
```

## Troubleshooting

### Common Issues

**Task Not Running:**
- Check Task Scheduler History (enable if disabled)
- Verify tasks are enabled in Task Scheduler
- Check Windows Event Viewer for task errors

**Registry Access Denied:**
- Ensure scheduled tasks run with highest privileges
- Verify Administrator account is being used

**Notifications Not Appearing:**
- Check Windows notification settings (Settings > System > Notifications)
- Run manual test: Task Scheduler → `Win10-BluetoothPolicy-TaskScheduler-ManualTest`
- Check log file for notification attempt results

**PowerShell Execution Policy:**
- Setup script bypasses execution policy automatically
- If needed: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

**Windows Update Trigger Not Working:**
- Verify Event ID 19 exists in System event log after updates
- Check XML subscription format in Task Scheduler
- Fallback: Regular scheduled runs still work

**Module Installation Fails:**
- Ensure internet connectivity for PowerShell Gallery
- TLS settings are configured automatically by the script

### Manual Testing
1. Open Task Scheduler (`taskschd.msc`)
2. Find `Win10-BluetoothPolicy-TaskScheduler-ManualTest`
3. Right-click → "Run"
4. Watch PowerShell window for output
5. Check log file for results

### Log Analysis
- Review `BluetoothPolicyAutomation.log` for error messages
- Look for both `[SCHEDULED]` and `[WU-TRIGGER]` entries
- Verify registry value changes are logged as `[SUCCESS]`

This automated system ensures your Bluetooth audio volume control continues working correctly after Windows updates without manual intervention.
