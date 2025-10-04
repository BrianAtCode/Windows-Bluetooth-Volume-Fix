# Windows Bluetooth Volume Fix
# Task Scheduler Implementation with Windows Update Completion Trigger
# Created: October 2025

param(
    [string]$LogPath = "C:\Windows\Logs\UpdateBluetoothPolicy.log",
    [switch]$TriggeredByWindowsUpdate
)

# Function to write log entries
function Write-PolicyLog {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $triggerContext = if ($TriggeredByWindowsUpdate) { "WU-TRIGGER" } else { "SCHEDULED" }
    $logEntry = "[$timestamp] [$Level] [$triggerContext] $Message"

    # Create log directory if it doesn't exist
    $logDir = Split-Path $LogPath -Parent
    if (-not (Test-Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    # Write to log file and console
    $logEntry | Out-File -FilePath $LogPath -Append -Encoding UTF8
    Write-Host $logEntry
}

# Function to show system notification for restart reminder
function Show-RestartNotification {
    param(
        [string]$Title = "Bluetooth Audio Settings Updated",
        [string]$Message = "Your Bluetooth AVRCP settings have been updated to fix audio volume control issues. Please restart your computer when convenient to ensure all changes take effect."
    )

    try {
        Write-PolicyLog "Attempting to show restart notification to user..." "INFO"

        # Method 1: Try using Windows Toast Notifications (Windows 10/11)
        try {
            # Load required assemblies for toast notifications
            [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
            [Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom.XmlDocument, ContentType = WindowsRuntime] > $null

            # Create toast XML template
            $toastXml = @"
<toast duration="long" scenario="reminder">
    <visual>
        <binding template="ToastGeneric">
            <text>$Title</text>
            <text>$Message</text>
            <image placement="appLogoOverride" hint-crop="circle" src="ms-appx:///Images/restart.png"/>
        </binding>
    </visual>
    <actions>
        <action content="Restart Now" arguments="restart" activationType="protocol" />
        <action content="Remind Later" arguments="dismiss" activationType="system" />
    </actions>
    <audio src="ms-winsoundevent:Notification.Default" loop="false" />
</toast>
"@

            # Create XML document and toast notification
            $xmlDoc = New-Object Windows.Data.Xml.Dom.XmlDocument
            $xmlDoc.LoadXml($toastXml)
            $toast = [Windows.UI.Notifications.ToastNotification]::new($xmlDoc)

            # Set expiration time (30 minutes from now)
            $toast.ExpirationTime = [DateTimeOffset]::Now.AddMinutes(30)

            # Create notifier and show toast
            $appId = "Microsoft.Windows.Computer"  # Use system app ID for better compatibility
            $notifier = [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId)
            $notifier.Show($toast)

            Write-PolicyLog "Toast notification displayed successfully" "SUCCESS"
            return $true

        } catch {
            Write-PolicyLog "Toast notification failed: $($_.Exception.Message)" "WARNING"
            # Fall back to balloon tip notification
        }

        # Method 2: Fallback to Balloon Tip Notification (Legacy but more reliable)
        try {
            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -AssemblyName System.Drawing

            # Create NotifyIcon object
            $balloon = New-Object System.Windows.Forms.NotifyIcon

            # Set up the balloon tip properties
            $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Process -Id $PID).Path)
            $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
            $balloon.BalloonTipTitle = $Title
            $balloon.BalloonTipText = $Message
            $balloon.Visible = $true

            # Register event handler for balloon tip click (optional restart action)
            $restartAction = {
                $result = [System.Windows.Forms.MessageBox]::Show(
                    "Would you like to restart your computer now to complete the Bluetooth audio fix?`n`nClick Yes to restart immediately, or No to restart later.",
                    "Restart Computer",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question
                )
                if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
                    Write-PolicyLog "User chose to restart computer immediately" "INFO"
                    # Restart computer with 30 second delay
                    shutdown /r /t 30 /c "Restarting to complete Bluetooth audio settings update..."
                } else {
                    Write-PolicyLog "User chose to restart later" "INFO"
                }
                $balloon.Dispose()
            }

            # Register the click event
            Register-ObjectEvent -InputObject $balloon -EventName BalloonTipClicked -Action $restartAction | Out-Null

            # Show balloon tip for 15 seconds
            $balloon.ShowBalloonTip(15000)

            # Clean up after a delay
            Start-Sleep -Seconds 2

            Write-PolicyLog "Balloon notification displayed successfully" "SUCCESS"

            # Schedule cleanup
            $cleanupTimer = New-Object System.Windows.Forms.Timer
            $cleanupTimer.Interval = 16000  # 16 seconds
            $cleanupTimer.Add_Tick({
                try {
                    $balloon.Dispose()
                    Get-EventSubscriber | Where-Object { $_.SourceObject -eq $balloon } | Unregister-Event
                    $cleanupTimer.Stop()
                    $cleanupTimer.Dispose()
                } catch {
                    # Ignore cleanup errors
                }
            })
            $cleanupTimer.Start()

            return $true

        } catch {
            Write-PolicyLog "Balloon notification failed: $($_.Exception.Message)" "WARNING"
        }

        # Method 3: Final fallback - Simple message box (synchronous)
        try {
            Add-Type -AssemblyName System.Windows.Forms
            [System.Windows.Forms.MessageBox]::Show(
                $Message,
                $Title,
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null

            Write-PolicyLog "Message box notification displayed successfully" "SUCCESS"
            return $true

        } catch {
            Write-PolicyLog "All notification methods failed: $($_.Exception.Message)" "ERROR"
            return $false
        }

    } catch {
        Write-PolicyLog "Error in Show-RestartNotification: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Function to check if running with elevated privileges
function Test-AdminPrivileges {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to detect recent Windows Update activity
function Get-RecentWindowsUpdateActivity {
    try {
        Write-PolicyLog "Checking recent Windows Update activity..." "INFO"

        # Check for recent Windows Update events (last 24 hours)
        $since = (Get-Date).AddHours(-24)

        # Look for successful installations (Event ID 19)
        $successfulUpdates = @()
        try {
            $successfulUpdates = Get-WinEvent -FilterHashtable @{
                LogName = 'System'
                ProviderName = 'Microsoft-Windows-WindowsUpdateClient'
                ID = 19
                StartTime = $since
            } -ErrorAction SilentlyContinue
        } catch {
            Write-PolicyLog "Could not retrieve Event ID 19 from System log: $($_.Exception.Message)" "WARNING"
        }

        # Alternative: Check Microsoft-Windows-WindowsUpdateClient/Operational log
        if (-not $successfulUpdates) {
            try {
                $successfulUpdates = Get-WinEvent -FilterHashtable @{
                    LogName = 'Microsoft-Windows-WindowsUpdateClient/Operational'
                    ID = 19
                    StartTime = $since
                } -ErrorAction SilentlyContinue
            } catch {
                Write-PolicyLog "Could not retrieve Event ID 19 from Operational log: $($_.Exception.Message)" "WARNING"
            }
        }

        if ($successfulUpdates) {
            Write-PolicyLog "Found $($successfulUpdates.Count) successful Windows Update(s) in the last 24 hours" "INFO"
            foreach ($update in $successfulUpdates | Select-Object -First 5) {
                $updateInfo = $update.Message -replace "`r`n", " " -replace "`n", " "
                Write-PolicyLog "  Recent Update: $updateInfo" "INFO"
            }
            return $true
        } else {
            Write-PolicyLog "No recent Windows Update installations found in the last 24 hours" "INFO"
            return $false
        }

    } catch {
        Write-PolicyLog "Error checking Windows Update activity: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Function to install PSWindowsUpdate module if needed
function Install-WindowsUpdateModule {
    try {
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            Write-PolicyLog "PSWindowsUpdate module not found. Installing..." "INFO"

            # Set TLS to 1.2 for compatibility with PowerShell Gallery
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

            # Install module with error handling
            Install-Module -Name PSWindowsUpdate -Force -Scope AllUsers -AllowClobber -ErrorAction Stop
            Write-PolicyLog "PSWindowsUpdate module installed successfully" "SUCCESS"
        } else {
            Write-PolicyLog "PSWindowsUpdate module already available" "INFO"
        }

        # Import the module
        Import-Module PSWindowsUpdate -Force -ErrorAction Stop
        Write-PolicyLog "PSWindowsUpdate module imported successfully" "SUCCESS"
        return $true

    } catch {
        Write-PolicyLog "Failed to install/import PSWindowsUpdate module: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Function to check Windows Updates
function Check-WindowsUpdates {
    try {
        Write-PolicyLog "Starting Windows Update scan..." "INFO"

        # Get available updates
        $availableUpdates = @()
        try {
            $availableUpdates = Get-WindowsUpdate -MicrosoftUpdate -ErrorAction SilentlyContinue
        } catch {
            Write-PolicyLog "Error getting updates via Get-WindowsUpdate: $($_.Exception.Message)" "WARNING"
            # Try alternative method using Windows Update API
            try {
                $updateSession = New-Object -ComObject Microsoft.Update.Session
                $updateSearcher = $updateSession.CreateUpdateSearcher()
                $searchResult = $updateSearcher.Search("IsInstalled=0")
                $availableUpdates = $searchResult.Updates
            } catch {
                Write-PolicyLog "Alternative update check also failed: $($_.Exception.Message)" "ERROR"
            }
        }

        if ($availableUpdates -and $availableUpdates.Count -gt 0) {
            Write-PolicyLog "Found $($availableUpdates.Count) available Windows Update(s)" "INFO"
            foreach ($update in $availableUpdates | Select-Object -First 10) {
                $title = if ($update.Title) { $update.Title } else { $update.ToString() }
                Write-PolicyLog "  Available Update: $title" "INFO"
            }
        } else {
            Write-PolicyLog "No Windows Updates available or system is up to date" "INFO"
        }

        # Get last update information using COM API
        try {
            $session = New-Object -ComObject "Microsoft.Update.Session"
            $history = $session.QueryHistory("", 0, 5)  # Get last 5 updates

            if ($history -and $history.Count -gt 0) {
                Write-PolicyLog "Recent update history (last 5):" "INFO"
                foreach ($historyItem in $history) {
                    $resultText = switch ($historyItem.ResultCode) {
                        2 { "Succeeded" }
                        3 { "Succeeded With Errors" }
                        4 { "Failed" }
                        5 { "Aborted" }
                        default { "Unknown ($($historyItem.ResultCode))" }
                    }
                    Write-PolicyLog "  $($historyItem.Date): $($historyItem.Title) - $resultText" "INFO"
                }
            }
        } catch {
            Write-PolicyLog "Could not retrieve update history: $($_.Exception.Message)" "WARNING"
        }

    } catch {
        Write-PolicyLog "Windows Update check failed: $($_.Exception.Message)" "ERROR"
    }
}

# Function to manage Bluetooth AVRCP registry with Windows update checking
function Manage-BluetoothAVRCP {
    try {
        Write-PolicyLog "Checking Bluetooth AVRCP CT registry value..." "INFO"

        $registryPath = "HKLM:\SYSTEM\ControlSet001\Control\Bluetooth\Audio\AVRCP\CT"
        $valueName = "DisableAbsoluteVolume"
        $changesMade = $false
        $shouldNotifyRestart = $false

        # Check if the registry path exists
        if (-not (Test-Path $registryPath)) {
            Write-PolicyLog "Registry path does not exist. Creating: $registryPath" "INFO"
            try {
                New-Item -Path $registryPath -Force | Out-Null
                Write-PolicyLog "Registry path created successfully" "SUCCESS"
                $changesMade = $true
            } catch {
                Write-PolicyLog "ERROR: Failed to create registry path: $($_.Exception.Message)" "ERROR"
                return $false
            }
        }

        # Check and modify the registry value
        try {
            $currentValue = Get-ItemProperty -Path $registryPath -Name $valueName -ErrorAction SilentlyContinue

            if ($currentValue) {
                $currentRegValue = $currentValue.$valueName
                Write-PolicyLog "Current $valueName value: $currentRegValue" "INFO"

                # If value is 0, set it to 1
                if ($currentRegValue -eq 0) {
                    Write-PolicyLog "Registry value is 0, changing to 1..." "INFO"
                    Set-ItemProperty -Path $registryPath -Name $valueName -Value 1 -Type DWord
                    Write-PolicyLog "Registry value updated successfully to 1" "SUCCESS"
                    $changesMade = $true
                    $shouldNotifyRestart = $true

                    # Verify the change
                    $newValue = (Get-ItemProperty -Path $registryPath -Name $valueName).$valueName
                    Write-PolicyLog "Verified new value: $newValue" "SUCCESS"

                } else {
                    Write-PolicyLog "Registry value is already set to $currentRegValue (not 0), no change needed" "INFO"
                }
            } else {
                Write-PolicyLog "Registry value $valueName does not exist, creating with value 1..." "INFO"
                New-ItemProperty -Path $registryPath -Name $valueName -Value 1 -PropertyType DWord -Force | Out-Null
                Write-PolicyLog "Registry value created with value 1" "SUCCESS"
                $changesMade = $true
                $shouldNotifyRestart = $true
            }

            # If triggered by Windows Update and changes were made, provide additional context
            if ($TriggeredByWindowsUpdate -and $changesMade) {
                Write-PolicyLog "IMPORTANT: Bluetooth AVRCP setting was corrected after Windows Update" "SUCCESS"
                Write-PolicyLog "This prevents potential Bluetooth audio volume control issues" "INFO"
            }

            # Show restart notification if registry was changed
            if ($shouldNotifyRestart) {
                Write-PolicyLog "Registry changes made - showing restart notification to user" "INFO"

                # Determine notification context
                $contextMessage = if ($TriggeredByWindowsUpdate) {
                    "Windows Update has been detected and your Bluetooth audio settings have been automatically corrected. Please restart your computer when convenient to ensure optimal Bluetooth audio performance."
                } else {
                    "Your Bluetooth AVRCP registry settings have been updated to prevent audio volume control issues. Please restart your computer when convenient to ensure all changes take effect properly."
                }

                # Show notification to user
                $notificationResult = Show-RestartNotification -Message $contextMessage
                if ($notificationResult) {
                    Write-PolicyLog "User notification for restart displayed successfully" "SUCCESS"
                } else {
                    Write-PolicyLog "Failed to display user notification for restart" "WARNING"
                }
            }

            return $true

        } catch {
            Write-PolicyLog "ERROR: Failed to modify registry value: $($_.Exception.Message)" "ERROR"
            return $false
        }

    } catch {
        Write-PolicyLog "Bluetooth AVRCP management failed: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# Main execution
try {
    $policyExecutionReason = if ($TriggeredByWindowsUpdate) { "Windows Update Completion Event" } else { "Scheduled Execution" }
    Write-PolicyLog "=== Windows Update Check and Bluetooth AVRCP Registry Policy Started ===" "INFO"
    Write-PolicyLog "Execution Method: Task Scheduler" "INFO"
    Write-PolicyLog "Execution Reason: $policyExecutionReason" "INFO"
    Write-PolicyLog "Target System: Personal PC" "INFO"

    # Check if running as Administrator
    if (-not (Test-AdminPrivileges)) {
        Write-PolicyLog "ERROR: Script must be run with Administrator privileges" "ERROR"
        Write-PolicyLog "Please ensure the scheduled task is configured to run with highest privileges" "ERROR"
        exit 1
    }

    Write-PolicyLog "Administrator privileges confirmed" "SUCCESS"

    # If triggered by Windows Update, check for recent update activity
    if ($TriggeredByWindowsUpdate) {
        $recentActivity = Get-RecentWindowsUpdateActivity
        if ($recentActivity) {
            Write-PolicyLog "Confirmed: Recent Windows Update activity detected" "SUCCESS"
        } else {
            Write-PolicyLog "No recent Windows Update activity found, but proceeding with registry check" "INFO"
        }
    }

    # Step 1: Install and check Windows Updates (unless specifically triggered by update completion)
    if (-not $TriggeredByWindowsUpdate) {
        if (Install-WindowsUpdateModule) {
            Check-WindowsUpdates
        } else {
            Write-PolicyLog "Skipping Windows Update check due to module installation failure" "WARNING"
        }
    } else {
        Write-PolicyLog "Skipping proactive Windows Update check (triggered by update completion)" "INFO"
    }

    # Step 2: Always manage Bluetooth AVRCP registry (this is the primary purpose when triggered by updates)
    if (Manage-BluetoothAVRCP) {
        Write-PolicyLog "Bluetooth AVRCP registry management completed successfully" "SUCCESS"
    } else {
        Write-PolicyLog "Bluetooth AVRCP registry management failed" "ERROR"
    }

    Write-PolicyLog "=== Policy execution completed ===" "SUCCESS"

} catch {
    Write-PolicyLog "CRITICAL ERROR: Script execution failed: $($_.Exception.Message)" "ERROR"
    Write-PolicyLog "Error details: $($_.Exception.StackTrace)" "ERROR"
    exit 1
} finally {
    # Cleanup COM objects if created
    try {
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    } catch {
        # Ignore cleanup errors
    }
}

Write-PolicyLog "Policy script execution finished" "INFO"
