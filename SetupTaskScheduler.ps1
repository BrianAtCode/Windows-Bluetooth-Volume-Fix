# Windows Bluetooth Volume Fix Setup 
# Windows Update Check + Bluetooth AVRCP Registry Management
# Created: October 2025

param(
    [string]$ScriptPath = "C:\Windows\Scripts\WindowsPolicyScript.ps1",
    [string]$TaskName = "Win10-BluetoothPolicy-TaskScheduler",
    [string]$Description = "Windows Bluetooth Volume Fix: Monitors Windows Updates and manages Bluetooth AVRCP registry settings with update completion trigger"
)

# Function to check admin privileges
function Test-AdminPrivileges {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to get PowerShell executable path
function Get-PowerShellPath {
    # Check if PowerShell 7+ is available
    $pwshPath = Get-Command pwsh.exe -ErrorAction SilentlyContinue
	
    if ($pwshPath -and $PSVersionTable.PSVersion.Major -ge 7) {
        Write-Host "PowerShell 7+ detected, using pwsh.exe" -ForegroundColor Cyan
        return "pwsh.exe"
    }
    else {
		Write-Host "executable file version is $($pwshPath.Version.Major) and the current ps version major is $($PSVersionTable.PSVersion.Major)."
        Write-Host "Using Windows PowerShell (powershell.exe)" -ForegroundColor Cyan
        return "powershell.exe"
    }
}

# Function to create the script directory and copy files
function Setup-ScriptEnvironment {
    Write-Host "Setting up script environment..." -ForegroundColor Yellow

    # Create script directory
    $scriptDir = Split-Path $ScriptPath -Parent
    if (-not (Test-Path $scriptDir)) {
        Write-Host "Creating script directory: $scriptDir" -ForegroundColor Green
        New-Item -ItemType Directory -Path $scriptDir -Force | Out-Null
    }

    # Check if the main script file exists in current directory
    $currentDir = Get-Location
    $sourceScript = Join-Path $currentDir "WindowsPolicyScript.ps1"
	write-host $currentDir.Path
	write-host $sourceScript
	
    if (Test-Path $sourceScript) {
        Write-Host "Copying script to: $ScriptPath" -ForegroundColor Green
        Copy-Item $sourceScript $ScriptPath -Force
    } else {
        Write-Host "WARNING: Main script file not found in current directory" -ForegroundColor Red
        Write-Host "Please ensure WindowsPolicyScript.ps1 is in the same directory as this setup script" -ForegroundColor Red
        return $false
    }

    # Create logs directory
    $logsDir = "C:\Windows\Logs"
    if (-not (Test-Path $logsDir)) {
        Write-Host "Creating logs directory: $logsDir" -ForegroundColor Green
        New-Item -ItemType Directory -Path $logsDir -Force | Out-Null
    }

    return $true
}

# Function to remove existing tasks if they exist
function Remove-ExistingTasks {
    param([string]$BaseName)

    $tasksToRemove = @(
        $BaseName,
        "$BaseName-ManualTest",
        "$BaseName-WindowsUpdateTrigger"
    )

    foreach ($taskName in $tasksToRemove) {
        try {
            $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
            if ($existingTask) {
                Write-Host "Removing existing task: $taskName" -ForegroundColor Yellow
                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
                Write-Host "Task removed: $taskName" -ForegroundColor Green
            }
        } catch {
            Write-Host "Error removing task $taskName : $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

# Function to create the main scheduled task - PowerShell 7.x compatible
function Create-MainPolicyTask {
    param(
        [string]$Name,
        [string]$Path,
        [string]$Desc
    )

    Write-Host "Creating main scheduled task: $Name" -ForegroundColor Yellow

    try {
        # Get appropriate PowerShell executable
        $psExePath = Get-PowerShellPath

        # Define the action - PowerShell script execution with proper parameters
        $actionArgs = @(
            "-ExecutionPolicy", "Bypass",
            "-NonInteractive", 
            "-WindowStyle", "Hidden",
            "-File", "`"$Path`""
        )
        $action = New-ScheduledTaskAction -Execute $psExePath -Argument ($actionArgs -join " ")

        # Define multiple triggers for comprehensive coverage
        Write-Host "Setting up triggers..." -ForegroundColor Cyan

        # 1. At startup (delayed by 2 minutes to ensure system is ready)
        $startupTrigger = New-ScheduledTaskTrigger -AtStartup

        # 2. Daily at 3:00 AM (maintenance window)
        $dailyTrigger = New-ScheduledTaskTrigger -Daily -At "03:00"

        # 3. Weekly on Sunday at 1:00 AM (weekly maintenance)
        $weeklyTrigger = New-ScheduledTaskTrigger -Weekly -WeeksInterval 1 -DaysOfWeek Sunday -At "01:00"

        # 4. On user logon (with delay)
        $logonTrigger = New-ScheduledTaskTrigger -AtLogOn

        # Apply delays using TimeSpan objects
        $startupTrigger.Delay = "PT2M"
        $logonTrigger.Delay = "PT1M"

        # Combine all triggers
        $triggers = @($startupTrigger, $dailyTrigger, $weeklyTrigger, $logonTrigger)

        # Define the principal - run as SYSTEM with highest privileges
        Write-Host "Configuring security context..." -ForegroundColor Cyan
        $principal = New-ScheduledTaskPrincipal -UserId "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest

        # Define settings for optimal execution
        # CRITICAL: Use TimeSpan objects for ALL duration parameters in PowerShell 7+
        Write-Host "Configuring task settings..." -ForegroundColor Cyan

        $settingsParams = @{
            AllowStartIfOnBatteries = $true
            DontStopIfGoingOnBatteries = $true
            StartWhenAvailable = $true
            MultipleInstances = 'IgnoreNew'
            ExecutionTimeLimit = (New-TimeSpan -Minutes 10)  # Use TimeSpan object, not "PT10M"
            RestartCount = 3
            RestartInterval = (New-TimeSpan -Minutes 5)  # Use TimeSpan object, not "PT5M"
        }

        $settings = New-ScheduledTaskSettingsSet @settingsParams

        # Allow task to run on demand
        $settings.AllowDemandStart = $true

        # Create the scheduled task object
        $task = New-ScheduledTask -Action $action -Trigger $triggers -Principal $principal -Settings $settings -Description $Desc

        # Register the task
        Write-Host "Registering main scheduled task..." -ForegroundColor Cyan
        Register-ScheduledTask -TaskName $Name -InputObject $task -Force | Out-Null

        Write-Host "Main scheduled task '$Name' created successfully!" -ForegroundColor Green
        return $true

    } catch {
        Write-Host "Error creating main scheduled task: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Error details: $($_.Exception.GetType().FullName)" -ForegroundColor Red
        Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
        return $false
    }
}

# Function to create Windows Update completion trigger task using schtasks (more reliable)
function Create-WindowsUpdateTriggerTask {
    param(
        [string]$BaseName,
        [string]$Path
    )

    $updateTaskName = "$BaseName-WindowsUpdateTrigger"
    Write-Host "Creating Windows Update completion trigger task: $updateTaskName" -ForegroundColor Yellow

    try {
        # Get appropriate PowerShell executable
        $psExePath = Get-PowerShellPath

        # Determine full path to PowerShell executable
        $psFullPath = (Get-Command $psExePath).Source
        if (-not $psFullPath) {
            $psFullPath = $psExePath
        }

        # Use schtasks for better event trigger compatibility
        Write-Host "Using schtasks method for event trigger..." -ForegroundColor Cyan

        # Create XML content for the task with proper escaping
        # Note: Using PT2M (ISO 8601) in XML is fine for schtasks, it only fails in PowerShell cmdlets
        $taskXML = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Triggered by Windows Update completion events to fix Bluetooth AVRCP registry</Description>
  </RegistrationInfo>
  <Triggers>
    <EventTrigger>
      <Enabled>true</Enabled>
      <Subscription>&lt;QueryList&gt;&lt;Query Id="0" Path="System"&gt;&lt;Select Path="System"&gt;*[System[Provider[@Name='Microsoft-Windows-WindowsUpdateClient'] and (EventID=19)]]&lt;/Select&gt;&lt;/Query&gt;&lt;/QueryList&gt;</Subscription>
      <Delay>PT2M</Delay>
    </EventTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>S-1-5-18</UserId>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <DisallowStartOnRemoteAppSession>false</DisallowStartOnRemoteAppSession>
    <UseUnifiedSchedulingEngine>true</UseUnifiedSchedulingEngine>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT5M</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>$psFullPath</Command>
      <Arguments>-ExecutionPolicy Bypass -NonInteractive -WindowStyle Hidden -File "$Path" -TriggeredByWindowsUpdate</Arguments>
    </Exec>
  </Actions>
</Task>
"@

        # Save XML to temp file
        $tempXMLPath = [System.IO.Path]::GetTempFileName() + ".xml"
        $taskXML | Out-File -FilePath $tempXMLPath -Encoding Unicode

        # Create task using schtasks
        Write-Host "Registering Windows Update trigger task..." -ForegroundColor Cyan
        $result = & schtasks /create /tn $updateTaskName /xml $tempXMLPath /f 2>&1

        # Clean up temp file
        Remove-Item $tempXMLPath -Force -ErrorAction SilentlyContinue

        if ($LASTEXITCODE -eq 0) {
            Write-Host "Windows Update trigger task '$updateTaskName' created successfully!" -ForegroundColor Green
            Write-Host "This task will run automatically after Windows Updates complete" -ForegroundColor Cyan
            return $true
        } else {
            Write-Host "schtasks failed with exit code: $LASTEXITCODE" -ForegroundColor Red
            Write-Host "Output: $result" -ForegroundColor Red
            return $false
        }

    } catch {
        Write-Host "Error creating Windows Update trigger task: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to create a manual test task
function Create-ManualTestTask {
    param([string]$MainTaskName)

    $testTaskName = "$MainTaskName-ManualTest"
    Write-Host "Creating manual test task: $testTaskName" -ForegroundColor Yellow

    try {
        # Get appropriate PowerShell executable
        $psExePath = Get-PowerShellPath

        # Same action as main task but with visible window for testing
        $actionArgs = @(
            "-ExecutionPolicy", "Bypass",
            "-NoExit",
            "-File", "`"$ScriptPath`""
        )
        $action = New-ScheduledTaskAction -Execute $psExePath -Argument ($actionArgs -join " ")

        # No automatic triggers - manual run only
        $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddYears(10)  # Far future date

        # Run as current user for testing
        $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
        $principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Highest

        # Basic settings with TimeSpan objects
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

        $task = New-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description "Manual test version of the Windows 10 Bluetooth Policy automation script"

        Register-ScheduledTask -TaskName $testTaskName -InputObject $task -Force | Out-Null

        Write-Host "Manual test task '$testTaskName' created successfully!" -ForegroundColor Green
        return $true

    } catch {
        Write-Host "Error creating manual test task: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Function to display task information
function Show-TaskInfo {
    param([string]$BaseName)

    $tasks = @(
        $BaseName,
        "$BaseName-WindowsUpdateTrigger",
        "$BaseName-ManualTest"
    )

    Write-Host "`n=== Task Information ===" -ForegroundColor Green

    foreach ($taskName in $tasks) {
        try {
            $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
            if ($task) {
                Write-Host "`nTask Name: $($task.TaskName)" -ForegroundColor White
                Write-Host "State: $($task.State)" -ForegroundColor White
                Write-Host "Description: $($task.Description)" -ForegroundColor White

                Write-Host "Triggers:" -ForegroundColor Cyan
                foreach ($trigger in $task.Triggers) {
                    $triggerType = $trigger.CimClass.CimClassName -replace "MSFT_TaskTrigger", ""
                    if ($triggerType -eq "Event") {
                        Write-Host "  - Event Trigger (Windows Update Completion)" -ForegroundColor White
                    } else {
                        Write-Host "  - $triggerType" -ForegroundColor White
                    }
                }
            }
        } catch {
            Write-Host "Could not retrieve info for task: $taskName" -ForegroundColor Red
        }
    }
}

# Main execution
Write-Host "`n=== Windows Bluetooth Volume Fix Setup ===" -ForegroundColor Magenta
Write-Host "Auto‑fix via Task Scheduler (Win10)" -ForegroundColor Magenta
Write-Host "Task Name: $TaskName" -ForegroundColor White
Write-Host "Script Path: $ScriptPath`n" -ForegroundColor White

# Display PowerShell version information
Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor Cyan
Write-Host "PowerShell Edition: $($PSVersionTable.PSEdition)" -ForegroundColor Cyan

# Check admin privileges
if (-not (Test-AdminPrivileges)) {
    Write-Host "`nERROR: This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "Please right-click on PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host "`nAdministrator privileges confirmed" -ForegroundColor Green

# Setup environment
if (-not (Setup-ScriptEnvironment)) {
    Write-Host "`nScript environment setup failed. Please check the errors above." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Remove existing tasks if present
Remove-ExistingTasks -BaseName $TaskName

# Create the main scheduled task
if (Create-MainPolicyTask -Name $TaskName -Path $ScriptPath -Desc $Description) {

    # Create Windows Update completion trigger task
    $updateTriggerSuccess = Create-WindowsUpdateTriggerTask -BaseName $TaskName -Path $ScriptPath

    # Create manual test task
    Create-ManualTestTask -MainTaskName $TaskName

    # Show task information
    Show-TaskInfo -BaseName $TaskName

    Write-Host "`n=== Setup Complete! ===" -ForegroundColor Green
    Write-Host "Your Windows 10 Bluetooth Policy automation is now active!" -ForegroundColor White
    Write-Host "`nThe policy will run:" -ForegroundColor Cyan
    Write-Host "  • At system startup (2 minute delay)" -ForegroundColor White
    Write-Host "  • Daily at 3:00 AM" -ForegroundColor White
    Write-Host "  • Weekly on Sundays at 1:00 AM" -ForegroundColor White
    Write-Host "  • At user logon (1 minute delay)" -ForegroundColor White
    if ($updateTriggerSuccess) {
        Write-Host "  • After Windows Update completion (2 minute delay)" -ForegroundColor Yellow
    } else {
        Write-Host "  • Windows Update trigger setup failed - will rely on scheduled runs" -ForegroundColor Red
    }

    Write-Host "`nTo test the script:" -ForegroundColor Yellow
    Write-Host "  1. Open Task Scheduler (taskschd.msc)" -ForegroundColor White
    Write-Host "  2. Find '$TaskName-ManualTest'" -ForegroundColor White
    Write-Host "  3. Right-click and select 'Run'" -ForegroundColor White

    Write-Host "`nLogs will be saved to:" -ForegroundColor Yellow
    Write-Host "  C:\Windows\Logs\BluetoothPolicyAutomation.log" -ForegroundColor White

} else {
    Write-Host "`nTask creation failed. Please check the errors above." -ForegroundColor Red
}

Write-Host "`nPress Enter to exit..." -ForegroundColor Gray
Read-Host
