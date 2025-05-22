function Write-Log {
    param (
        [string]$Message
    )

    $logFile = "C:\_Logfiles\CopyBitLockerFilesDuringAUTOPILOT.log"
    $logDir = "C:\_Logfiles"

    try {
        if (-not (Test-Path -Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }

        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "[$timestamp] $Message"
        Add-Content -Path $logFile -Value $logEntry -ErrorAction Stop
    } catch {
        Write-Host "Failed to write to log file: $_"
        Write-Host $logEntry
    }
}


Write-Log "Starting CopyBitLockerFilesDuringAUTOPILOT.ps1"


$tempFolder = "C:\Temp\BitLockerPinSetup"
Write-Log "Creating temporary folder: $tempFolder"

try {
    if (-not (Test-Path -Path $tempFolder)) {
        New-Item -ItemType Directory -Path $tempFolder -Force | Out-Null
    }
} catch {
    Write-Log "Failed to create temporary folder: $_"
    exit 1
}


$sourceFiles = @(
    "SetBitLockerPin.ps1",
    "Popup.ps1",
    "ServiceUI.exe",
    "PIN-W11-BitLocker-0.png",
    "SetBitLockerPin.png",
    "Wallpaper.png"
)

foreach ($file in $sourceFiles) {
    $sourcePath = Join-Path -Path $PSScriptRoot -ChildPath $file
    $destPath = Join-Path -Path $tempFolder -ChildPath $file

    try {
        if (Test-Path -Path $sourcePath) {
            Copy-Item -Path $sourcePath -Destination $destPath -Force -ErrorAction Stop
            Write-Log "Copied $file to $destPath"
        } else {
            Write-Log "Source file $file not found in $PSScriptRoot"
            exit 1
        }
    } catch {
        Write-Log "Failed to copy $file to $destPath : $_"
        exit 1
    }
}


$taskName = "SetBitLockerPinAtLogon"
$taskDescription = "Execute SetBitLockerPin.ps1 when Windows Hello for Business enrollment completes (Event ID 300)"
$taskPath = "\"  

$eventQuery = @"
<QueryList>
  <Query Id="0" Path="Microsoft-Windows-User Device Registration/Admin">
    <Select Path="Microsoft-Windows-User Device Registration/Admin">*[System[(EventID=300)]]</Select>
  </Query>
</QueryList>
"@

try {
    
    $taskService = New-Object -ComObject Schedule.Service
    $taskService.Connect()

    
    $rootFolder = $taskService.GetFolder($taskPath)
    try {
        $rootFolder.DeleteTask($taskName, 0)
        Write-Log "Existing task $taskName deleted"
    } catch {
        Write-Log "No existing task $taskName found, proceeding with creation"
    }

    
    $taskDefinition = $taskService.NewTask(0)
    $taskDefinition.RegistrationInfo.Description = $taskDescription
    $taskDefinition.Principal.UserId = "NT AUTHORITY\SYSTEM"
    $taskDefinition.Principal.LogonType = 3  # TASK_LOGON_SERVICE_ACCOUNT
    $taskDefinition.Principal.RunLevel = 1   # TASK_RUNLEVEL_HIGHEST
    $taskDefinition.Settings.Enabled = $true
    $taskDefinition.Settings.AllowDemandStart = $true
    $taskDefinition.Settings.StartWhenAvailable = $true
    $taskDefinition.Settings.RunOnlyIfNetworkAvailable = $false
    $taskDefinition.Settings.StopIfGoingOnBatteries = $false
    $taskDefinition.Settings.DisallowStartIfOnBatteries = $false

    
    $trigger = $taskDefinition.Triggers.Create(0)  # 0 = TASK_TRIGGER_EVENT
    $trigger.Id = "EventTriggerId"
    $trigger.Subscription = $eventQuery
    $trigger.Enabled = $true

    
    $action = $taskDefinition.Actions.Create(0)  # 0 = TASK_ACTION_EXEC
    $action.Path = "powershell.exe"
    $action.Arguments = "-ExecutionPolicy Bypass -File `"$tempFolder\SetBitLockerPin.ps1`""

    
    $rootFolder.RegisterTaskDefinition($taskName, $taskDefinition, 6, $null, $null, 3)  # 6 = TASK_CREATE_OR_UPDATE, 3 = TASK_LOGON_SERVICE_ACCOUNT
    Write-Log "Scheduled task $taskName created successfully at root with event trigger (Event ID 300)"
} catch {
    Write-Log "Failed to create scheduled task $taskName : $_"
    exit 1
}


Write-Log "CopyBitLockerFilesDuringAUTOPILOT.ps1 completed successfully"
exit 0
