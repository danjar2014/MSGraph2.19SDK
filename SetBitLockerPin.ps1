# Function to write to log file
function Write-Log {
    param (
        [string]$Message
    )

    $logFile = "C:\_Logfiles\MainSetPreBootPinBitlocker.log"
    $logDir = "C:\_Logfiles"

    try {
        # Create log directory if it doesn't exist
        if (-not (Test-Path -Path $logDir)) {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
        }

        # Create log entry with timestamp
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "[$timestamp] $Message"

        # Append to log file
        Add-Content -Path $logFile -Value $logEntry -ErrorAction Stop
    } catch {
        # If logging fails, write to console as fallback
        Write-Host "Failed to write to log file: $_"
        Write-Host $logEntry
    }
}

# Log the start of the script
Write-Log "Starting SetBitLockerPin.ps1"

# Check if running with elevated privileges
$isElevated = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isElevated) {
    Write-Log "Script is not running with elevated privileges. Exiting."
    exit 1
}
Write-Log "Script is running with elevated privileges."


Write-Log "Update Scheduled task $taskName with logon trigger"
$taskName = "SetBitLockerPinAtLogon"
try {

    $task = Get-ScheduledTask -TaskName $taskName
    $trigger = New-ScheduledTaskTrigger -AtLogOn
    $currenttriggers = @($task.Triggers)
    $NewTriggers = $currenttriggers + $trigger
    Set-ScheduledTask -TaskName $taskName -Trigger $NewTriggers
    Write-Log "Scheduled task $taskName updated successfully with logon trigger"

} catch {
    Write-Log "Failed to update scheduled task $taskName : $_"
    exit 1
}


# Define paths
$popupScriptPath = "C:\Temp\BitLockerPinSetup\Popup.ps1"
$serviceUIPath = "C:\Temp\BitLockerPinSetup\ServiceUI.exe"
$pathPINFile = $(Join-Path -Path "$env:SystemRoot\tracing" -ChildPath "168ba6df825678e4da1a.tmp")

# Check if ServiceUI.exe exists
if (-not (Test-Path -Path $serviceUIPath)) {
    Write-Log "ServiceUI.exe not found at $serviceUIPath. Exiting."
    exit 1
}
Write-Log "ServiceUI.exe found at $serviceUIPath."

# Check if Popup.ps1 exists
if (-not (Test-Path -Path $popupScriptPath)) {
    Write-Log "Popup.ps1 not found at $popupScriptPath. Exiting."
    exit 1
}
Write-Log "Popup.ps1 found at $popupScriptPath."

# Execute Popup.ps1 using ServiceUI.exe
Write-Log "Executing Popup.ps1 using ServiceUI.exe to prompt for BitLocker PIN."
try {
    $env:PinComplexityLevel = "High"  # Example: Set complexity level
    & $serviceUIPath -process:Explorer.exe "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -WindowStyle Hidden -Ex bypass -file "$popupScriptPath"
    $exitCode = $LASTEXITCODE
    Write-Log "Popup.ps1 executed with exit code: $exitCode"
} catch {
    Write-Log "Error executing Popup.ps1: $_"
    exit 1
}

# Check if the PIN file was created and the exit code is 0
if ($exitCode -eq 0 -and (Test-Path -Path $pathPINFile)) {
    Write-Log "PIN file found at $pathPINFile."
    
    # Read the encrypted PIN from the file
    try {
        $encodedText = Get-Content -Path $pathPINFile -Raw
        if ($encodedText.Length -gt 0) {
            Write-Log "PIN file contains data. Decrypting PIN."
            
            # Using DPAPI with a random generated shared 256-bit key to decrypt the PIN
            $key = (43,155,164,59,21,127,28,43,81,18,198,145,127,51,72,55,39,23,228,166,146,237,41,131,176,14,4,67,230,81,212,214)
            $secure = ConvertTo-SecureString $encodedText -Key $key
            $PIN = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure))
            Write-Log "PIN successfully decrypted from file."
            
            # Set the BitLocker PIN
            try {
                Write-Log "Setting BitLocker PIN for drive $env:SystemDrive."
                Add-BitLockerKeyProtector -MountPoint $env:SystemDrive -Pin $(ConvertTo-SecureString $PIN -AsPlainText -Force) -TpmAndPinProtector
                Write-Log "BitLocker PIN set successfully for drive $env:SystemDrive."
            } catch {
                Write-Log "Error setting BitLocker PIN: $_"
                exit 1
            }
        } else {
            Write-Log "PIN file is empty. Exiting."
            exit 1
        }
    } catch {
        Write-Log "Error decrypting PIN from file: $_"
        exit 1
    }
} else {
    Write-Log "Popup.ps1 failed (exit code: $exitCode) or PIN file not found at $pathPINFile. Exiting."
    exit 1
}

# Cleanup
try {
    Remove-Item -Path $pathPINFile -Force -ErrorAction SilentlyContinue
    Write-Log "PIN file deleted successfully."
} catch {
    Write-Log "Error deleting PIN file: $_"
}


# Run Intune sync
Start-Process -FilePath "C:\Program Files (x86)\Microsoft Intune Management Extension\Microsoft.Management.Services.IntuneWindowsAgent.exe" -ArgumentList "/sync" -NoNewWindow
Write-Log "Intune Sync ..."

# Delete the temporary folder
$tempFolder = "C:\Temp"
try {
    if (Test-Path -Path $tempFolder) {
        Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction Stop
        Write-Log "Temporary folder deleted: $tempFolder"
    } else {
        Write-Log "Temporary folder $tempFolder does not exist. Skipping deletion."
    }
} catch {
    Write-Log "Failed to delete temporary folder $tempFolder : $_"
}

# Delete the scheduled task
$taskName = "SetBitLockerPinAtLogon"
$taskPath = "\"
try {
    if (Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $taskName -TaskPath $taskPath -Confirm:$false -ErrorAction Stop
        Write-Log "Scheduled task $taskName deleted successfully."
    } else {
        Write-Log "Scheduled task $taskName does not exist. Skipping deletion."
    }
} catch {
    Write-Log "Failed to delete scheduled task $taskName : $_"
}

# Log completion
Write-Log "SetBitLockerPin.ps1 completed successfully."
