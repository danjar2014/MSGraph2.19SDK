$tempFolder = "C:\Temp\BitLockerPinSetup"
$requiredFiles = @(
    "SetBitLockerPin.ps1",
    "Popup.ps1",
    "ServiceUI.exe",
    "PIN-W11-BitLocker-0.png",
    "SetBitLockerPin.png",
    "Wallpaper.png"
)


$allFilesPresent = $true
foreach ($file in $requiredFiles) {
    $filePath = Join-Path -Path $tempFolder -ChildPath $file
    if (-not (Test-Path -Path $filePath)) {
        $allFilesPresent = $false
        Write-Output "File $file not found in $tempFolder"
        break
    }
}


$taskName = "SetBitLockerPinAtLogon"
$taskPath = "\"
$taskExists = Get-ScheduledTask -TaskName $taskName -TaskPath $taskPath -ErrorAction SilentlyContinue

$keyProtector = (Get-BitLockerVolume -MountPoint $env:SystemDrive).KeyProtector | Where-Object { $_.KeyProtectorType -eq 'TpmPin' }


if (($allFilesPresent -and $taskExists) -or $keyProtector) {
    Write-Output "Files copied and scheduled task created successfully or TPMPIN is Set successfully"
    exit 0
} else {
    Write-Output "Files, scheduled task or TPMPIN is missing"
    exit 1
}
