$keyProtector = (Get-BitLockerVolume -MountPoint $env:SystemDrive).KeyProtector | Where-Object { $_.KeyProtectorType -eq 'TpmPin' }

if ($null -ne $keyProtector) {
    Write-Output "TPM+PIN protector already configured"
    exit 1
} else {
    Write-Output "No TPM+PIN protector found"
    exit 0
}
