#=============================================================================================================================
#
# Script Name:     Check_EFI_FreeSpace.ps1
# Description:     Check free space on EFI partition
# Notes:           Needs to be run with admin privileges
# Version/Date:    0.0.1 21/07/2025
#
#=============================================================================================================================

try
{   
    $efiVolume = Get-Volume -UniqueId "\\?\Volume$(((Get-Partition).Where{$_.GptType -eq '{c12a7328-f81f-11d2-ba4b-00a0c93ec93b}'}.Guid))\"
    $freeSpace = [Math]::Round($efiVolume.SizeRemaining / 1MB, 2)
    $Space = [Math]::Round($efiVolume.Size / 1MB, 2)
    
    # Check if free space is less than 25 MB
    if ($freeSpace -lt 25) {
            Write-Host "ERROR: EFI volume $freeSpace MB of $space MB available"
            exit 1
        } else {
            Write-Host "PASS: EFI volume $freeSpace MB of $space MB available"
            exit 0
    }
}
catch 
{
    $errMsg = $_.Exception.Message
    Write-Error $errMsg
    exit 1
}
