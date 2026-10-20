#=============================================================================================================================
#
# Script Name:     Remove_Software_0.0.1.ps1
# Description:     Remove_Software_0.0.1.ps1
# Notes:          
# Version/Date:    0.0.1 06/09/2021
#
#=============================================================================================================================

# Define Variables
$SoftwareNames = @('Microsoft Silverlight')
$UninstallKeys = Get-ChildItem -path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\"
$UninstallKeys64 = Get-ChildItem -path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
$SoftwareGuid = 0
$SoftwareGuid64 = 0
$exitcode = 1337

Write-Host "Software to delete:"
ForEach ($SoftwareName in $SoftwareNames) {
    Write-Host $SoftwareName
} 

Function Test-RegistryValue($regkey, $name) {
    $exists = Get-ItemProperty -Path "$regkey" -Name "$name" -ErrorAction SilentlyContinue
    If (($null -ne $exists) -and ($exists.Length -ne 0)) {
        Return $true
    }
    Return $false
}

#Check for Software in Registry
ForEach ($SoftwareName in $SoftwareNames) {
    Write-Host "Starting the search for: $SoftwareName"
        ForEach ( $UninstallKey in $UninstallKeys )
        {
            if(Test-RegistryValue "Registry::$UninstallKey" "DisplayName"){
                if ((Get-ItemPropertyValue -Path "Registry::$UninstallKey" -Name "DisplayName") -contains $SoftwareName) {
        
                    Write-Host "SoftwareGuid: $($UninstallKey.PSChildName)"
                    $SoftwareGuid = $UninstallKey.PSChildName
                }
            }
        }
        ForEach ( $UninstallKey64 in $UninstallKeys64 )
        {
            if(Test-RegistryValue "Registry::$UninstallKey64" "DisplayName"){
                if ((Get-ItemPropertyValue -Path "Registry::$UninstallKey64" -Name "DisplayName") -contains $SoftwareName) {
        
                    Write-Host "SoftwareGuid: $($UninstallKey64.PSChildName)"
                    $SoftwareGuid64 = $UninstallKey64.PSChildName
                }
            
            }
        }
        #Uninstall the correct way
        
        if($SoftwareGuid){
            $SoftwareMSIUninstallArguments = @(
            "/x"
            "$SoftwareGuid"
            "/qn"
            "/norestart"
            "/l*v"
            '"C:\Windows\Temp\software32_uninstall_script.log"'
            )
            Write-Host "Found 32bit: $SoftwareName -> Uninstalling"
            $exitcode = (Start-Process "msiexec.exe" -ArgumentList $SoftwareMSIUninstallArguments -Wait -NoNewWindow -Passthru).ExitCode
        }
        
        if($SoftwareGuid64){
            $Software64MSIUninstallArguments = @(
            "/x"
            "$SoftwareGuid64"
            "/qn"
            "/norestart"
            "/l*v"
            '"C:\Windows\Temp\software64_uninstall_script.log"'
            )
            Write-Host "Found 64bit: $SoftwareName -> Uninstalling"
            $exitcode = (Start-Process "msiexec.exe" -ArgumentList $Software64MSIUninstallArguments -Wait -NoNewWindow -Passthru).ExitCode
        }
        
        Write-Host "Removed $SoftwareName if found"
        
        $SoftwareGuid = 0
        $SoftwareGuid64 = 0
}

#intune sync
Write-Host "Starting Intune Sync"
Get-ScheduledTask | Where-Object {$_.TaskName -eq 'Schedule to run OMADMClient by client'} | Start-ScheduledTask

#done
Write-Host "All done"
exit $exitcode