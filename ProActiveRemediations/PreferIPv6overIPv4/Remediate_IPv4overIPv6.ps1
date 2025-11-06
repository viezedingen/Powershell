$registryPath = 'HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters'
$regKey = "DisabledComponents"
$Value = "0x20"

try {
    IF(!(Test-Path $registryPath)){
    New-Item -Path $registryPath -Force | Out-Null
    New-ItemProperty -Path $registryPath -name $regkey -PropertyType DWord -Value $Value -Force | Out-Null
    Write-Output "Registry key created at $registryPath with value $Value"
    } ELSE {
    New-ItemProperty -Path $registryPath -name $regkey -PropertyType DWord -Value $Value -Force | Out-Null
    Write-Output "Registry key $regkey updated at $registryPath with value $Value"
    }
    Exit 0
}
catch {
    $errMsg = $_.Exception.Message
    Write-Error $errMsg
    exit 1
}

