$Path = "registry::HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters"
$Name1 = "DisabledComponents"
$Value1 = "0x20"

Try {
    $Registry1 = (Get-ItemProperty -Path $Path -Name $Name1 -ErrorAction Stop | Select-Object -ExpandProperty $Name1)
    If ($Registry1 -eq $Value1){
        Write-Output "IPv4 is preferred over IPv6"
        Exit 0
    } 
    Write-Output "Regkey is set to $Registry1, which is not the expected value of $Value1"
    Exit 1
} 
Catch {
    Write-Warning "Registry key not found or error accessing it."
    Exit 1
}