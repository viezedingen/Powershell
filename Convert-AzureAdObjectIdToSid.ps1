
function Convert-AzureAdObjectIdToSid {
    param (
        [String] $ObjectId
    )

    $bytes = [Guid]::Parse($ObjectId).ToByteArray()
    $array = New-Object 'UInt32[]' 4
    [Buffer]::BlockCopy($bytes, 0, $array, 0, 16)
    $sid = "S-1-12-1-$($array -join '-')"
    return $sid
}

# Example usage:
$objectId = "d267901d-90af-4433-97f7-96a1a587f39c"
$sid = Convert-AzureAdObjectIdToSid -ObjectId $objectId
Write-Output $sid
