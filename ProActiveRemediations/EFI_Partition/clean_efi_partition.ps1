# WARNING: This script modifies the EFI partition. Use only if you understand the implications.

# Get the EFI partition and assign a drive letter
$efiPartition = Get-Partition | Where-Object { $_.GptType -eq "{C12A7328-F81F-11D2-BA4B-00A0C93EC93B}" }
if ($efiPartition) {
    $driveLetter = "Z:"
    $null = mountvol $driveLetter "$($efiPartition.AccessPaths[0])"
} else {
    Write-Error "EFI partition not found."
    exit 1
}

# Clean up unnecessary files (example: HP files or old boot entries)
$unwantedPaths = @(
    "$driveLetter\EFI\HP",
    "$driveLetter\EFI\Boot\old",
    "$driveLetter\EFI\Microsoft\Boot\BCD_Backup"
)

foreach ($path in $unwantedPaths) {
    if (Test-Path $path) {
        Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
        Write-Output "Removed: $path"
    }
}

# Clean up unneeded font files except wgl4_boot.ttf
$fontFiles = Get-ChildItem "$driveLetter\EFI\Microsoft\Boot\Fonts" -Filter *.ttf
$excludedFonts = @("wgl4_boot.ttf")

foreach ($file in $fontFiles) {
    if ($excludedFonts -notcontains $file.Name) {
        Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
        Write-Output "Removed font file: $file.Name"
    }
}

# Optional: Unmount the EFI partition
mountvol $driveLetter /D
Write-Output "EFI partition cleanup complete."
exit 0