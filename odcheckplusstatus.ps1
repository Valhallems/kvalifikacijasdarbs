$Hostname = $env:COMPUTERNAME
$OneDriveStatusFilePath = "$env:USERPROFILE\OneDriveStatus_$Hostname.txt"
$SharedFolderPath = "\\192.168.1.10\Users\oxbow\Desktop\FileShare"

# OneDrive Status
$OneDriveStatus = if (Get-Process -Name OneDrive -ErrorAction SilentlyContinue) { "Running" } else { "Not Running" }
$SyncStatus = if (Test-Path "$env:USERPROFILE\OneDrive") { "Up to Date" } else { "OneDrive folder not found" }

$FreeSpace = "{0:N2} GB" -f ((Get-PSDrive -Name C).Free / 1GB)

# SSD Health Status using WMIC
$SSDHealthOutput = try {
    wmic diskdrive get model,mediaType,status /format:csv | ConvertFrom-Csv
} catch {
    @([PSCustomObject]@{
        Model = "N/A"
        MediaType = "N/A"
        Status = "Error retrieving SSD health: $_"
    })
}

# Format Disk Health Output
$FormattedDiskDetails = ""
foreach ($disk in $SSDHealthOutput) {
    $FormattedDiskDetails += "Disk Model: $($disk.Model)`nDisk Type: $($disk.MediaType)`nDisk Status: $($disk.Status)`n---`n"
}

# Write to File
@"
Timestamp: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Hostname: $Hostname
OneDrive Status: $OneDriveStatus
Sync Status: $SyncStatus
Free Disk Space: $FreeSpace

Disk Health Status:
$FormattedDiskDetails
"@ | Set-Content -Path $OneDriveStatusFilePath

# Copy to Shared Folder
try {
    Copy-Item -Path $OneDriveStatusFilePath -Destination $SharedFolderPath -Force
    Write-Host "OneDrive status file successfully copied to $SharedFolderPath."
} catch {
    Write-Host "Failed to copy the OneDrive status file: $_"
}
