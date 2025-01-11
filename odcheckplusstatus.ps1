$Hostname = $env:COMPUTERNAME
$OneDriveStatusFilePath = "$env:USERPROFILE\OneDriveStatus_$Hostname.txt"
$SharedFolderPath = "\\192.168.1.10\Users\oxbow\Desktop\FileShare"
$SyncDiagnosticsFilePath = Join-Path -Path "$env:LOCALAPPDATA\Microsoft\OneDrive\logs\Personal" -ChildPath "SyncDiagnostics.log"

$OneDriveStatus = if (Get-Process -Name OneDrive -ErrorAction SilentlyContinue) { "Running" } else { "Not Running" }
$SyncStatusMessage = "Sync diagnostics log file not found"
$FilesToUpload = $null
$FilesToDownload = $null

if (Test-Path $SyncDiagnosticsFilePath) {
   
    $LogContent = Get-Content -Path $SyncDiagnosticsFilePath -Raw
    $FilesToUpload = if ($LogContent -match "FilesToUpload\s*=\s*(\d+)") { [int]$Matches[1] } else { $null }
    $FilesToDownload = if ($LogContent -match "FilesToDownload\s*=\s*(\d+)") { [int]$Matches[1] } else { $null }

    if ($FilesToUpload -eq 0 -and $FilesToDownload -eq 0) {
        $SyncStatusMessage = "Up to Date"
    } elseif ($FilesToUpload -gt 0 -and $FilesToDownload -eq 0) {
        $SyncStatusMessage = "Upload Pending"
    } elseif ($FilesToDownload -gt 0 -and $FilesToUpload -eq 0) {
        $SyncStatusMessage = "Download Pending"
    } else {
        $SyncStatusMessage = "Syncing"
    }
}

$FreeSpace = "{0:N2} GB" -f ((Get-PSDrive -Name C).Free / 1GB)

$SSDHealthOutput = try {
    wmic diskdrive get model,mediaType,status /format:csv | ConvertFrom-Csv
} catch {
    @([PSCustomObject]@{
        Model = "N/A"
        MediaType = "N/A"
        Status = "Error retrieving SSD health: $_"
    })
}

$FormattedDiskDetails = ""
foreach ($disk in $SSDHealthOutput) {
    $FormattedDiskDetails += "Disk Model: $($disk.Model)`nDisk Type: $($disk.MediaType)`nDisk Status: $($disk.Status)`n---`n"
}

@"
Timestamp: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Hostname: $Hostname
OneDrive Status: $OneDriveStatus
Sync Status: $SyncStatusMessage
Files to Upload: $FilesToUpload
Files to Download: $FilesToDownload
Free Disk Space: $FreeSpace

Disk Health Status:
$FormattedDiskDetails
"@ | Set-Content -Path $OneDriveStatusFilePath

try {
    Copy-Item -Path $OneDriveStatusFilePath -Destination $SharedFolderPath -Force
    Write-Host "OneDrive status file successfully copied to $SharedFolderPath."
} catch {
    Write-Host "Failed to copy the OneDrive status file: $_"
}