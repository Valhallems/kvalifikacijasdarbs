$Hostname = $env:COMPUTERNAME
$OneDriveStatusFilePath = "$env:USERPROFILE\OneDriveStatus_$Hostname.txt"
$SharedFolderPath = "\\192.168.1.10\Users\oxbow\Desktop\FileShare"

$OneDriveStatus = if (Get-Process -Name OneDrive -ErrorAction SilentlyContinue) { "Running" } else { "Not Running" }

$SyncDiagnosticsFilePath = Join-Path -Path "$env:LOCALAPPDATA\Microsoft\OneDrive\logs\Personal" -ChildPath "SyncDiagnostics.log"
$SyncStatus = "Log file not found"
$FilesToUpload = $null
$FilesToDownload = $null

if (Test-Path $SyncDiagnosticsFilePath) {
    $LogContent = Get-Content -Path $SyncDiagnosticsFilePath -Raw
    $FilesToUpload = if ($LogContent -match "FilesToUpload\s*=\s*(\d+)") { [int]$Matches[1] } else { $null }
    $FilesToDownload = if ($LogContent -match "FilesToDownload\s*=\s*(\d+)") { [int]$Matches[1] } else { $null }

    if ($FilesToUpload -eq 0 -and $FilesToDownload -eq 0) {
        $SyncStatus = "Up to Date"
    } elseif ($FilesToUpload -gt 0 -and $FilesToDownload -eq 0) {
        $SyncStatus = "Upload Pending"
    } elseif ($FilesToDownload -gt 0 -and $FilesToUpload -eq 0) {
        $SyncStatus = "Download Pending"
    } else {
        $SyncStatus = "Syncing"
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
Sync Status: $SyncStatus
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

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$fileShareFolder = $SharedFolderPath
$csvFile = "$env:USERPROFILE\Desktop\OneDriveStatusHistory.csv"

if (-not (Test-Path $csvFile)) {
    "Date,Hostname,OneDrive Status,Sync Status,Free Space,Disk Model,Disk Type,Disk Status" | Out-File -FilePath $csvFile -Encoding UTF8
}

$form = New-Object System.Windows.Forms.Form -Property @{ Text = "OneDrive and Disk Status Monitor"; Size = New-Object System.Drawing.Size(800, 600) }

$dataGridView = New-Object System.Windows.Forms.DataGridView -Property @{ 
    Size = New-Object System.Drawing.Size(750, 400); Location = New-Object System.Drawing.Point(20, 20); ReadOnly = $true; ColumnCount = 7 }
$dataGridView.Columns[0].Name, $dataGridView.Columns[1].Name, $dataGridView.Columns[2].Name, $dataGridView.Columns[3].Name, $dataGridView.Columns[4].Name, $dataGridView.Columns[5].Name, $dataGridView.Columns[6].Name = "Hostname", "OneDrive Status", "Sync Status", "Free Space", "Disk Model", "Disk Type", "Disk Status"
$form.Controls.Add($dataGridView)

$lastRefreshLabel = New-Object System.Windows.Forms.Label -Property @{ Location = New-Object System.Drawing.Point(20, 440) }
$form.Controls.Add($lastRefreshLabel)

$refreshButton = New-Object System.Windows.Forms.Button -Property @{ Text = "Refresh"; Location = New-Object System.Drawing.Point(20, 480) }
$form.Controls.Add($refreshButton)

$exportButton = New-Object System.Windows.Forms.Button -Property @{ Text = "Export Status to CSV"; Location = New-Object System.Drawing.Point(120, 480) }
$form.Controls.Add($exportButton)

function Refresh-Status {
    $dataGridView.Rows.Clear()
    $textFiles = Get-ChildItem -Path $fileShareFolder -Filter "*.txt"

    foreach ($file in $textFiles) {
        $content = Get-Content -Path $file.FullName
        $parsedData = @{ Hostname = ""; Status = ""; SyncStatus = ""; FreeSpace = ""; DiskModel = "None"; DiskType = "None"; DiskStatus = "None" }

        foreach ($line in $content) {
            if ($line -match "^Hostname: (.+)$") { $parsedData.Hostname = $matches[1] }
            if ($line -match "^OneDrive Status: (.+)$") { $parsedData.Status = $matches[1] }
            if ($line -match "^Sync Status: (.+)$") { $parsedData.SyncStatus = $matches[1] }
            if ($line -match "^Free Disk Space: (.+)$") { $parsedData.FreeSpace = $matches[1] }
            if ($line -match "^Disk Model: (.+)$") { $parsedData.DiskModel = $matches[1] }
            if ($line -match "^Disk Type: (.+)$") { $parsedData.DiskType = $matches[1] }
            if ($line -match "^Disk Status: (.+)$") { $parsedData.DiskStatus = $matches[1] }
        }

        $row = $dataGridView.Rows.Add($parsedData.Hostname, $parsedData.Status, $parsedData.SyncStatus, $parsedData.FreeSpace, $parsedData.DiskModel, $parsedData.DiskType, $parsedData.DiskStatus)

        $dataGridView.Rows[$row].Cells[1].Style.BackColor = if ($parsedData.Status -eq "Not Running") { [System.Drawing.Color]::Red } else { [System.Drawing.Color]::LightGreen }
        $dataGridView.Rows[$row].Cells[6].Style.BackColor = if ($parsedData.DiskStatus -eq "OK") { [System.Drawing.Color]::LightGreen } else { [System.Drawing.Color]::Red }

        "$((Get-Date).ToString("yyyy-MM-dd HH:mm:ss")),$($parsedData.Hostname),$($parsedData.Status),$($parsedData.SyncStatus),$($parsedData.FreeSpace),$($parsedData.DiskModel),$($parsedData.DiskType),$($parsedData.DiskStatus)" | Out-File -FilePath $csvFile -Append -Encoding UTF8
    }

    $lastRefreshLabel.Text = "Last Refresh: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
}

$exportButton.Add_Click({
    $currentViewFile = "$env:USERPROFILE\Desktop\CurrentOneDriveStatus.csv"
    $data = $dataGridView.Rows | Where-Object { -not $_.IsNewRow } | ForEach-Object {
        [PSCustomObject]@{
            Hostname       = $_.Cells[0].Value
            OneDriveStatus = $_.Cells[1].Value
            SyncStatus     = $_.Cells[2].Value
            FreeSpace      = $_.Cells[3].Value
            DiskModel      = $_.Cells[4].Value
            DiskType       = $_.Cells[5].Value
            DiskStatus     = $_.Cells[6].Value
        }
    }
    $data | Export-Csv -Path $currentViewFile -NoTypeInformation
    [System.Windows.Forms.MessageBox]::Show("Status exported to $currentViewFile")
})

$refreshButton.Add_Click({ Refresh-Status })

Refresh-Status
$form.ShowDialog()