# Simple PowerCLI Installer with File Browser
Add-Type -AssemblyName System.Windows.Forms

$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.Title = "Select PowerCLI Offline Bundle (ZIP file)"
$dialog.Filter = "ZIP files (*.zip)|*.zip"
$dialog.InitialDirectory = [Environment]::GetFolderPath("Downloads")

if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    $BundlePath = $dialog.FileName
    Write-Host "Installing from: $BundlePath" -ForegroundColor Green
    
    # Rest of installation logic here
    # ... (use the installation functions from the main script)
} else {
    Write-Host "No file selected." -ForegroundColor Yellow
}
