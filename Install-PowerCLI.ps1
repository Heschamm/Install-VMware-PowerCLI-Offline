# PowerCLI Offline Installer with File Browser
# Requires PowerShell 

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Show-FileBrowser {
    param([string]$InitialDirectory = "$env:USERPROFILE\Downloads")
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = $InitialDirectory
    $OpenFileDialog.Filter = "ZIP files (*.zip)|*.zip|All files (*.*)|*.*"
    $OpenFileDialog.FilterIndex = 1
    $OpenFileDialog.Title = "Select PowerCLI Offline Bundle"
    $OpenFileDialog.ShowHelp = $true
    
    $result = $OpenFileDialog.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $OpenFileDialog.FileName
    } else {
        return $null
    }
}

function Install-PowerCLIFromBundle {
    param([string]$BundlePath)
    
    Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor Cyan
    Write-Host "Installing PowerCLI from selected bundle..." -ForegroundColor Green

    # Define paths
    $ExtractPath = "$env:TEMP\PowerCLI_13.3.0_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    $RepositoryName = "LocalPowerCLI_$(Get-Random)"

    try {
        # Check if bundle exists
        if (-not (Test-Path $BundlePath)) {
            Write-Error "Selected PowerCLI bundle not found: $BundlePath"
            return $false
        }

        Write-Host "Selected bundle: $BundlePath" -ForegroundColor Green
        Write-Host "File size: $([math]::Round((Get-Item $BundlePath).Length / 1MB, 2)) MB" -ForegroundColor White

        # Clean up any previous extraction
        if (Test-Path $ExtractPath) {
            Remove-Item $ExtractPath -Recurse -Force -ErrorAction SilentlyContinue
        }

        # Extract the bundle
        Write-Host "Extracting PowerCLI bundle..." -ForegroundColor Yellow
        [System.IO.Compression.ZipFile]::ExtractToDirectory($BundlePath, $ExtractPath)

        # Check what was extracted
        Write-Host "Contents extracted to: $ExtractPath" -ForegroundColor Green
        $Items = Get-ChildItem $ExtractPath
        Write-Host "Found $($Items.Count) items in bundle:" -ForegroundColor Cyan
        $Items | ForEach-Object { Write-Host "  - $($_.Name) ($($_.GetType().Name))" }

        # Look for .nupkg files
        $NupkgFiles = Get-ChildItem $ExtractPath -Filter "*.nupkg" -Recurse
        if ($NupkgFiles.Count -eq 0) {
            Write-Host "No .nupkg files found. Checking for module directories..." -ForegroundColor Yellow
            
            # Alternative approach - look for module folders
            $ModuleDirs = Get-ChildItem $ExtractPath -Directory | Where-Object { 
                $_.Name -like "VMware.*" -or (Test-Path "$($_.FullName)\*.psd1")
            }
            
            if ($ModuleDirs) {
                Write-Host "Found $($ModuleDirs.Count) module directories" -ForegroundColor Green
                InstallFromDirectories -ModuleDirs $ModuleDirs
            } else {
                throw "No recognizable module structure found in the bundle."
            }
        } else {
            Write-Host "Found $($NupkgFiles.Count) .nupkg files" -ForegroundColor Green
            InstallFromNupkg -NupkgFiles $NupkgFiles -ExtractPath $ExtractPath -RepositoryName $RepositoryName
        }

        # Verify installation
        return (VerifyInstallation)
    }
    catch {
        Write-Error "Installation failed: $($_.Exception.Message)"
        Write-Host "Full error: $($_.Exception)" -ForegroundColor Red
        return $false
    }
    finally {
        # Cleanup
        if (Test-Path $ExtractPath) {
            Write-Host "Cleaning up temporary files..." -ForegroundColor Yellow
            Remove-Item $ExtractPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

function InstallFromNupkg {
    param($NupkgFiles, $ExtractPath, $RepositoryName)
    
    Write-Host "Installing from .nupkg files..." -ForegroundColor Green
    
    # Register local repository
    if (Get-PSRepository -Name $RepositoryName -ErrorAction SilentlyContinue) {
        Unregister-PSRepository -Name $RepositoryName
    }
    
    Register-PSRepository -Name $RepositoryName -SourceLocation $ExtractPath -InstallationPolicy Trusted
    
    # Show available modules
    Write-Host "Available modules in bundle:" -ForegroundColor Cyan
    $NupkgFiles | ForEach-Object { 
        $moduleName = $_.BaseName -replace '\.\d+\.\d+\.\d+.*$', ''  # Remove version from filename
        Write-Host "  - $moduleName" -ForegroundColor White
    }

    # Install the main PowerCLI module (it should handle dependencies)
    Write-Host "`nInstalling VMware.PowerCLI module..." -ForegroundColor Cyan
    $installResult = Install-Module -Name VMware.PowerCLI -Repository $RepositoryName -Scope CurrentUser -Force -AllowClobber -ErrorAction SilentlyContinue
    
    if (-not $installResult) {
        Write-Host "Trying to install core modules individually..." -ForegroundColor Yellow
        
        # Install core modules individually
        $CoreModules = @(
            "VMware.Vim",
            "VMware.VimAutomation.Core", 
            "VMware.VimAutomation.Common",
            "VMware.VimAutomation.Sdk",
            "VMware.PowerCLI"
        )
        
        foreach ($Module in $CoreModules) {
            $matchingFile = $NupkgFiles | Where-Object { $_.Name -match $Module }
            if ($matchingFile) {
                Write-Host "Installing $Module..." -ForegroundColor Cyan
                Install-Module -Name $Module -Repository $RepositoryName -Scope CurrentUser -Force -AllowClobber -ErrorAction SilentlyContinue
            }
        }
    }
    
    # Clean up repository
    Unregister-PSRepository -Name $RepositoryName
}

function InstallFromDirectories {
    param($ModuleDirs)
    
    Write-Host "Installing from module directories..." -ForegroundColor Green
    
    $ModulePath = "$env:USERPROFILE\Documents\WindowsPowerShell\Modules"
    if (-not (Test-Path $ModulePath)) {
        New-Item -ItemType Directory -Path $ModulePath -Force
    }
    
    foreach ($ModuleDir in $ModuleDirs) {
        $ModuleName = $ModuleDir.Name
        $Destination = Join-Path $ModulePath $ModuleName
        
        Write-Host "Copying $ModuleName to modules directory..." -ForegroundColor Cyan
        
        # Remove existing version
        if (Test-Path $Destination) {
            Write-Host "  Removing existing version..." -ForegroundColor Yellow
            Remove-Item $Destination -Recurse -Force
        }
        
        # Copy new version
        Copy-Item -Path $ModuleDir.FullName -Destination $Destination -Recurse -Force
        
        # Verify the module can be loaded
        $Manifest = Get-ChildItem $Destination -Filter "*.psd1" | Select-Object -First 1
        if ($Manifest) {
            Write-Host "  ✓ $ModuleName" -ForegroundColor Green
        } else {
            Write-Host "  ⚠ $ModuleName (no manifest found)" -ForegroundColor Yellow
        }
    }
}

function VerifyInstallation {
    Write-Host "`nVerifying installation..." -ForegroundColor Green
    
    # Refresh module cache
    Get-Module -ListAvailable | Out-Null
    
    $VMModules = Get-Module -Name VMware.* -ListAvailable
    if ($VMModules) {
        Write-Host "✓ Successfully installed $($VMModules.Count) VMware modules:" -ForegroundColor Green
        $VMModules | Sort-Object Name | ForEach-Object { 
            Write-Host "  - $($_.Name) v$($_.Version)" -ForegroundColor White
        }
        
        # Test import
        Write-Host "`nTesting module import..." -ForegroundColor Yellow
        try {
            Import-Module VMware.PowerCLI -Force -ErrorAction Stop
            Write-Host "✓ VMware.PowerCLI imported successfully!" -ForegroundColor Green
            
            # Show available commands
            $CmdCount = (Get-Command -Module VMware.PowerCLI).Count
            Write-Host "  Available commands: $CmdCount" -ForegroundColor White
            return $true
        }
        catch {
            Write-Host "⚠ Could not import VMware.PowerCLI: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "Trying individual module import..." -ForegroundColor Yellow
            
            # Try importing core modules individually
            try {
                Import-Module VMware.Vim -ErrorAction SilentlyContinue
                Import-Module VMware.VimAutomation.Core -ErrorAction SilentlyContinue
                Write-Host "✓ Core modules imported individually" -ForegroundColor Green
                return $true
            }
            catch {
                Write-Host "✗ Failed to import modules individually" -ForegroundColor Red
                return $false
            }
        }
    } else {
        Write-Host "✗ No VMware modules found after installation" -ForegroundColor Red
        return $false
    }
}

function Show-ConfigurationMenu {
    Write-Host "`nPowerCLI Configuration Menu" -ForegroundColor Magenta
    Write-Host "==========================" -ForegroundColor Magenta
    
    $choice = Read-Host "Would you like to configure PowerCLI settings? (Y/N)"
    if ($choice -eq 'Y' -or $choice -eq 'y') {
        try {
            Import-Module VMware.PowerCLI -Force -ErrorAction Stop
            
            Write-Host "`nConfiguring PowerCLI settings..." -ForegroundColor Green
            Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $false -Confirm:$false
            Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false
            Set-PowerCLIConfiguration -Scope User -Confirm:$false
            
            Write-Host "`nCurrent PowerCLI Configuration:" -ForegroundColor Cyan
            Get-PowerCLIConfiguration
            
            Write-Host "`n✓ PowerCLI configuration completed!" -ForegroundColor Green
        }
        catch {
            Write-Host "⚠ Could not configure PowerCLI: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

# Main execution
Clear-Host
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host "    PowerCLI Offline Installer" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan
Write-Host ""

# Show file browser dialog
Write-Host "Please select the PowerCLI offline bundle file..." -ForegroundColor Yellow
$SelectedBundle = Show-FileBrowser

if ($SelectedBundle) {
    Write-Host "`nSelected file: $SelectedBundle" -ForegroundColor Green
    
    # Confirm installation
    $confirm = Read-Host "`nProceed with PowerCLI installation? (Y/N)"
    if ($confirm -eq 'Y' -or $confirm -eq 'y') {
        $success = Install-PowerCLIFromBundle -BundlePath $SelectedBundle
        
        if ($success) {
            Write-Host "`n PowerCLI installation completed successfully!" -ForegroundColor Green
            Show-ConfigurationMenu
        } else {
            Write-Host "`n PowerCLI installation failed." -ForegroundColor Red
        }
    } else {
        Write-Host "Installation cancelled by user." -ForegroundColor Yellow
    }
} else {
    Write-Host "No file selected. Installation cancelled." -ForegroundColor Yellow
}

Write-Host "`nPress any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
