Add-Type -AssemblyName System.Windows.Forms
$ErrorActionPreference = "Stop"

$repoApiUrl = "https://api.github.com/repos/Interkarma/daggerfall-unity/releases/latest"
$dataUrl = "https://drive.usercontent.google.com/download?id=0B0i8ZocaUWLGWHc1WlF3dHNUNTQ&export=download&confirm=yes&uuid=5151e262-ed9f-44c8-8ebe-eab55f22c78e"

$tempUnityZip = "$env:TEMP\DaggerfallUnity-latest.zip"
$tempDataZip = "$env:TEMP\DaggerfallGameFiles.zip"

$headers = @{
    "User-Agent" = "Mozilla/5.0"
}

Write-Host "Fetching latest Daggerfall Unity release info..." -ForegroundColor Yellow
$releaseInfo = Invoke-RestMethod -Uri $repoApiUrl -Headers $headers

$version = $releaseInfo.tag_name.TrimStart("v")
Write-Host "Latest version detected: $version" -ForegroundColor Green

$asset = $releaseInfo.assets | Where-Object { $_.name -like "dfu_windows_64bit-*.zip" }

if (!$asset) {
    Write-Error "Could not find dfu_windows_64bit zip in latest release assets."
    exit
}

$unityUrl = $asset.browser_download_url
Write-Host "Downloading Daggerfall Unity $version..." -ForegroundColor Yellow

# Ask user to pick folder
$folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$folderDialog.Description = "Select where to create the Daggerfall Unity versioned folder"
$null = $folderDialog.ShowDialog()
$parentDir = $folderDialog.SelectedPath

if ([string]::IsNullOrWhiteSpace($parentDir)) {
    Write-Host "No folder selected. Exiting..."
    exit
}

# Create versioned install directory
$installDir = Join-Path $parentDir "DaggerfallUnity-$version"
$gameDataDir = Join-Path $installDir "Daggerfall"

# Confirm if folder exists
if (Test-Path -Path $installDir) {
    Add-Type -AssemblyName System.Windows.Forms
    $result = [System.Windows.Forms.MessageBox]::Show(
        "The folder '$installDir' already exists. Overwrite it?",
        "Folder Exists",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Question
    )
    if ($result -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Aborted by user."
        exit
    }
    Write-Host "Clearing existing folder..."
    Remove-Item -Recurse -Force "$installDir\*"
} else {
    New-Item -ItemType Directory -Path $installDir | Out-Null
}

# Download helper
function Download-File($url, $outPath) {
    Write-Host "Downloading... (this may take a few minutes)" -ForegroundColor Yellow
    Invoke-WebRequest -Uri $url -OutFile $outPath -Headers @{ "User-Agent" = "Mozilla/5.0" }
}

# Download files
Download-File $unityUrl $tempUnityZip

Write-Host "Downloading official Daggerfall game files..." -ForegroundColor Yellow
Download-File $dataUrl $tempDataZip

# Extract
Write-Host "Extracting Daggerfall Unity..." -ForegroundColor Yellow
Expand-Archive -Path $tempUnityZip -DestinationPath $installDir -Force

Write-Host "Extracting Daggerfall game files..." -ForegroundColor Yellow
if (!(Test-Path -Path $gameDataDir)) {
    New-Item -ItemType Directory -Path $gameDataDir | Out-Null
}
Expand-Archive -Path $tempDataZip -DestinationPath $gameDataDir -Force

# Cleanup temp files
Write-Host "Cleaning up temporary files..." -ForegroundColor Yellow
Remove-Item $tempUnityZip
Remove-Item $tempDataZip

# Final message
Write-Host ""
Write-Host "Daggerfall Unity $version installed successfully in $installDir." -ForegroundColor Green
Write-Host "Official Daggerfall game files extracted to $gameDataDir."
Write-Host "On first launch, point Daggerfall Unity to the '$gameDataDir' folder when it asks for game files."

# Final confirmation form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Installation Complete"
$form.Size = New-Object System.Drawing.Size(400,260)
$form.StartPosition = "CenterScreen"

$checkboxDesktopShortcut = New-Object System.Windows.Forms.CheckBox
$checkboxDesktopShortcut.Text = "Create Desktop shortcut"
$checkboxDesktopShortcut.Location = New-Object System.Drawing.Point(30,30)
$checkboxDesktopShortcut.Size = New-Object System.Drawing.Size(300,30)

$checkboxStartMenuShortcut = New-Object System.Windows.Forms.CheckBox
$checkboxStartMenuShortcut.Text = "Create Start Menu shortcut"
$checkboxStartMenuShortcut.Location = New-Object System.Drawing.Point(30,60)
$checkboxStartMenuShortcut.Size = New-Object System.Drawing.Size(300,30)

$checkboxOpenFolder = New-Object System.Windows.Forms.CheckBox
$checkboxOpenFolder.Text = "Open install folder"
$checkboxOpenFolder.Location = New-Object System.Drawing.Point(30,90)
$checkboxOpenFolder.Size = New-Object System.Drawing.Size(300,30)

$checkboxLaunchGame = New-Object System.Windows.Forms.CheckBox
$checkboxLaunchGame.Text = "Launch Daggerfall Unity"
$checkboxLaunchGame.Location = New-Object System.Drawing.Point(30,120)
$checkboxLaunchGame.Size = New-Object System.Drawing.Size(300,30)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "OK"
$okButton.Location = New-Object System.Drawing.Point(150,170)
$okButton.Size = New-Object System.Drawing.Size(100,30)

# Setup WScriptShell
$WScriptShell = New-Object -ComObject WScript.Shell

$okButton.Add_Click({
    if ($checkboxDesktopShortcut.Checked) {
        $desktopShortcut = $WScriptShell.CreateShortcut((Join-Path $env:USERPROFILE\Desktop "Daggerfall Unity.lnk"))
        $desktopShortcut.TargetPath = (Join-Path $installDir "DaggerfallUnity.exe")
        $desktopShortcut.WorkingDirectory = $installDir
        $desktopShortcut.IconLocation = (Join-Path $installDir "DaggerfallUnity.exe")
        $desktopShortcut.Save()
    }
    if ($checkboxStartMenuShortcut.Checked) {
        $startMenuFolder = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Daggerfall Unity"
        if (!(Test-Path -Path $startMenuFolder)) {
            New-Item -ItemType Directory -Path $startMenuFolder | Out-Null
        }
        $startShortcut = $WScriptShell.CreateShortcut((Join-Path $startMenuFolder "Daggerfall Unity.lnk"))
        $startShortcut.TargetPath = (Join-Path $installDir "DaggerfallUnity.exe")
        $startShortcut.WorkingDirectory = $installDir
        $startShortcut.IconLocation = (Join-Path $installDir "DaggerfallUnity.exe")
        $startShortcut.Save()
    }
    if ($checkboxOpenFolder.Checked) {
        Start-Process "explorer.exe" -ArgumentList "`"$installDir`""
    }
    if ($checkboxLaunchGame.Checked) {
        Start-Process (Join-Path $installDir "DaggerfallUnity.exe")
    }
    $form.Close()
})

$form.Controls.Add($checkboxDesktopShortcut)
$form.Controls.Add($checkboxStartMenuShortcut)
$form.Controls.Add($checkboxOpenFolder)
$form.Controls.Add($checkboxLaunchGame)
$form.Controls.Add($okButton)

$form.Topmost = $true
$form.ShowDialog()

