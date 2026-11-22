param([switch]$SkipLanguageTool)

$ErrorActionPreference = 'Stop'
$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$vendorDir = Join-Path $projectDir 'vendor'

if (-not (Get-Command pyw.exe -ErrorAction SilentlyContinue)) {
    throw 'Python for Windows is required. Install Python, then run setup.ps1 again.'
}

if (-not $SkipLanguageTool) {
    if (-not (Get-Command java.exe -ErrorAction SilentlyContinue)) {
        throw 'Java is required for offline LanguageTool. Install Java, then run setup.ps1 again.'
    }
    $existingServer = Get-ChildItem -LiteralPath $vendorDir -Filter 'languagetool-server.jar' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $existingServer) {
        New-Item -ItemType Directory -Force -Path $vendorDir | Out-Null
        $archive = Join-Path $vendorDir 'LanguageTool-latest-snapshot.zip'
        Write-Host 'Downloading LanguageTool from its official site (about 250 MB)...'
        Invoke-WebRequest -Uri 'https://languagetool.org/download/snapshots/LanguageTool-latest-snapshot.zip' -OutFile $archive
        Write-Host 'Extracting LanguageTool...'
        Expand-Archive -LiteralPath $archive -DestinationPath $vendorDir -Force
    }
}

$desktop = [Environment]::GetFolderPath('Desktop')
$shortcutPath = Join-Path $desktop "Kira's Pawn.lnk"
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = (Get-Command pyw.exe).Source
$shortcut.Arguments = '"' + (Join-Path $projectDir 'writing_mode.py') + '"'
$shortcut.WorkingDirectory = $projectDir
$shortcut.Description = "Kira's Pawn - writing, research, dictionary, and local grammar"
$shortcut.IconLocation = 'C:\Windows\System32\shell32.dll,70'
$shortcut.Save()

Write-Host "Created $shortcutPath"
Write-Host 'Double-click the shortcut to begin.'
