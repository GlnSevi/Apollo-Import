param(
    [string]$SourceExe = "Apollo-Import-GUI-v0.1-onefile.exe"
)

$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourcePath = Join-Path $scriptRoot $SourceExe
if (-not (Test-Path -LiteralPath $sourcePath)) {
    throw "Anwendungsdatei nicht gefunden: $sourcePath"
}

$installDir = Join-Path $env:LocalAppData "Programs\\Apollo Import GUI"
$targetExe = Join-Path $installDir "Apollo Import GUI.exe"
$startMenuDir = Join-Path $env:APPDATA "Microsoft\\Windows\\Start Menu\\Programs\\Apollo Import GUI"
$desktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) "Apollo Import GUI.lnk"
$startMenuShortcut = Join-Path $startMenuDir "Apollo Import GUI.lnk"

New-Item -ItemType Directory -Path $installDir -Force | Out-Null
New-Item -ItemType Directory -Path $startMenuDir -Force | Out-Null

Copy-Item -LiteralPath $sourcePath -Destination $targetExe -Force

$wshShell = New-Object -ComObject WScript.Shell
foreach ($shortcutPath in @($desktopShortcut, $startMenuShortcut)) {
    $shortcut = $wshShell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $targetExe
    $shortcut.WorkingDirectory = $installDir
    $shortcut.IconLocation = $targetExe
    $shortcut.Save()
}
