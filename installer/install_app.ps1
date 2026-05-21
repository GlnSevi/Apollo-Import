param(
    [string]$SourceExe = ""
)

$ErrorActionPreference = "Stop"

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourcePath = $null
$resolvedSourceExe = $SourceExe.Trim()

if ($resolvedSourceExe) {
    $sourcePath = Join-Path $scriptRoot $resolvedSourceExe
} else {
    $candidate = Get-ChildItem -LiteralPath $scriptRoot -Filter "*onefile.exe" -File |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($null -ne $candidate) {
        $sourcePath = $candidate.FullName
    }
}

$displaySource = if ($resolvedSourceExe) { $resolvedSourceExe } else { "*onefile.exe" }
if (-not (Test-Path -LiteralPath $sourcePath)) {
    throw "Anwendungsdatei nicht gefunden: $displaySource"
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
