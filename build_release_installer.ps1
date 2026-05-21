param(
    [string]$Version = "v0.1.2"
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$iexpress = Join-Path $env:SystemRoot "System32\\iexpress.exe"
if (-not (Test-Path -LiteralPath $iexpress)) {
    throw "IExpress wurde nicht gefunden: $iexpress"
}

$onefileSpec = Join-Path $root "Apollo-Import-GUI-v0.1-onefile.spec"
$distRoot = Join-Path $root "dist"
$stageDir = Join-Path $root "build\\installer-stage"
$sedPath = Join-Path $root "build\\Apollo-Import-GUI-setup.sed"
$onefileExeName = "Apollo-Import-GUI-v0.1.2-onefile.exe"
$onefileExePath = Join-Path $distRoot $onefileExeName
$installerExePath = Join-Path $distRoot "Apollo-Import-GUI-$Version-setup.exe"

python -m py_compile apollo_import_gui.py
pyinstaller --noconfirm $onefileSpec

if (-not (Test-Path -LiteralPath $onefileExePath)) {
    throw "Onefile-EXE wurde nicht erzeugt: $onefileExePath"
}

if (Test-Path -LiteralPath $stageDir) {
    Remove-Item -LiteralPath $stageDir -Recurse -Force
}
New-Item -ItemType Directory -Path $stageDir -Force | Out-Null

Copy-Item -LiteralPath $onefileExePath -Destination (Join-Path $stageDir $onefileExeName) -Force
Copy-Item -LiteralPath (Join-Path $root "installer\\install.cmd") -Destination (Join-Path $stageDir "install.cmd") -Force
Copy-Item -LiteralPath (Join-Path $root "installer\\install_app.ps1") -Destination (Join-Path $stageDir "install_app.ps1") -Force

$sedContent = @"
[Version]
Class=IEXPRESS
SEDVersion=3
[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=1
HideExtractAnimation=0
UseLongFileName=1
InsideCompressed=0
CAB_FixedSize=0
CAB_ResvCodeSigning=0
RebootMode=N
InstallPrompt=
DisplayLicense=
FinishMessage=Die Installation wurde abgeschlossen.
TargetName=$installerExePath
FriendlyName=Apollo Import GUI Setup
AppLaunched=install.cmd
PostInstallCmd=<None>
AdminQuietInstCmd=install.cmd
UserQuietInstCmd=install.cmd
SourceFiles=SourceFiles
[Strings]
FILE0=$onefileExeName
FILE1=install.cmd
FILE2=install_app.ps1
[SourceFiles]
SourceFiles0=$stageDir
[SourceFiles0]
%FILE0%=
%FILE1%=
%FILE2%=
"@

$sedContent | Set-Content -LiteralPath $sedPath -Encoding ASCII

if (Test-Path -LiteralPath $installerExePath) {
    Remove-Item -LiteralPath $installerExePath -Force
}

& $iexpress /N $sedPath

if (-not (Test-Path -LiteralPath $installerExePath)) {
    throw "Installer-EXE wurde nicht erzeugt: $installerExePath"
}

Write-Host $installerExePath
