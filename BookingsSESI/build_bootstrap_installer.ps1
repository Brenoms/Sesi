param(
    [string]$OutputDir = "dist_installer"
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$PayloadBuilder = Join-Path $ProjectRoot "build_exe_installer.ps1"
$InstallerWorkDir = "BookingsSESI_InstallerBuild_" + [guid]::NewGuid().ToString("N")
$PayloadZip = Join-Path (Join-Path $env:TEMP $InstallerWorkDir) "payload.zip"
$VenvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$OutputDirFull = Join-Path $ProjectRoot $OutputDir

if (-not (Test-Path $VenvPython)) {
    throw "Python da .venv não encontrado em $VenvPython"
}

& powershell -ExecutionPolicy Bypass -File $PayloadBuilder -OutputDir $OutputDir -InstallerWorkDir $InstallerWorkDir

if (-not (Test-Path $PayloadZip)) {
    throw "Payload ZIP não encontrado em $PayloadZip"
}

New-Item -ItemType Directory -Path $OutputDirFull -Force | Out-Null

& $VenvPython -m PyInstaller --clean --noconfirm (Join-Path $ProjectRoot "installer_gui.spec")

$BuiltInstaller = Join-Path $ProjectRoot "dist\BookingsSESI-Instalador.exe"
if (-not (Test-Path $BuiltInstaller)) {
    throw "Instalador grafico nao encontrado em $BuiltInstaller"
}

Copy-Item -LiteralPath $BuiltInstaller `
    -Destination (Join-Path $OutputDirFull "BookingsSESI-Instalador.exe") `
    -Force
