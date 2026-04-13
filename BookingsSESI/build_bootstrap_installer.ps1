param(
    [string]$OutputDir = "dist_installer"
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$PayloadBuilder = Join-Path $ProjectRoot "build_exe_installer.ps1"
$InstallerWorkDir = "BookingsSESI_InstallerBuild_" + [guid]::NewGuid().ToString("N")
$PayloadZip = Join-Path (Join-Path $env:TEMP $InstallerWorkDir) "payload.zip"
$PyInstallerPayloadDir = Join-Path $env:TEMP "BookingsSESI_InstallerBuild"
$PyInstallerPayloadZip = Join-Path $PyInstallerPayloadDir "payload.zip"
$PyInstallerWorkDir = Join-Path $env:TEMP ("BookingsSESI_PyInstaller_" + [guid]::NewGuid().ToString("N"))
$PyInstallerDistDir = Join-Path $ProjectRoot "dist_gui"
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
New-Item -ItemType Directory -Path $PyInstallerPayloadDir -Force | Out-Null
Copy-Item -LiteralPath $PayloadZip -Destination $PyInstallerPayloadZip -Force
New-Item -ItemType Directory -Path $PyInstallerWorkDir -Force | Out-Null
New-Item -ItemType Directory -Path $PyInstallerDistDir -Force | Out-Null

& $VenvPython -m PyInstaller --noconfirm --distpath $PyInstallerDistDir --workpath $PyInstallerWorkDir (Join-Path $ProjectRoot "installer_gui.spec")

$BuiltInstaller = Join-Path $PyInstallerDistDir "BookingsSESI-Instalador.exe"
if (-not (Test-Path $BuiltInstaller)) {
    throw "Instalador grafico nao encontrado em $BuiltInstaller"
}

Copy-Item -LiteralPath $BuiltInstaller `
    -Destination (Join-Path $OutputDirFull "BookingsSESI-Instalador.exe") `
    -Force
