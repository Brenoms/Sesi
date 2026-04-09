param(
    [string]$OutputDir = "dist_installer",
    [string]$InstallerWorkDir = ""
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($InstallerWorkDir)) {
    $InstallerWorkDir = "BookingsSESI_InstallerBuild_" + [guid]::NewGuid().ToString("N")
}
$InstallerWork = Join-Path $env:TEMP $InstallerWorkDir
$PayloadDir = Join-Path $InstallerWork "payload"
$ExtractCmd = Join-Path $InstallerWork "run_installer.cmd"
$PayloadZip = Join-Path $InstallerWork "payload.zip"
$SedFile = Join-Path $InstallerWork "BookingsSESI_Installer.sed"
$IExpress = Join-Path $env:WINDIR "System32\iexpress.exe"
$OutputDirFull = Join-Path $ProjectRoot $OutputDir
$TargetExeTemp = Join-Path $InstallerWork "BookingsSESI-Instalador.exe"
$TargetExe = Join-Path $OutputDirFull "BookingsSESI-Instalador.exe"
$VenvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"

$requiredPayloadItems = @(
    "bookings_site_automator.py",
    "selectors_bookings_sesi.json",
    "bookings_jobs_exemplo.json",
    "bookings_sesi.ico",
    "imagens\sesi_logo_app.png",
    "sesi_logo_recortada.png",
    "imagens\sesi_logo_vermelha.png"
)

if (-not (Test-Path $VenvPython)) {
    throw "Python da .venv nao encontrado em $VenvPython"
}

if (Test-Path $InstallerWork) {
    Remove-Item -LiteralPath $InstallerWork -Recurse -Force
}
New-Item -ItemType Directory -Path $PayloadDir -Force | Out-Null
New-Item -ItemType Directory -Path $OutputDirFull -Force | Out-Null

foreach ($item in $requiredPayloadItems) {
    $source = Join-Path $ProjectRoot $item
    if (-not (Test-Path $source)) {
        throw "Item obrigatorio nao encontrado: $source"
    }
    $destination = Join-Path $PayloadDir ([System.IO.Path]::GetFileName($item))
    Copy-Item -LiteralPath $source -Destination $destination -Recurse -Force
}

$pythonBase = (& $VenvPython -c "import sys; print(sys.base_prefix)").Trim()
if (-not $pythonBase) {
    throw "Nao foi possivel identificar o Python base para montar o runtime portatil."
}
if (-not (Test-Path $pythonBase)) {
    throw "Diretorio do Python base nao encontrado: $pythonBase"
}

$portableRuntimeDir = Join-Path $PayloadDir "python_runtime"
New-Item -ItemType Directory -Path $portableRuntimeDir -Force | Out-Null

$runtimeDirectories = @(
    "DLLs",
    "Lib",
    "libs",
    "tcl"
)

foreach ($dir in $runtimeDirectories) {
    $source = Join-Path $pythonBase $dir
    if (-not (Test-Path $source)) {
        throw "Diretorio obrigatorio do runtime nao encontrado: $source"
    }
    Copy-Item -LiteralPath $source -Destination (Join-Path $portableRuntimeDir $dir) -Recurse -Force
}

$runtimeFiles = @(
    "python.exe",
    "pythonw.exe",
    "python3.dll",
    "python313.dll",
    "vcruntime140.dll",
    "vcruntime140_1.dll",
    "LICENSE.txt"
)

foreach ($file in $runtimeFiles) {
    $source = Join-Path $pythonBase $file
    if (-not (Test-Path $source)) {
        throw "Arquivo obrigatorio do runtime nao encontrado: $source"
    }
    Copy-Item -LiteralPath $source -Destination (Join-Path $portableRuntimeDir $file) -Force
}

$sitePackagesSource = Join-Path $ProjectRoot ".venv\Lib\site-packages"
$sitePackagesDestination = Join-Path $portableRuntimeDir "Lib\site-packages"
if (-not (Test-Path $sitePackagesSource)) {
    throw "site-packages da .venv nao encontrado em $sitePackagesSource"
}
New-Item -ItemType Directory -Path $sitePackagesDestination -Force | Out-Null
Get-ChildItem -LiteralPath $sitePackagesSource -Force | ForEach-Object {
    Copy-Item -LiteralPath $_.FullName -Destination $sitePackagesDestination -Recurse -Force
}

$browserSource = Join-Path $env:LOCALAPPDATA "ms-playwright"
if (-not (Test-Path $browserSource)) {
    throw "Pasta do Chromium do Playwright nao encontrada em $browserSource"
}
Copy-Item -LiteralPath $browserSource -Destination (Join-Path $PayloadDir "ms-playwright") -Recurse -Force

Copy-Item -LiteralPath (Join-Path $ProjectRoot "install_from_package.ps1") -Destination (Join-Path $PayloadDir "install_from_package.ps1") -Force

if (Test-Path $PayloadZip) {
    Remove-Item -LiteralPath $PayloadZip -Force
}
Push-Location $PayloadDir
try {
    & tar.exe -a -c -f $PayloadZip *
} finally {
    Pop-Location
}

$cmdContent = @"
@echo off
setlocal
set "WORKDIR=%TEMP%\BookingsSESI_Install"
if exist "%WORKDIR%" rmdir /s /q "%WORKDIR%"
mkdir "%WORKDIR%"
powershell -NoProfile -ExecutionPolicy Bypass -Command "Expand-Archive -LiteralPath '%~dp0payload.zip' -DestinationPath '%WORKDIR%' -Force"
powershell -NoProfile -ExecutionPolicy Bypass -File "%WORKDIR%\install_from_package.ps1"
endlocal
"@
Set-Content -LiteralPath $ExtractCmd -Value $cmdContent -Encoding ASCII

$sedContent = @"
[Version]
Class=IEXPRESS
SEDVersion=3
[Options]
PackagePurpose=InstallApp
ShowInstallProgramWindow=1
HideExtractAnimation=1
UseLongFileName=1
InsideCompressed=0
CAB_FixedSize=0
CAB_ResvCodeSigning=0
RebootMode=N
InstallPrompt=
DisplayLicense=
FinishMessage=Instalacao concluida.
TargetName=$TargetExeTemp
FriendlyName=SESI Reservas Recorrentes Instalador
AppLaunched=run_installer.cmd
PostInstallCmd=<None>
AdminQuietInstCmd=
UserQuietInstCmd=
SourceFiles=SourceFiles
[SourceFiles]
SourceFiles0=$InstallerWork
[SourceFiles0]
%FILE0%=
%FILE1%=
[Strings]
FILE0=run_installer.cmd
FILE1=payload.zip
"@
Set-Content -LiteralPath $SedFile -Value $sedContent -Encoding ASCII

& $IExpress /N $SedFile | Out-Null

if ((Test-Path $TargetExeTemp) -and (Test-Path $TargetExe)) {
    Remove-Item -LiteralPath $TargetExe -Force
}
if (Test-Path $TargetExeTemp) {
    Copy-Item -LiteralPath $TargetExeTemp -Destination $TargetExe -Force
}

if (-not (Test-Path $TargetExe)) {
    throw "Falha ao gerar o instalador EXE."
}

Write-Host "Instalador gerado em: $TargetExe"
