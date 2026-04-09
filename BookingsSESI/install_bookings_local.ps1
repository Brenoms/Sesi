param(
    [string]$InstallDir = "$env:LOCALAPPDATA\BookingsSESI"
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$VenvPython = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
$PythonBase = (& $VenvPython -c "import sys; print(sys.base_prefix)").Trim()
$AppDisplayName = "SESI Reservas Recorrentes"
$DesktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) "$AppDisplayName.lnk"
$StartMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\$AppDisplayName"
$StartMenuShortcut = Join-Path $StartMenuDir "$AppDisplayName.lnk"
$UninstallShortcut = Join-Path $StartMenuDir "Desinstalar $AppDisplayName.lnk"
$LauncherCmd = Join-Path $InstallDir "Iniciar $AppDisplayName.cmd"
$UninstallCmd = Join-Path $InstallDir "Desinstalar $AppDisplayName.cmd"
$UninstallScript = Join-Path $InstallDir "uninstall_bookings_local.ps1"
$IconFile = Join-Path $InstallDir "bookings_sesi.ico"

Write-Host "Instalando em: $InstallDir"

if (-not (Test-Path $VenvPython)) {
    throw "Python da .venv nao encontrado em $VenvPython"
}
if (-not (Test-Path $PythonBase)) {
    throw "Diretorio do Python base nao encontrado em $PythonBase"
}

New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
New-Item -ItemType Directory -Path $StartMenuDir -Force | Out-Null

$itemsToCopy = @(
    "bookings_site_automator.py",
    "selectors_bookings_sesi.json",
    "bookings_jobs_exemplo.json",
    "bookings_sesi.ico",
    "imagens\sesi_logo_app.png",
    "sesi_logo_recortada.png",
    "imagens\sesi_logo_vermelha.png"
)

foreach ($item in $itemsToCopy) {
    $source = Join-Path $ProjectRoot $item
    if (-not (Test-Path $source)) {
        throw "Item obrigatorio nao encontrado: $source"
    }

    $destination = Join-Path $InstallDir ([System.IO.Path]::GetFileName($item))
    if (Test-Path $destination) {
        Remove-Item -LiteralPath $destination -Recurse -Force
    }
    Copy-Item -LiteralPath $source -Destination $destination -Recurse -Force
}

$browserSource = Join-Path $env:LOCALAPPDATA "ms-playwright"
if (-not (Test-Path $browserSource)) {
    throw "Pasta do Chromium do Playwright nao encontrada em $browserSource"
}
$browserDestination = Join-Path $InstallDir "ms-playwright"
if (Test-Path $browserDestination) {
    Remove-Item -LiteralPath $browserDestination -Recurse -Force
}
Copy-Item -LiteralPath $browserSource -Destination $browserDestination -Recurse -Force

$runtimeDestination = Join-Path $InstallDir "python_runtime"
if (Test-Path $runtimeDestination) {
    Remove-Item -LiteralPath $runtimeDestination -Recurse -Force
}
New-Item -ItemType Directory -Path $runtimeDestination -Force | Out-Null

foreach ($dir in @("DLLs", "Lib", "libs", "tcl")) {
    Copy-Item -LiteralPath (Join-Path $PythonBase $dir) -Destination (Join-Path $runtimeDestination $dir) -Recurse -Force
}

foreach ($file in @("python.exe", "pythonw.exe", "python3.dll", "python313.dll", "vcruntime140.dll", "vcruntime140_1.dll", "LICENSE.txt")) {
    Copy-Item -LiteralPath (Join-Path $PythonBase $file) -Destination (Join-Path $runtimeDestination $file) -Force
}

$sitePackagesSource = Join-Path $ProjectRoot ".venv\Lib\site-packages"
$sitePackagesDestination = Join-Path $runtimeDestination "Lib\site-packages"
New-Item -ItemType Directory -Path $sitePackagesDestination -Force | Out-Null
Get-ChildItem -LiteralPath $sitePackagesSource -Force | ForEach-Object {
    Copy-Item -LiteralPath $_.FullName -Destination $sitePackagesDestination -Recurse -Force
}

$launcherContent = @"
@echo off
cd /d "%~dp0"
set "APP_HOME=%~dp0"
set "PYTHONHOME=%APP_HOME%python_runtime"
set "PYTHONPATH=%APP_HOME%python_runtime\Lib;%APP_HOME%python_runtime\Lib\site-packages"
set "TCL_LIBRARY=%APP_HOME%python_runtime\tcl\tcl8.6"
set "TK_LIBRARY=%APP_HOME%python_runtime\tcl\tk8.6"
set "PLAYWRIGHT_BROWSERS_PATH=%APP_HOME%ms-playwright"
start "" "%APP_HOME%python_runtime\pythonw.exe" "%APP_HOME%bookings_site_automator.py"
"@
Set-Content -LiteralPath $LauncherCmd -Value $launcherContent -Encoding ASCII

$uninstallCmdContent = @"
@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall_bookings_local.ps1"
"@
Set-Content -LiteralPath $UninstallCmd -Value $uninstallCmdContent -Encoding ASCII

$uninstallContent = @"
param()
`$appDisplayName = "$AppDisplayName"
`$desktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) "`$appDisplayName.lnk"
`$startMenuDir = Join-Path `$env:APPDATA "Microsoft\Windows\Start Menu\Programs\`$appDisplayName"
`$startMenuShortcut = Join-Path `$startMenuDir "`$appDisplayName.lnk"
`$uninstallShortcut = Join-Path `$startMenuDir "Desinstalar `$appDisplayName.lnk"
if (Test-Path `$desktopShortcut) { Remove-Item -LiteralPath `$desktopShortcut -Force }
if (Test-Path `$startMenuShortcut) { Remove-Item -LiteralPath `$startMenuShortcut -Force }
if (Test-Path `$uninstallShortcut) { Remove-Item -LiteralPath `$uninstallShortcut -Force }
if (Test-Path `$startMenuDir) { Remove-Item -LiteralPath `$startMenuDir -Recurse -Force }
if (Test-Path "$InstallDir") { Remove-Item -LiteralPath "$InstallDir" -Recurse -Force }
"@
Set-Content -LiteralPath $UninstallScript -Value $uninstallContent -Encoding UTF8

$wshell = New-Object -ComObject WScript.Shell

$desktop = $wshell.CreateShortcut($DesktopShortcut)
$desktop.TargetPath = $LauncherCmd
$desktop.WorkingDirectory = $InstallDir
$desktop.IconLocation = "$IconFile,0"
$desktop.Save()

$startMenu = $wshell.CreateShortcut($StartMenuShortcut)
$startMenu.TargetPath = $LauncherCmd
$startMenu.WorkingDirectory = $InstallDir
$startMenu.IconLocation = "$IconFile,0"
$startMenu.Save()

$uninstallMenu = $wshell.CreateShortcut($UninstallShortcut)
$uninstallMenu.TargetPath = $UninstallCmd
$uninstallMenu.WorkingDirectory = $InstallDir
$uninstallMenu.IconLocation = "$IconFile,0"
$uninstallMenu.Save()

Write-Host "Instalacao concluida."
Write-Host "Atalho criado na area de trabalho e no menu Iniciar."
