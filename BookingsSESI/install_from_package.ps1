param(
    [string]$InstallDir = "$env:LOCALAPPDATA\BookingsSESI",
    [switch]$NoLaunch
)

$ErrorActionPreference = "Stop"
$PayloadRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$AppDisplayName = "SESI Reservas Recorrentes"
$DesktopShortcut = Join-Path ([Environment]::GetFolderPath("Desktop")) "$AppDisplayName.lnk"
$StartMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\$AppDisplayName"
$StartMenuShortcut = Join-Path $StartMenuDir "$AppDisplayName.lnk"
$UninstallShortcut = Join-Path $StartMenuDir "Desinstalar $AppDisplayName.lnk"
$LauncherCmd = Join-Path $InstallDir "Iniciar $AppDisplayName.cmd"
$UninstallCmd = Join-Path $InstallDir "Desinstalar $AppDisplayName.cmd"
$UninstallScript = Join-Path $InstallDir "uninstall_bookings_sesi.ps1"
$IconFile = Join-Path $InstallDir "bookings_sesi.ico"

$requiredItems = @(
    "bookings_site_automator.py",
    "selectors_bookings_sesi.json",
    "bookings_jobs_exemplo.json",
    "python_runtime",
    "ms-playwright",
    "bookings_sesi.ico",
    "sesi_logo_app.png",
    "sesi_logo_recortada.png",
    "sesi_logo_vermelha.png"
)

foreach ($item in $requiredItems) {
    $source = Join-Path $PayloadRoot $item
    if (-not (Test-Path $source)) {
        throw "Item obrigatorio nao encontrado no pacote: $item"
    }
}

New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
New-Item -ItemType Directory -Path $StartMenuDir -Force | Out-Null

$existingTargets = @(
    "bookings_site_automator.py",
    "selectors_bookings_sesi.json",
    "bookings_jobs_exemplo.json",
    "python_runtime",
    "ms-playwright",
    "playwright_auth",
    "bookings_app_settings.json",
    "bookings_sesi.ico",
    "sesi_logo_app.png",
    "sesi_logo_recortada.png",
    "sesi_logo_vermelha.png",
    "Iniciar $AppDisplayName.cmd",
    "Desinstalar $AppDisplayName.cmd",
    "uninstall_bookings_sesi.ps1"
)

foreach ($target in $existingTargets) {
    $dest = Join-Path $InstallDir $target
    if (Test-Path $dest) {
        Remove-Item -LiteralPath $dest -Recurse -Force
    }
}

foreach ($item in $requiredItems) {
    $source = Join-Path $PayloadRoot $item
    $destination = Join-Path $InstallDir $item
    Copy-Item -LiteralPath $source -Destination $destination -Recurse -Force
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
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0uninstall_bookings_sesi.ps1"
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

$shell = New-Object -ComObject WScript.Shell

$desktop = $shell.CreateShortcut($DesktopShortcut)
$desktop.TargetPath = $LauncherCmd
$desktop.WorkingDirectory = $InstallDir
$desktop.IconLocation = "$IconFile,0"
$desktop.Save()

$startMenu = $shell.CreateShortcut($StartMenuShortcut)
$startMenu.TargetPath = $LauncherCmd
$startMenu.WorkingDirectory = $InstallDir
$startMenu.IconLocation = "$IconFile,0"
$startMenu.Save()

$uninstallMenu = $shell.CreateShortcut($UninstallShortcut)
$uninstallMenu.TargetPath = $UninstallCmd
$uninstallMenu.WorkingDirectory = $InstallDir
$uninstallMenu.IconLocation = "$IconFile,0"
$uninstallMenu.Save()

if (-not $NoLaunch) {
    Start-Process -FilePath $LauncherCmd
}
