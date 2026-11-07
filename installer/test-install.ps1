param([Parameter(Mandatory = $true)][string]$Installer)
$ErrorActionPreference = 'Stop'

$dataDir = Join-Path $env:LOCALAPPDATA 'PHDTech\Commission Tool'
$registryKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{E2A3D376-B63E-4C84-9B61-90CE5FDC8D22}_is1'
$machineRegistryKey = 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{E2A3D376-B63E-4C84-9B61-90CE5FDC8D22}_is1'
if ((Test-Path $dataDir) -or (Test-Path $registryKey) -or (Test-Path $machineRegistryKey)) {
    throw 'Installer smoke tests require a clean Windows account; existing application data was found.'
}
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('commission-installer-' + [Guid]::NewGuid().ToString('N'))
$installed = Join-Path $testRoot 'Installed App With Spaces'
$workbooks = Join-Path $testRoot 'Existing Workbooks'
$settings = Join-Path $dataDir 'settings.json'
$profileMarker = Join-Path $dataDir 'connectwise-profile\synthetic-marker.txt'
$workbookMarker = Join-Path $workbooks 'Synthetic Workbook.xlsx'
$legacyMarker = Join-Path $installed 'Synthetic Legacy Workbook.xlsx'
$createdData = $false

function Install-Preview {
    $arguments = @('/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART', "/DIR=`"$installed`"")
    $process = Start-Process -FilePath $Installer -ArgumentList $arguments -Wait -PassThru
    if ($process.ExitCode -ne 0) { throw "Installer failed with exit code $($process.ExitCode)." }
}
function Test-InstalledPackage {
    $launcher = Start-Process -FilePath (Join-Path $installed 'PHDTech Commission Tool.exe') -ArgumentList '--self-check' -Wait -PassThru
    if ($launcher.ExitCode -ne 0) { throw 'Installed desktop executable failed its self-check.' }
    & (Join-Path $installed 'runtime\node.exe') (Join-Path $PSScriptRoot 'smoke-package.js') $installed
    if ($LASTEXITCODE -ne 0) { throw 'Installed runtime, workbook writer or browser failed its smoke test.' }
}
function Assert-UserData {
    if ((Get-Content -LiteralPath $settings -Raw) -ne $script:settingsBefore) { throw 'Installer changed existing settings.' }
    if ((Get-Content -LiteralPath $profileMarker -Raw) -ne 'Synthetic profile marker') { throw 'Installer changed browser-profile data.' }
    if ((Get-Content -LiteralPath $workbookMarker -Raw) -ne 'Synthetic workbook preservation marker') { throw 'Installer changed the workbook folder.' }
    if ((Get-Content -LiteralPath $legacyMarker -Raw) -ne 'Synthetic legacy workbook marker') { throw 'Installer changed an untracked workbook inside the application folder.' }
}

try {
    New-Item -ItemType Directory -Path (Split-Path $profileMarker -Parent), $workbooks -Force | Out-Null
    $createdData = $true
    [IO.File]::WriteAllText($settings, (@{ version = 1; workbookDir = $workbooks } | ConvertTo-Json))
    [IO.File]::WriteAllText($profileMarker, 'Synthetic profile marker')
    [IO.File]::WriteAllText($workbookMarker, 'Synthetic workbook preservation marker')
    $script:settingsBefore = Get-Content -LiteralPath $settings -Raw
    Install-Preview
    [IO.File]::WriteAllText($legacyMarker, 'Synthetic legacy workbook marker')
    Test-InstalledPackage
    Assert-UserData
    # Reinstall with the same application ID, exercising the update path.
    Install-Preview
    Test-InstalledPackage
    Assert-UserData
    $uninstaller = Get-ChildItem -LiteralPath $installed -Filter 'unins*.exe' | Select-Object -First 1
    if (-not $uninstaller) { throw 'The installer did not provide an uninstaller.' }
    $process = Start-Process -FilePath $uninstaller.FullName -ArgumentList '/VERYSILENT', '/SUPPRESSMSGBOXES', '/NORESTART' -Wait -PassThru
    if ($process.ExitCode -ne 0) { throw "Uninstaller failed with exit code $($process.ExitCode)." }
    if (Test-Path (Join-Path $installed 'app\commission-menu.js')) { throw 'Application files remain after uninstall.' }
    Assert-UserData
    Write-Output 'Install, reinstall, installed writer/browser, uninstall and user-data preservation checks passed.'
} finally {
    if ($createdData -and (Test-Path $dataDir)) { Remove-Item -LiteralPath $dataDir -Recurse -Force }
    if (Test-Path $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force -ErrorAction SilentlyContinue }
}
