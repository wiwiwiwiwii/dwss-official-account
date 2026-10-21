param([switch]$InstallDependencies)

$ErrorActionPreference = 'Stop'
$projectDirectory = $PSScriptRoot
$workspaceDirectory = Split-Path -Parent $projectDirectory
$environmentDirectory = Join-Path $workspaceDirectory '.build-venv'
$buildPython = Join-Path $environmentDirectory 'Scripts\python.exe'

if ($InstallDependencies) {
    if (-not (Test-Path -LiteralPath $buildPython)) {
        & python -m venv --system-site-packages $environmentDirectory
        if ($LASTEXITCODE -ne 0) { throw 'Failed to create the build environment.' }
    }
    & $buildPython -m pip install --index-url https://pypi.org/simple -r (Join-Path $projectDirectory 'requirements-build.txt')
    if ($LASTEXITCODE -ne 0) {
        & $buildPython -m pip install --index-url https://pypi.tuna.tsinghua.edu.cn/simple -r (Join-Path $projectDirectory 'requirements-build.txt')
        if ($LASTEXITCODE -ne 0) { throw 'Dependency installation failed on both sources.' }
    }
}
if (-not (Test-Path -LiteralPath $buildPython)) {
    throw 'Run build.ps1 -InstallDependencies first.'
}
# EXE 按版本命名（v1.1.0 → 报告转公众号v1_1.exe，v1.1.0 beta 2 → 报告转公众号v1_1_beta2.exe，补丁号非零时追加 _N），旧版本文件不被覆盖
$appSource = Get-Content -LiteralPath (Join-Path $projectDirectory 'app.py') -Raw -Encoding utf8
if ($appSource -notmatch 'APP_VERSION = "v(\d+)\.(\d+)\.(\d+)(?: beta (\d+))?"') { throw 'APP_VERSION not found in app.py.' }
$versionTag = "v$($Matches[1])_$($Matches[2])" + $(if ($Matches[3] -ne '0') { "_$($Matches[3])" } else { '' }) + $(if ($Matches[4]) { "_beta$($Matches[4])" } else { '' })
$exeName = "报告转公众号$versionTag"
$env:WORKBENCH_EXE_NAME = $exeName
& $buildPython -m PyInstaller --noconfirm --distpath (Join-Path $workspaceDirectory 'release') --workpath (Join-Path $workspaceDirectory '.build-output') (Join-Path $projectDirectory 'workbench.spec')
if ($LASTEXITCODE -ne 0) { throw 'Build failed. If the EXE is running, close the workbench before rebuilding.' }
$releaseExe = Join-Path $workspaceDirectory "release\$exeName.exe"
$fixedEntry = Join-Path $workspaceDirectory "$exeName.exe"
# 固定分发入口与 release 保持同一文件（README 的双副本契约）
Copy-Item -LiteralPath $releaseExe -Destination $fixedEntry -Force
Get-FileHash -Algorithm SHA256 -LiteralPath $releaseExe, $fixedEntry
