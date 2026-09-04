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
& $buildPython -m PyInstaller --noconfirm --distpath (Join-Path $workspaceDirectory 'release') --workpath (Join-Path $workspaceDirectory '.build-output') (Join-Path $projectDirectory 'workbench.spec')
if ($LASTEXITCODE -ne 0) { throw 'Build failed. If the EXE is running, close the workbench before rebuilding.' }
$releaseExe = Join-Path $workspaceDirectory 'release\报告转公众号.exe'
$fixedEntry = Join-Path $workspaceDirectory '报告转公众号.exe'
# 固定分发入口与 release 保持同一文件（README 的双副本契约）
Copy-Item -LiteralPath $releaseExe -Destination $fixedEntry -Force
Get-FileHash -Algorithm SHA256 -LiteralPath $releaseExe, $fixedEntry
