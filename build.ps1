param(
    [string]$PythonPath = 'python'
)

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
$pythonCommand = Get-Command $PythonPath -ErrorAction SilentlyContinue
if (-not $pythonCommand) { throw "找不到 Python：$PythonPath。请用 -PythonPath 指定非 Conda 的 Python 3.10+。" }
$PythonPath = $pythonCommand.Source
$buildVenv = '.venv-build'
if (Test-Path $buildVenv) { Remove-Item $buildVenv -Recurse -Force }
& $PythonPath -m venv $buildVenv
& .\.venv-build\Scripts\python.exe -m pip install --upgrade pip setuptools wheel
& .\.venv-build\Scripts\python.exe -m pip install '.[build]'
& .\.venv-build\Scripts\pyinstaller.exe --noconfirm --clean --onefile --noconsole `
    --name KzzMonitor `
    --collect-all akshare `
    --collect-all py_mini_racer `
    --collect-all pystray `
    --collect-all winotify `
    --hidden-import openpyxl `
    launcher.py
New-Item -ItemType Directory -Force -Path release | Out-Null
Copy-Item dist\KzzMonitor.exe release\KzzMonitor.exe -Force
if (Test-Path '可转债监控.xlsx') { Copy-Item '可转债监控.xlsx' release\ -Force }
Copy-Item README.md release\ -Force
& .\release\KzzMonitor.exe --version-probe
if ($LASTEXITCODE -ne 0) { throw 'EXE 启动自检失败。' }
Write-Host "构建并自检完成: $PSScriptRoot\release\KzzMonitor.exe"
