$ErrorActionPreference = 'Stop'
$base = $PSScriptRoot
$exe = Join-Path $base 'KzzMonitor.exe'
if (-not (Test-Path $exe)) { $exe = Join-Path $base 'release\KzzMonitor.exe' }
if (Test-Path $exe) {
    $program = $exe
    $arguments = 'run'
    $workingDirectory = Split-Path $exe
} else {
    $program = Join-Path $base '.venv\Scripts\pythonw.exe'
    if (-not (Test-Path $program)) { throw '请先运行 build.ps1，或创建 .venv 并安装项目。' }
    $arguments = '-m kzz_monitor.cli run'
    $workingDirectory = $base
}
$action = New-ScheduledTaskAction -Execute $program -Argument $arguments -WorkingDirectory $workingDirectory
$logon = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$daily = New-ScheduledTaskTrigger -Daily -At '09:20'
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 2)
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
Register-ScheduledTask -TaskName 'KzzMonitor' -Action $action -Trigger @($logon, $daily) -Settings $settings -Principal $principal -Description '可转债监控：登录及交易日开盘前启动' -Force | Out-Null
Write-Host '已安装计划任务 KzzMonitor（登录时及每天 09:20 启动）。'
