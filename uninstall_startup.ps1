$task = Get-ScheduledTask -TaskName 'KzzMonitor' -ErrorAction SilentlyContinue
if ($task) {
    Unregister-ScheduledTask -TaskName 'KzzMonitor' -Confirm:$false
    Write-Host '已移除计划任务 KzzMonitor。'
} else {
    Write-Host '计划任务 KzzMonitor 不存在。'
}
