<#
작업 스케줄러에 'FDA_Weekly_WarningLetter_Mail' 을 등록합니다.
매주 수요일 09:00 실행. PC가 꺼져 있었으면 켜진 뒤 바로 실행합니다.

    .\install_task.ps1            등록 (이미 있으면 덮어씀)
    .\install_task.ps1 -Remove    제거
#>
param([switch]$Remove)

$taskName = 'FDA_Weekly_WarningLetter_Mail'

if ($Remove) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-Host "작업을 제거했습니다: $taskName"
    exit 0
}

$script = Join-Path $PSScriptRoot 'run_fda_letters.ps1'
if (-not (Test-Path (Join-Path $PSScriptRoot 'fda-mail.config.ps1'))) {
    Write-Host "먼저 fda-mail.config.ps1 을 만들어 주세요 (example 파일 참조)." -ForegroundColor Red
    exit 2
}

$action = New-ScheduledTaskAction -Execute 'powershell.exe' `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$script`""

# FDA 는 경고장 목록을 통상 화요일(미 동부시간)에 갱신하므로 수요일 아침에 받습니다.
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Wednesday -At '09:00'

$settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30) `
    -MultipleInstances IgnoreNew

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings `
    -Description 'FDA Warning Letter 신규 건 메일 발송 (gmpai letters). 식약처 행정처분 메일과 같은 계정.' `
    -Force | Out-Null

$next = (Get-ScheduledTaskInfo -TaskName $taskName).NextRunTime
Write-Host "등록 완료: $taskName"
Write-Host "  실행 파일: $script"
Write-Host "  다음 실행: $next"
Write-Host "작업 스케줄러(taskschd.msc)에서 MFDS_Daily_Disposition_Mail 옆에 보입니다."
