Set-StrictMode -Version Latest
$ErrorActionPreference="Stop"
Import-Module (Join-Path $PSScriptRoot "src\F1-Mailer.Core.psm1") -Force

$s=Get-F1Settings
$runs=Get-F1UpcomingRunDates -FirstRunDate ([string]$s.first_run_date) -ScheduleTime ([string]$s.schedule_time) -IntervalDays ([int]$s.interval_days) -From (Get-Date) -Count 5

Write-Host "F1 OUTLOOK 14-DAY MAILER"
Write-Host "PROSSIME 5 ESECUZIONI:"
foreach($r in $runs){ Write-Host (" - {0:dd/MM/yyyy HH:mm}" -f $r) }
