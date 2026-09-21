Set-StrictMode -Version Latest
$ErrorActionPreference="Stop"
Import-Module (Join-Path $PSScriptRoot "src\F1-Mailer.Core.psm1") -Force

try {
    $s=Get-F1Settings
    $task=Get-F1TaskStatus

    if(-not $task.Exists){ throw "TASK_F1_MANCANTE" }
    if(-not $task.Enabled){ throw "TASK_F1_DISABILITATA" }
    if(-not $task.StartWhenAvailable){ throw "TASK_START_WHEN_AVAILABLE_FALSE" }
    if($null -ne $task.IntervalDays -and [int]$task.IntervalDays -ne [int]$s.interval_days){
        throw "TASK_INTERVAL_MISMATCH"
    }

    Write-Host "TASK: ATTIVA"
    Write-Host "NOME: F1 OUTLOOK MONTHLY MAILER"
    Write-Host "CADENZA CONFIGURATA: $($s.interval_days) GIORNI"
    Write-Host "INTERVALLO TASK: $($task.IntervalDays)"
    Write-Host "START WHEN AVAILABLE: $($task.StartWhenAvailable)"
    Write-Host "START BOUNDARY: $($task.StartBoundary)"
    Write-Host "ULTIMA ESECUZIONE: $($task.LastRunTime)"
    Write-Host "PROSSIMA ESECUZIONE: $($task.NextRunTime)"
    exit 0
} catch {
    Write-Host "VERIFICA TASK: ERRORE"
    Write-Host $_.Exception.Message
    exit 1
}
