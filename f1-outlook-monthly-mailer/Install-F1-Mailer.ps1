Set-StrictMode -Version Latest
$ErrorActionPreference="Stop"

if ($env:OS -ne "Windows_NT") { throw "WINDOWS_REQUIRED" }

$base=Join-Path $env:LOCALAPPDATA "F1OutlookMonthlyMailer"
$appDir=Join-Path $base "app"
New-Item -ItemType Directory -Path $appDir -Force | Out-Null

$source=(Resolve-Path $PSScriptRoot).Path
$dest=[IO.Path]::GetFullPath($appDir)
if ($source.TrimEnd("\") -ne $dest.TrimEnd("\")) {
    Get-ChildItem -Path $PSScriptRoot -Force | Where-Object { $_.Name -notin @(".git","dist","build") } | ForEach-Object {
        Copy-Item -Path $_.FullName -Destination $appDir -Recurse -Force
    }
}

Import-Module (Join-Path $appDir "src\F1-Mailer.Core.psm1") -Force
Initialize-F1LocalStore | Out-Null

$settingsPath=Join-Path $base "settings.json"
if (-not (Test-Path $settingsPath)) {
    Copy-Item (Join-Path $appDir "settings.example.json") $settingsPath
}
$s=Get-F1Settings

if ([int]$s.schedule_day -lt 1 -or [int]$s.schedule_day -gt 28) { throw "schedule_day deve essere compreso tra 1 e 28" }
if ([string]$s.schedule_time -notmatch '^([01]\d|2[0-3]):[0-5]\d$') { throw "schedule_time deve essere HH:mm" }

Write-Host "Verifica Outlook Classic e account F1..."
$d=Get-F1OutlookDiagnostics -Sender $s.sender

if(-not $d.F1AccountPresent){
    Write-Host ""
    Write-Host "=================================================="
    Write-Host "CONFIGURAZIONE OUTLOOK F1 NECESSARIA"
    Write-Host "=================================================="
    Write-Host "Outlook Classic: $(if($d.ClassicInstalled){'DISPONIBILE'}else{'NON TROVATO'})"
    Write-Host "Account richiesto: $($s.sender)"
    Write-Host "Account attualmente rilevati:"
    if($d.AccountCount -eq 0){
        Write-Host " - nessuno"
    } else {
        $d.Accounts | ForEach-Object {
            $v=if($_.SmtpAddress){$_.SmtpAddress}else{$_.DisplayName}
            Write-Host " - $v"
        }
    }
    Write-Host "L'account F1 deve essere aggiunto al profilo Outlook Classic."
    Write-Host "Avvio configurazione Outlook..."
    Write-Host "=================================================="

    & (Join-Path $appDir "Setup-F1-Outlook-Account.ps1")
    $setupExit=$LASTEXITCODE
    $d=Get-F1OutlookDiagnostics -Sender $s.sender
    if(-not $d.F1AccountPresent){
        Write-Host ""
        Write-Host "INSTALLAZIONE IN ATTESA: Outlook Classic non vede ancora $($s.sender)."
        Write-Host "Riesegui Install-F1-Mailer.ps1 dopo aver completato l'aggiunta dell'account."
        exit $(if($setupExit){$setupExit}else{2})
    }
}

$ctx=Get-F1OutlookContext -Sender $s.sender -ContactsFolderName $s.contacts_folder -CreateFolder -RequiredCategory $s.required_category -BlockedCategory $s.blocked_category
Write-Host "OUTLOOK CLASSIC: OK"
Write-Host "ACCOUNT F1: OK - $($s.sender)"
Write-Host "CARTELLA CONTATTI: OK - $($s.contacts_folder)"
Write-Host "F1-CONSENSO: OK"
Write-Host "F1-DISCRITTO: OK"

$user=[Security.Principal.WindowsIdentity]::GetCurrent().Name
$xml=New-F1TaskXml -User $user -AppDir $appDir -ScheduleDay ([int]$s.schedule_day) -ScheduleTime ([string]$s.schedule_time)
Register-ScheduledTask -TaskName "F1 OUTLOOK MONTHLY MAILER" -Xml $xml -Force | Out-Null
$task=Get-F1TaskStatus

Write-Host ""
Write-Host "INSTALLAZIONE COMPLETATA"
Write-Host "Versione: 1.0.2"
Write-Host "Task: F1 OUTLOOK MONTHLY MAILER"
Write-Host "Task attiva: $($task.Enabled)"
Write-Host "StartWhenAvailable: $($task.StartWhenAvailable)"
Write-Host "Prossima esecuzione: $($task.NextRunTime)"
Write-Host "Configurazione: $settingsPath"
Write-Host "Pannello: $(Join-Path $appDir 'F1-Control-Panel.ps1')"
Write-Host ""
Write-Host "L'installazione NON ha inviato alcuna email."
Write-Host "Per una prova eseguire Test-Email.ps1."
