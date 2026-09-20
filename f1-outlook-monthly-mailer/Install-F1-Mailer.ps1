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
$ctx=Get-F1OutlookContext -Sender $s.sender -ContactsFolderName $s.contacts_folder -CreateFolder -RequiredCategory $s.required_category -BlockedCategory $s.blocked_category
Write-Host "OUTLOOK CLASSIC: OK"
Write-Host "ACCOUNT F1: OK - $($s.sender)"
Write-Host "CARTELLA CONTATTI: OK - $($s.contacts_folder)"

$time=[datetime]::ParseExact([string]$s.schedule_time,"HH:mm",[Globalization.CultureInfo]::InvariantCulture)
$now=Get-Date
$day=[int]$s.schedule_day
$start=Get-Date -Year $now.Year -Month $now.Month -Day $day -Hour $time.Hour -Minute $time.Minute -Second 0
if ($start -le $now) { $start=$start.AddMonths(1) }

$main=Join-Path $appDir "F1-Mailer.ps1"
$arg='-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}"' -f $main
$xmlArg=[Security.SecurityElement]::Escape($arg)
$xmlWork=[Security.SecurityElement]::Escape($appDir)
$user=[Security.SecurityElement]::Escape([Security.Principal.WindowsIdentity]::GetCurrent().Name)
$startText=$start.ToString("yyyy-MM-ddTHH:mm:ss")
$months="<January/><February/><March/><April/><May/><June/><July/><August/><September/><October/><November/><December/>"

$xml=@"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo><Description>Invio mensile F1 tramite Outlook Classic.</Description></RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>$startText</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByMonth>
        <DaysOfMonth><Day>$day</Day></DaysOfMonth>
        <Months>$months</Months>
      </ScheduleByMonth>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>$user</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <ExecutionTimeLimit>PT4H</ExecutionTimeLimit>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>powershell.exe</Command>
      <Arguments>$xmlArg</Arguments>
      <WorkingDirectory>$xmlWork</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"@

Register-ScheduledTask -TaskName "F1 OUTLOOK MONTHLY MAILER" -Xml $xml -Force | Out-Null
$info=Get-ScheduledTaskInfo -TaskName "F1 OUTLOOK MONTHLY MAILER"

Write-Host ""
Write-Host "INSTALLAZIONE COMPLETATA"
Write-Host "Task: F1 OUTLOOK MONTHLY MAILER"
Write-Host "Prossima esecuzione: $($info.NextRunTime)"
Write-Host "Configurazione: $settingsPath"
Write-Host "Pannello: $(Join-Path $appDir 'F1-Control-Panel.ps1')"
Write-Host ""
Write-Host "L'installazione NON ha inviato alcuna email."
Write-Host "Per una prova eseguire Test-Email.ps1."
