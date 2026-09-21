Set-StrictMode -Version Latest
$ErrorActionPreference="Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module (Join-Path $PSScriptRoot "src\F1-Mailer.Core.psm1") -Force

$form=New-Object Windows.Forms.Form
$form.Text="F1 OUTLOOK 14-DAY MAILER"
$form.Size=New-Object Drawing.Size(710,720)
$form.StartPosition="CenterScreen"
$form.Font=New-Object Drawing.Font("Segoe UI",10)

$title=New-Object Windows.Forms.Label
$title.Text="F1 OUTLOOK 14-DAY MAILER v1.1.1"
$title.Font=New-Object Drawing.Font("Segoe UI",16,[Drawing.FontStyle]::Bold)
$title.AutoSize=$true
$title.Location=New-Object Drawing.Point(24,20)
$form.Controls.Add($title)

$status=New-Object Windows.Forms.TextBox
$status.Multiline=$true
$status.ReadOnly=$true
$status.ScrollBars="Vertical"
$status.Location=New-Object Drawing.Point(24,65)
$status.Size=New-Object Drawing.Size(645,285)
$form.Controls.Add($status)

function Refresh-Panel {
    try {
        $s=Get-F1Settings
        $st=Get-F1State
        $now=Get-Date
        $cycleDate=Get-F1CycleDate -Date $now -FirstRunDate ([string]$s.first_run_date) -IntervalDays ([int]$s.interval_days)
        $campaign=Get-F1CampaignKey -Date $cycleDate
        $d=Get-F1OutlookDiagnostics -Sender $s.sender
        $task=Get-F1TaskStatus

        $sent=@($st.records | Where-Object { $_.campaign_key -eq $campaign -and $_.status -eq "SENT" }).Count
        $failed=@($st.records | Where-Object { $_.campaign_key -eq $campaign -and $_.status -eq "FAILED" }).Count
        $uncertain=@($st.records | Where-Object { $_.campaign_key -eq $campaign -and $_.status -eq "UNCERTAIN" }).Count

        $contacts="NON DISPONIBILI"
        $consents="NON DISPONIBILI"
        $unsubscribed="NON DISPONIBILI"

        if($d.F1AccountPresent){
            try {
                $ctx=Get-F1OutlookContext -Sender $s.sender -ContactsFolderName $s.contacts_folder -RequiredCategory $s.required_category -BlockedCategory $s.blocked_category
                $list=@(Get-F1Contacts $ctx.ContactsFolder)
                $contacts=$list.Count
                $consents=@($list | Where-Object {
                    (Test-F1Category $_.Categories $s.required_category) -and
                    -not (Test-F1Category $_.Categories $s.blocked_category)
                }).Count
                $unsubscribed=@($list | Where-Object {
                    Test-F1Category $_.Categories $s.blocked_category
                }).Count
            } catch {}
        }

        $next=$task.NextRunTime
        if(-not $next){
            $next=Get-F1NextScheduledRun -FirstRunDate ([string]$s.first_run_date) -ScheduleTime ([string]$s.schedule_time) -IntervalDays ([int]$s.interval_days) -From $now
        }
        $daysToNext=[Math]::Max(0,[Math]::Ceiling(($next-$now).TotalDays))

        $status.Text=@"
VERSIONE: 1.1.1
OUTLOOK CLASSIC: $(if($d.ClassicInstalled -and $d.MapiAvailable){'OK'}else{'ERRORE'})
ACCOUNT F1: $(if($d.F1AccountPresent){'OK'}else{'MANCANTE'})
ACCOUNT TROVATI: $($d.AccountCount)
CONTATTI: $contacts
CONSENSI: $consents
DISISCRITTI: $unsubscribed
CICLO ATTUALE: $campaign
INTERVALLO: $($s.interval_days) GIORNI
PRIMA ESECUZIONE: $($s.first_run_date) $($s.schedule_time)
EMAIL INVIATE NEL CICLO: $sent
EMAIL FALLITE: $failed
EMAIL UNCERTAIN: $uncertain
ULTIMA ESECUZIONE: $($st.last_run_at)
PROSSIMA ESECUZIONE: $next
GIORNI ALLA PROSSIMA: $daysToNext
TASK: $(if($task.Exists -and $task.Enabled){'ATTIVA'}elseif($task.Exists){'DISABILITATA'}else{'MANCANTE'})
START WHEN AVAILABLE: $($task.StartWhenAvailable)
PAUSA: $($st.paused)
"@
    } catch {
        $status.Text=$_.Exception.Message
    }
}

function Add-Button([string]$text,[int]$x,[int]$y,[scriptblock]$action) {
    $b=New-Object Windows.Forms.Button
    $b.Text=$text
    $b.Location=New-Object Drawing.Point($x,$y)
    $b.Size=New-Object Drawing.Size(205,42)
    $b.Add_Click($action)
    $form.Controls.Add($b)
}

function Launch-Script([string]$name) {
    $path=Join-Path $PSScriptRoot $name
    $args='-NoProfile -ExecutionPolicy Bypass -NoExit -File "{0}"' -f $path
    Start-Process -FilePath "powershell.exe" -ArgumentList $args
}

Add-Button "DIAGNOSTICA OUTLOOK" 24 370 { Launch-Script "Diagnose-Outlook.ps1" }
Add-Button "CONFIGURA ACCOUNT F1" 242 370 { Launch-Script "Setup-F1-Outlook-Account.ps1" }
Add-Button "TEST OUTLOOK" 460 370 { Launch-Script "Test-Outlook.ps1" }

Add-Button "INVIA TEST" 24 425 { Launch-Script "Test-Email.ps1" }
Add-Button "VERIFICA EMAIL TEST" 242 425 { Launch-Script "Verify-Test-Email.ps1" }
Add-Button "ESEGUI ORA" 460 425 { Launch-Script "Run-F1-Mailer-Now.ps1" }

Add-Button "PAUSA" 24 480 {
    $st=Get-F1State
    $st.paused=$true
    Save-F1State $st
    Refresh-Panel
}
Add-Button "RIPRENDI" 242 480 {
    $st=Get-F1State
    $st.paused=$false
    Save-F1State $st
    Refresh-Panel
}
Add-Button "APRI LOG" 460 480 { Launch-Script "Open-Logs.ps1" }

Add-Button "APRI CONTATTI F1" 24 535 {
    try {
        $s=Get-F1Settings
        $c=Get-F1OutlookContext -Sender $s.sender -ContactsFolderName $s.contacts_folder -CreateFolder -RequiredCategory $s.required_category -BlockedCategory $s.blocked_category
        $c.ContactsFolder.Display()
    } catch {
        [Windows.Forms.MessageBox]::Show($_.Exception.Message,"F1") | Out-Null
    }
}
Add-Button "IMPOSTAZIONI" 242 535 {
    Start-Process notepad.exe -ArgumentList (Get-F1SettingsPath)
}
Add-Button "AGGIORNA" 460 535 { Refresh-Panel }

Add-Button "VERIFICA TASK 14 GIORNI" 24 590 { Launch-Script "Verify-Task-14Days.ps1" }
Add-Button "MOSTRA PROSSIME ESECUZIONI" 242 590 { Launch-Script "Show-Next-Runs.ps1" }

Refresh-Panel
[void]$form.ShowDialog()
