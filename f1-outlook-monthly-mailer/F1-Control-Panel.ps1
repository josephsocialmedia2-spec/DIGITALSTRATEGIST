Set-StrictMode -Version Latest
$ErrorActionPreference="Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module (Join-Path $PSScriptRoot "src\F1-Mailer.Core.psm1") -Force

$form=New-Object Windows.Forms.Form
$form.Text="F1 EMAIL ENGINE"
$form.Size=New-Object Drawing.Size(690,650)
$form.StartPosition="CenterScreen"
$form.Font=New-Object Drawing.Font("Segoe UI",10)

$title=New-Object Windows.Forms.Label
$title.Text="F1 OUTLOOK MONTHLY MAILER v1.0.1"
$title.Font=New-Object Drawing.Font("Segoe UI",16,[Drawing.FontStyle]::Bold)
$title.AutoSize=$true
$title.Location=New-Object Drawing.Point(24,20)
$form.Controls.Add($title)

$status=New-Object Windows.Forms.TextBox
$status.Multiline=$true
$status.ReadOnly=$true
$status.ScrollBars="Vertical"
$status.Location=New-Object Drawing.Point(24,65)
$status.Size=New-Object Drawing.Size(625,245)
$form.Controls.Add($status)

function Refresh-Panel {
    try {
        $s=Get-F1Settings
        $st=Get-F1State
        $campaign=Get-F1CampaignKey
        $d=Get-F1OutlookDiagnostics -Sender $s.sender
        $task=Get-F1TaskStatus
        $sent=@($st.records | Where-Object { $_.campaign_key -eq $campaign -and $_.status -eq "SENT" }).Count
        $failed=@($st.records | Where-Object { $_.campaign_key -eq $campaign -and $_.status -eq "FAILED" }).Count
        $contacts="NON DISPONIBILI"
        if($d.F1AccountPresent){
            try {
                $ctx=Get-F1OutlookContext -Sender $s.sender -ContactsFolderName $s.contacts_folder -RequiredCategory $s.required_category -BlockedCategory $s.blocked_category
                $contacts=@(Get-F1Contacts $ctx.ContactsFolder).Count
            } catch {}
        }
        $status.Text=@"
OUTLOOK CLASSIC: $(if($d.ClassicInstalled -and $d.MapiAvailable){'OK'}else{'ERRORE'})
ACCOUNT F1: $(if($d.F1AccountPresent){'OK'}else{'MANCANTE'})
ACCOUNT TROVATI: $($d.AccountCount)
STATO OUTLOOK: $($d.Status)
CAMPAGNA: $campaign
CONTATTI: $contacts
INVIATI: $sent
FALLITI: $failed
PAUSA: $($st.paused)
ULTIMA ESECUZIONE: $($st.last_run_at)
TASK: $(if($task.Exists -and $task.Enabled){'ATTIVA'}elseif($task.Exists){'DISABILITATA'}else{'MANCANTE'})
START WHEN AVAILABLE: $($task.StartWhenAvailable)
PROSSIMA ESECUZIONE: $($task.NextRunTime)
"@
    } catch { $status.Text=$_.Exception.Message }
}

function Add-Button([string]$text,[int]$x,[int]$y,[scriptblock]$action) {
    $b=New-Object Windows.Forms.Button
    $b.Text=$text
    $b.Location=New-Object Drawing.Point($x,$y)
    $b.Size=New-Object Drawing.Size(195,42)
    $b.Add_Click($action)
    $form.Controls.Add($b)
}

function Launch-Script([string]$name) {
    $path=Join-Path $PSScriptRoot $name
    $args='-NoProfile -ExecutionPolicy Bypass -NoExit -File "{0}"' -f $path
    Start-Process -FilePath "powershell.exe" -ArgumentList $args
}

Add-Button "DIAGNOSTICA OUTLOOK" 24 330 { Launch-Script "Diagnose-Outlook.ps1" }
Add-Button "CONFIGURA ACCOUNT F1" 235 330 { Launch-Script "Setup-F1-Outlook-Account.ps1" }
Add-Button "TEST OUTLOOK" 446 330 { Launch-Script "Test-Outlook.ps1" }

Add-Button "INVIA TEST" 24 385 { Launch-Script "Test-Email.ps1" }
Add-Button "VERIFICA EMAIL TEST" 235 385 { Launch-Script "Verify-Test-Email.ps1" }
Add-Button "ESEGUI ORA" 446 385 { Launch-Script "Run-F1-Mailer-Now.ps1" }

Add-Button "PAUSA" 24 440 {
    $st=Get-F1State; $st.paused=$true; Save-F1State $st; Refresh-Panel
}
Add-Button "RIPRENDI" 235 440 {
    $st=Get-F1State; $st.paused=$false; Save-F1State $st; Refresh-Panel
}
Add-Button "APRI LOG" 446 440 { Launch-Script "Open-Logs.ps1" }

Add-Button "APRI CONTATTI F1" 24 495 {
    try {
        $s=Get-F1Settings
        $c=Get-F1OutlookContext -Sender $s.sender -ContactsFolderName $s.contacts_folder -CreateFolder -RequiredCategory $s.required_category -BlockedCategory $s.blocked_category
        $c.ContactsFolder.Display()
    } catch { [Windows.Forms.MessageBox]::Show($_.Exception.Message,"F1") | Out-Null }
}
Add-Button "IMPOSTAZIONI" 235 495 { Start-Process notepad.exe -ArgumentList (Get-F1SettingsPath) }
Add-Button "AGGIORNA" 446 495 { Refresh-Panel }

Refresh-Panel
[void]$form.ShowDialog()
