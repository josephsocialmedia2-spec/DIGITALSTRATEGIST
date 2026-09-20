Set-StrictMode -Version Latest
$ErrorActionPreference="Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module (Join-Path $PSScriptRoot "src\F1-Mailer.Core.psm1") -Force

$form=New-Object Windows.Forms.Form
$form.Text="F1 EMAIL ENGINE"
$form.Size=New-Object Drawing.Size(640,520)
$form.StartPosition="CenterScreen"
$form.Font=New-Object Drawing.Font("Segoe UI",10)

$title=New-Object Windows.Forms.Label
$title.Text="F1 OUTLOOK MONTHLY MAILER"
$title.Font=New-Object Drawing.Font("Segoe UI",16,[Drawing.FontStyle]::Bold)
$title.AutoSize=$true
$title.Location=New-Object Drawing.Point(24,20)
$form.Controls.Add($title)

$status=New-Object Windows.Forms.TextBox
$status.Multiline=$true
$status.ReadOnly=$true
$status.ScrollBars="Vertical"
$status.Location=New-Object Drawing.Point(24,65)
$status.Size=New-Object Drawing.Size(575,190)
$form.Controls.Add($status)

function Refresh-Panel {
    try {
        $s=Get-F1Settings
        $st=Get-F1State
        $campaign=Get-F1CampaignKey
        $sent=@($st.records | Where-Object { $_.campaign_key -eq $campaign -and $_.status -eq "SENT" }).Count
        $failed=@($st.records | Where-Object { $_.campaign_key -eq $campaign -and $_.status -eq "FAILED" }).Count
        $next="NON INSTALLATA"
        try { $next=(Get-ScheduledTaskInfo -TaskName "F1 OUTLOOK MONTHLY MAILER").NextRunTime } catch {}
        $status.Text=@"
OUTLOOK: da verificare con TEST OUTLOOK
ACCOUNT: $($s.sender)
CAMPAGNA: $campaign
INVIATI: $sent
FALLITI: $failed
PAUSA: $($st.paused)
ULTIMA ESECUZIONE: $($st.last_run_at)
PROSSIMA ESECUZIONE: $next
CONTATTI: $($s.contacts_folder)
"@
    } catch { $status.Text=$_.Exception.Message }
}

function Add-Button([string]$text,[int]$x,[int]$y,[scriptblock]$action) {
    $b=New-Object Windows.Forms.Button
    $b.Text=$text
    $b.Location=New-Object Drawing.Point($x,$y)
    $b.Size=New-Object Drawing.Size(180,42)
    $b.Add_Click($action)
    $form.Controls.Add($b)
}

function Launch-Script([string]$name) {
    $path=Join-Path $PSScriptRoot $name
    $args='-NoProfile -ExecutionPolicy Bypass -File "{0}"' -f $path
    Start-Process -FilePath "powershell.exe" -ArgumentList $args
}

Add-Button "TEST OUTLOOK" 24 280 { Launch-Script "Test-Outlook.ps1" }
Add-Button "INVIA TEST" 220 280 { Launch-Script "Test-Email.ps1" }
Add-Button "ESEGUI ORA" 416 280 { Launch-Script "Run-F1-Mailer-Now.ps1" }

Add-Button "PAUSA" 24 335 {
    $st=Get-F1State; $st.paused=$true; Save-F1State $st; Refresh-Panel
}
Add-Button "RIPRENDI" 220 335 {
    $st=Get-F1State; $st.paused=$false; Save-F1State $st; Refresh-Panel
}
Add-Button "APRI LOG" 416 335 { Launch-Script "Open-Logs.ps1" }

Add-Button "APRI CONTATTI OUTLOOK" 24 390 {
    try {
        $s=Get-F1Settings
        $c=Get-F1OutlookContext -Sender $s.sender -ContactsFolderName $s.contacts_folder -CreateFolder -RequiredCategory $s.required_category -BlockedCategory $s.blocked_category
        $c.ContactsFolder.Display()
    } catch { [Windows.Forms.MessageBox]::Show($_.Exception.Message,"F1") | Out-Null }
}
Add-Button "IMPOSTAZIONI" 220 390 {
    Start-Process notepad.exe -ArgumentList (Get-F1SettingsPath)
}
Add-Button "AGGIORNA" 416 390 { Refresh-Panel }

Refresh-Panel
[void]$form.ShowDialog()
