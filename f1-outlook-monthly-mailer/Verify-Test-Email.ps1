param([int]$WaitSeconds=120)
Set-StrictMode -Version Latest
$ErrorActionPreference="Stop"
Import-Module (Join-Path $PSScriptRoot "src\F1-Mailer.Core.psm1") -Force

$subject="TEST F1 - OUTLOOK 14 DAY MAILER"
$since=(Get-Date).AddMinutes(-30)

try {
    $s=Get-F1Settings
    $c=Get-F1OutlookContext -Sender $s.sender -ContactsFolderName $s.contacts_folder -RequiredCategory $s.required_category -BlockedCategory $s.blocked_category

    $sentFolder=$c.Store.GetDefaultFolder(5)
    $sent=Find-F1MailBySubject -Folder $sentFolder -Subject $subject -Since $since
    if($sent){ Write-Host "POSTA INVIATA F1: OK" }
    else { Write-Host "POSTA INVIATA F1: NON TROVATA" }

    $inbox=$c.Store.GetDefaultFolder(6)
    $deadline=(Get-Date).AddSeconds([Math]::Max(0,$WaitSeconds))
    $received=$null
    do {
        $received=Find-F1MailBySubject -Folder $inbox -Subject $subject -Since $since -Sender $s.sender
        if($received){ break }
        if((Get-Date) -lt $deadline){ Start-Sleep -Seconds 5 }
    } while((Get-Date) -lt $deadline)

    if($received){ Write-Host "EMAIL TEST RICEVUTA: SI" }
    else { Write-Host "EMAIL TEST RICEVUTA: NO" }

    if($sent -and $received){ exit 0 }
    exit 2
} catch {
    Write-Host "VERIFICA EMAIL TEST: ERRORE"
    Write-Host $_.Exception.Message
    exit 1
}
