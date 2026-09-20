Set-StrictMode -Version Latest
$ErrorActionPreference="Stop"
Import-Module (Join-Path $PSScriptRoot "src\F1-Mailer.Core.psm1") -Force
try {
    $s=Get-F1Settings
    $d=Get-F1OutlookDiagnostics -Sender $s.sender
    if(-not $d.ClassicInstalled){ throw "OUTLOOK_CLASSIC_NECESSARIO" }
    if(-not $d.MapiAvailable){ throw "OUTLOOK_PROFILE_ERROR" }
    if(-not $d.F1AccountPresent){ throw "CLASSIC_OUTLOOK_PROFILE_MISSING_F1" }
    $c=Get-F1OutlookContext -Sender $s.sender -ContactsFolderName $s.contacts_folder -CreateFolder -RequiredCategory $s.required_category -BlockedCategory $s.blocked_category
    $count=@(Get-F1Contacts $c.ContactsFolder).Count
    Write-Host "OUTLOOK CLASSIC: OK"
    Write-Host "ACCOUNT F1: OK - $($s.sender)"
    Write-Host "CONTATTI F1: OK - $count contatti"
    Write-Host "F1-CONSENSO: OK"
    Write-Host "F1-DISCRITTO: OK"
    exit 0
} catch {
    Write-Host "TEST OUTLOOK: ERRORE"
    Write-Host $_.Exception.Message
    exit 1
}
