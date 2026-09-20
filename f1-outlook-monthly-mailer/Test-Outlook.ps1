Set-StrictMode -Version Latest
$ErrorActionPreference="Stop"
Import-Module (Join-Path $PSScriptRoot "src\F1-Mailer.Core.psm1") -Force
try {
    $s=Get-F1Settings
    $c=Get-F1OutlookContext -Sender $s.sender -ContactsFolderName $s.contacts_folder -CreateFolder -RequiredCategory $s.required_category -BlockedCategory $s.blocked_category
    $count=@(Get-F1Contacts $c.ContactsFolder).Count
    Write-Host "OUTLOOK CLASSIC: OK"
    Write-Host "ACCOUNT F1: OK - $($s.sender)"
    Write-Host "CONTATTI F1: OK - $count contatti"
    exit 0
} catch {
    Write-Host "TEST OUTLOOK: ERRORE"
    Write-Host $_.Exception.Message
    exit 1
}
