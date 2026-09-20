Set-StrictMode -Version Latest
$ErrorActionPreference="Stop"
Import-Module (Join-Path $PSScriptRoot "src\F1-Mailer.Core.psm1") -Force
try {
    $s=Get-F1Settings
    $c=Get-F1OutlookContext -Sender $s.sender -ContactsFolderName $s.contacts_folder -CreateFolder -RequiredCategory $s.required_category -BlockedCategory $s.blocked_category
    $template=Get-Content (Join-Path $PSScriptRoot "templates\agent-pricing.html") -Raw -Encoding UTF8
    $html=Render-F1Template -Template $template -FirstName "F1" -LastName "" -Email $s.sender -CampaignKey "TEST-F1-OUTLOOK"
    Send-F1OutlookMail -Context $c -To $s.sender -Subject "TEST F1 — OUTLOOK MONTHLY MAILER" -Html $html
    Write-F1Log -Action "TEST_SENT" -Email $s.sender -Result "SENT"
    Write-Host "EMAIL TEST CONSEGNATA A OUTLOOK PER L'INVIO: $($s.sender)"
    exit 0
} catch {
    Write-F1Log -Action "TEST_SENT" -Result "ERROR" -Error $_.Exception.Message
    Write-Host "EMAIL TEST NON INVIATA: $($_.Exception.Message)"
    exit 1
}
