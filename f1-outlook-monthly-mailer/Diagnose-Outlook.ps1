Set-StrictMode -Version Latest
$ErrorActionPreference="Stop"
Import-Module (Join-Path $PSScriptRoot "src\F1-Mailer.Core.psm1") -Force

$sender="F1IMMOBILIARESUSA@OUTLOOK.IT"
try {
    $s=Get-F1Settings
    $sender=[string]$s.sender
} catch {}

$d=Get-F1OutlookDiagnostics -Sender $sender

Write-Host "=================================================="
Write-Host "F1 OUTLOOK DIAGNOSTICA"
Write-Host "=================================================="
Write-Host "WINDOWS USER: $($d.WindowsUser)"
Write-Host "OUTLOOK CLASSIC INSTALLATO: $(if($d.ClassicInstalled){'SI'}else{'NO'})"
Write-Host "PATH OUTLOOK.EXE: $($d.ClassicPath)"
Write-Host "VERSIONE OUTLOOK: $($d.ClassicVersion)"
Write-Host "PROCESSO OUTLOOK CLASSIC ATTIVO: $(if($d.ClassicProcessRunning){'SI'}else{'NO'})"
Write-Host "NUOVO OUTLOOK INSTALLATO: $(if($d.NewOutlookInstalled){'SI'}else{'NO'})"
Write-Host "MAPI DISPONIBILE: $(if($d.MapiAvailable){'SI'}else{'NO'})"
Write-Host "NUMERO ACCOUNT COM: $($d.AccountCount)"
Write-Host ""
foreach($a in $d.Accounts){
    Write-Host "ACCOUNT #$($a.Index)"
    Write-Host "  DisplayName: $($a.DisplayName)"
    Write-Host "  SmtpAddress: $($a.SmtpAddress)"
    Write-Host "  UserName: $($a.UserName)"
    Write-Host "  AccountType: $($a.AccountType)"
    Write-Host "  DeliveryStore: $($a.DeliveryStore)"
    Write-Host "  StoreID: $($a.StoreID)"
}
Write-Host ""
Write-Host "ACCOUNT RICHIESTO: $($d.RequiredAccount)"
Write-Host "ACCOUNT F1 PRESENTE: $(if($d.F1AccountPresent){'SI'}else{'NO'})"
Write-Host "STATO: $($d.Status)"

if($d.F1AccountPresent){ exit 0 }
exit 2
