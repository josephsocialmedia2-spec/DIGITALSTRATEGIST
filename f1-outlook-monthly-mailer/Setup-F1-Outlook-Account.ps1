Set-StrictMode -Version Latest
$ErrorActionPreference="Stop"
Import-Module (Join-Path $PSScriptRoot "src\F1-Mailer.Core.psm1") -Force

$sender="F1IMMOBILIARESUSA@OUTLOOK.IT"
try { $sender=[string](Get-F1Settings).sender } catch {}

$d=Get-F1OutlookDiagnostics -Sender $sender
if($d.F1AccountPresent){
    Write-Host "ACCOUNT F1 GIA PRESENTE IN OUTLOOK CLASSIC: $sender"
    exit 0
}

Write-Host "=================================================="
Write-Host "CONFIGURAZIONE OUTLOOK F1 NECESSARIA"
Write-Host "=================================================="
Write-Host "Account richiesto: $sender"
Write-Host "Stato rilevato: $($d.Status)"
Write-Host ""
if($d.AccountCount -gt 0){
    Write-Host "Account attualmente rilevati da Outlook Classic:"
    $d.Accounts | ForEach-Object {
        $v=if($_.SmtpAddress){$_.SmtpAddress}else{$_.DisplayName}
        Write-Host " - $v"
    }
}
Write-Host ""

if(-not $d.ClassicInstalled){
    Write-Host "OUTLOOK CLASSIC NECESSARIO"
    Write-Host "Il Nuovo Outlook non puo essere usato da questo programma tramite COM/MAPI."
    exit 3
}

$classicPath=$d.ClassicPath
if($classicPath -like "COM disponibile*"){
    $classicPath=Get-F1ClassicOutlookPath
}

$opened=$false
if($classicPath -and (Test-Path -LiteralPath $classicPath)){
    try {
        Start-Process -FilePath $classicPath -ArgumentList "/manageprofiles"
        $opened=$true
    } catch {}
    if(-not $opened){
        $cpl=Join-Path (Split-Path $classicPath -Parent) "MLCFG32.CPL"
        if(Test-Path -LiteralPath $cpl){
            try {
                Start-Process -FilePath "control.exe" -ArgumentList ('"{0}"' -f $cpl)
                $opened=$true
            } catch {}
        }
    }
    if(-not $opened){
        try {
            Start-Process -FilePath $classicPath
            $opened=$true
        } catch {}
    }
}

Write-Host "AZIONE:"
Write-Host "Aggiungi al profilo di OUTLOOK CLASSIC l'account:"
Write-Host $sender
Write-Host ""
Write-Host "Password e verifica MFA devono essere inserite solo nelle finestre ufficiali Microsoft."
Write-Host "Quando l'account e stato aggiunto, chiudi e riapri Outlook Classic."
Read-Host "Poi premi INVIO qui per ricontrollare"

$d2=Get-F1OutlookDiagnostics -Sender $sender
if($d2.F1AccountPresent){
    Write-Host "ACCOUNT F1 TROVATO IN OUTLOOK CLASSIC: $sender"
    exit 0
}

Write-Host "ACCOUNT F1 ANCORA NON PRESENTE NEL PROFILO MAPI."
Write-Host "Stato: $($d2.Status)"
Write-Host "Account COM rilevati:"
$d2.Accounts | ForEach-Object {
    $v=if($_.SmtpAddress){$_.SmtpAddress}else{$_.DisplayName}
    Write-Host " - $v"
}
exit 2
