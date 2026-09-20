$ErrorActionPreference="SilentlyContinue"
Unregister-ScheduledTask -TaskName "F1 OUTLOOK MONTHLY MAILER" -Confirm:$false
Write-Host "Attivita pianificata rimossa."
Write-Host "I dati locali e i log restano in: $env:LOCALAPPDATA\F1OutlookMonthlyMailer"
