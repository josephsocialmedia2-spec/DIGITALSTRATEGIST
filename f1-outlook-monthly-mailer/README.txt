F1 OUTLOOK 14-DAY MAILER v1.1.0

OBIETTIVO
Invio automatico Agent Pricing ogni 14 giorni tramite Windows + PowerShell + Outlook Classic COM/MAPI + Windows Task Scheduler.

MITTENTE OBBLIGATORIO
F1IMMOBILIARESUSA@OUTLOOK.IT

TASK
Nome storico mantenuto: F1 OUTLOOK MONTHLY MAILER
Nuova cadenza: ogni 14 giorni
StartWhenAvailable: true

DEFAULT
first_run_date: 2026-09-22
schedule_time: 09:00
interval_days: 14

CICLO
F1-AGENT-PRICING-YYYY-MM-DD

CONTATTI
F1 CAMPAGNA EMAIL
F1-CONSENSO = autorizzato
F1-DISCRITTO = blocco assoluto

TEST
Test-Outlook.ps1
Test-Email.ps1
Verify-Test-Email.ps1
Verify-Task-14Days.ps1
Show-Next-Runs.ps1

OGGETTO TEST
TEST F1 - OUTLOOK 14 DAY MAILER

INSTALLAZIONE
Eseguire Install-F1-Mailer.ps1.
L'installer conserva settings.json, state.json, log e report e aggiorna il vecchio task senza duplicarlo.

CONFIGURAZIONE
%LOCALAPPDATA%\F1OutlookMonthlyMailer\settings.json

STATO
%LOCALAPPDATA%\F1OutlookMonthlyMailer\state.json

LOG
%LOCALAPPDATA%\F1OutlookMonthlyMailer\logs

REPORT
%LOCALAPPDATA%\F1OutlookMonthlyMailer\reports

La campagna reale NON parte durante installazione o test.
