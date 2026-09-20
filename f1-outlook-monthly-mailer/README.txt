F1 OUTLOOK MONTHLY MAILER v1.0.1

TECNOLOGIA
Windows + PowerShell + Outlook Classic COM/MAPI + Windows Task Scheduler.
Non usa Supabase, Graph, Azure, SMTP, Python, Node.js o Power Automate.

MITTENTE OBBLIGATORIO
F1IMMOBILIARESUSA@OUTLOOK.IT
Nessun fallback su altri account.

INSTALLAZIONE / AGGIORNAMENTO
1. Estrai lo ZIP.
2. Esegui Install-F1-Mailer.ps1 con PowerShell.
3. L'installer conserva settings.json, state.json, log e report esistenti.
4. Se F1 manca dal profilo MAPI, viene avviato Setup-F1-Outlook-Account.ps1.

DIAGNOSTICA
Diagnose-Outlook.ps1 distingue Outlook Classic, Nuovo Outlook e account MAPI.
Stati principali:
NEW_OUTLOOK_ONLY
CLASSIC_OUTLOOK_AVAILABLE
CLASSIC_OUTLOOK_PROFILE_MISSING_F1
F1_ACCOUNT_FOUND

TEST
Test-Outlook.ps1 verifica Outlook Classic, account F1, cartella contatti e categorie.
Test-Email.ps1 invia UNA sola email da F1 a F1.
Verify-Test-Email.ps1 controlla Posta inviata e Posta in arrivo F1.

CONTATTI
Cartella dello store F1: F1 CAMPAGNA EMAIL
Categoria necessaria: F1-CONSENSO
Categoria di blocco: F1-DISCRITTO

CONFIGURAZIONE
%LOCALAPPDATA%\F1OutlookMonthlyMailer\settings.json

LOG
%LOCALAPPDATA%\F1OutlookMonthlyMailer\logs

REPORT
%LOCALAPPDATA%\F1OutlookMonthlyMailer\reports

PANNELLO
F1-Control-Panel.ps1

La campagna reale non parte durante l'installazione.
