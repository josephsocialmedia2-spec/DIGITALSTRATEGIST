# F1 Outlook Monthly Mailer v1.0.2

Applicazione locale Windows che invia la campagna mensile F1 esclusivamente tramite Outlook Classic e l'account F1IMMOBILIARESUSA@OUTLOOK.IT.

## Architettura

Windows Task Scheduler -> PowerShell -> Outlook Classic COM/MAPI -> account F1 -> contatti dello store F1 -> invio -> stato locale -> report.

Non usa Supabase, Microsoft Graph, Azure, SMTP, Python, Node.js, Power Automate o servizi di invio esterni.

## Novita 1.0.2

Hotfix invio COM: SendUsingAccount viene ora assegnato, il messaggio viene salvato come draft, l'account viene riletto e l'assegnazione viene ritentata una volta prima di bloccare l'invio.

- Diagnose-Outlook.ps1 distingue Nuovo Outlook, Outlook Classic, disponibilita MAPI e account realmente visibili a COM.
- Setup-F1-Outlook-Account.ps1 apre il percorso di gestione profili di Outlook Classic e richiede esclusivamente l'aggiunta dell'account F1.
- ricerca account robusta con priorita a SmtpAddress e confronto case-insensitive;
- nessun fallback su account differenti;
- SendUsingAccount viene controllato prima di Send();
- contatti e categorie vengono cercati nello store associato all'account F1;
- Verify-Test-Email.ps1 verifica Posta inviata F1 e ricezione nella Posta in arrivo F1;
- pannello aggiornato con diagnostica Outlook, task e nuovi pulsanti;
- installer idempotente: conserva settings.json, state.json, log e report.

## Installazione / aggiornamento

Estrai la Release v1.0.2 ed esegui Install-F1-Mailer.ps1.

Se F1IMMOBILIARESUSA@OUTLOOK.IT e presente solo nel Nuovo Outlook, l'installer non invia e avvia Setup-F1-Outlook-Account.ps1. Il requisito valido e che Outlook Classic/MAPI mostri l'account F1.

Password, MFA e credenziali Microsoft non vengono richieste dal programma e devono essere inserite soltanto nelle finestre ufficiali Microsoft.

## Diagnostica

Esegui Diagnose-Outlook.ps1.

Stati principali:

- NEW_OUTLOOK_ONLY
- CLASSIC_OUTLOOK_AVAILABLE
- CLASSIC_OUTLOOK_PROFILE_MISSING_F1
- F1_ACCOUNT_FOUND

La diagnostica mostra utente Windows, percorso e versione di Outlook Classic, processo attivo, disponibilita MAPI, numero account e per ogni account DisplayName, SmtpAddress, UserName, AccountType, DeliveryStore e StoreID quando disponibili.

## Test Outlook

Test-Outlook.ps1 deve restituire:

- OUTLOOK CLASSIC: OK
- ACCOUNT F1: OK
- CONTATTI F1: OK
- F1-CONSENSO: OK
- F1-DISCRITTO: OK

## Test email reale

Test-Email.ps1 invia una sola email:

Da: F1IMMOBILIARESUSA@OUTLOOK.IT
A: F1IMMOBILIARESUSA@OUTLOOK.IT
Oggetto: TEST F1 - OUTLOOK MONTHLY MAILER

Successivamente esegui Verify-Test-Email.ps1. Il progetto considera il test completato soltanto se trova il messaggio nella Posta inviata dello store F1 e nella Posta in arrivo F1.

## Contatti e consenso

Cartella contatti dello store F1:

F1 CAMPAGNA EMAIL

Categorie:

- F1-CONSENSO: contatto autorizzato;
- F1-DISCRITTO: blocco assoluto.

## Task mensile

Task:

F1 OUTLOOK MONTHLY MAILER

L'installer registra o aggiorna la stessa attivita senza duplicarla. E configurata con StartWhenAvailable=true.

Configurazione locale:

%LOCALAPPDATA%\F1OutlookMonthlyMailer\settings.json

Parametri:

- schedule_day
- schedule_time
- delay_seconds
- max_emails_per_run

## Stato, log e report

Stato anti-duplicazione:

%LOCALAPPDATA%\F1OutlookMonthlyMailer\state.json

Log:

%LOCALAPPDATA%\F1OutlookMonthlyMailer\logs

Report:

%LOCALAPPDATA%\F1OutlookMonthlyMailer\reports

## Pannello

F1-Control-Panel.ps1

Include:

- DIAGNOSTICA OUTLOOK
- CONFIGURA ACCOUNT F1
- TEST OUTLOOK
- INVIA TEST
- VERIFICA EMAIL TEST
- ESEGUI ORA
- PAUSA
- RIPRENDI
- APRI LOG
- APRI CONTATTI F1
- IMPOSTAZIONI

La campagna reale non viene inviata durante installazione o aggiornamento.
