# F1 Outlook 14-Day Mailer v1.1.1

Sistema locale Windows per inviare automaticamente la campagna Agent Pricing ogni 14 giorni tramite Outlook Classic.

## Architettura

Windows Task Scheduler -> PowerShell -> Outlook Classic COM/MAPI -> F1IMMOBILIARESUSA@OUTLOOK.IT -> F1 CAMPAGNA EMAIL -> consenso/disiscrizioni -> anti-duplicazione -> invio -> state.json -> log -> report.

GitHub gestisce codice, test, packaging e release. GitHub Actions non invia email reali.

## Cadenza

La schedulazione usa un CalendarTrigger nativo di Windows Task Scheduler con:

- ScheduleByDay
- DaysInterval = 14
- StartWhenAvailable = true
- first_run_date configurabile
- schedule_time configurabile

Impostazioni di default:

- first_run_date = 2026-09-22
- schedule_time = 09:00
- interval_days = 14

Esempio:

- 22/09/2026 09:00
- 06/10/2026 09:00
- 20/10/2026 09:00
- 03/11/2026 09:00
- 17/11/2026 09:00

Il task mantiene il nome storico F1 OUTLOOK MONTHLY MAILER per aggiornare l'installazione precedente senza creare duplicati. L'interfaccia usa il nome F1 OUTLOOK 14-DAY MAILER.

## Migrazione automatica

L'installer e idempotente.

Conserva:

- settings.json
- state.json
- log
- report

Se un vecchio settings.json non contiene interval_days o first_run_date, viene migrato automaticamente:

- interval_days = 14
- first_run_date = giorno successivo all'aggiornamento
- schedule_time esistente, oppure 09:00 se assente

Il vecchio task viene sovrascritto con lo stesso nome e convertito alla schedulazione a 14 giorni.

## Ciclo campagna

La chiave non e piu mensile.

Formato:

F1-AGENT-PRICING-YYYY-MM-DD

La data e la data pianificata del ciclo di 14 giorni.

Ogni contatto puo ricevere una sola email per ciclo. Gli stati SENT, SENDING e UNCERTAIN impediscono il reinvio automatico nello stesso ciclo.

## Contatti

Account obbligatorio:

F1IMMOBILIARESUSA@OUTLOOK.IT

Cartella:

F1 CAMPAGNA EMAIL

Categorie:

- F1-CONSENSO: autorizzato
- F1-DISCRITTO: blocco assoluto

F1-DISCRITTO prevale sempre su F1-CONSENSO.

Prima di ogni ciclo il programma cerca DISISCRIVIMI nelle risposte e aggiorna il contatto.

## Hotfix 1.1.1

Su alcune installazioni Outlook Classic accetta il setter SendUsingAccount ma il getter restituisce null. Il mailer crea ora il messaggio direttamente nella cartella Bozze dello store F1, assegna comunque l'account F1 e, se il getter resta null, verifica dopo l'invio che il messaggio sia comparso nella Posta inviata dello store F1. Se la conferma manca, lo stato resta UNCERTAIN e il messaggio non viene reinviato automaticamente.

## Invio sicuro

SendUsingAccount deve corrispondere all'account F1.

Il programma:

1. assegna l'account F1;
2. salva il draft;
3. rilegge SendUsingAccount;
4. ritenta una volta l'assegnazione se necessario;
5. blocca l'invio in caso di account differente;
6. se Send() restituisce un errore dopo l'avvio dell'invio, registra UNCERTAIN e non reinvia automaticamente.

## Limiti e batch

max_emails_per_run mantiene il limite operativo di batch, default 100. Se i contatti eleggibili superano il batch, il processo continua automaticamente con i batch successivi nello stesso ciclo senza perdere i rimanenti.

delay_seconds regola la pausa tra un invio e il successivo.

## Test email

Test-Email.ps1 invia una sola email:

Da: F1IMMOBILIARESUSA@OUTLOOK.IT
A: F1IMMOBILIARESUSA@OUTLOOK.IT
Oggetto: TEST F1 - OUTLOOK 14 DAY MAILER

Verify-Test-Email.ps1 verifica:

- Posta inviata F1
- Posta in arrivo F1

La campagna completa non viene eseguita durante installazione o test.

## Pannello

F1-Control-Panel.ps1 mostra:

- versione
- Outlook Classic
- account F1
- account Outlook rilevati
- contatti
- consensi
- disiscritti
- ciclo corrente
- intervallo
- ultima esecuzione
- prossima esecuzione
- giorni alla prossima esecuzione
- inviati/falliti/uncertain
- stato task
- StartWhenAvailable
- pausa

Pulsanti aggiuntivi:

- VERIFICA TASK 14 GIORNI
- MOSTRA PROSSIME ESECUZIONI

## File locali

Configurazione:

%LOCALAPPDATA%\F1OutlookMonthlyMailer\settings.json

Stato:

%LOCALAPPDATA%\F1OutlookMonthlyMailer\state.json

Log:

%LOCALAPPDATA%\F1OutlookMonthlyMailer\logs

Report:

%LOCALAPPDATA%\F1OutlookMonthlyMailer\reports

## Tecnologie

Consentite:

Windows, PowerShell, Outlook Classic COM/MAPI, Windows Task Scheduler.

Non usa Supabase, Microsoft Graph, Azure, SMTP, Python, Node.js, Power Automate, Zapier, Make, SendGrid, Brevo o Mailchimp.
