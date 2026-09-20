# F1 Outlook Monthly Mailer

Versione 1.0.0. Programma locale Windows per inviare una campagna mensile con Outlook Classic usando esclusivamente l'account F1IMMOBILIARESUSA@OUTLOOK.IT.

## Tecnologia

Solo Windows, Windows PowerShell, Outlook Classic e Windows Task Scheduler. Non usa Supabase, Microsoft Graph, Azure, SMTP, Python, Node.js, Power Automate o servizi di invio esterni.

## Requisiti

- Windows 10 o Windows 11.
- Microsoft Outlook classico installato e configurato.
- L'account F1IMMOBILIARESUSA@OUTLOOK.IT deve essere presente nel profilo Outlook.
- L'utente Windows deve essere connesso al momento dell'esecuzione. L'attività usa InteractiveToken perché Outlook COM lavora nel profilo desktop dell'utente.
- Connessione Internet disponibile per Outlook e per caricare grafica/CTA della mail.

Il nuovo Outlook non espone la stessa automazione COM. Se il test restituisce OUTLOOK_CLASSIC_NECESSARIO, installare/attivare Outlook classico.

## Installazione

1. Scarica la Release v1.0.0.
2. Estrai lo ZIP.
3. Clic destro su Install-F1-Mailer.ps1.
4. Esegui con PowerShell.

L'installer copia l'app in %LOCALAPPDATA%\F1OutlookMonthlyMailer\app, crea settings.json, verifica Outlook Classic e l'account F1, crea la cartella Contatti F1 CAMPAGNA EMAIL, crea le categorie F1-CONSENSO e F1-DISCRITTO e registra l'attività mensile F1 OUTLOOK MONTHLY MAILER.

L'installazione non invia la campagna.

## Preparare i destinatari

Apri Outlook > Persone/Contatti. Inserisci i contatti nella sottocartella F1 CAMPAGNA EMAIL.

Assegna F1-CONSENSO esclusivamente ai contatti per i quali possiedi il consenso appropriato. Un contatto con F1-DISCRITTO non viene mai inviato anche se conserva F1-CONSENSO.

Prima di ogni campagna il programma controlla la Posta in arrivo e, se trova una risposta contenente DISISCRIVIMI, applica F1-DISCRITTO al relativo contatto.

## Test

Esegui Test-Outlook.ps1. Il risultato atteso è OUTLOOK CLASSIC: OK, ACCOUNT F1: OK e CONTATTI F1: OK.

Per inviare una sola prova a F1IMMOBILIARESUSA@OUTLOOK.IT, esegui Test-Email.ps1. Il test non entra nello stato della campagna mensile.

## Campagna mensile

La chiave mensile è F1-AGENT-PRICING-YYYY-MM.

Prima di ogni invio vengono ricontrollati: indirizzo valido, F1-CONSENSO, assenza di F1-DISCRITTO, assenza di SENT/SENDING/UNCERTAIN nella stessa campagna e account F1.

Lo stato locale si trova in %LOCALAPPDATA%\F1OutlookMonthlyMailer\state.json.

Se il processo viene interrotto dopo aver marcato un destinatario SENDING, quel destinatario non viene reinviato automaticamente al riavvio. Questa scelta conservativa evita duplicazioni in un sistema locale privo di transazioni con Outlook.

## Configurazione

File: %LOCALAPPDATA%\F1OutlookMonthlyMailer\settings.json

Valori:
- schedule_day: giorno del mese, da 1 a 28;
- schedule_time: ora HH:mm;
- delay_seconds: pausa tra invii;
- max_emails_per_run: limite per singola esecuzione.

Dopo aver cambiato giorno o ora, riesegui Install-F1-Mailer.ps1 per aggiornare Task Scheduler.

## Pannello

Esegui F1-Control-Panel.ps1.

Pulsanti: TEST OUTLOOK, INVIA TEST, ESEGUI ORA, PAUSA, RIPRENDI, APRI LOG, APRI CONTATTI OUTLOOK, IMPOSTAZIONI.

ESEGUI ORA esegue una campagna reale sui contatti eleggibili.

## Log e report

Log: %LOCALAPPDATA%\F1OutlookMonthlyMailer\logs

Report: %LOCALAPPDATA%\F1OutlookMonthlyMailer\reports

Gli indirizzi nei log vengono mascherati.

## Disinstallazione

Esegui Uninstall-F1-Mailer.ps1. Rimuove l'attività pianificata. I log e lo stato locale non vengono cancellati automaticamente.

## Limiti tecnici

Outlook COM accoda/invia tramite il profilo Outlook locale. Il programma non può garantire una ricevuta SMTP o la consegna nella casella destinataria: Send() conferma che Outlook ha accettato il comando. La verifica finale della prima email test deve essere fatta nella casella Outlook F1.

Nessuna password Microsoft viene letta o salvata dal programma.
