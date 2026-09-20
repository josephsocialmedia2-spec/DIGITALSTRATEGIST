Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-F1BasePath {
    if (-not $env:LOCALAPPDATA) { throw "LOCALAPPDATA_NOT_AVAILABLE" }
    Join-Path $env:LOCALAPPDATA "F1OutlookMonthlyMailer"
}

function Initialize-F1LocalStore {
    $base = Get-F1BasePath
    foreach ($dir in @($base, (Join-Path $base "logs"), (Join-Path $base "reports"))) {
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    }
    $statePath = Join-Path $base "state.json"
    if (-not (Test-Path $statePath)) {
        @{
            version = 1
            paused = $false
            last_run_at = $null
            last_unsubscribe_scan_at = $null
            records = @()
        } | ConvertTo-Json -Depth 8 | Set-Content -Path $statePath -Encoding UTF8
    }
    return $base
}

function Get-F1SettingsPath { Join-Path (Get-F1BasePath) "settings.json" }
function Get-F1StatePath { Join-Path (Get-F1BasePath) "state.json" }

function Get-F1Settings {
    Initialize-F1LocalStore | Out-Null
    $path = Get-F1SettingsPath
    if (-not (Test-Path $path)) { throw "SETTINGS_NOT_FOUND" }
    $s = Get-Content $path -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($required in @("schedule_day","schedule_time","delay_seconds","max_emails_per_run","contacts_folder","required_category","blocked_category","sender")) {
        if (-not ($s.PSObject.Properties.Name -contains $required)) { throw "SETTINGS_INVALID_$required" }
    }
    return $s
}

function Save-F1JsonAtomic {
    param([Parameter(Mandatory)]$Value,[Parameter(Mandatory)][string]$Path)
    $tmp = "$Path.tmp"
    $Value | ConvertTo-Json -Depth 12 | Set-Content -Path $tmp -Encoding UTF8
    Move-Item -Path $tmp -Destination $Path -Force
}

function Get-F1State {
    Initialize-F1LocalStore | Out-Null
    $path = Get-F1StatePath
    try {
        $s = Get-Content $path -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($null -eq $s.records) { $s | Add-Member -NotePropertyName records -NotePropertyValue @() -Force }
        if (-not ($s.PSObject.Properties.Name -contains "paused")) { $s | Add-Member -NotePropertyName paused -NotePropertyValue $false }
        return $s
    } catch {
        throw "STATE_CORRUPT: $($_.Exception.Message)"
    }
}

function Save-F1State { param($State) Save-F1JsonAtomic -Value $State -Path (Get-F1StatePath) }

function Normalize-F1Email {
    param([string]$Email)
    if ($null -eq $Email) { return "" }
    return $Email.Trim().ToLowerInvariant()
}

function Test-F1Email {
    param([string]$Email)
    $e = Normalize-F1Email $Email
    return [bool]($e -match '^[^@\s]+@[^@\s]+\.[^@\s]+$')
}

function Mask-F1Email {
    param([string]$Email)
    $e = Normalize-F1Email $Email
    if ($e -notmatch '^(.)([^@]*)@(.+)$') { return "***" }
    return "$($Matches[1])***@$($Matches[3])"
}

function Get-F1CampaignKey {
    param([datetime]$Date = (Get-Date))
    return "F1-AGENT-PRICING-{0:yyyy-MM}" -f $Date
}

function Test-F1Category {
    param([string]$Categories,[string]$Category)
    if ([string]::IsNullOrWhiteSpace($Categories)) { return $false }
    $items = $Categories -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    return [bool]($items | Where-Object { $_.Equals($Category,[StringComparison]::OrdinalIgnoreCase) })
}

function Add-F1CategoryText {
    param([string]$Categories,[string]$Category)
    $items = @()
    if (-not [string]::IsNullOrWhiteSpace($Categories)) {
        $items = @($Categories -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    }
    if (-not ($items | Where-Object { $_.Equals($Category,[StringComparison]::OrdinalIgnoreCase) })) {
        $items += $Category
    }
    return ($items -join ", ")
}

function Write-F1Log {
    param([string]$Action,[string]$Email="",[string]$Result="OK",[string]$Error="")
    $base = Initialize-F1LocalStore
    $path = Join-Path (Join-Path $base "logs") ("{0:yyyy-MM-dd}.log" -f (Get-Date))
    $sep = [char]9
    $line = "{0:o}{1}{2}{1}{3}{1}{4}{1}{5}" -f (Get-Date),$sep,$Action,(Mask-F1Email $Email),$Result,($Error -replace "[\r\n]+"," ")
    Add-Content -Path $path -Value $line -Encoding UTF8
}

function Get-F1Record {
    param($State,[string]$CampaignKey,[string]$Email)
    $e = Normalize-F1Email $Email
    return @($State.records | Where-Object {
        $_.campaign_key -eq $CampaignKey -and (Normalize-F1Email $_.email) -eq $e
    }) | Select-Object -First 1
}

function Set-F1Record {
    param($State,[string]$CampaignKey,[string]$Email,[string]$Status,[string]$Error="",[int]$RetryCount=0)
    $e = Normalize-F1Email $Email
    $existing = Get-F1Record -State $State -CampaignKey $CampaignKey -Email $e
    $now = (Get-Date).ToString("o")
    if ($existing) {
        $existing.status = $Status
        $existing.error = $Error
        $existing.retry_count = $RetryCount
        $existing.updated_at = $now
        if ($Status -eq "SENT") { $existing.sent_at = $now }
    } else {
        $record = [pscustomobject]@{
            campaign_key = $CampaignKey
            email = $e
            status = $Status
            sent_at = $(if ($Status -eq "SENT") { $now } else { $null })
            updated_at = $now
            error = $Error
            retry_count = $RetryCount
        }
        $State.records = @($State.records) + $record
    }
    Save-F1State $State
    return $State
}

function Test-F1AlreadySent {
    param($State,[string]$CampaignKey,[string]$Email)
    $r = Get-F1Record -State $State -CampaignKey $CampaignKey -Email $Email
    return [bool]($r -and $r.status -eq "SENT")
}

function Test-F1DoNotResend {
    param($State,[string]$CampaignKey,[string]$Email)
    $r = Get-F1Record -State $State -CampaignKey $CampaignKey -Email $Email
    if (-not $r) { return $false }
    return $r.status -in @("SENT","SENDING","UNCERTAIN")
}

function Get-F1OutlookApplication {
    try {
        return New-Object -ComObject Outlook.Application
    } catch {
        throw "OUTLOOK_CLASSIC_NECESSARIO: impossibile creare Outlook.Application COM. Installare/configurare Outlook classico."
    }
}

function Get-F1OutlookContext {
    param([string]$Sender,[string]$ContactsFolderName,[switch]$CreateFolder,[string]$RequiredCategory="F1-CONSENSO",[string]$BlockedCategory="F1-DISCRITTO")
    $app = Get-F1OutlookApplication
    $ns = $app.GetNamespace("MAPI")
    if ($null -eq $ns) { throw "OUTLOOK_PROFILE_ERROR" }
    $account = $null
    for ($i=1; $i -le $ns.Accounts.Count; $i++) {
        $a = $ns.Accounts.Item($i)
        $smtp = ""
        try { $smtp = [string]$a.SmtpAddress } catch {}
        if ((Normalize-F1Email $smtp) -eq (Normalize-F1Email $Sender)) { $account = $a; break }
    }
    if ($null -eq $account) { throw "ACCOUNT_F1_NON_TROVATO_IN_OUTLOOK" }

    $contactsRoot = $ns.GetDefaultFolder(10)
    $contactsFolder = $null
    for ($i=1; $i -le $contactsRoot.Folders.Count; $i++) {
        $f = $contactsRoot.Folders.Item($i)
        if ([string]$f.Name -eq $ContactsFolderName) { $contactsFolder = $f; break }
    }
    if ($null -eq $contactsFolder -and $CreateFolder) {
        $contactsFolder = $contactsRoot.Folders.Add($ContactsFolderName,10)
    }
    if ($null -eq $contactsFolder) { throw "CARTELLA_CONTATTI_F1_NON_TROVATA" }

    if ($CreateFolder) {
        foreach ($catName in @($RequiredCategory,$BlockedCategory)) {
            $found = $false
            for ($i=1; $i -le $ns.Categories.Count; $i++) {
                if ([string]$ns.Categories.Item($i).Name -eq $catName) { $found = $true; break }
            }
            if (-not $found) {
                try { $null = $ns.Categories.Add($catName) } catch { Write-F1Log -Action "CREATE_CATEGORY" -Result "WARN" -Error $_.Exception.Message }
            }
        }
    }

    [pscustomobject]@{ App=$app; Namespace=$ns; Account=$account; ContactsFolder=$contactsFolder }
}

function Get-F1Contacts {
    param($ContactsFolder)
    $out = @()
    for ($i=1; $i -le $ContactsFolder.Items.Count; $i++) {
        try {
            $item = $ContactsFolder.Items.Item($i)
            if ([string]$item.MessageClass -ne "IPM.Contact") { continue }
            $email = Normalize-F1Email ([string]$item.Email1Address)
            if (-not $email) { continue }
            $out += [pscustomobject]@{
                ComObject = $item
                Email = $email
                FirstName = [string]$item.FirstName
                LastName = [string]$item.LastName
                Categories = [string]$item.Categories
            }
        } catch {
            Write-F1Log -Action "READ_CONTACT" -Result "WARN" -Error $_.Exception.Message
        }
    }
    return $out
}

function Get-F1SenderAddress {
    param($MailItem)
    try {
        $addr = [string]$MailItem.SenderEmailAddress
        if ([string]$MailItem.SenderEmailType -eq "EX") {
            try {
                $ex = $MailItem.Sender.GetExchangeUser()
                if ($ex -and $ex.PrimarySmtpAddress) { $addr = [string]$ex.PrimarySmtpAddress }
            } catch {}
        }
        return Normalize-F1Email $addr
    } catch { return "" }
}

function Sync-F1Unsubscribes {
    param($Context,$Settings,$State)
    $contacts = Get-F1Contacts $Context.ContactsFolder
    $map = @{}
    foreach ($c in $contacts) { $map[$c.Email] = $c }
    $inbox = $Context.Account.DeliveryStore.GetDefaultFolder(6)
    $items = $inbox.Items
    $items.Sort("[ReceivedTime]",$true)
    $cutoff = (Get-Date).AddDays(-90)
    if ($State.last_unsubscribe_scan_at) {
        try { $cutoff = [datetime]::Parse([string]$State.last_unsubscribe_scan_at) } catch {}
    }
    $checked = 0
    for ($i=1; $i -le $items.Count; $i++) {
        if ($checked -ge 5000) { break }
        $checked++
        $m = $items.Item($i)
        try {
            if ($m.ReceivedTime -lt $cutoff) { break }
            $hay = (([string]$m.Subject) + [Environment]::NewLine + ([string]$m.Body)).ToUpperInvariant()
            if ($hay -notlike "*DISISCRIVIMI*") { continue }
            $sender = Get-F1SenderAddress $m
            if (-not $sender) { continue }
            if ($map.ContainsKey($sender)) {
                $c = $map[$sender]
                $c.ComObject.Categories = Add-F1CategoryText -Categories $c.Categories -Category $Settings.blocked_category
                $c.ComObject.Save()
                Write-F1Log -Action "UNSUBSCRIBE" -Email $sender -Result "BLOCKED"
            } else {
                Write-F1Log -Action "UNSUBSCRIBE" -Email $sender -Result "NO_CONTACT"
            }
        } catch {
            Write-F1Log -Action "UNSUBSCRIBE_SCAN" -Result "WARN" -Error $_.Exception.Message
        }
    }
    $State.last_unsubscribe_scan_at = (Get-Date).ToString("o")
    Save-F1State $State
}

function Render-F1Template {
    param([string]$Template,[string]$FirstName,[string]$LastName,[string]$Email,[string]$CampaignKey)
    $name = if ([string]::IsNullOrWhiteSpace($FirstName)) { "" } else { [System.Net.WebUtility]::HtmlEncode($FirstName.Trim()) }
    $greeting = if ($name) { "Buongiorno $name," } else { "Buongiorno," }
    $result = $Template
    $values = @{
        "{{GREETING}}" = $greeting
        "{{NOME}}" = $name
        "{{COGNOME}}" = [System.Net.WebUtility]::HtmlEncode([string]$LastName)
        "{{EMAIL}}" = [System.Net.WebUtility]::HtmlEncode([string]$Email)
        "{{CAMPAIGN_KEY}}" = [System.Net.WebUtility]::HtmlEncode([string]$CampaignKey)
    }
    foreach ($k in $values.Keys) { $result = $result.Replace($k,[string]$values[$k]) }
    return $result
}

function Send-F1OutlookMail {
    param($Context,[string]$To,[string]$Subject,[string]$Html)
    $mail = $Context.App.CreateItem(0)
    $mail.To = $To
    $mail.Subject = $Subject
    $mail.HTMLBody = $Html
    $mail.SendUsingAccount = $Context.Account
    $mail.Send()
}

function Acquire-F1Mutex {
    $created = $false
    $mutex = New-Object System.Threading.Mutex($true,"Local\F1OutlookMonthlyMailer",[ref]$created)
    if (-not $created) {
        $mutex.Dispose()
        throw "MAILER_ALREADY_RUNNING"
    }
    return $mutex
}

function Write-F1Report {
    param([hashtable]$Stats,[string]$CampaignKey,[datetime]$Started,[datetime]$Finished)
    $base = Initialize-F1LocalStore
    $path = Join-Path (Join-Path $base "reports") ("report-{0:yyyy-MM-dd-HHmm}.txt" -f $Finished)
    @"
F1 OUTLOOK MONTHLY MAILER
Campagna: $CampaignKey
Data: $($Started.ToString("yyyy-MM-dd"))
Ora inizio: $($Started.ToString("HH:mm:ss"))
Ora fine: $($Finished.ToString("HH:mm:ss"))
Contatti trovati: $($Stats.contacts)
Consenso valido: $($Stats.authorized)
Disiscritti: $($Stats.unsubscribed)
Invalidi: $($Stats.invalid)
Gia inviati/bloccati da stato: $($Stats.already)
Inviati: $($Stats.sent)
Falliti: $($Stats.failed)
Rimanenti oltre limite: $($Stats.remaining)
"@ | Set-Content -Path $path -Encoding UTF8
    return $path
}

Export-ModuleMember -Function *-F1*
