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

function Get-F1ClassicOutlookPath {
    $candidates = New-Object System.Collections.Generic.List[string]
    foreach ($regPath in @(
        "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE",
        "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE",
        "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"
    )) {
        try {
            $key = Get-Item -LiteralPath $regPath -ErrorAction Stop
            $value = [string]$key.GetValue("")
            if ($value) { $candidates.Add($value.Trim('"')) }
        } catch {}
    }
    $roots=@($env:ProgramFiles,[Environment]::GetEnvironmentVariable("ProgramFiles(x86)"))
    foreach ($root in $roots) {
        if ($root) {
            $candidates.Add((Join-Path $root "Microsoft Office\root\Office16\OUTLOOK.EXE"))
            $candidates.Add((Join-Path $root "Microsoft Office\Office16\OUTLOOK.EXE"))
        }
    }
    try {
        $cmd = Get-Command outlook.exe -ErrorAction Stop
        if ($cmd.Source) { $candidates.Add([string]$cmd.Source) }
    } catch {}
    foreach ($p in ($candidates | Select-Object -Unique)) {
        if ($p -and (Test-Path -LiteralPath $p)) { return (Resolve-Path -LiteralPath $p).Path }
    }
    return $null
}

function Test-F1NewOutlookInstalled {
    try {
        $pkg = Get-AppxPackage -Name "Microsoft.OutlookForWindows" -ErrorAction SilentlyContinue
        if ($pkg) { return $true }
    } catch {}
    try {
        if (Get-Process -Name "olk" -ErrorAction SilentlyContinue) { return $true }
    } catch {}
    return $false
}

function Get-F1OutlookApplication {
    try {
        return New-Object -ComObject Outlook.Application
    } catch {
        throw "OUTLOOK_CLASSIC_NECESSARIO: impossibile creare Outlook.Application COM."
    }
}

function Get-F1OutlookAccountSnapshot {
    param($Namespace)
    $result = @()
    if ($null -eq $Namespace) { return $result }
    for ($i=1; $i -le $Namespace.Accounts.Count; $i++) {
        $a = $Namespace.Accounts.Item($i)
        $displayName=""; $smtp=""; $userName=""; $accountType=""; $storeName=""; $storeId=""
        try { $displayName=[string]$a.DisplayName } catch {}
        try { $smtp=[string]$a.SmtpAddress } catch {}
        try { $userName=[string]$a.UserName } catch {}
        try { $accountType=[string]$a.AccountType } catch {}
        try { $storeName=[string]$a.DeliveryStore.DisplayName } catch {}
        try { $storeId=[string]$a.DeliveryStore.StoreID } catch {}
        $result += [pscustomobject]@{
            Index=$i
            ComObject=$a
            DisplayName=$displayName
            SmtpAddress=$smtp
            UserName=$userName
            AccountType=$accountType
            DeliveryStore=$storeName
            StoreID=$storeId
        }
    }
    return $result
}

function Get-F1AccountIdentityValues {
    param($Account)
    $values = New-Object System.Collections.Generic.List[string]
    foreach ($prop in @("SmtpAddress","DisplayName","UserName")) {
        try {
            $v = [string]$Account.$prop
            if ($v) { $values.Add($v) }
        } catch {}
    }
    try {
        $v=[string]$Account.DeliveryStore.DisplayName
        if ($v) { $values.Add($v) }
    } catch {}
    return @($values | ForEach-Object { Normalize-F1Email $_ } | Where-Object { $_ } | Select-Object -Unique)
}

function Test-F1AccountMatches {
    param($Account,[string]$Sender)
    if ($null -eq $Account) { return $false }
    $target = Normalize-F1Email $Sender
    try {
        $smtp = Normalize-F1Email ([string]$Account.SmtpAddress)
        if ($smtp -and $smtp -eq $target) { return $true }
    } catch {}
    return [bool](Get-F1AccountIdentityValues $Account | Where-Object { $_ -eq $target })
}

function Find-F1OutlookAccount {
    param($Namespace,[string]$Sender)
    if ($null -eq $Namespace) { return $null }
    $target = Normalize-F1Email $Sender
    for ($i=1; $i -le $Namespace.Accounts.Count; $i++) {
        $a=$Namespace.Accounts.Item($i)
        try {
            if ((Normalize-F1Email ([string]$a.SmtpAddress)) -eq $target) { return $a }
        } catch {}
    }
    for ($i=1; $i -le $Namespace.Accounts.Count; $i++) {
        $a=$Namespace.Accounts.Item($i)
        if (Test-F1AccountMatches -Account $a -Sender $Sender) { return $a }
    }
    return $null
}

function Get-F1OutlookDiagnostics {
    param([string]$Sender="F1IMMOBILIARESUSA@OUTLOOK.IT")
    $classicPath = Get-F1ClassicOutlookPath
    $newInstalled = Test-F1NewOutlookInstalled
    $classicProcess = [bool](Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue)
    $version=""
    if ($classicPath) {
        try { $version=(Get-Item -LiteralPath $classicPath).VersionInfo.FileVersion } catch {}
    }
    $mapi=$false
    $accounts=@()
    $f1Found=$false
    try {
        $app=Get-F1OutlookApplication
        $ns=$app.GetNamespace("MAPI")
        if ($ns) {
            $mapi=$true
            $accounts=@(Get-F1OutlookAccountSnapshot $ns)
            $f1Found=[bool](Find-F1OutlookAccount -Namespace $ns -Sender $Sender)
            if (-not $classicPath) { $classicPath="COM disponibile; percorso EXE non risolto" }
        }
    } catch {}
    $classicInstalled=[bool]$classicPath
    $status = if ($f1Found) {
        "F1_ACCOUNT_FOUND"
    } elseif ($classicInstalled -and $mapi) {
        "CLASSIC_OUTLOOK_PROFILE_MISSING_F1"
    } elseif ($classicInstalled) {
        "CLASSIC_OUTLOOK_AVAILABLE"
    } elseif ($newInstalled) {
        "NEW_OUTLOOK_ONLY"
    } else {
        "OUTLOOK_CLASSIC_NOT_FOUND"
    }
    [pscustomobject]@{
        WindowsUser=[Security.Principal.WindowsIdentity]::GetCurrent().Name
        ClassicInstalled=$classicInstalled
        ClassicPath=$classicPath
        ClassicVersion=$version
        ClassicProcessRunning=$classicProcess
        NewOutlookInstalled=$newInstalled
        MapiAvailable=$mapi
        AccountCount=$accounts.Count
        Accounts=$accounts
        F1AccountPresent=$f1Found
        RequiredAccount=$Sender
        Status=$status
    }
}

function Get-F1OutlookContext {
    param([string]$Sender,[string]$ContactsFolderName,[switch]$CreateFolder,[string]$RequiredCategory="F1-CONSENSO",[string]$BlockedCategory="F1-DISCRITTO")
    $app = Get-F1OutlookApplication
    $ns = $app.GetNamespace("MAPI")
    if ($null -eq $ns) { throw "OUTLOOK_PROFILE_ERROR" }
    $account = Find-F1OutlookAccount -Namespace $ns -Sender $Sender
    if ($null -eq $account) {
        $detected=@(Get-F1OutlookAccountSnapshot $ns | ForEach-Object {
            if ($_.SmtpAddress) { $_.SmtpAddress } elseif ($_.DisplayName) { $_.DisplayName } else { "account#$($_.Index)" }
        })
        throw "CLASSIC_OUTLOOK_PROFILE_MISSING_F1: richiesto=$Sender; rilevati=$($detected -join ', ')"
    }

    $store=$null
    try { $store=$account.DeliveryStore } catch {}
    if ($null -eq $store) { throw "F1_DELIVERY_STORE_NOT_AVAILABLE" }
    try { $contactsRoot=$store.GetDefaultFolder(10) } catch { throw "F1_STORE_CONTACTS_UNAVAILABLE" }
    if ($null -eq $contactsRoot) { throw "F1_STORE_CONTACTS_UNAVAILABLE" }

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
        $categories=$null
        try { $categories=$store.Categories } catch {}
        if ($null -eq $categories) { throw "F1_STORE_CATEGORIES_UNAVAILABLE" }
        foreach ($catName in @($RequiredCategory,$BlockedCategory)) {
            $found = $false
            for ($i=1; $i -le $categories.Count; $i++) {
                if ([string]$categories.Item($i).Name -eq $catName) { $found = $true; break }
            }
            if (-not $found) {
                try { $null = $categories.Add($catName) } catch { throw "CREATE_CATEGORY_FAILED_$catName" }
            }
        }
    }

    [pscustomobject]@{
        App=$app
        Namespace=$ns
        Account=$account
        Store=$store
        ContactsFolder=$contactsFolder
        Sender=$Sender
    }
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
    $inbox = $Context.Store.GetDefaultFolder(6)
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
    if (-not (Test-F1AccountMatches -Account $Context.Account -Sender $Context.Sender)) {
        throw "SEND_ACCOUNT_MISMATCH"
    }
    $mail = $Context.App.CreateItem(0)
    $mail.To = $To
    $mail.Subject = $Subject
    $mail.HTMLBody = $Html
    $mail.SendUsingAccount = $Context.Account
    if (-not (Test-F1AccountMatches -Account $mail.SendUsingAccount -Sender $Context.Sender)) {
        throw "SEND_ACCOUNT_MISMATCH"
    }
    $mail.Send()
}

function Find-F1MailBySubject {
    param($Folder,[string]$Subject,[datetime]$Since,[string]$Sender="")
    if ($null -eq $Folder) { return $null }
    $items=$Folder.Items
    try { $items.Sort("[ReceivedTime]",$true) } catch {
        try { $items.Sort("[SentOn]",$true) } catch {}
    }
    $limit=[Math]::Min([int]$items.Count,500)
    for ($i=1; $i -le $limit; $i++) {
        $m=$items.Item($i)
        try {
            $when=$null
            try { $when=[datetime]$m.ReceivedTime } catch {}
            if (-not $when) { try { $when=[datetime]$m.SentOn } catch {} }
            if ($when -and $when -lt $Since) { continue }
            if ([string]$m.Subject -ne $Subject) { continue }
            if ($Sender) {
                if ((Get-F1SenderAddress $m) -ne (Normalize-F1Email $Sender)) { continue }
            }
            return $m
        } catch {}
    }
    return $null
}

function Get-F1TaskStatus {
    $task=$null; $info=$null
    try { $task=Get-ScheduledTask -TaskName "F1 OUTLOOK MONTHLY MAILER" -ErrorAction Stop } catch {}
    if ($task) { try { $info=Get-ScheduledTaskInfo -TaskName "F1 OUTLOOK MONTHLY MAILER" -ErrorAction Stop } catch {} }
    [pscustomobject]@{
        Exists=[bool]$task
        State=$(if($task){[string]$task.State}else{"MISSING"})
        Enabled=$(if($task){[bool]$task.Settings.Enabled}else{$false})
        StartWhenAvailable=$(if($task){[bool]$task.Settings.StartWhenAvailable}else{$false})
        LastRunTime=$(if($info){$info.LastRunTime}else{$null})
        NextRunTime=$(if($info){$info.NextRunTime}else{$null})
        LastTaskResult=$(if($info){$info.LastTaskResult}else{$null})
    }
}

function New-F1TaskXml {
    param(
        [string]$User,
        [string]$AppDir,
        [int]$ScheduleDay,
        [string]$ScheduleTime,
        [datetime]$Now=(Get-Date)
    )
    if ($ScheduleDay -lt 1 -or $ScheduleDay -gt 28) { throw "schedule_day deve essere compreso tra 1 e 28" }
    if ($ScheduleTime -notmatch '^([01]\d|2[0-3]):[0-5]\d$') { throw "schedule_time deve essere HH:mm" }
    $time=[datetime]::ParseExact($ScheduleTime,"HH:mm",[Globalization.CultureInfo]::InvariantCulture)
    $start=Get-Date -Year $Now.Year -Month $Now.Month -Day $ScheduleDay -Hour $time.Hour -Minute $time.Minute -Second 0
    if ($start -le $Now) { $start=$start.AddMonths(1) }
    $main=Join-Path $AppDir "F1-Mailer.ps1"
    $arg='-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}"' -f $main
    $xmlArg=[Security.SecurityElement]::Escape($arg)
    $xmlWork=[Security.SecurityElement]::Escape($AppDir)
    $xmlUser=[Security.SecurityElement]::Escape($User)
    $startText=$start.ToString("yyyy-MM-ddTHH:mm:ss")
    $months="<January/><February/><March/><April/><May/><June/><July/><August/><September/><October/><November/><December/>"
    return @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo><Description>Invio mensile F1 tramite Outlook Classic.</Description></RegistrationInfo>
  <Triggers>
    <CalendarTrigger>
      <StartBoundary>$startText</StartBoundary>
      <Enabled>true</Enabled>
      <ScheduleByMonth>
        <DaysOfMonth><Day>$ScheduleDay</Day></DaysOfMonth>
        <Months>$months</Months>
      </ScheduleByMonth>
    </CalendarTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>$xmlUser</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <ExecutionTimeLimit>PT4H</ExecutionTimeLimit>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>powershell.exe</Command>
      <Arguments>$xmlArg</Arguments>
      <WorkingDirectory>$xmlWork</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"@
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
