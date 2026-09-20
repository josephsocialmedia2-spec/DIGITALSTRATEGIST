Set-StrictMode -Version Latest
$ErrorActionPreference="Stop"
$root=Split-Path $PSScriptRoot -Parent
$env:LOCALAPPDATA=Join-Path $PSScriptRoot "tmp-local"
if (Test-Path $env:LOCALAPPDATA) { Remove-Item $env:LOCALAPPDATA -Recurse -Force }
New-Item -ItemType Directory -Path $env:LOCALAPPDATA -Force | Out-Null

Import-Module (Join-Path $root "src\F1-Mailer.Core.psm1") -Force
Initialize-F1LocalStore | Out-Null
Copy-Item (Join-Path $root "settings.example.json") (Get-F1SettingsPath)

$failures=@()
function Assert-True([bool]$Condition,[string]$Name) {
    if ($Condition) { Write-Host "PASS $Name" }
    else { Write-Host "FAIL $Name"; $script:failures += $Name }
}
function Assert-Eq($Actual,$Expected,[string]$Name) {
    Assert-True ($Actual -eq $Expected) "$Name (actual=$Actual expected=$Expected)"
}

Assert-True (Test-F1Email "test@example.com") "valid email"
Assert-True (-not (Test-F1Email "bad@")) "invalid email"
Assert-Eq (Normalize-F1Email " TEST@Example.COM ") "test@example.com" "normalize email"
Assert-Eq (Get-F1CampaignKey ([datetime]"2026-09-20")) "F1-AGENT-PRICING-2026-09" "campaign key"
Assert-True (Test-F1Category "Altro; F1-CONSENSO" "f1-consenso") "required category case insensitive"
Assert-True (-not (Test-F1Category "F1-CONSENSO" "F1-DISCRITTO")) "blocked category absent"
Assert-True ((Add-F1CategoryText "F1-CONSENSO" "F1-DISCRITTO") -like "*F1-DISCRITTO*") "add category"

$state=Get-F1State
$state=Set-F1Record $state "F1-AGENT-PRICING-2026-09" "A@B.IT" "SENT" "" 0
$state=Get-F1State
Assert-True (Test-F1AlreadySent $state "F1-AGENT-PRICING-2026-09" "a@b.it") "anti duplicate"
Assert-True (Test-F1DoNotResend $state "F1-AGENT-PRICING-2026-09" "A@B.IT") "do not resend sent"

$template=Get-Content (Join-Path $root "templates\agent-pricing.html") -Raw
$html=Render-F1Template $template "Mario" "Rossi" "mario@example.com" "F1-AGENT-PRICING-2026-09"
Assert-True ($html -like "*Buongiorno Mario,*") "template greeting"
Assert-True ($html -like "*F1-AGENT-PRICING-2026-09*") "template campaign key"
Assert-True ($html -like "*vendere-casa-f1.jpg*") "template image"
Assert-True ($html -like "*DISISCRIVIMI*") "template unsubscribe"

$f1Store=[pscustomobject]@{DisplayName="f1immobiliaresusa@outlook.it";StoreID="store-f1"}
$f1Account=[pscustomobject]@{
    SmtpAddress="F1IMMOBILIARESUSA@OUTLOOK.IT"
    DisplayName="F1 Immobiliare Susa"
    UserName="f1immobiliaresusa"
    AccountType=0
    DeliveryStore=$f1Store
}
$otherStore=[pscustomobject]@{DisplayName="contatti.realmediapro@outlook.it";StoreID="store-other"}
$otherAccount=[pscustomobject]@{
    SmtpAddress="contatti.realmediapro@outlook.it"
    DisplayName="Real Media Pro"
    UserName="contatti.realmediapro"
    AccountType=0
    DeliveryStore=$otherStore
}

Assert-True (Test-F1AccountMatches $f1Account "f1immobiliaresusa@outlook.it") "F1 account exact"
Assert-True (Test-F1AccountMatches $f1Account "F1IMMOBILIARESUSA@OUTLOOK.IT") "F1 account case insensitive"
Assert-True (-not (Test-F1AccountMatches $otherAccount "f1immobiliaresusa@outlook.it")) "different account rejected"

$accounts=[pscustomobject]@{Count=2;Data=@($otherAccount,$f1Account)}
$accounts | Add-Member -MemberType ScriptMethod -Name Item -Value { param($i) return $this.Data[$i-1] }
$ns=[pscustomobject]@{Accounts=$accounts}
$found=Find-F1OutlookAccount $ns "f1immobiliaresusa@outlook.it"
Assert-True ($null -ne $found) "diagnostic account mock found"
Assert-Eq $found.SmtpAddress "F1IMMOBILIARESUSA@OUTLOOK.IT" "find prioritizes F1"

$accountsMissing=[pscustomobject]@{Count=1;Data=@($otherAccount)}
$accountsMissing | Add-Member -MemberType ScriptMethod -Name Item -Value { param($i) return $this.Data[$i-1] }
$nsMissing=[pscustomobject]@{Accounts=$accountsMissing}
Assert-True ($null -eq (Find-F1OutlookAccount $nsMissing "f1immobiliaresusa@outlook.it")) "missing F1 rejected"

$mockMail=[pscustomobject]@{To="";Subject="";HTMLBody="";SendUsingAccount=$null;Sent=$false}
$mockMail | Add-Member -MemberType ScriptMethod -Name Send -Value { $this.Sent=$true }
$mockApp=[pscustomobject]@{Mail=$mockMail}
$mockApp | Add-Member -MemberType ScriptMethod -Name CreateItem -Value { param($kind) return $this.Mail }
$mockContext=[pscustomobject]@{App=$mockApp;Account=$f1Account;Sender="F1IMMOBILIARESUSA@OUTLOOK.IT"}
Send-F1OutlookMail $mockContext "x@example.com" "Subject" "<b>HTML</b>"
Assert-Eq $mockMail.To "x@example.com" "COM mock To"
Assert-Eq $mockMail.Subject "Subject" "COM mock Subject"
Assert-True $mockMail.Sent "COM mock Send"
Assert-True (Test-F1AccountMatches $mockMail.SendUsingAccount "f1immobiliaresusa@outlook.it") "SendUsingAccount F1"

$badMail=[pscustomobject]@{To="";Subject="";HTMLBody="";SendUsingAccount=$null;Sent=$false}
$badMail | Add-Member -MemberType ScriptMethod -Name Send -Value { $this.Sent=$true }
$badApp=[pscustomobject]@{Mail=$badMail}
$badApp | Add-Member -MemberType ScriptMethod -Name CreateItem -Value { param($kind) return $this.Mail }
$badContext=[pscustomobject]@{App=$badApp;Account=$otherAccount;Sender="F1IMMOBILIARESUSA@OUTLOOK.IT"}
$mismatch=$false
try { Send-F1OutlookMail $badContext "x@example.com" "Bad" "<b>Bad</b>" } catch { $mismatch=$_.Exception.Message -like "*SEND_ACCOUNT_MISMATCH*" }
Assert-True $mismatch "wrong sender blocked"

$s=Get-F1Settings
Assert-Eq $s.sender "F1IMMOBILIARESUSA@OUTLOOK.IT" "settings sender"
Assert-Eq ([int]$s.max_emails_per_run) 100 "settings max"

$taskXml=New-F1TaskXml -User "TEST\User" -AppDir "C:\F1Mailer" -ScheduleDay 1 -ScheduleTime "09:00" -Now ([datetime]"2026-09-20T12:00:00")
Assert-True ($taskXml -like "*<StartWhenAvailable>true</StartWhenAvailable>*") "task StartWhenAvailable"
Assert-True ($taskXml -like "*<CalendarTrigger>*") "task calendar trigger"
Assert-True ($taskXml -like "*<Day>1</Day>*") "task monthly day"
Assert-True ($taskXml -like "*F1-Mailer.ps1*") "task points to mailer"

$parseErrors=@()
Get-ChildItem $root -Recurse -Include *.ps1,*.psm1 | ForEach-Object {
    $tokens=$null; $errors=$null
    [System.Management.Automation.Language.Parser]::ParseFile($_.FullName,[ref]$tokens,[ref]$errors) | Out-Null
    if ($errors.Count -gt 0) { $parseErrors += "$($_.Name): $($errors[0].Message)" }
}
if ($parseErrors.Count -gt 0) { $parseErrors | ForEach-Object { Write-Host "PARSE_ERROR $_" } }
Assert-True ($parseErrors.Count -eq 0) "PowerShell syntax"

Remove-Item $env:LOCALAPPDATA -Recurse -Force -ErrorAction SilentlyContinue
if ($failures.Count -gt 0) {
    Write-Host "FAILED TESTS: $($failures -join ', ')"
    exit 1
}
Write-Host "ALL TESTS PASSED"
