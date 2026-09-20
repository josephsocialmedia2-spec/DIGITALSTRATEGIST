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

$mockMail=[pscustomobject]@{To="";Subject="";HTMLBody="";SendUsingAccount=$null;Sent=$false}
$mockMail | Add-Member -MemberType ScriptMethod -Name Send -Value { $this.Sent=$true }
$mockApp=[pscustomobject]@{Mail=$mockMail}
$mockApp | Add-Member -MemberType ScriptMethod -Name CreateItem -Value { param($kind) return $this.Mail }
$mockAccount=[pscustomobject]@{SmtpAddress="f1immobiliaresusa@outlook.it"}
$mockContext=[pscustomobject]@{App=$mockApp;Account=$mockAccount}
Send-F1OutlookMail $mockContext "x@example.com" "Subject" "<b>HTML</b>"
Assert-Eq $mockMail.To "x@example.com" "COM mock To"
Assert-Eq $mockMail.Subject "Subject" "COM mock Subject"
Assert-True $mockMail.Sent "COM mock Send"

$s=Get-F1Settings
Assert-Eq $s.sender "F1IMMOBILIARESUSA@OUTLOOK.IT" "settings sender"
Assert-Eq ([int]$s.max_emails_per_run) 100 "settings max"

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
