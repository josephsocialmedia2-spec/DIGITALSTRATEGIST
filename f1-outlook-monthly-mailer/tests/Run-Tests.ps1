Set-StrictMode -Version Latest
$ErrorActionPreference="Stop"
$root=Split-Path $PSScriptRoot -Parent
$env:LOCALAPPDATA=Join-Path $PSScriptRoot "tmp-local"
if(Test-Path $env:LOCALAPPDATA){Remove-Item $env:LOCALAPPDATA -Recurse -Force}
New-Item -ItemType Directory -Path $env:LOCALAPPDATA -Force | Out-Null

Import-Module (Join-Path $root "src\F1-Mailer.Core.psm1") -Force
Initialize-F1LocalStore | Out-Null
Copy-Item (Join-Path $root "settings.example.json") (Get-F1SettingsPath)

$failures=@()
function Assert-True([bool]$c,[string]$n){
  if($c){Write-Host "PASS $n"}else{Write-Host "FAIL $n";$script:failures+=$n}
}
function Assert-Eq($a,$e,[string]$n){Assert-True ($a -eq $e) "$n (actual=$a expected=$e)"}

Assert-True (Test-F1Email "test@example.com") "valid email"
Assert-True (-not (Test-F1Email "bad@")) "invalid email"
Assert-Eq (Normalize-F1Email " TEST@Example.COM ") "test@example.com" "normalize"

$anchor="2026-09-22"
Assert-Eq (Get-F1CampaignKey ([datetime]"2026-09-22")) "F1-AGENT-PRICING-2026-09-22" "campaign key"
$runs=@(Get-F1UpcomingRunDates -FirstRunDate $anchor -ScheduleTime "09:00" -IntervalDays 14 -From ([datetime]"2026-09-21T12:00:00") -Count 5)
$expected=@("2026-09-22 09:00","2026-10-06 09:00","2026-10-20 09:00","2026-11-03 09:00","2026-11-17 09:00")
for($i=0;$i -lt $expected.Count;$i++){Assert-Eq ($runs[$i].ToString("yyyy-MM-dd HH:mm")) $expected[$i] "run $($i+1)"}

Assert-True (Test-F1Category "F1-CONSENSO" "f1-consenso") "consent case insensitive"
Assert-True (Test-F1Category "F1-CONSENSO;F1-DISCRITTO" "F1-DISCRITTO") "blocked category present"

$state=Get-F1State
$key="F1-AGENT-PRICING-2026-09-22"
$state=Set-F1Record $state $key "sent@example.com" "SENT"
$state=Set-F1Record $state $key "sending@example.com" "SENDING"
$state=Set-F1Record $state $key "uncertain@example.com" "UNCERTAIN"
$state=Get-F1State
Assert-True (Test-F1DoNotResend $state $key "sent@example.com") "no resend SENT"
Assert-True (Test-F1DoNotResend $state $key "sending@example.com") "no resend SENDING"
Assert-True (Test-F1DoNotResend $state $key "uncertain@example.com") "no resend UNCERTAIN"

$template=Get-Content (Join-Path $root "templates\agent-pricing.html") -Raw
$html=Render-F1Template $template "Mario" "Rossi" "mario@example.com" $key
Assert-True ($html -like "*Buongiorno Mario,*") "template greeting"
Assert-True ($html -like "*$key*") "template cycle"
Assert-True ($html -like "*DISISCRIVIMI*") "unsubscribe text"

$f1Store=[pscustomobject]@{DisplayName="f1immobiliaresusa@outlook.it";StoreID="store-f1"}
$f1Account=[pscustomobject]@{SmtpAddress="F1IMMOBILIARESUSA@OUTLOOK.IT";DisplayName="F1";UserName="f1";AccountType=0;DeliveryStore=$f1Store}
$otherAccount=[pscustomobject]@{SmtpAddress="contatti.realmediapro@outlook.it";DisplayName="Real";UserName="real";AccountType=0;DeliveryStore=[pscustomobject]@{DisplayName="real";StoreID="other"}}
Assert-True (Test-F1AccountMatches $f1Account "f1immobiliaresusa@outlook.it") "account F1"
Assert-True (-not (Test-F1AccountMatches $otherAccount "f1immobiliaresusa@outlook.it")) "wrong account rejected"

$accounts=[pscustomobject]@{Count=2;Data=@($otherAccount,$f1Account)}
$accounts | Add-Member ScriptMethod Item {param($i)$this.Data[$i-1]}
$ns=[pscustomobject]@{Accounts=$accounts}
Assert-Eq (Find-F1OutlookAccount $ns "f1immobiliaresusa@outlook.it").SmtpAddress "F1IMMOBILIARESUSA@OUTLOOK.IT" "find F1"

$sentItem=[pscustomobject]@{Subject="S";To="x@example.com";SentOn=(Get-Date)}
Assert-True (Test-F1MailRecipientMatches $sentItem "x@example.com") "recipient match"
$sentItems=[pscustomobject]@{Count=1;Data=@($sentItem)}
$sentItems | Add-Member ScriptMethod Sort {param($p,$d)}
$sentItems | Add-Member ScriptMethod Item {param($i)$this.Data[$i-1]}
$sentFolder=[pscustomobject]@{Items=$sentItems}
Assert-True ($null -ne (Find-F1SentMail -Folder $sentFolder -Subject "S" -To "x@example.com" -Since (Get-Date).AddMinutes(-1))) "sent item verification"

$mail=[pscustomobject]@{To="";Subject="";HTMLBody="";SendUsingAccount=$null;Sent=$false}
$mail | Add-Member ScriptMethod Save {}
$mail | Add-Member ScriptMethod Send {$this.Sent=$true}
$draftItems=[pscustomobject]@{Mail=$mail}
$draftItems | Add-Member ScriptMethod Add {param($kind)$this.Mail}
$draftFolder=[pscustomobject]@{Items=$draftItems}
$store=[pscustomobject]@{Drafts=$draftFolder}
$store | Add-Member ScriptMethod GetDefaultFolder {param($id)if($id -eq 16){$this.Drafts}else{$null}}
$ctx=[pscustomobject]@{Account=$f1Account;Store=$store;Sender="F1IMMOBILIARESUSA@OUTLOOK.IT"}
$result=Send-F1OutlookMail $ctx "x@example.com" "Subject" "<b>HTML</b>"
Assert-True $mail.Sent "mock send"
Assert-Eq $result.Verification "SENDUSINGACCOUNT" "getter verification"

$bad=$false
try{Send-F1OutlookMail ([pscustomobject]@{Account=$otherAccount;Store=$store;Sender="F1IMMOBILIARESUSA@OUTLOOK.IT"}) "x@example.com" "Bad" "x"}catch{$bad=$_.Exception.Message -like "*SEND_ACCOUNT_MISMATCH*"}
Assert-True $bad "SEND_ACCOUNT_MISMATCH"

$s=Get-F1Settings
Assert-Eq ([int]$s.interval_days) 14 "settings interval"
Assert-Eq $s.first_run_date "2026-09-22" "first run"
Assert-Eq $s.schedule_time "09:00" "schedule time"

$xml=New-F1TaskXml -User "TEST\User" -AppDir "C:\F1Mailer" -FirstRunDate "2026-09-22" -ScheduleTime "09:00" -IntervalDays 14
Assert-True ($xml -like "*<DaysInterval>14</DaysInterval>*") "task 14 days"
Assert-True ($xml -like "*<StartWhenAvailable>true</StartWhenAvailable>*") "start when available"
Assert-True ($xml -like "*<StartBoundary>2026-09-22T09:00:00</StartBoundary>*") "start boundary"
Assert-True (-not ($xml -like "*<ScheduleByMonth>*")) "monthly trigger removed"
Assert-True ($xml -like "*F1-Mailer.ps1*") "task target"

$parseErrors=@()
Get-ChildItem $root -Recurse -Include *.ps1,*.psm1 | ForEach-Object {
  $tokens=$null;$errors=$null
  [System.Management.Automation.Language.Parser]::ParseFile($_.FullName,[ref]$tokens,[ref]$errors)|Out-Null
  if($errors.Count -gt 0){$parseErrors+="$($_.Name): $($errors[0].Message)"}
}
$parseErrors|ForEach-Object{Write-Host "PARSE_ERROR $_"}
Assert-True ($parseErrors.Count -eq 0) "PowerShell syntax"

Remove-Item $env:LOCALAPPDATA -Recurse -Force -ErrorAction SilentlyContinue
if($failures.Count -gt 0){Write-Host "FAILED TESTS: $($failures -join ', ')";exit 1}
Write-Host "ALL TESTS PASSED"
