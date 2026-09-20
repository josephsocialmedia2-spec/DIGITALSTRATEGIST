param([switch]$DryRun)
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
Import-Module (Join-Path $PSScriptRoot "src\F1-Mailer.Core.psm1") -Force

$mutex = $null
$started = Get-Date
$campaign = Get-F1CampaignKey
$stats = @{
    contacts=0; authorized=0; unsubscribed=0; invalid=0; already=0
    sent=0; failed=0; remaining=0
}

try {
    $mutex = Acquire-F1Mutex
    $settings = Get-F1Settings
    $state = Get-F1State
    if ($state.paused -eq $true) {
        Write-F1Log -Action "RUN" -Result "PAUSED"
        exit 0
    }

    $context = Get-F1OutlookContext -Sender $settings.sender -ContactsFolderName $settings.contacts_folder -RequiredCategory $settings.required_category -BlockedCategory $settings.blocked_category
    Sync-F1Unsubscribes -Context $context -Settings $settings -State $state
    $state = Get-F1State
    $contacts = @(Get-F1Contacts $context.ContactsFolder)
    $stats.contacts = $contacts.Count

    $eligible = @()
    foreach ($c in $contacts) {
        if (-not (Test-F1Email $c.Email)) {
            $stats.invalid++
            Write-F1Log -Action "ELIGIBILITY" -Email $c.Email -Result "INVALID_EMAIL"
            continue
        }
        if (Test-F1Category -Categories $c.Categories -Category $settings.blocked_category) {
            $stats.unsubscribed++
            Write-F1Log -Action "ELIGIBILITY" -Email $c.Email -Result "F1-DISCRITTO"
            continue
        }
        if (-not (Test-F1Category -Categories $c.Categories -Category $settings.required_category)) {
            Write-F1Log -Action "ELIGIBILITY" -Email $c.Email -Result "NO_F1_CONSENSO"
            continue
        }
        $stats.authorized++
        if (Test-F1DoNotResend -State $state -CampaignKey $campaign -Email $c.Email) {
            $stats.already++
            continue
        }
        $eligible += $c
    }

    $max = [Math]::Max(1,[int]$settings.max_emails_per_run)
    $batch = @($eligible | Select-Object -First $max)
    $stats.remaining = [Math]::Max(0,$eligible.Count-$batch.Count)
    $template = Get-Content (Join-Path $PSScriptRoot "templates\agent-pricing.html") -Raw -Encoding UTF8
    $subject = "Stai pensando di vendere casa? Scopri quanto può valere davvero"

    foreach ($c in $batch) {
        $state = Get-F1State
        if (Test-F1DoNotResend -State $state -CampaignKey $campaign -Email $c.Email) {
            $stats.already++
            continue
        }

        $currentCategories = [string]$c.ComObject.Categories
        if (-not (Test-F1Category $currentCategories $settings.required_category) -or
            (Test-F1Category $currentCategories $settings.blocked_category)) {
            Write-F1Log -Action "PRE_SEND_GATE" -Email $c.Email -Result "BLOCKED"
            continue
        }

        $html = Render-F1Template -Template $template -FirstName $c.FirstName -LastName $c.LastName -Email $c.Email -CampaignKey $campaign

        if ($DryRun) {
            Write-F1Log -Action "SEND" -Email $c.Email -Result "DRY_RUN"
            continue
        }

        $state = Set-F1Record -State $state -CampaignKey $campaign -Email $c.Email -Status "SENDING" -RetryCount 0
        $sent = $false
        for ($attempt=1; $attempt -le 3; $attempt++) {
            try {
                Send-F1OutlookMail -Context $context -To $c.Email -Subject $subject -Html $html
                $state = Get-F1State
                $state = Set-F1Record -State $state -CampaignKey $campaign -Email $c.Email -Status "SENT" -RetryCount ($attempt-1)
                Write-F1Log -Action "SEND" -Email $c.Email -Result "SENT"
                $stats.sent++
                $sent = $true
                break
            } catch {
                $msg = $_.Exception.Message
                Write-F1Log -Action "SEND" -Email $c.Email -Result "RETRY_$attempt" -Error $msg
                if ($attempt -lt 3) { Start-Sleep -Seconds ([Math]::Min(60,5*$attempt)) }
                else {
                    $state = Get-F1State
                    $state = Set-F1Record -State $state -CampaignKey $campaign -Email $c.Email -Status "FAILED" -Error $msg -RetryCount 3
                    $stats.failed++
                }
            }
        }

        if ($sent -and [int]$settings.delay_seconds -gt 0) {
            Start-Sleep -Seconds ([int]$settings.delay_seconds)
        }
    }

    $state = Get-F1State
    $state.last_run_at = (Get-Date).ToString("o")
    Save-F1State $state
    $report = Write-F1Report -Stats $stats -CampaignKey $campaign -Started $started -Finished (Get-Date)
    Write-F1Log -Action "RUN" -Result "COMPLETED" -Error $report
    Write-Host "F1 OUTLOOK MONTHLY MAILER COMPLETATO"
    Write-Host "Campagna: $campaign"
    Write-Host "Inviati: $($stats.sent)  Falliti: $($stats.failed)  Rimanenti: $($stats.remaining)"
    Write-Host "Report: $report"
}
catch {
    Write-F1Log -Action "RUN" -Result "ERROR" -Error $_.Exception.Message
    Write-Error $_.Exception.Message
    exit 1
}
finally {
    if ($mutex) {
        try { $mutex.ReleaseMutex() } catch {}
        $mutex.Dispose()
    }
}
