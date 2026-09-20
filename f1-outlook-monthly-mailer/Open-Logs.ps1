$base=Join-Path $env:LOCALAPPDATA "F1OutlookMonthlyMailer"
$logs=Join-Path $base "logs"
if (-not (Test-Path $logs)) { New-Item -ItemType Directory -Path $logs -Force | Out-Null }
Start-Process explorer.exe -ArgumentList $logs
