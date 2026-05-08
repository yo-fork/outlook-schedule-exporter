param(
    [string]$ConfigPath = (Join-Path (Split-Path -Parent $PSScriptRoot) "config.json"),
    [int]$IntervalMinutes = 30,
    [string]$TaskName = "Export Outlook Schedule"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ($IntervalMinutes -lt 5) {
    throw "IntervalMinutes should be 5 or greater."
}

$scriptPath = Join-Path $PSScriptRoot "Export-OutlookSchedule.ps1"

if (-not (Test-Path -LiteralPath $scriptPath)) {
    throw "Script not found: $scriptPath"
}

if (-not (Test-Path -LiteralPath $ConfigPath)) {
    throw "Config file not found: $ConfigPath"
}

$powerShellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"

# Do not force ExecutionPolicy Bypass here. Use the organization's allowed policy.
$arguments = "-NoProfile -File `"$scriptPath`" -ConfigPath `"$ConfigPath`""

$action = New-ScheduledTaskAction -Execute $powerShellExe -Argument $arguments

$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1)
$trigger.Repetition = New-ScheduledTaskRepetitionSettings `
    -Interval (New-TimeSpan -Minutes $IntervalMinutes) `
    -Duration ([TimeSpan]::MaxValue)

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Description "Export Outlook calendar to local or shared HTML/CSV files." `
    -Force | Out-Null

Write-Host "Registered scheduled task: $TaskName"
Write-Host "Interval: every $IntervalMinutes minutes"
Write-Host "Script: $scriptPath"
Write-Host "Config: $ConfigPath"
