param(
    [string]$TaskName = "HR Job Scraper v4 Daily Email",
    [string]$RunTime = "11:00"
)

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$BatchPath = Join-Path $ProjectDir "run_daily.bat"

if (-not (Test-Path $BatchPath)) {
    throw "Missing runner: $BatchPath"
}

$Action = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$BatchPath`"" -WorkingDirectory $ProjectDir
$Trigger = New-ScheduledTaskTrigger -Daily -At $RunTime
$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Settings $Settings `
    -Description "Runs HR job scraper and emails the Excel report through Brevo." `
    -Force

Write-Host "Scheduled task created/updated: $TaskName"
Write-Host "Run time: $RunTime daily"
Write-Host "Runner: $BatchPath"
