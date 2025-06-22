$src = $args[0]
if ($src.length -lt 1) {
    Write-Host "Specify exe path."
    exit 1
}

$action = New-ScheduledTaskAction -Execute $src
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -Priority 5

$taskName = "Keyhac-on-startup"

if ($null -ne (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue)) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
}

Register-ScheduledTask -TaskName $taskName `
    -Action $action `
    -Trigger $trigger `
    -Description "Run keyhac on startup." `
    -Settings $settings
