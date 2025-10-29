$exe = $args[0]
if ($exe.length -lt 1) {
    Write-Host "Specify exe path."
}
else {
    if (Test-Path -Path $exe) {

        $config = Get-Content -Path $($PSScriptRoot | Join-Path -ChildPath "config.json") | ConvertFrom-Json
        $taskPath = $config.taskPath
        if (-not $taskPath.StartsWith("\")) {
            $taskPath = "\" + $taskPath
        }
        if (-not $taskPath.EndsWith("\")) {
            $taskPath = $taskPath + "\"
        }

        $dest = $env:APPDATA | Join-Path -ChildPath $config.appDirName
        if (-not (Test-Path $dest -PathType Container)) {
            New-Item -Path $dest -ItemType Directory > $null
        }

        $src = $PSScriptRoot | Join-Path -ChildPath "run.ps1" | Copy-Item -Destination $dest -PassThru


        $action = New-ScheduledTaskAction -Execute powershell.exe -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$src`" `"$exe`""
        $trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
        $settings = New-ScheduledTaskSettingsSet -Hidden -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
        Register-ScheduledTask -TaskName "startup" `
            -TaskPath $taskPath `
            -Action $action `
            -Trigger $trigger `
            -Description "Run keyhac on startup." `
            -Settings $settings `
            -Force
    }
    else {
        Write-Host "Invalid path."
    }
}