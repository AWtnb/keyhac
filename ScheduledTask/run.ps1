function logWrite {
    $log = $input -join ""
    $log = (Get-Date -Format "yyyyMMdd-HH:mm:ss ") + $log
    $log | Out-File -FilePath ($env:USERPROFILE | Join-Path -ChildPath "Desktop\Keyhac-startup-error.log") -Append
}

$proc = Get-Process -ProcessName "Keyhac" -ErrorAction SilentlyContinue
if ($proc) {
    Write-Host "Keyhac already running."
    [System.Environment]::exit(0)
}

if (($args.Count -lt 1) -or ($args[0].Trim().Length -lt 1)) {
    "PowerShell script to start Keyhac is not specified." | logWrite
    [System.Environment]::exit(1)
}

$src = $args[0].trim()
if (Test-Path $src) {
    try {
        Start-Process -FilePath $src -ErrorAction Stop
        Write-Host "Starting Keyhac.exe..."
        [System.Environment]::exit(0)
    }
    catch {
        $_ | logWrite
        [System.Environment]::exit(1)
    }
}
else {
    "PowerShell script to start Keyhac not found." | logWrite
    [System.Environment]::exit(1)
}