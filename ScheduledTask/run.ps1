param([parameter(Mandatory)]$src)

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

if (-not $src) {
    "Keyhac path not specified." | logWrite
    [System.Environment]::exit(1)
}

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
    "Keyhac path not found." | logWrite
    [System.Environment]::exit(1)
}