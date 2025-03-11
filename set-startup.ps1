$d = $args[0]
if ($d.length -lt 1) {
    Write-Host "Specify exe path."
} else {
    $wsShell = New-Object -ComObject WScript.Shell
    $startup = $env:USERPROFILE | Join-Path -ChildPath "AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
    $shortcutPath = $startup | Join-Path -ChildPath ((Get-Item $d).BaseName + ".lnk")
    $shortcut = $WsShell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $d
    $shortcut.Save()
    "Created shortcut on startup: {0}" -f $shortcutPath | Write-Host
}