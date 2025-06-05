# https://github.com/smzht/fakeymacs/blob/master/keyhac.bat
$bat = @'
@echo off

start "" /high "_"
'@

$d = $args[0]
if ($d.length -lt 1) {
    Write-Host "Specify exe path."
}
else {
    $bat = $bat -replace "_", $d
    $bat | Out-File -FilePath ($env:USERPROFILE | Join-Path -ChildPath "AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\keyhac.bat") -Encoding utf8
}