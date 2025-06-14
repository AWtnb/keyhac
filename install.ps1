$d = "Keyhac"
$dataPath = $env:APPDATA | Join-Path -ChildPath $d
$srcPath = $PSScriptRoot | Join-Path -ChildPath $d
New-Item -Path $dataPath -Value $srcPath -ItemType Junction -Confirm