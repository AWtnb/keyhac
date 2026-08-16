$d = "Keyhac"
$dataPath = $env:APPDATA | Join-Path -ChildPath $d
Remove-Item -Path $dataPath -Force -Recurse -ErrorAction SilentlyContinue > $null
$srcPath = $PSScriptRoot | Join-Path -ChildPath $d
New-Item -Path $dataPath -Value $srcPath -ItemType Junction -Confirm -Force