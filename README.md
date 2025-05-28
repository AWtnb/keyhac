# README

[keyhac](https://github.com/crftwr/keyhac) customization.

Environment:

- [CorvusSKK](https://github.com/nathancorvussolis/corvusskk)
- JIS keyboard


## Install

```PowerShell
$d = "Keyhac"; New-Item -Path ($env:APPDATA | Join-Path -ChildPath $d) -Value ($pwd.Path | Join-Path -ChildPath $d) -ItemType Junction
```

Running [`set-startup.ps1`](set-startup.ps1) with `keyhac.exe` path makes shortcut (.lnk) to windows startup:

```
.\set-startup.ps1 "$env:USERPROFILE\Sync\portable_app\keyhac\keyhac.exe"
```

## Notes

### `keyhac.ini`

Set `visible = 0` on `[CONSOLE]` section of `keyhac.ini` to hide console window on startup.

```ini
[CONSOLE]
visible = 0
```


### `.stignore`

With [Syncthing](https://syncthing.net/), append below on `.stignore` to skip syncing local history.

```
(?d)keyhac.ini
```




