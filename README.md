# keyhac customization

With [Syncthing](https://syncthing.net/), append below on `.stignore` to skip syncing local history.

```
keyhac.ini
```


## Install

```PowerShell
$d = "Keyhac"; New-Item -Path ($env:APPDATA | Join-Path -ChildPath $d) -Value ($pwd.Path | Join-Path -ChildPath $d) -ItemType Junction
```

Running [`set-startup.ps1`](set-startup.ps1) with `keyhac.exe` path makes shortcut (.lnk) to windows startup:

```
.\set-startup.ps1 "$env:USERPROFILE\Sync\portable_app\keyhac\keyhac.exe"
```

Environment:

- [CorvusSKK](https://github.com/nathancorvussolis/corvusskk)
- JIS keyboard

---

https://sites.google.com/site/craftware/keyhac-en

https://github.com/crftwr/keyhac
