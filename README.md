# README

[keyhac](https://github.com/crftwr/keyhac) customization.

Environment:

- [CorvusSKK](https://github.com/nathancorvussolis/corvusskk)
- JIS keyboard


## Install

Run [`install.ps1`](./install.ps1) to create junction of `Keyhac` to AppData.

### Optional

Running [`set-startup.ps1`](./set-startup.ps1) with `keyhac.exe` path makes `keyhac.bat` on startup, which runs keyhac with high priority:

```PowerShell
# EXAMPLE
.\set-startup.ps1 "$env:USERPROFILE\Sync\portable_app\keyhac\keyhac.exe"
```

Running [`set-startmenu.ps1`](./set-startmenu.ps1) makes start menu to edit this repository on VSCode.

Set `visible = 0` on `[CONSOLE]` section of `keyhac.ini` to hide console window on startup.

```ini
[CONSOLE]
visible = 0
```

With [Syncthing](https://syncthing.net/), append below on `.stignore` to skip syncing local history.

```
(?d)keyhac.ini
```




