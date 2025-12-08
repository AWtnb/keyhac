# README

[keyhac](https://github.com/crftwr/keyhac-win) customization.

Environment:

- [CorvusSKK](https://github.com/nathancorvussolis/corvusskk)
- JIS keyboard


## Install

Run [`install.ps1`](./install.ps1) to create junction of `Keyhac` to AppData.

### Optional

Running [`ScheduledTask/install.ps1`](./ScheduledTask/install.ps1) with `keyhac.exe` path copies `ScheduledTask/run.ps1` to `$env:AppData\KeyhacStarter` and registers scheduled task to run it at logon:

```PowerShell
# EXAMPLE
.\ScheduledTask\install.ps1 "$env:USERPROFILE\Personal\tools\portable_apps\keyhac\keyhac.exe"
```

Running [`set-startmenu.ps1`](./set-startmenu.ps1) makes start menu to edit this repository on VSCode.

With [Syncthing](https://syncthing.net/), append below on `.stignore` to skip syncing local history.

```
(?d)keyhac.ini
```




