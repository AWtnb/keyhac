# keyhac customization

With [Syncthing](https://syncthing.net/), append below on `.stignore` to skip syncing local history.

```
keyhac.ini
```


## Install

```PowerShell
$d = "Keyhac"; New-Item -Path ($env:APPDATA | Join-Path -ChildPath $d) -Value ($pwd.Path | Join-Path -ChildPath $d) -ItemType Junction
```

Environment:

- [CorvusSKK](https://github.com/nathancorvussolis/corvusskk)
- JIS keyboard

---

https://sites.google.com/site/craftware/keyhac-en

https://github.com/crftwr/keyhac
