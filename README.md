# README

[keyhac](https://github.com/crftwr/keyhac) customization.

Environment:

- [CorvusSKK](https://github.com/nathancorvussolis/corvusskk)
- JIS keyboard


## Install

Run [`install.ps1`](./install.ps1) to create junction of `Keyhac` to AppData.

```
powershell .\install.ps1
```

## Development Setup

This project uses [uv](https://docs.astral.sh/uv/) to match the Python version bundled with keyhac.

To set up the development environment, run:

```
uv sync
```


