# SAM Offer Generator update system

## Modules

The update system is based on the existing distribution modules:

```text
SAM-Offer-Generator-App-No-Runtime.zip
SAM-Offer-Generator-Runtime-Win64.zip
SAM-Offer-Generator-windows-portable.zip
```

Normal app updates download only:

```text
SAM-Offer-Generator-App-No-Runtime.zip
```

The heavy runtime module is not downloaded unless Python, PyInstaller, PySide6, Qt, or requirements change.

## Files

```text
core/update_client.py     Checks GitHub Release, downloads update ZIP, starts updater
updater.py                External updater process
config/update.json        Repository and asset names
```

## User flow

```text
GUI -> Обновления -> check latest release -> download app module -> start updater -> quit GUI
updater.exe -> wait for GUI process -> merge ZIP into app folder -> restart app
```

## Important limitations

`updater.exe` skips replacing itself while it is running. If the updater itself needs to be changed, use a full portable reinstall or update it manually.

The updater does not elevate privileges. It works only when the application folder is writable by the current Windows user.
