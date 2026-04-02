# NetworkInterfaceSwitcher
Quickly switch between two network interfaces

Running as a Windows Service
----------------------------

This application can run either as a desktop UI or as a Windows Service. The executable chooses mode at startup:

- If started from an interactive desktop session it runs the WinForms UI.
- If started by the Service Control Manager it runs as a service and enforces the configured interface pair periodically.

To install the service (run from an elevated PowerShell):

```powershell
./scripts/install-service.ps1 -ExePath 'C:\full\path\to\NetworkInterfaceSwitcher.exe'
```

To uninstall the service:

```powershell
./scripts/uninstall-service.ps1
```

Configuration
-------------

The UI stores selected interfaces in the current user registry key `HKCU:\SOFTWARE\NetworkInterfaceSwitcher`.
The UI will attempt to also write to `HKLM:\SOFTWARE\NetworkInterfaceSwitcher` so a service running as LocalSystem can read the configured values.
You can manually set the two string values `Interface1` and `Interface2` under that key if needed.

Helper scripts
--------------

- `scripts/install-service.ps1` - Installs the service. Options:
  - `-CopyToProgramFiles` copy the executable into Program Files before installing
  - `-Interface1` and `-Interface2` optionally set initial HKLM configuration
- `scripts/uninstall-service.ps1` - Stops and deletes the service
- `scripts/set-config-hklm.ps1` - Writes `Interface1`/`Interface2` to HKLM and restarts the service (requires elevation)

Service logging
---------------

The service writes basic logs to `%ProgramData%\NetworkInterfaceSwitcher\service.log` which can help debugging when Event Log entries are not present.

How to use
----------

* Run UI (no UAC)
* Double-click the executable or start from an interactive session — the app starts the WinForms UI as before.
* Select Interface1 and Interface2 and use the UI to switch. The UI saves to HKCU and attempts to save to HKLM (ignored if not elevated).
* Install and run service (requires admin/elevation)
* From elevated PowerShell:
* Install: .\scripts\install-service.ps1 -ExePath 'C:\full\path\NetworkInterfaceSwitcher.exe'
* Optional: add -CopyToProgramFiles to copy the exe to Program Files before installing.
* Optionally pass -Interface1 'NameA' -Interface2 'NameB' to write HKLM config during install.
* Uninstall: .\scripts\uninstall-service.ps1
* You can also set config manually or use .\scripts\set-config-hklm.ps1 -Interface1 'NameA' -Interface2 'NameB'.
* Service name: NetworkInterfaceSwitcherService. It runs as LocalSystem by default and enforces the configured pair periodically.
