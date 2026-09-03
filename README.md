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

Switching without a UAC prompt
-------------------------------

The UI never calls `netsh` itself and never requests elevation. Instead, clicking "Switch Interfaces"
sends a request over a local named pipe (`NetworkInterfaceSwitcherPipe`) to the installed service, which
runs as LocalSystem and performs the actual `netsh` call. Any locally logged-on user can send a request;
only the service can act on it, so no UAC prompt is ever shown to the user - but this also means
**switching only works while the service is installed and running**. If the service isn't running, the UI
shows an error explaining that instead of silently failing.

Configuration
-------------

The UI stores the last-selected dropdown pair in the current user registry key `HKCU:\SOFTWARE\NetworkInterfaceSwitcher`
purely for convenience across UI restarts.

The service is the source of truth for enforcement and keeps its own state in `HKLM:\SOFTWARE\NetworkInterfaceSwitcher`:
`Interface1`, `Interface2` (the configured pair) and `ActiveInterface` (which of the two should be enabled).
It writes these itself whenever it handles a pipe switch request. You can also set them manually, or via the
install script below, to seed the initial state before any switch has happened.

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
* Select Interface1 and Interface2 and use the UI to switch. The actual switch is performed by the service
  (see "Switching without a UAC prompt" above) — the service must be installed and running for this to work.
* Install and run service (requires admin/elevation)
* From elevated PowerShell:
* Install: .\scripts\install-service.ps1 -ExePath 'C:\full\path\NetworkInterfaceSwitcher.exe'
* Optional: add -CopyToProgramFiles to copy the exe to Program Files before installing.
* Optionally pass -Interface1 'NameA' -Interface2 'NameB' to write HKLM config during install.
* Uninstall: .\scripts\uninstall-service.ps1
* You can also set config manually or use .\scripts\set-config-hklm.ps1 -Interface1 'NameA' -Interface2 'NameB'.
* Service name: NetworkInterfaceSwitcherService. It runs as LocalSystem by default and enforces the configured pair periodically.
