param(
    [string]$Interface1,
    [string]$Interface2,
    [string]$ActiveInterface,
    [string]$ServiceName = 'NetworkInterfaceSwitcherService'
)

if (-not $Interface1 -and -not $Interface2 -and -not $ActiveInterface) {
    Write-Error "Provide at least one of -Interface1, -Interface2, or -ActiveInterface"
    exit 1
}

New-Item -Path 'HKLM:\SOFTWARE\NetworkInterfaceSwitcher' -Force | Out-Null
if ($Interface1) { Set-ItemProperty -Path 'HKLM:\SOFTWARE\NetworkInterfaceSwitcher' -Name 'Interface1' -Value $Interface1 }
if ($Interface2) { Set-ItemProperty -Path 'HKLM:\SOFTWARE\NetworkInterfaceSwitcher' -Name 'Interface2' -Value $Interface2 }
if ($ActiveInterface) { Set-ItemProperty -Path 'HKLM:\SOFTWARE\NetworkInterfaceSwitcher' -Name 'ActiveInterface' -Value $ActiveInterface }

Write-Host "Configuration written to HKLM. Restarting service $ServiceName"
& sc.exe stop $ServiceName | Out-Null
Start-Sleep -Seconds 1
& sc.exe start $ServiceName | Out-Null
Write-Host "Done."
