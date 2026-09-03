<#
Installs the NetworkInterfaceSwitcher executable as a Windows Service.
Run in an elevated PowerShell session.
Usage:
  .\install-service.ps1 -ExePath 'C:\path\to\NetworkInterfaceSwitcher.exe' -ServiceName 'NetworkInterfaceSwitcherService'
If ExePath is omitted the script will look for the release build under ../bin/Release/net8.0-windows/ relative to this script.
Pass -Interface1/-Interface2 to seed the configured pair, and optionally -ActiveInterface to state
which one should be enforced as enabled (defaults to Interface1 if omitted).
#>
param(
    [string]$ExePath = (Join-Path -Path (Resolve-Path (Join-Path $PSScriptRoot '..\bin\Release\net8.0-windows')) -ChildPath 'NetworkInterfaceSwitcher.exe'),
    [string]$ServiceName = 'NetworkInterfaceSwitcherService',
    [string]$InstallFolder = "$env:ProgramFiles\\NetworkInterfaceSwitcher",
    [switch]$CopyToProgramFiles,
    [string]$Interface1,
    [string]$Interface2,
    [string]$ActiveInterface
)

if (-not (Test-Path $ExePath)) {
    Write-Error "Executable not found: $ExePath"
    exit 1
}

if ($CopyToProgramFiles) {
    if (-not (Test-Path $InstallFolder)) { New-Item -ItemType Directory -Path $InstallFolder -Force | Out-Null }
    $dest = Join-Path $InstallFolder (Split-Path $ExePath -Leaf)
    Copy-Item -Path $ExePath -Destination $dest -Force
    $ExePath = $dest
}

Write-Host "Installing service '$ServiceName' for executable: $ExePath"

# Create the service (binPath requires the exe quoted)
$binPath = "`"$ExePath`""
& sc.exe create $ServiceName binPath= $binPath start= auto DisplayName= "Network Interface Switcher Service" obj= LocalSystem | Out-Null

# Set a human-readable description
& sc.exe description $ServiceName "Service that enforces the configured network interface pair." | Out-Null

if ($Interface1 -or $Interface2) {
    try {
        Write-Host "Writing configuration to HKLM..."
        New-Item -Path 'HKLM:\SOFTWARE\NetworkInterfaceSwitcher' -Force | Out-Null
        if ($Interface1) { Set-ItemProperty -Path 'HKLM:\SOFTWARE\NetworkInterfaceSwitcher' -Name 'Interface1' -Value $Interface1 }
        if ($Interface2) { Set-ItemProperty -Path 'HKLM:\SOFTWARE\NetworkInterfaceSwitcher' -Name 'Interface2' -Value $Interface2 }

        # Explicitly seed which of the two should be enforced as active. Defaults to Interface1 to
        # match the service's own fallback when ActiveInterface is unset, but stating it here avoids
        # relying on that undocumented default.
        $effectiveActive = $ActiveInterface
        if (-not $effectiveActive) { $effectiveActive = $Interface1 }
        if ($effectiveActive) { Set-ItemProperty -Path 'HKLM:\SOFTWARE\NetworkInterfaceSwitcher' -Name 'ActiveInterface' -Value $effectiveActive }
    } catch {
        Write-Warning "Failed to write HKLM configuration: $_"
    }
}

Write-Host "Service created. Starting service..."
& sc.exe start $ServiceName

Write-Host "Service install complete. Use 'sc stop $ServiceName' and 'sc delete $ServiceName' to remove."
