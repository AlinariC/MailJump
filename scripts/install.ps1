param(
    [string]$TrayPath = "..\src\MailJumpTray\bin\Release\net9.0-windows\MailJumpTray.exe",
    [string]$StubPath = "..\src\MAPIStub\x64\Release\MAPI32.dll",
    [string]$InstallPath = "$env:ProgramFiles\MailJump"
)

if (!(Test-Path $TrayPath)) { throw "MailJumpTray.exe not found at $TrayPath" }
if (!(Test-Path $StubPath)) { throw "MAPI32.dll not found at $StubPath" }

if (!(Test-Path $InstallPath)) { New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null }

Copy-Item $TrayPath (Join-Path $InstallPath 'MailJumpTray.exe') -Force
Copy-Item $StubPath (Join-Path $InstallPath 'MAPI32.dll') -Force

# copy uninstall script to install location
Copy-Item (Join-Path $PSScriptRoot 'uninstall.ps1') (Join-Path $InstallPath 'uninstall.ps1') -Force

$regBase = 'HKCU:\Software\Clients\Mail'
New-Item -Path $regBase -Force | Out-Null
Set-ItemProperty -Path $regBase -Name '(default)' -Value 'MailJump'

$clientKey = Join-Path $regBase 'MailJump'
New-Item -Path $clientKey -Force | Out-Null
Set-ItemProperty -Path $clientKey -Name 'DLLPath' -Value (Join-Path $InstallPath 'MAPI32.dll')
Set-ItemProperty -Path $clientKey -Name 'ApplicationName' -Value 'MailJump'

$uninstallKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\MailJump'
New-Item -Path $uninstallKey -Force | Out-Null
Set-ItemProperty -Path $uninstallKey -Name 'DisplayName' -Value 'MailJump'
Set-ItemProperty -Path $uninstallKey -Name 'DisplayIcon' -Value (Join-Path $InstallPath 'MailJumpTray.exe')
Set-ItemProperty -Path $uninstallKey -Name 'UninstallString' -Value "powershell.exe -ExecutionPolicy Bypass -File `"$InstallPath\uninstall.ps1`""
Set-ItemProperty -Path $uninstallKey -Name 'InstallLocation' -Value $InstallPath
Set-ItemProperty -Path $uninstallKey -Name 'Publisher' -Value 'MailJump Project'

Write-Host "MailJump installed to $InstallPath"
