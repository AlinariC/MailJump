param(
    [string]$InstallPath = "$env:ProgramFiles\MailJump"
)

Remove-Item -Path $InstallPath -Recurse -Force -ErrorAction SilentlyContinue

$regBase = 'HKCU:\Software\Clients\Mail'
$clientKey = Join-Path $regBase 'MailJump'
if (Test-Path $clientKey) { Remove-Item $clientKey -Recurse -Force }
if ((Get-ItemProperty $regBase -ErrorAction SilentlyContinue).'('(default)')' -eq 'MailJump') {
    Remove-ItemProperty -Path $regBase -Name '(default)' -ErrorAction SilentlyContinue
}

$uninstallKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\MailJump'
if (Test-Path $uninstallKey) { Remove-Item $uninstallKey -Recurse -Force }

Write-Host 'MailJump uninstalled'
