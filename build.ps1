# Build script for MailJump
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Ensure-Admin {
    $current = [Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
    if (-not $current.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Error "Run this script from an elevated PowerShell prompt."
        exit 1
    }
}

function Install-UsingWinget($id, $override = $null) {
    if (-not (winget list --exact --id $id | Select-String $id)) {
        Write-Host "Installing $id"
        $args = @('install','--id',$id,'--accept-package-agreements','--accept-source-agreements','--silent')
        if ($override) { $args += '--override'; $args += $override }
        winget @args
    } else {
        Write-Host "$id already installed"
    }
}

Ensure-Admin

# Install dependencies
Install-UsingWinget 'Microsoft.VisualStudio.2022.BuildTools' '--add Microsoft.VisualStudio.Workload.VCTools --quiet --norestart'
Install-UsingWinget 'Microsoft.DotNet.SDK.9'
Install-UsingWinget 'NSIS.NSIS'

# Refresh PATH after potential installs
$env:Path = [System.Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [System.Environment]::GetEnvironmentVariable('Path','User')

# Build MailJumpTray
Write-Host 'Building MailJumpTray...'
dotnet publish './src/MailJumpTray/MailJumpTray.csproj' -c Release -r win-x64 --self-contained -p:PublishSingleFile=true

# Build MAPIStub
Write-Host 'Building MAPIStub...'
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
$installPath = & $vswhere -latest -products '*' -requires Microsoft.Component.MSBuild -property installationPath
$vcvars = Join-Path $installPath 'VC/Auxiliary/Build/vcvars64.bat'
$buildCmd = 'cl.exe /nologo /LD MAPIStub.cpp /FeMAPI32.dll /link /DEF:MAPIStub.def'
Push-Location 'src/MAPIStub'
cmd /c "`"$vcvars`" && $buildCmd"
Pop-Location

# Build NSIS installer
Write-Host 'Building installer...'
$makensis = Get-Command makensis.exe -ErrorAction SilentlyContinue
if ($makensis) {
    Push-Location 'installer'
    & $makensis 'MailJumpInstaller.nsi'
    Pop-Location
} else {
    Write-Warning 'NSIS not found; installer not built.'
}

Write-Host 'Build complete.'
