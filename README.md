# MailJump

MailJump is a simple helper application that provides a minimal MAPI interface for Microsoft Word or other applications that expect a MAPI email client. The tray application receives simplified MAPI requests over a named pipe and invokes **New Outlook** using the `mailto:` protocol to draft the message.

This repository contains two components:

* **MailJumpTray** – a Windows Forms tray application that waits for MAPI requests on a named pipe and launches New Outlook.
* **MAPIStub** – a minimal `MAPI32.dll` replacement exporting `MAPISendMail`. It serialises the request and forwards it to the tray application through the named pipe.

The stub only implements the features required for Word's "Send to Mail Recipient" command and is not a full MAPI implementation.

## Building

1. **MailJumpTray** requires the .NET 9 SDK or later. Build the project with `dotnet build` targeting `x64`.
2. **MAPIStub** can be built with any 64-bit Visual Studio toolchain using the provided `MAPIStub.def` to export `MAPISendMail`.

After building, place `MAPI32.dll` from the `MAPIStub` project somewhere in your `PATH` or alongside Word so that it is loaded when a MAPI call is made. Run `MailJumpTray.exe` to start the tray application before using Word's email features.

This project is a basic proof of concept and now supports multiple attachments, though it still only handles a single primary recipient.

## Installer

A PowerShell installer script is available in the `scripts` folder to simplify deployment.
Run the script after building both modules to copy them to `C:\Program Files\MailJump`,
set MailJump as the default MAPI client and register uninstall information.

```powershell
cd scripts
./install.ps1
```

The installer also places `uninstall.ps1` in the install directory so you can remove
MailJump by running that script later.

For distributing MailJump as a single executable installer you can build the
NSIS package located in the `installer` folder. Install the NSIS toolset and
run `makensis` to generate `MailJumpInstaller.exe`:

```bash
cd installer
makensis MailJumpInstaller.nsi
```

The resulting `MailJumpInstaller.exe` installs MailJump to
`C:\Program Files\MailJump` and registers it as the default MAPI client.
