# MailJump

MailJump is a simple helper application that provides a minimal MAPI interface for Microsoft Word or other applications that expect a MAPI email client. The tray application receives simplified MAPI requests over a named pipe and invokes **New Outlook** using the `mailto:` protocol to draft the message.

This repository contains two components:

* **MailJumpTray** – a Windows Forms tray application that waits for MAPI requests on a named pipe and launches New Outlook.
* **MAPIStub** – a minimal `MAPI32.dll` replacement exporting `MAPISendMail`. It serialises the request and forwards it to the tray application through the named pipe.

The stub only implements the features required for Word's "Send to Mail Recipient" command and is not a full MAPI implementation.

## Building

1. **MailJumpTray** requires the .NET 6 SDK or later. Build the project with `dotnet build` targeting `x64`.
2. **MAPIStub** can be built with any 64-bit Visual Studio toolchain using the provided `MAPIStub.def` to export `MAPISendMail`.

After building, place `MAPI32.dll` from the `MAPIStub` project somewhere in your `PATH` or alongside Word so that it is loaded when a MAPI call is made. Run `MailJumpTray.exe` to start the tray application before using Word's email features.

This project is a basic proof of concept and now supports multiple attachments, though it still only handles a single primary recipient.
