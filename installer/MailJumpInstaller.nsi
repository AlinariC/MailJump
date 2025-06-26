!include "MUI2.nsh"

Name "MailJump"
OutFile "MailJumpInstaller.exe"
InstallDir "$PROGRAMFILES\MailJump"
RequestExecutionLevel admin

Section "Install"
    SetOutPath "$INSTDIR"
    File "../src/MailJumpTray/bin/Release/net9.0-windows/win-x64/publish/MailJumpTray.exe"
    File "../src/MAPIStub/MAPI32.dll"

    WriteRegStr HKCU "Software\Clients\Mail" "" "MailJump"
    WriteRegStr HKCU "Software\Clients\Mail\MailJump" "DLLPath" "$INSTDIR\MAPI32.dll"
    WriteRegStr HKCU "Software\Clients\Mail\MailJump" "ApplicationName" "MailJump"

    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\MailJump" "DisplayName" "MailJump"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\MailJump" "DisplayIcon" "$INSTDIR\MailJumpTray.exe"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\MailJump" "UninstallString" "$INSTDIR\Uninstall.exe"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\MailJump" "InstallLocation" "$INSTDIR"
    WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\MailJump" "Publisher" "MailJump Project"

    WriteUninstaller "$INSTDIR\Uninstall.exe"
SectionEnd

Section "Uninstall"
    Delete "$INSTDIR\MailJumpTray.exe"
    Delete "$INSTDIR\MAPI32.dll"
    Delete "$INSTDIR\Uninstall.exe"
    RMDir "$INSTDIR"

    ReadRegStr $0 HKCU "Software\Clients\Mail" ""
    StrCmp $0 "MailJump" 0 +2
        DeleteRegValue HKCU "Software\Clients\Mail" ""
    DeleteRegKey HKCU "Software\Clients\Mail\MailJump"

    DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\MailJump"
SectionEnd
