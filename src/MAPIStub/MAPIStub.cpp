#include <windows.h>
#include <mapi.h>
#include <string>
#include <fstream>
#include <sstream>

extern "C" ULONG FAR PASCAL MAPISendMail(LHANDLE lhSession, ULONG ulUIParam,
    lpMapiMessage lpMessage, FLAGS flFlags, ULONG ulReserved)
{
    std::ostringstream oss;
    if (lpMessage && lpMessage->lpszSubject)
        oss << "{\"subject\":\"" << lpMessage->lpszSubject << "\",";
    if (lpMessage && lpMessage->lpszNoteText)
        oss << "\"body\":\"" << lpMessage->lpszNoteText << "\",";
    if (lpMessage && lpMessage->lpRecips && lpMessage->nRecipCount > 0)
        oss << "\"to\":\"" << lpMessage->lpRecips[0].lpszAddress << "\",";
    oss << "\"attachment\":null}";
    std::string json = oss.str();
    HANDLE hPipe = CreateFileA(R"\\.\pipe\MailJumpPipe", GENERIC_WRITE, 0, NULL, OPEN_EXISTING, 0, NULL);
    if (hPipe != INVALID_HANDLE_VALUE)
    {
        DWORD written;
        WriteFile(hPipe, json.c_str(), (DWORD)json.size(), &written, NULL);
        CloseHandle(hPipe);
    }
    return SUCCESS_SUCCESS;
}
