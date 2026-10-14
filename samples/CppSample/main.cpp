// TlsMailSender C++ 샘플
// Native AOT로 게시된 TlsMailSender.dll(net10.0-windows, NativeLib=Shared)을
// LoadLibraryW/GetProcAddress로 동적 로드하여 호출하는 예제.
//
// 빌드 전 준비:
//   dotnet publish ..\..\TlsMailSender\TlsMailSender.csproj -f net10.0-windows -r win-x64 -c Release
//   위 게시 결과의 TlsMailSender.dll 을 이 실행 파일과 같은 폴더에 복사한다.
//
// 사용법:
//   TlsMailSenderCppSample.exe <smtpServer> <port> <user> <pass> <useTls 0|1> <from> <to> <subject> <body> [attachments] [fromDisplayName] [toDisplayName] [replyTo]
//   attachments 는 '|' 로 구분합니다. 표시 이름 또는 Reply-To 를 주면 SendMailWithAliasReplyTo 를 호출합니다.

#include "TlsMailSender.h"

#include <cstdio>
#include <cstdlib>
#include <cwchar>

static void PrintUsage()
{
    wprintf(L"TlsMailSender C++ Sample\n\n");
    wprintf(L"사용법:\n");
    wprintf(L"  TlsMailSenderCppSample.exe <smtpServer> <port> <user> <pass> <useTls 0|1> <from> <to> <subject> <body> [attachments] [fromDisplayName] [toDisplayName] [replyTo]\n\n");
    wprintf(L"attachments 는 '|' 로 구분한 경로입니다. 생략하면 첨부가 없습니다.\n");
    wprintf(L"fromDisplayName / toDisplayName / replyTo 중 하나라도 주면 SendMailWithAliasReplyTo 를 호출합니다.\n");
}

int wmain(int argc, wchar_t* argv[])
{
    if (argc < 10)
    {
        PrintUsage();
        return 1;
    }

    const wchar_t* smtpServer = argv[1];
    int smtpPort = _wtoi(argv[2]);
    const wchar_t* smtpUser = argv[3];
    const wchar_t* smtpPass = argv[4];
    int useTls = _wtoi(argv[5]);
    const wchar_t* from = argv[6];
    const wchar_t* to = argv[7];
    const wchar_t* subject = argv[8];
    const wchar_t* body = argv[9];
    const wchar_t* attachments = (argc > 10 && argv[10][0] != L'\0') ? argv[10] : nullptr;

    HMODULE hModule = LoadLibraryW(L"TlsMailSender.dll");
    if (hModule == nullptr)
    {
        wprintf(L"TlsMailSender.dll 로드 실패 (오류 코드: %lu)\n", GetLastError());
        wprintf(L"게시된 Native AOT 결과물을 이 실행 파일과 같은 폴더에 복사했는지 확인하세요.\n");
        return 1;
    }

    auto reload = reinterpret_cast<TlsMailSender_ReloadAllowedCerts>(GetProcAddress(hModule, "ReloadAllowedCerts"));
    if (reload != nullptr)
        reload();

    int result = 0;
    if (argc > 11)
    {
        auto send = reinterpret_cast<TlsMailSender_SendMailWithAliasReplyTo>(GetProcAddress(hModule, "SendMailWithAliasReplyTo"));
        if (send == nullptr)
        {
            wprintf(L"SendMailWithAliasReplyTo 심볼을 찾을 수 없습니다.\n");
            FreeLibrary(hModule);
            return 1;
        }

        const wchar_t* fromDisplayName = (argc > 11 && argv[11][0] != L'\0') ? argv[11] : nullptr;
        const wchar_t* toDisplayName = (argc > 12 && argv[12][0] != L'\0') ? argv[12] : nullptr;
        const wchar_t* replyTo = (argc > 13 && argv[13][0] != L'\0') ? argv[13] : nullptr;
        result = send(
            smtpServer, smtpPort, smtpUser, smtpPass, useTls,
            from, to, subject, body, attachments,
            fromDisplayName, toDisplayName, replyTo);
    }
    else
    {
        auto send = reinterpret_cast<TlsMailSender_SendMail>(GetProcAddress(hModule, "SendMail"));
        if (send == nullptr)
        {
            wprintf(L"SendMail 심볼을 찾을 수 없습니다.\n");
            FreeLibrary(hModule);
            return 1;
        }

        result = send(
            smtpServer, smtpPort, smtpUser, smtpPass, useTls,
            from, to, subject, body, attachments);
    }

    if (result == 1)
        wprintf(L"성공\n");
    else
        wprintf(L"실패: 반환값 %d. C:\\temp\\TlsMailSender.log 를 확인하세요.\n", result);

    fflush(stdout);
    FreeLibrary(hModule);
    return result == 1 ? 0 : 1;
}
