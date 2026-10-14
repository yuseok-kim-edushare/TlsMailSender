#pragma once

#include <windows.h>

// TlsMailSender Native AOT 공유 라이브러리의 C ABI.
// net10.0-windows, NativeLib=Shared 로 게시한 TlsMailSender.dll 을
// LoadLibraryW / GetProcAddress 로 호출합니다.
//
// 문자열은 UTF-16(널 종료)이며 Windows wchar_t* 와 같습니다.
// 호출 규약은 __stdcall 입니다. EntryPoint 이름은 장식 없는 그대로입니다.
// 반환값: 1 = 성공, 0 = 실패. 실패 상세는 C:\temp\TlsMailSender.log 에 남습니다.
//
// attachments 는 '|' 로 구분한 경로 목록입니다. null 이면 첨부가 없습니다.
// fromDisplayName, toDisplayName, replyTo 는 null 을 허용합니다.

typedef int(__stdcall* TlsMailSender_SendMail)(
    const wchar_t* smtpServer,
    int smtpPort,
    const wchar_t* smtpUser,
    const wchar_t* smtpPass,
    int useTls,
    const wchar_t* from,
    const wchar_t* to,
    const wchar_t* subject,
    const wchar_t* body,
    const wchar_t* attachments);

typedef int(__stdcall* TlsMailSender_SendMailWithAliasReplyTo)(
    const wchar_t* smtpServer,
    int smtpPort,
    const wchar_t* smtpUser,
    const wchar_t* smtpPass,
    int useTls,
    const wchar_t* from,
    const wchar_t* to,
    const wchar_t* subject,
    const wchar_t* body,
    const wchar_t* attachments,
    const wchar_t* fromDisplayName,
    const wchar_t* toDisplayName,
    const wchar_t* replyTo);

typedef void(__stdcall* TlsMailSender_ReloadAllowedCerts)();
